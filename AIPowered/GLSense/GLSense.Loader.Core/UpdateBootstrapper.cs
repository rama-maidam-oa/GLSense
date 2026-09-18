// UpdateBootstrapper.cs in GLSense.Loader.Core
//
// Pre-AppDomain-load bootstrap: decides which release of GLSense.Addin.Core to load,
// gets it onto disk if it isn't there yet, and maintains ReleaseHistory.json - the
// permanent catalog of every release ever adopted on this machine. See
// docs/superpowers/specs/2026-08-30-hotreload-release-history-design.md.
//
// Decision tree:
//   1. ReleaseHistory.json does not exist -> first-ever run on this machine. If
//      Manifest\ has both manifest.json and a zip (the MSI's bundled seed), wipe any
//      stray Versions\ content, extract+catalog it (source "Install"), delete the
//      Manifest folder, done. If no seed is present, fall through to step 2/3 with an
//      empty catalog.
//   2. ReleaseHistory.json exists and Manifest\ has both manifest.json and a zip:
//      compare the manifest's version+releaseDate against every existing catalog
//      entry. An exact match means this is a reinstall of an already-known release -
//      reconcile the catalog (drop entries whose folder no longer exists) before
//      extracting. Either way, extract+catalog normally (the caller's `source`
//      parameter is used as-is - see ResolveVersionToLoad's own doc comment).
//   3. No zip in Manifest\, but Versions\{FolderName}\ for the currently active
//      release already has DLLs -> reuse it, nothing to do.
//   4. Nothing usable anywhere -> return null so the caller can skip loading the
//      AppDomain instead of crashing Excel.
using GLSense.Contracts;
using GLSense.Shared;
using System;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Security.Cryptography;

namespace GLSense.Loader.Core
{
    public class UpdateBootstrapper
    {
        /// <summary>
        /// Resolves which release to load, extracting/cataloguing a new one if
        /// Manifest\ has a zip waiting. `source` records WHY this call is happening -
        /// "Install" (default) for the automatic Excel-startup path and for an ordinary
        /// local dev-loop rebuild picked up there, or "Online"/"Offline" when this is
        /// called after RibReload's picker window staged a validated release into
        /// Manifest\. Never pass "Online"/"Offline" from the startup path.
        /// </summary>
        public ResolvedRelease ResolveVersionToLoad(IGLSenseContext context, string source = "Install")
        {
            var logger = context.Logger;
            var paths = context.Paths;

            try
            {
                paths.Refresh();

                bool catalogExists = File.Exists(paths.ReleaseHistoryFile);

                if (!catalogExists)
                {
                    logger?.LogDebug("UpdateBootstrapper: ReleaseHistory.json does not exist - treating this as the first-ever run on this machine.");

                    if (Directory.Exists(paths.ManifestDirectory) &&
                        File.Exists(paths.ManifestFile) &&
                        Directory.GetFiles(paths.ManifestDirectory, "*.zip").Any())
                    {
                        if (Directory.Exists(paths.VersionsPath))
                        {
                            try
                            {
                                logger?.LogDebug($"UpdateBootstrapper: wiping stray Versions\\ content before first-ever seed ('{paths.VersionsPath}').");
                                Directory.Delete(paths.VersionsPath, true);
                            }
                            catch (Exception ex)
                            {
                                // Non-fatal: if some stray content is locked (e.g. an
                                // orphaned Excel.exe - see CLAUDE.md section 29), proceed
                                // anyway rather than aborting the whole fresh-install
                                // seed. Worst case, unrelated stray folders remain
                                // alongside the newly seeded one - harmless clutter,
                                // consistent with this feature's own "never auto-prune
                                // Versions\" design.
                                logger?.LogException(ex, "UpdateBootstrapper: failed to wipe stray Versions\\ content - proceeding with the seed anyway.");
                            }
                        }

                        var seeded = ExtractManifestZipAndAdopt(context, "Install");
                        DeleteManifestFolder(context);
                        return seeded;
                    }

                    logger?.LogDebug("UpdateBootstrapper: no seed manifest+zip found on first-ever run - falling through with an empty catalog.");
                }

                if (Directory.Exists(paths.ManifestDirectory) && File.Exists(paths.ManifestFile))
                {
                    string zipPath = Directory.GetFiles(paths.ManifestDirectory, "*.zip").FirstOrDefault();
                    if (zipPath != null)
                    {
                        string candidateVersion = paths.LatestVersion;
                        string candidateReleaseDate = paths.LatestReleaseDate;

                        var existing = ReleaseHistoryStore.ReadAll(paths.ReleaseHistoryFile);
                        bool isKnownReinstall = existing.Any(e =>
                            string.Equals(e.Version, candidateVersion, StringComparison.OrdinalIgnoreCase) &&
                            string.Equals(e.ReleaseDate, candidateReleaseDate, StringComparison.OrdinalIgnoreCase));

                        if (isKnownReinstall)
                        {
                            logger?.LogDebug($"UpdateBootstrapper: manifest in Manifest\\ matches an existing catalog entry (version={candidateVersion}, releaseDate={candidateReleaseDate}) - reconciling before extracting.");
                            ReleaseHistoryStore.Reconcile(paths.ReleaseHistoryFile, paths.VersionsPath);
                        }

                        // A dev machine's GLSense\post_build.cmd copies a fresh zip+manifest
                        // into Manifest\ on every host rebuild, so this branch can win over
                        // whatever an Online fetch just downloaded/cataloged - the staged
                        // local dev build gets adopted here instead, silently, with only a
                        // LogDebug line (below) explaining why. That's out of scope to change
                        // here (which release wins is unaffected), but when this happens right
                        // after an Online fetch it's worth being loud about in the log.
                        if (string.Equals(source, "Online", StringComparison.OrdinalIgnoreCase))
                        {
                            logger?.LogWarn($"UpdateBootstrapper: a manifest+zip is already staged in '{paths.ManifestDirectory}' (version={candidateVersion}, releaseDate={candidateReleaseDate}) - adopting THAT staged release instead of the one just fetched via Online. If this wasn't expected, check whether a local dev rebuild (GLSense\\post_build.cmd) staged it.");
                        }

                        return ExtractManifestZipAndAdopt(context, source);
                    }
                }

                // Resolve "what's already installed" from the catalog itself, NOT from
                // paths.LatestVersion/LatestReleaseDate (manifest.json) - that file can
                // be (and, after a fresh-install seed, always is) deleted by this class,
                // and PathProvider lazily recreates it with unrelated hardcoded default
                // values the moment it's missing. Recomputing FolderName from those
                // defaults would silently point at a folder that doesn't exist, making
                // the add-in fail to load on the very next ordinary launch after a
                // successful install. The catalog is durable and unaffected by
                // manifest.json's lifecycle, so it's the correct source of truth here.
                var catalogEntries = ReleaseHistoryStore.ReadAll(paths.ReleaseHistoryFile);
                var activeEntry = catalogEntries
                    .Where(e => !string.IsNullOrWhiteSpace(e.FolderName) &&
                                Directory.Exists(Path.Combine(paths.VersionsPath, e.FolderName)) &&
                                Directory.GetFiles(Path.Combine(paths.VersionsPath, e.FolderName), "*.dll").Any())
                    .OrderByDescending(e => ParseReleaseDateOrMin(e.ReleaseDate))
                    .FirstOrDefault();

                if (activeEntry != null)
                {
                    logger?.LogDebug($"UpdateBootstrapper: no zip present, using catalog's most recent valid entry '{activeEntry.FolderName}'.");
                    return new ResolvedRelease { Version = activeEntry.Version, ReleaseDate = activeEntry.ReleaseDate, FolderName = activeEntry.FolderName };
                }

                logger?.LogError($"UpdateBootstrapper: no zip in '{paths.ManifestDirectory}' and no usable entry in the catalog - nothing to load.");
                return null;
            }
            catch (Exception ex)
            {
                logger?.LogException(ex, "UpdateBootstrapper.ResolveVersionToLoad");
                return null;
            }
        }

        /// <summary>
        /// Verifies zipBytes against the manifest's own checksum, extracts into
        /// Versions\{folderName}\, writes manifestJson as that folder's own
        /// manifest.json snapshot, and appends a ReleaseEntry to the catalog. Shared
        /// by ExtractManifestZipAndAdopt (which reads manifestJson/zipBytes from the
        /// Manifest\ folder on disk - see that method) and GLReloadSourcePicker's
        /// Online bulk-download loop (which fetches both over the network - see
        /// docs/superpowers/specs/2026-09-18-online-reload-multiversion-design.md).
        /// Returns null (does not throw) on a missing/unparseable version entry or a
        /// checksum mismatch - the caller decides what "one release in a batch
        /// failed" means for its own flow; nothing is extracted or catalogued in
        /// that case.
        /// </summary>
        public ResolvedRelease ExtractAndCatalog(IGLSenseContext context, string manifestJson, byte[] zipBytes, string source)
        {
            var logger = context.Logger;
            var paths = context.Paths;

            var parsed = new VersionParser(logger).ParseVersionJson(manifestJson);
            var info = parsed.AllVersions?.FirstOrDefault();
            if (info == null || string.IsNullOrWhiteSpace(info.Version))
            {
                logger?.LogError("UpdateBootstrapper.ExtractAndCatalog: manifest JSON did not contain a usable version entry.");
                return null;
            }

            string version = info.Version;
            string releaseDate = info.ReleaseDate;
            string folderName = !string.IsNullOrWhiteSpace(info.FolderName)
                ? info.FolderName
                : ReleaseHistoryStore.BuildFolderName(version, releaseDate);

            // folderName can be server-supplied on the Online path - validate it before
            // it ever reaches Path.Combine/a recursive Directory.Delete below. A single
            // Path.GetFileName(folderName) != folderName check catches path separators,
            // ".." traversal, and rooted paths all at once (any of those change what
            // GetFileName returns from the original string). Also refuse to ever
            // delete/overwrite the folder the currently-loaded release lives in.
            if (!string.Equals(Path.GetFileName(folderName), folderName, StringComparison.Ordinal) ||
                string.Equals(folderName, context.ActiveFolderName, StringComparison.OrdinalIgnoreCase))
            {
                logger?.LogError($"UpdateBootstrapper.ExtractAndCatalog: rejected folderName '{folderName}' for '{version}' ({releaseDate}) - either not a bare folder name or matches the currently active release's folder. Not extracting.");
                return null;
            }

            string checksum = info.Checksum ?? string.Empty;
            string notes = string.IsNullOrWhiteSpace(info.Notes) ? "Published by GLSense.Addin.Core" : info.Notes;

            string actualChecksum;
            using (var sha256 = SHA256.Create())
            {
                actualChecksum = BitConverter.ToString(sha256.ComputeHash(zipBytes)).Replace("-", "");
            }

            if (!string.IsNullOrWhiteSpace(checksum) &&
                !string.Equals(actualChecksum, checksum, StringComparison.OrdinalIgnoreCase))
            {
                logger?.LogError($"UpdateBootstrapper.ExtractAndCatalog: checksum mismatch for '{version}' ({releaseDate}) - expected {checksum}, got {actualChecksum}. Not extracting.");
                return null;
            }

            string versionFolder = Path.Combine(paths.VersionsPath, folderName);
            if (Directory.Exists(versionFolder))
                Directory.Delete(versionFolder, true);
            Directory.CreateDirectory(versionFolder);

            string tempZipPath = Path.Combine(Path.GetTempPath(), $"GLSenseOnline_{Guid.NewGuid():N}.zip");
            try
            {
                File.WriteAllBytes(tempZipPath, zipBytes);
                ZipFile.ExtractToDirectory(tempZipPath, versionFolder);
            }
            finally
            {
                if (File.Exists(tempZipPath))
                    File.Delete(tempZipPath);
            }

            File.WriteAllText(Path.Combine(versionFolder, "manifest.json"), manifestJson);

            var entry = new ReleaseEntry
            {
                Version = version,
                ReleaseDate = releaseDate,
                FolderName = folderName,
                Checksum = checksum,
                Notes = notes,
                Source = source
            };
            ReleaseHistoryStore.Append(paths.ReleaseHistoryFile, entry);

            logger?.LogDebug($"UpdateBootstrapper.ExtractAndCatalog: extracted and catalogued '{folderName}' (source={source}).");

            return new ResolvedRelease { Version = version, ReleaseDate = releaseDate, FolderName = folderName };
        }

        private ResolvedRelease ExtractManifestZipAndAdopt(IGLSenseContext context, string source)
        {
            var logger = context.Logger;
            var paths = context.Paths;

            string zipPath = Directory.GetFiles(paths.ManifestDirectory, "*.zip").First();
            string manifestJson = File.ReadAllText(paths.ManifestFile);
            byte[] zipBytes = File.ReadAllBytes(zipPath);

            logger?.LogDebug($"UpdateBootstrapper: extracting '{zipPath}' (source={source}).");

            var resolved = ExtractAndCatalog(context, manifestJson, zipBytes, source);
            if (resolved == null)
            {
                logger?.LogError($"UpdateBootstrapper: ExtractAndCatalog failed for '{zipPath}' - leaving it in place for the next launch to retry.");
                return null;
            }

            // Delete the zip only after the catalog append has genuinely succeeded -
            // if anything above throws, the zip is still there so the next launch can
            // retry the full extract+catalog sequence, instead of being left with DLLs
            // on disk but no catalog entry and no way to retry (the zip already gone).
            File.Delete(zipPath);

            logger?.LogDebug($"UpdateBootstrapper: extracted, catalogued (source={source}), and deleted '{zipPath}'. Adopting '{resolved.FolderName}'.");

            return resolved;
        }

        // "Newest wins" catalog lookups must sort by parsed date, not raw string - that
        // was only ever chronologically correct because every LOCAL writer emits the
        // fixed yyyy-MM-ddTHH:mm:ss format. The Online path is the first to ingest
        // releaseDate strings from an external server, which might use a different
        // format - falls back to DateTime.MinValue for an unparseable value, same
        // defensive shape as OnlineReleaseClassifier.ParseReleaseDateOrMin (a different
        // project/namespace - GLSense.Views vs. this one - so not shared/reused
        // directly here).
        private static DateTime ParseReleaseDateOrMin(string releaseDate)
        {
            return DateTime.TryParse(releaseDate, out var dt) ? dt : DateTime.MinValue;
        }

        private void DeleteManifestFolder(IGLSenseContext context)
        {
            try
            {
                if (Directory.Exists(context.Paths.ManifestDirectory))
                    Directory.Delete(context.Paths.ManifestDirectory, true);
            }
            catch (Exception ex)
            {
                context.Logger?.LogException(ex, "UpdateBootstrapper.DeleteManifestFolder");
            }
        }
    }
}
