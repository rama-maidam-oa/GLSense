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
//
// ExtractAndCatalog is idempotent by design: if Versions\{folderName}\ already has
// DLLs on disk (e.g. this exact release was already extracted earlier - via Reload's
// Offline picker, an earlier Online fetch, or simply because it's the release
// currently/previously active this Excel session), it skips the delete+re-extract and
// just catalogs it. This matters because a native DLL that folder holds (e.g.
// e_sqlite3.dll) is never released once loaded via P/Invoke/LoadLibrary, even after
// the AppDomain that loaded it unloads - so deleting that folder again would throw
// UnauthorizedAccessException. Re-loading a release via Release History, then coming
// back to Reload and picking the SAME (already-extracted) build again, must not throw.
using GLSense.Contracts;
using GLSense.Shared;
using System;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Security.Cryptography;
using System.Threading;

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
            // embedded ".." traversal, and rooted paths all at once (any of those change
            // what GetFileName returns from the original string) - but it does NOT catch
            // a bare "." or ".." with no separator at all, since GetFileName performs no
            // dot-segment canonicalization on a separator-free input (GetFileName("..")
            // returns ".." unchanged). Path.Combine(VersionsPath, "..") resolves to
            // VersionsPath's own PARENT (the whole AddinCore install folder, including
            // Manifest\/ReleaseHistory.json/every other release), and "." resolves to
            // VersionsPath itself - either would be wiped by the Directory.Delete below
            // if not explicitly rejected here too. Also refuse to ever delete/overwrite
            // the folder the currently-loaded release lives in.
            if (!string.Equals(Path.GetFileName(folderName), folderName, StringComparison.Ordinal) ||
                folderName == "." || folderName == ".." ||
                string.Equals(folderName, context.ActiveFolderName, StringComparison.OrdinalIgnoreCase))
            {
                logger?.LogError($"UpdateBootstrapper.ExtractAndCatalog: rejected folderName '{folderName}' for '{version}' ({releaseDate}) - either not a bare folder name, a bare dot-segment, or matches the currently active release's folder. Not extracting.");
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

            // If this folder already exists with content, don't wipe and re-extract it -
            // it's the identical (version, releaseDate) release, already on disk from an
            // earlier extraction. Deleting it unconditionally used to crash here with
            // UnauthorizedAccessException on a native DLL (e_sqlite3.dll) whenever that
            // folder had ever been the active release during this Excel process: native
            // DLLs loaded via P/Invoke/LoadLibrary are never released when an AppDomain
            // unloads (a well-known .NET Framework limitation - unlike managed assemblies,
            // which are shadow-copied), so the file stays locked for the life of the
            // process even after switching to a different release via Reload/Release
            // History. Treat an already-populated folder as already-extracted instead of
            // trying to recreate identical content; ReleaseHistoryStore.Append below
            // already dedupes an identical catalog entry, so this stays idempotent.
            string versionFolder = Path.Combine(paths.VersionsPath, folderName);
            bool alreadyExtracted = Directory.Exists(versionFolder) &&
                Directory.GetFiles(versionFolder, "*.dll").Any();

            if (alreadyExtracted)
            {
                logger?.LogDebug($"UpdateBootstrapper.ExtractAndCatalog: '{folderName}' already exists on disk with content - skipping re-extraction (its files may still be locked by a native DLL loaded earlier in this process). Cataloging only.");
            }
            else
            {
                // This folder exists but has no DLLs yet (a stray/partial leftover, not a
                // genuine prior extraction) - try to clear it, but a transient lock (e.g.
                // antivirus still scanning a file that was only just written, or a handle
                // not yet released a beat after AppDomain.Unload/ReloadAddinCore's own
                // pre-flight window-close) shouldn't take the whole reload down either.
                // Retry briefly, then fall back to extracting into the folder as-is -
                // ZipFile.ExtractToDirectory overwrites whatever files it needs to, and any
                // genuinely unrelated leftover file is harmless clutter, consistent with
                // this feature's own "never auto-prune Versions\" design (see the first-run
                // seed's identical non-fatal wipe above).
                if (Directory.Exists(versionFolder))
                {
                    const int maxAttempts = 3;
                    for (int attempt = 1; attempt <= maxAttempts; attempt++)
                    {
                        try
                        {
                            Directory.Delete(versionFolder, true);
                            break;
                        }
                        catch (Exception ex) when (attempt < maxAttempts)
                        {
                            logger?.LogWarn($"UpdateBootstrapper.ExtractAndCatalog: attempt {attempt}/{maxAttempts} to clear stray '{folderName}' failed ({ex.GetType().Name}: {ex.Message}) - retrying shortly.");
                            Thread.Sleep(200);
                        }
                        catch (Exception ex)
                        {
                            logger?.LogException(ex, $"UpdateBootstrapper.ExtractAndCatalog: could not clear stray '{folderName}' after {maxAttempts} attempts - extracting into it as-is instead of failing the reload.");
                        }
                    }
                }

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
            }

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

            // Delete the zip only after the catalog append has genuinely succeeded - if
            // anything above throws, the zip is still there so the next launch can retry
            // the full extract+catalog sequence, instead of being left with DLLs on disk
            // but no catalog entry and no way to retry (the zip already gone). A failure
            // to delete it here (e.g. still locked by an antivirus scan of the file this
            // process just finished writing) must NOT undo the successful extract+catalog
            // above by throwing out of this method - that would report a working reload as
            // a failure. Leaving the zip behind on that path is harmless and safe to retry:
            // the next call into ExtractAndCatalog for this same manifest finds the target
            // folder already has DLLs and simply skips re-extraction (see that method's own
            // idempotency comment).
            try
            {
                File.Delete(zipPath);
                logger?.LogDebug($"UpdateBootstrapper: extracted, catalogued (source={source}), and deleted '{zipPath}'. Adopting '{resolved.FolderName}'.");
            }
            catch (Exception ex)
            {
                logger?.LogException(ex, $"UpdateBootstrapper: extracted and catalogued '{resolved.FolderName}' (source={source}), but could not delete '{zipPath}' - leaving it in place, harmless to retry. Adopting '{resolved.FolderName}' anyway.");
            }

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
