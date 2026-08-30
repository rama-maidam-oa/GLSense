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
                            logger?.LogDebug($"UpdateBootstrapper: wiping stray Versions\\ content before first-ever seed ('{paths.VersionsPath}').");
                            Directory.Delete(paths.VersionsPath, true);
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

                        return ExtractManifestZipAndAdopt(context, source);
                    }
                }

                string localVersion = paths.LatestVersion;
                string localReleaseDate = paths.LatestReleaseDate;
                string folderName = ReleaseHistoryStore.BuildFolderName(localVersion, localReleaseDate);
                string versionFolder = Path.Combine(paths.VersionsPath, folderName);
                bool haveLocalDlls = Directory.Exists(versionFolder) && Directory.GetFiles(versionFolder, "*.dll").Any();

                if (haveLocalDlls)
                {
                    logger?.LogDebug($"UpdateBootstrapper: no zip present, but '{versionFolder}' already has DLLs - using installed version '{localVersion}'.");
                    return new ResolvedRelease { Version = localVersion, ReleaseDate = localReleaseDate, FolderName = folderName };
                }

                logger?.LogError($"UpdateBootstrapper: no zip in '{paths.ManifestDirectory}' and no usable install at '{versionFolder}' - nothing to load.");
                return null;
            }
            catch (Exception ex)
            {
                logger?.LogException(ex, "UpdateBootstrapper.ResolveVersionToLoad");
                return null;
            }
        }

        private ResolvedRelease ExtractManifestZipAndAdopt(IGLSenseContext context, string source)
        {
            var logger = context.Logger;
            var paths = context.Paths;

            string version = paths.LatestVersion;
            string releaseDate = paths.LatestReleaseDate;
            string folderName = ReleaseHistoryStore.BuildFolderName(version, releaseDate);
            string versionFolder = Path.Combine(paths.VersionsPath, folderName);
            string zipPath = Directory.GetFiles(paths.ManifestDirectory, "*.zip").First();

            logger?.LogDebug($"UpdateBootstrapper: extracting '{zipPath}' into '{versionFolder}' (source={source}).");

            if (Directory.Exists(versionFolder))
                Directory.Delete(versionFolder, true);
            Directory.CreateDirectory(versionFolder);

            ZipFile.ExtractToDirectory(zipPath, versionFolder);
            File.Delete(zipPath);

            // Per-version manifest snapshot - a permanent, self-contained record of
            // exactly what this folder is, independent of the transient copy in
            // Manifest\ (which the caller may delete afterward - e.g. the fresh-install
            // path).
            File.Copy(paths.ManifestFile, Path.Combine(versionFolder, "manifest.json"), true);

            var entry = new ReleaseEntry
            {
                Version = version,
                ReleaseDate = releaseDate,
                FolderName = folderName,
                Checksum = paths.LatestChecksum,
                Notes = string.IsNullOrWhiteSpace(paths.LatestNotes) ? "Published by GLSense.Addin.Core" : paths.LatestNotes,
                Source = source
            };
            ReleaseHistoryStore.Append(paths.ReleaseHistoryFile, entry);

            logger?.LogDebug($"UpdateBootstrapper: extracted, catalogued (source={source}), and deleted '{zipPath}'. Adopting '{folderName}'.");

            return new ResolvedRelease { Version = version, ReleaseDate = releaseDate, FolderName = folderName };
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
