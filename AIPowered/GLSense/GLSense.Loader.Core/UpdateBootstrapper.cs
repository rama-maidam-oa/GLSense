// UpdateBootstrapper.cs in GLSense.Loader.Core
//
// Pre-AppDomain-load bootstrap: decides which version of GLSense.Addin.Core to load,
// and gets it onto disk if it isn't there yet, BEFORE AddinDomainLoader ever creates the
// AppDomain. Lives here (not in GLSense.Addin.Core) because Addin.Core isn't loaded yet
// at this point and can't be responsible for replacing itself - see
// AddinModule_OnRibbonLoaded / AddinModule.ReloadAddinCore (GLSense host project) for
// the call sites.
//
// FOLDER-ONLY, no remote/network step (deliberately simplified for local testing - see
// CLAUDE.md section 17). This used to also check a local HTTP host
// (GLSense.LocalUpdateHost) for a newer release and download it; that was removed
// because it kept failing to connect in practice (the host console app is an easy-to-
// forget extra manual step) and added a lot of moving parts for something that's still
// just being tested. The three-tier design's "online" tier is expected to come back
// later, once this simpler folder-driven flow is confirmed working end to end.
//
// Decision tree:
//   1. Manifest folder doesn't exist at all -> nothing to bootstrap from, return null.
//      Defensive only - PathProvider.Ensure() already creates this folder before this
//      ever runs, so in practice this should never actually be hit.
//   2. Manifest folder has BOTH manifest.json AND a .zip -> extract the zip into
//      Versions\V{version}\ (version read from the local manifest.json), delete the
//      zip, done. This is now the ONLY way a version ever gets installed/updated -
//      post_build.cmd drops a fresh zip + manifest.json here on every build (see
//      GLSense.Addin.Core\post_build.cmd).
//   3. Manifest folder has ONLY manifest.json (no zip) -> if Versions\V{version}\
//      already has DLLs on disk, use them (already installed, nothing to do).
//      Otherwise there is nothing usable - return null so the caller can skip loading
//      the AppDomain instead of crashing.
using GLSense.Contracts;
using System;
using System.IO;
using System.IO.Compression;
using System.Linq;

namespace GLSense.Loader.Core
{
    public class UpdateBootstrapper
    {
        public string ResolveVersionToLoad(IGLSenseContext context)
        {
            var logger = context.Logger;
            var paths = context.Paths;

            try
            {
                // PathProvider.LatestVersion/LatestReleaseDate are cached (static) fields,
                // only re-parsed from manifest.json when Refresh()/InitializeVersion() runs.
                // The PathProvider instance backing this call was constructed once, at
                // AddinModule_OnRibbonLoaded time - on a manual Reload (RibReload_OnClick ->
                // ReloadAddinCore), a rebuild could easily have happened in between,
                // overwriting manifest.json on disk with a new version/releaseDate that this
                // cache doesn't know about yet. Refresh unconditionally so every read below
                // (and GLAbout's version/build-date display, which reads the same cached
                // fields via ServiceLocator.Version/ReleaseDate) reflects what's actually on
                // disk right now, not whatever was true at Excel startup.
                paths.Refresh();

                if (!Directory.Exists(paths.ManifestDirectory))
                {
                    logger?.LogError($"UpdateBootstrapper: Manifest folder does not exist ('{paths.ManifestDirectory}') - nothing to bootstrap from.");
                    return null;
                }

                if (!File.Exists(paths.ManifestFile))
                {
                    // Shouldn't normally happen - PathProvider seeds a default manifest.json
                    // as soon as the folder is ensured - but if it's ever deleted out from
                    // under us, there's nothing local to read.
                    logger?.LogError($"UpdateBootstrapper: manifest.json not found in '{paths.ManifestDirectory}' - nothing to bootstrap from.");
                    return null;
                }

                string zipPath = Directory.GetFiles(paths.ManifestDirectory, "*.zip").FirstOrDefault();
                if (zipPath != null)
                {
                    return ExtractLocalZipAndAdopt(context, zipPath);
                }

                string localVersion = paths.LatestVersion;
                string versionFolder = Path.Combine(paths.VersionsPath, $"V{localVersion}");
                bool haveLocalDlls = Directory.Exists(versionFolder) && Directory.GetFiles(versionFolder, "*.dll").Any();

                if (haveLocalDlls)
                {
                    logger?.LogDebug($"UpdateBootstrapper: no zip present, but '{versionFolder}' already has DLLs - using installed version '{localVersion}'.");
                    return localVersion;
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

        private string ExtractLocalZipAndAdopt(IGLSenseContext context, string zipPath)
        {
            var logger = context.Logger;
            var paths = context.Paths;

            string version = paths.LatestVersion;
            string versionFolder = Path.Combine(paths.VersionsPath, $"V{version}");

            logger?.LogDebug($"UpdateBootstrapper: found zip '{zipPath}' alongside manifest.json - extracting to '{versionFolder}'.");

            if (Directory.Exists(versionFolder))
                Directory.Delete(versionFolder, true);
            Directory.CreateDirectory(versionFolder);

            ZipFile.ExtractToDirectory(zipPath, versionFolder);
            File.Delete(zipPath);

            logger?.LogDebug($"UpdateBootstrapper: extracted and deleted '{zipPath}'. Adopting version '{version}'.");
            return version;
        }
    }
}
