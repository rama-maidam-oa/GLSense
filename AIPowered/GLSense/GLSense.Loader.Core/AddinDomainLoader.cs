//GLSense.Loader.Core/AddinDomainLoader.cs
using GLSense.Contracts;
using System;
using System.IO;

namespace GLSense.Loader.Core
{
    public class AddinDomainLoader
    {
        private IGLSenseContext _context;
        private AppDomain _domain;
        private IGLSenseAddin _instance;

        public IGLSenseAddin Load(IGLSenseContext context)
        {
            _context = context;

            string versionFolderName = $"V{_context.Version}";
            string dllPath = Path.Combine(_context.Paths.VersionsPath, versionFolderName);

            //LogFilesInFolder(dllPath);  this is used for debugging

            _context.Logger?.LogDebug($"AddinDomainLoader.Load: preparing to load add-in version '{_context.Version}' from '{dllPath}'.");

            try
            {
                var setup = new AppDomainSetup
                {
                    ApplicationBase = dllPath,
                    ShadowCopyFiles = "true",
                    CachePath = Path.Combine(dllPath, "ShadowCache"),
                    PrivateBinPath = dllPath,
                    PrivateBinPathProbe = dllPath
                };

                _domain = AppDomain.CreateDomain($"GLSenseDomain_{_context.Version}", null, setup);
                _context.Logger?.LogDebug($"AddinDomainLoader.Load: created AppDomain 'GLSenseDomain_{_context.Version}'.");

                // No need for AssemblyResolve here - it's handled in RemoteLoader
                var loader = (RemoteLoader)_domain.CreateInstanceAndUnwrap(
                    typeof(RemoteLoader).Assembly.FullName,
                    typeof(RemoteLoader).FullName);
                _context.Logger?.LogDebug("AddinDomainLoader.Load: RemoteLoader created and unwrapped in the new AppDomain.");

                _instance = loader.Create(dllPath, context);
                _context.Logger?.LogDebug("AddinDomainLoader.Load: add-in instance created by RemoteLoader; initializing.");

                _instance.Initialize(context);
                _context.Logger?.LogDebug($"AddinDomainLoader.Load: add-in version '{_context.Version}' initialized successfully.");

                return _instance;
            }
            catch (Exception ex)
            {
                _context.Logger?.LogError($"AddinDomainLoader.Load: failed to load add-in version '{_context.Version}' from '{dllPath}'.", ex);
                throw;
            }
        }

        private void LogFilesInFolder(string folderPath)
        {
            try
            {
                if (!Directory.Exists(folderPath))
                {
                    _context.Logger?.LogError($"Folder does not exist: {folderPath}");
                    return;
                }

                var allFiles = Directory.GetFiles(folderPath, "*.*");
                _context.Logger?.LogDebug($"Files in {folderPath}:");
                foreach (var file in allFiles)
                {
                    _context.Logger?.LogDebug($"  - {Path.GetFileName(file)}");
                }
            }
            catch (Exception ex)
            {
                _context.Logger?.LogError($"Error logging files: {ex.Message}", ex);
            }
        }

        public void Unload(IGLSenseContext ctx)
        {
            try
            {
                if (_domain != null)
                {
                    ctx.Logger?.LogDebug("AddinDomainLoader.Unload: unloading AppDomain.");
                    AppDomain.Unload(_domain);
                    _domain = null;
                    ctx.Logger?.LogDebug("AddinDomainLoader.Unload: AppDomain unloaded successfully.");
                }
                else
                {
                    ctx.Logger?.LogDebug("AddinDomainLoader.Unload: no AppDomain to unload (already null).");
                }
            }
            catch (Exception ex)
            {
                ctx.Logger?.LogError($"AddinDomainLoader.Unload: error unloading add-in AppDomain: {ex.Message}", ex);
            }
        }
    }
}
