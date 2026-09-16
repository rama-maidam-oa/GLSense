//GLSense.Loader.Core/AddinDomainLoader.cs
using GLSense.Contracts;
using System;
using System.IO;
using System.Threading.Tasks;

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

            string dllPath = Path.Combine(_context.Paths.VersionsPath, _context.ActiveFolderName);

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

        // AppDomain.Unload() has no built-in timeout - if any thread is still executing
        // (or blocked in a COM call, e.g. a live Excel RCW) inside the domain being
        // unloaded, it can block indefinitely waiting to abort that thread. Since this is
        // called from Excel's own shutdown event handler (AddinModule_AddinBeginShutdown),
        // a hang here previously meant Excel's own shutdown never completed either -
        // Excel.exe lingering as an orphaned process in Task Manager even after its
        // window closed. Running the unload on its own Task and bounding the wait means
        // this method - and therefore Excel's shutdown sequence - can never be blocked
        // for more than UnloadTimeoutSeconds, regardless of whether the unload itself
        // ever actually completes. The Task's own thread is a background/thread-pool
        // thread (IsBackground=true by default), so even if AppDomain.Unload() truly
        // never returns, it can't by itself keep the process alive past this method
        // returning - only a live foreground thread can do that.
        private const int UnloadTimeoutSeconds = 5;

        public void Unload(IGLSenseContext ctx)
        {
            try
            {
                if (_domain == null)
                {
                    ctx.Logger?.LogDebug("AddinDomainLoader.Unload: no AppDomain to unload (already null).");
                    return;
                }

                var domainToUnload = _domain;
                _domain = null;

                ctx.Logger?.LogDebug("AddinDomainLoader.Unload: unloading AppDomain.");

                var unloadTask = Task.Run(() => AppDomain.Unload(domainToUnload));
                if (unloadTask.Wait(TimeSpan.FromSeconds(UnloadTimeoutSeconds)))
                {
                    ctx.Logger?.LogDebug("AddinDomainLoader.Unload: AppDomain unloaded successfully.");
                }
                else
                {
                    ctx.Logger?.LogError($"AddinDomainLoader.Unload: AppDomain.Unload did not complete within {UnloadTimeoutSeconds}s - abandoning the wait so shutdown can proceed. The domain may still be torn down asynchronously in the background.");
                }
            }
            catch (Exception ex)
            {
                ctx.Logger?.LogError($"AddinDomainLoader.Unload: error unloading add-in AppDomain: {ex.Message}", ex);
            }
        }
    }
}
