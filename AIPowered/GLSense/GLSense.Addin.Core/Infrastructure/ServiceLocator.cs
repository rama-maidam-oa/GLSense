// GLSense.Addin.Core/Infrastructure/ServiceLocator.cs
using GLSense.Addin.Core.Helpers;
using GLSense.Contracts;
using System;
using System.Runtime.InteropServices;

namespace GLSense.Addin.Core.Infrastructure
{
    public static class ServiceLocator
    {
        private static IGLSenseContext _context;
        private static bool _isInitialized;
        private static SQLiteHelper _sqliteHelper;

        // Addin.Core's own Excel object (see IGLSenseContext.AcquireExcelUnknown). Created on the
        // thread that runs Initialize - Excel's main thread - so it belongs to Excel's apartment
        // like the host's object does. Released in Reset (shutdown and hot reload).
        private static Microsoft.Office.Interop.Excel.Application _localExcelApp;
        private static bool _localExcelUnavailable;
        private static readonly object _excelLock = new object();

        /// <summary>
        /// Gets the full context. All services are accessible through this.
        /// </summary>
        public static IGLSenseContext Context
        {
            get
            {
                if (!_isInitialized)
                    throw new InvalidOperationException("ServiceLocator has not been initialized. Call ServiceLocator.Initialize(context) first.");
                return _context;
            }
        }

        /// <summary>
        /// Gets the Logger from the context
        /// </summary>
        public static ILogger Logger
        {
            get
            {
                if (!_isInitialized)
                    throw new InvalidOperationException("ServiceLocator has not been initialized. Call ServiceLocator.Initialize(context) first.");
                return _context.Logger;
            }
        }

        /// <summary>
        /// Gets the RibbonController from the context
        /// </summary>
        public static IRibbonController RibbonController
        {
            get
            {
                if (!_isInitialized)
                    throw new InvalidOperationException("ServiceLocator has not been initialized. Call ServiceLocator.Initialize(context) first.");
                return _context.RibbonController;
            }
        }

        /// <summary>
        /// Gets the Edge Add-in instance from the context
        /// </summary>
        public static object EdgeAddinInstance
        {
            get
            {
                if (!_isInitialized)
                    throw new InvalidOperationException("ServiceLocator has not been initialized. Call ServiceLocator.Initialize(context) first.");
                return _context.EdgeAddinInstance;
            }
        }

        /// <summary>
        /// Gets the ExcelApp from the context
        /// </summary>
        public static Microsoft.Office.Interop.Excel.Application ExcelApp
        {
            get
            {
                if (!_isInitialized)
                    throw new InvalidOperationException("ServiceLocator has not been initialized. Call ServiceLocator.Initialize(context) first.");
                return GetLocalExcelApp() ?? _context.ExcelApp as Microsoft.Office.Interop.Excel.Application;
            }
        }

        /// <summary>
        /// Gets the PathProvider from the context
        /// </summary>
        public static IPathProvider Paths
        {
            get
            {
                if (!_isInitialized)
                    throw new InvalidOperationException("ServiceLocator has not been initialized. Call ServiceLocator.Initialize(context) first.");
                return _context.Paths;
            }
        }

        /// <summary>
        /// Gets the ExcelHandle from the context
        /// </summary>
        public static IntPtr ExcelHandle
        {
            get
            {
                if (!_isInitialized)
                    throw new InvalidOperationException("ServiceLocator has not been initialized. Call ServiceLocator.Initialize(context) first.");
                return _context.ExcelHandle;
            }
        }

        /// <summary>
        /// Gets the Version from the context
        /// </summary>
        public static string Version
        {
            get
            {
                if (!_isInitialized)
                    throw new InvalidOperationException("ServiceLocator has not been initialized. Call ServiceLocator.Initialize(context) first.");
                return _context.Version;
            }
        }

        /// <summary>
        /// Gets the ReleaseDate from the context
        /// </summary>
        public static string ReleaseDate
        {
            get
            {
                if (!_isInitialized)
                    throw new InvalidOperationException("ServiceLocator has not been initialized. Call ServiceLocator.Initialize(context) first.");
                return _context.ReleaseDate;
            }
        }

        /// <summary>
        /// Gets the folder name (under Versions\) that the currently loaded release's DLLs
        /// actually live in - see IGLSenseContext.ActiveFolderName. This is the correct,
        /// current way to locate any file shipped alongside the release (e.g. native DLLs);
        /// never reconstruct a "V{Version}"-style path from Version alone, since a release's
        /// folder name also includes its release date (see ReleaseHistoryStore.BuildFolderName).
        /// </summary>
        public static string ActiveFolderName
        {
            get
            {
                if (!_isInitialized)
                    throw new InvalidOperationException("ServiceLocator has not been initialized. Call ServiceLocator.Initialize(context) first.");
                return _context.ActiveFolderName;
            }
        }

        /// <summary>
        /// Gets the AllVersions from the context
        /// </summary>
        public static System.Collections.Generic.IReadOnlyList<VersionInfo> AllVersions
        {
            get
            {
                if (!_isInitialized)
                    throw new InvalidOperationException("ServiceLocator has not been initialized. Call ServiceLocator.Initialize(context) first.");
                return _context.AllVersions;
            }
        }

        /// <summary>
        /// Gets the singleton SQLiteHelper instance
        /// </summary>
        public static SQLiteHelper Database
        {
            get
            {
                if (!_isInitialized)
                    throw new InvalidOperationException("ServiceLocator has not been initialized. Call ServiceLocator.Initialize(context) first.");

                if (_sqliteHelper == null)
                {
                    _sqliteHelper = SQLiteHelper.Instance;
                    // Auto-initialize the database
                    _sqliteHelper.InitializeDatabase();
                }
                return _sqliteHelper;
            }
        }

        /// <summary>
        /// Check if ribbon controller is available
        /// </summary>
        public static bool IsRibbonAvailable => _isInitialized && _context?.RibbonController != null;

        /// <summary>
        /// Check if ServiceLocator is initialized
        /// </summary>
        public static bool IsInitialized => _isInitialized;

        /// <summary>
        /// Initialize the ServiceLocator with context
        /// </summary>
        // Returns Addin.Core's local Excel object, creating it on first use. If the host can't
        // provide the pointer (e.g. an older host) it returns null and the caller falls back to the
        // host's proxied object, as before.
        private static Microsoft.Office.Interop.Excel.Application GetLocalExcelApp()
        {
            if (_localExcelApp != null || _localExcelUnavailable)
                return _localExcelApp;

            lock (_excelLock)
            {
                if (_localExcelApp != null || _localExcelUnavailable)
                    return _localExcelApp;

                IntPtr unknown = IntPtr.Zero;
                try
                {
                    unknown = _context.AcquireExcelUnknown();
                    if (unknown == IntPtr.Zero)
                    {
                        _localExcelUnavailable = true;
                        _context?.Logger?.LogWarn("ServiceLocator: host returned no Excel pointer - using the host's Excel object.");
                        return null;
                    }

                    _localExcelApp = (Microsoft.Office.Interop.Excel.Application)Marshal.GetObjectForIUnknown(unknown);
                    _context?.Logger?.LogDebug("ServiceLocator: using Addin.Core's own Excel connection (no cross-AppDomain proxy).");
                }
                catch (Exception ex)
                {
                    _localExcelUnavailable = true;
                    _context?.Logger?.LogWarn($"ServiceLocator: could not create a local Excel connection, using the host's Excel object: {ex.GetType().Name}: {ex.Message}");
                }
                finally
                {
                    if (unknown != IntPtr.Zero)
                        Marshal.Release(unknown);
                }

                return _localExcelApp;
            }
        }

        public static void Initialize(IGLSenseContext context)
        {
            if (context == null)
                throw new ArgumentNullException(nameof(context));

            _context = context;
            _isInitialized = true;

            // Create the local Excel object now, on Excel's main thread (the host calls Initialize
            // from AddinModule_OnRibbonLoaded / ReloadAddinCore on that thread).
            _localExcelApp = null;
            _localExcelUnavailable = false;
            GetLocalExcelApp();

            // Context (and therefore Logger) is now set, so it's safe to log from here on.
            // Initialize the database automatically
            try
            {
                Logger?.LogDebug("ServiceLocator.Initialize: context set - initializing SQLite database...");
                var db = Database; // This triggers initialization
                Logger?.LogDebug("ServiceLocator.Initialize: SQLite database initialized successfully.");
            }
            catch (Exception ex)
            {
                // Deliberately swallowed (not rethrown) - a failed SQLite auto-init must not
                // prevent the rest of Initialize() (WPF app/ribbon setup) from running. Logged
                // via LogException instead of the previous LogError so the full stack trace
                // survives into the log file - this is exactly the kind of startup failure the
                // user won't be able to reproduce/re-debug later.
                Logger?.LogException(ex, "ServiceLocator.Initialize: Failed to initialize SQLiteHelper");
            }
        }

        /// <summary>
        /// Reset the ServiceLocator (useful for shutdown)
        /// </summary>
        public static void Reset()
        {
            // Log BEFORE tearing down _context/_isInitialized (the Logger property throws
            // once _isInitialized is false), and go through _context directly rather than the
            // Logger property so this never throws even if Reset() is called on an already-
            // uninitialized ServiceLocator.
            try
            {
                if (_isInitialized)
                {
                    _context?.Logger?.LogDebug("ServiceLocator.Reset: clearing context and cached services.");
                }
            }
            catch
            {
                // Logging itself must never throw or block shutdown/reset.
            }

            // Release Addin.Core's own Excel object and every Excel object still referenced in
            // this AppDomain, before it is unloaded (shutdown or hot reload). Same double collect
            // FinalWorkingCode does at shutdown; here it runs in the domain that holds them.
            try
            {
                if (_localExcelApp != null)
                    Marshal.FinalReleaseComObject(_localExcelApp);
            }
            catch
            {
                // Must never block shutdown/reset.
            }
            _localExcelApp = null;
            _localExcelUnavailable = false;

            _context = null;
            _isInitialized = false;
            _sqliteHelper = null;

            try
            {
                GC.Collect();
                GC.WaitForPendingFinalizers();
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
            catch
            {
                // Must never block shutdown/reset.
            }
        }

        /// <summary>
        /// Gets the Edge Add-in instance (searches if not cached)
        /// </summary>
        public static object GetEdgeAddinInstance()
        {
            if (!_isInitialized)
                throw new InvalidOperationException("ServiceLocator has not been initialized.");

            try
            {
                return _context.GetEdgeAddinInstance();
            }
            catch (Exception ex)
            {
                // Additive only - still rethrows (callers may depend on this throwing), just
                // makes sure the failure is captured in the log first.
                _context?.Logger?.LogException(ex, "ServiceLocator.GetEdgeAddinInstance");
                throw;
            }
        }
    }
}
