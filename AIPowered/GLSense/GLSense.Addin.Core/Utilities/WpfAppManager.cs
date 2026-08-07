using GLSense.Addin.Core.Helpers;
using GLSense.Addin.Core.Infrastructure;
using System;
using System.Windows;
using System.Windows.Threading;

namespace GLSense.Addin.Core.Utilities
{
    public static class WpfAppManager
    {
        private static bool _isInitialized;
        private static readonly object _lock = new object();

        public static void EnsureApplication()
        {
            if (_isInitialized)
                return;

            lock (_lock)
            {
                if (_isInitialized)
                    return;

                try
                {
                    ServiceLocator.Logger?.LogDebug("WpfAppManager.EnsureApplication: starting one-time WPF Application initialization.");

                    // Set this thread's DPI awareness context to Per-Monitor-V2 and leave it
                    // set. IMPORTANT: this must NOT be reverted (no "using"/Dispose) here.
                    // Every actual WPF window on this thread - every BaseWindow-derived
                    // dialog, and ConfiguratorPaneHost's borderless Window for the Balance
                    // Configurator - is created later via separate InvokeOnWpfThread calls,
                    // well after this method returns, all on this same dedicated WPF
                    // dispatcher thread. Neither `new Application()` nor wiring
                    // Dispatcher.UnhandledException below creates any native window handle -
                    // only Window.Show()/EnsureHandle() elsewhere does, later. If the context
                    // is reverted immediately after this block (as it previously was, via a
                    // "using" scope), it has no effect on any real window's HWND creation, so
                    // every window ends up DPI-aware in name only - it gets bitmap-scaled by
                    // the OS instead of rendering natively per-monitor, which is what caused
                    // windows (including the Balance Configurator) to intermittently render
                    // "zoomed in"/blurry for some users, especially at >100% display scaling.
                    // Keeping the context active for the life of this thread ensures every
                    // window ever created on it is genuinely Per-Monitor-V2 DPI aware from
                    // HWND creation on. Mirrors FinalWorkingCode's WpfAppManager.cs, which
                    // documents this exact same fix for the identical reason.
                    DpiAwarenessHelper.SetPerMonitorAware();

                    if (Application.Current == null)
                    {
                        var app = new Application();
                        app.ShutdownMode = ShutdownMode.OnExplicitShutdown;
                        ServiceLocator.Logger?.LogDebug("WpfAppManager: Created new Application instance with DPI awareness.");
                    }
                    else
                    {
                        ServiceLocator.Logger?.LogDebug("WpfAppManager: Using existing Application instance.");
                    }

                    // Set DPI awareness on the application
                    if (Application.Current != null)
                    {
                        Application.Current.Dispatcher.UnhandledException += (s, e) =>
                        {
                            // This is every unhandled exception thrown on the WPF Dispatcher
                            // thread across every GLSense window (Login, CubeDetails, all
                            // Group C/H/I dialogs, etc.) - log the full structured dump, not
                            // just the message, since e.Handled=true means it would
                            // otherwise vanish silently.
                            ServiceLocator.Logger?.LogException(e.Exception, "WpfAppManager: Unhandled WPF Dispatcher exception (suppressed, UI kept alive)");
                            e.Handled = true;
                        };
                    }

                    _isInitialized = true;
                    ServiceLocator.Logger?.LogDebug("WpfAppManager: Initialization completed with PerMonitorV2 DPI awareness.");
                }
                catch (Exception ex)
                {
                    ServiceLocator.Logger?.LogException(ex, "WpfAppManager.EnsureApplication: Initialization failed");
                    throw;
                }
            }
        }

        public static void InvokeOnWpfThread(Action action)
        {
            if (action == null)
                return;

            try
            {
                EnsureApplication();

                if (Application.Current != null && !Application.Current.Dispatcher.HasShutdownStarted)
                {
                    if (Application.Current.Dispatcher.CheckAccess())
                    {
                        action();
                    }
                    else
                    {
                        Application.Current.Dispatcher.Invoke(action);
                    }
                }
                else
                {
                    if (Dispatcher.CurrentDispatcher.CheckAccess())
                    {
                        action();
                    }
                    else
                    {
                        Dispatcher.CurrentDispatcher.Invoke(action);
                    }
                }
            }
            catch (Exception ex)
            {
                ServiceLocator.Logger?.LogException(ex, "WpfAppManager.InvokeOnWpfThread");
                throw;
            }
        }

        public static bool IsInitialized => _isInitialized;
    }
}