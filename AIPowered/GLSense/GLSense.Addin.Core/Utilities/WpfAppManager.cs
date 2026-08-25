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
                            //
                            // Explicit LogWarn marker (kept permanently - see the GLLogin
                            // "blank combo box" investigation): a XamlParseException/
                            // StaticResource-resolution failure thrown while a control's
                            // Style/ControlTemplate is being applied would land here, get
                            // swallowed by e.Handled=true, and leave that control's area
                            // blank with nothing else visibly wrong - this line makes that
                            // scenario unmistakable in the log instead of just another
                            // LogException entry among many.
                            ServiceLocator.Logger?.LogWarn($"WpfAppManager: UNHANDLED WPF DISPATCHER EXCEPTION being suppressed - {e.Exception.GetType().Name}: {e.Exception.Message}");
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

        /// <summary>
        /// Explicitly shuts down the WPF Application created by EnsureApplication().
        /// ShutdownMode is OnExplicitShutdown (deliberately - see EnsureApplication's own
        /// comment: the Application/Dispatcher needs to stay alive across every window
        /// open/close during a normal session), which means nothing ever terminates its
        /// Dispatcher message loop unless something calls this. Without it, on a real
        /// Excel shutdown the thread hosting this Application/Dispatcher is still alive
        /// and actively pumping when AddinDomainLoader.Unload() tries to unload this
        /// AppDomain - a live thread executing inside a domain being unloaded is exactly
        /// what makes AppDomain.Unload() hang/block, which in turn blocks
        /// AddinModule_AddinBeginShutdown from returning, which blocks Excel's own
        /// shutdown sequence - the most likely explanation for Excel.exe lingering as an
        /// orphaned process after its window closes. Called from AddinEntry.Shutdown()
        /// right before the AppDomain unload is attempted.
        /// </summary>
        public static void Shutdown()
        {
            try
            {
                if (Application.Current == null || Application.Current.Dispatcher.HasShutdownStarted)
                {
                    ServiceLocator.Logger?.LogDebug("WpfAppManager.Shutdown: no live Application to shut down.");
                    return;
                }

                ServiceLocator.Logger?.LogDebug("WpfAppManager.Shutdown: shutting down the WPF Application/Dispatcher.");

                // Application.Shutdown() must run on the Dispatcher thread that owns it -
                // InvokeOnWpfThread already handles the "already on that thread vs. needs
                // Invoke" branching correctly.
                InvokeOnWpfThread(() => Application.Current?.Shutdown());

                _isInitialized = false;
            }
            catch (Exception ex)
            {
                ServiceLocator.Logger?.LogException(ex, "WpfAppManager.Shutdown");
            }
        }
    }
}