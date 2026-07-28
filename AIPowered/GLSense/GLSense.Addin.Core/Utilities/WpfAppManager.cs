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

                    // Set PerMonitorV2 DPI awareness BEFORE creating Application
                    using (DpiAwarenessHelper.SetPerMonitorAware())
                    {
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