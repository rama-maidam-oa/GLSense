using System;
using System.Windows;
using System.Windows.Threading;

namespace GLSense.Utilities
{
    public static class WpfAppManager
    {
        private static readonly object _lock = new object();
        private static bool _dispatcherInitialized = false;

        public static void EnsureApplication()
        {
            LogUtility.LogInfo("[DPI-DIAG] WpfAppManager.EnsureApplication: start");
            if (Application.Current != null)
            {
                LogUtility.LogInfo("[DPI-DIAG] WpfAppManager.EnsureApplication: Application.Current already set, returning early");
                return;
            }

            lock (_lock)
            {
                if (Application.Current == null)
                {
                    LogUtility.LogDebug("WpfAppManager.EnsureApplication: no Application.Current, creating dedicated WPF Application on this thread.");
                    LogUtility.LogInfo("[DPI-DIAG] WpfAppManager.EnsureApplication: acquired lock, no Application.Current, creating dedicated thread/Application");
                    Dispatcher _wpfDispatcher;

                    // Set this thread's DPI awareness context to Per-Monitor-V2 and leave it set.
                    // IMPORTANT: this must NOT be reverted (no "using"/Dispose) here. The actual
                    // native window handles for every DpiAwareWindow are created lazily, well after
                    // this method returns (typically on Show()/ShowDialog()), all on this same
                    // dedicated WPF dispatcher thread. If the context is reverted immediately after
                    // this block (as it previously was), it has no effect on any real window's HWND
                    // creation, so the window ends up DPI-aware in name only - it gets bitmap-scaled
                    // by the OS instead of rendering natively per-monitor, which is what caused
                    // incorrect sizing/scaling and windows overflowing the screen's work area.
                    // Keeping the context active for the life of this thread ensures every window
                    // ever created on it is genuinely Per-Monitor-V2 DPI aware from HWND creation on.
                    DpiAwarenessHelper.SetPerMonitorAware();
                    LogUtility.LogInfo("[DPI-DIAG] WpfAppManager.EnsureApplication: SetPerMonitorAware done, before new Application()");

                    var app = new Application
                    {
                        ShutdownMode = ShutdownMode.OnExplicitShutdown
                    };
                    LogUtility.LogInfo("[DPI-DIAG] WpfAppManager.EnsureApplication: Application created");

                    // Store the dispatcher for later use
                    _wpfDispatcher = app.Dispatcher;

                    // Initialize dispatcher properly
                    if (!_dispatcherInitialized)
                    {
                        LogUtility.LogInfo("[DPI-DIAG] WpfAppManager.EnsureApplication: before dummy-control dispatcher.Invoke (cross-thread boundary)");
                        // Create a dummy control on the UI thread to initialize dispatcher
                        _wpfDispatcher.Invoke(() =>
                        {
                            var dummy = new System.Windows.Controls.Control();
                        });
                        LogUtility.LogInfo("[DPI-DIAG] WpfAppManager.EnsureApplication: dummy-control dispatcher.Invoke returned OK");
                        _dispatcherInitialized = true;
                    }

                    app.DispatcherUnhandledException += OnDispatcherUnhandledException;
                    LogUtility.LogDebug("WpfAppManager.EnsureApplication: WPF Application created and dispatcher initialized.");
                    LogUtility.LogInfo("[DPI-DIAG] WpfAppManager.EnsureApplication: end");
                }
            }
        }

        private static void OnDispatcherUnhandledException(object sender,
            DispatcherUnhandledExceptionEventArgs e)
        {
            LogUtility.LogException(e.Exception, "WpfAppManager.OnDispatcherUnhandledException", forceLog: true);
            if (e.Exception is System.IO.IOException)
            {
                // Don't mark as handled for IO exceptions so we can see them
            }
            else
            {
                e.Handled = true; // Prevent application crash for non-IO exceptions
            }
        }

        public static void InvokeOnWpfThread(Action action)
        {
            LogUtility.LogInfo($"[DPI-DIAG] WpfAppManager.InvokeOnWpfThread: start, callingThread={System.Threading.Thread.CurrentThread.ManagedThreadId}");
            if (Application.Current != null)
            {
                var dispatcher = Application.Current.Dispatcher;
                if (dispatcher.CheckAccess())
                {
                    LogUtility.LogInfo("[DPI-DIAG] WpfAppManager.InvokeOnWpfThread: already on dispatcher thread, calling SafeExecute directly");
                    SafeExecute(action);
                }
                else
                {
                    LogUtility.LogWarn($"[DPI-DIAG] WpfAppManager.InvokeOnWpfThread: before BLOCKING dispatcher.Invoke (cross-thread boundary, dispatcherThread={dispatcher.Thread?.ManagedThreadId})");
                    dispatcher.Invoke(() => SafeExecute(action));
                    LogUtility.LogWarn("[DPI-DIAG] WpfAppManager.InvokeOnWpfThread: dispatcher.Invoke returned OK");
                }
            }
            else
            {
                LogUtility.LogWarn("[DPI-DIAG] WpfAppManager.InvokeOnWpfThread: Application.Current is null, action NOT executed");
            }
            LogUtility.LogInfo("[DPI-DIAG] WpfAppManager.InvokeOnWpfThread: end");
        }
        public static T InvokeOnWpfThread<T>(Func<T> func)
        {
            if (Application.Current != null)
            {
                var dispatcher = Application.Current.Dispatcher;
                if (dispatcher.CheckAccess())
                {
                    return SafeExecute(func);
                }
                else
                {
                    return dispatcher.Invoke(() => SafeExecute(func));
                }
            }
            return default(T);
        }

        private static T SafeExecute<T>(Func<T> func)
        {
            try
            {
                return func();
            }
            catch (System.IO.IOException ex)
            {
                LogUtility.LogException(ex, "WpfAppManager.SafeExecute<T> (IO, retrying)");
                // Retry once after a small delay
                System.Threading.Thread.Sleep(100);
                try
                {
                    return func();
                }
                catch (Exception retryEx)
                {
                    LogUtility.LogException(retryEx, "WpfAppManager.SafeExecute<T> (retry failed)");
                    throw;
                }
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex, "WpfAppManager.SafeExecute<T>");
                throw;
            }
        }
        private static void SafeExecute(Action action)
        {
            try
            {
                action();
            }
            catch (System.IO.IOException ex)
            {
                LogUtility.LogException(ex, "WpfAppManager.SafeExecute (IO, retrying)");
                // Retry once after a small delay
                System.Threading.Thread.Sleep(100);
                try
                {
                    action();
                }
                catch (Exception retryEx)
                {
                    LogUtility.LogException(retryEx, "WpfAppManager.SafeExecute (retry failed)");
                    throw;
                }
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex, "WpfAppManager.SafeExecute");
                throw;
            }
        }
    }
}