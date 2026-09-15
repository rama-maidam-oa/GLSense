using System;
using System.Threading;
using System.Windows;
using System.Windows.Threading;

namespace GLSense.Utilities
{
    public static class WpfAppManager
    {
        private static readonly object _lock = new object();
        private static bool _dispatcherInitialized = false;

        /// <summary>Installs a DispatcherSynchronizationContext for the CURRENT thread, so
        /// that any `await` executed on this thread reliably resumes back on this same
        /// thread afterward, instead of on an arbitrary threadpool thread.
        ///
        /// Root cause this fixes: Excel's main STA thread - where every COM event handler,
        /// ribbon click, and (once GLBalanceConfiguratorForm/GLConfiguratorPane construct
        /// their WPF content on it) WPF DependencyObject lives - never had a
        /// SynchronizationContext installed on it at all (confirmed via [ThreadDiag]
        /// logging: SynchronizationContext.Current was null there). A normal WPF app gets
        /// this for free from Application.Run(); this add-in never calls that on Excel's
        /// own thread, so `await` had nothing to marshal its continuation through and fell
        /// back to the default threadpool scheduler. That surfaced repeatedly as
        /// InvalidOperationException/NotSupportedException ("a different thread owns it" /
        /// "changes to its SourceCollection from a thread different from the Dispatcher
        /// thread") from unrelated-looking places across GLConfiguratorViewModel
        /// (LoadDataAsync, UpdateParameterSummary via ApplyDefaultSelections) and
        /// GLBalanceConfigurator (ReLoadConfigurator) - every one of them just a different
        /// `await` in the same call tree losing the thread the same way. Call this once,
        /// early, on Excel's own thread (AddinModule_AddinInitialize) rather than patching
        /// every individual call site as each one surfaces.</summary>
        public static void EnsureMainThreadSynchronizationContext()
        {
            try
            {
                if (SynchronizationContext.Current is DispatcherSynchronizationContext)
                {
                    return;
                }

                SynchronizationContext.SetSynchronizationContext(
                    new DispatcherSynchronizationContext(Dispatcher.CurrentDispatcher));
                LogUtility.LogDebug("WpfAppManager.EnsureMainThreadSynchronizationContext: installed DispatcherSynchronizationContext for this thread.");
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex, "WpfAppManager.EnsureMainThreadSynchronizationContext");
            }
        }

        public static void EnsureApplication()
        {
            if (Application.Current != null)
            {
                return;
            }

            lock (_lock)
            {
                if (Application.Current == null)
                {
                    LogUtility.LogDebug("WpfAppManager.EnsureApplication: no Application.Current, creating dedicated WPF Application on this thread.");
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

                    var app = new Application
                    {
                        ShutdownMode = ShutdownMode.OnExplicitShutdown
                    };

                    // Store the dispatcher for later use
                    _wpfDispatcher = app.Dispatcher;

                    // Initialize dispatcher properly
                    if (!_dispatcherInitialized)
                    {
                        // Create a dummy control on the UI thread to initialize dispatcher
                        _wpfDispatcher.Invoke(() =>
                        {
                            var dummy = new System.Windows.Controls.Control();
                        });
                        _dispatcherInitialized = true;
                    }

                    app.DispatcherUnhandledException += OnDispatcherUnhandledException;
                    LogUtility.LogDebug("WpfAppManager.EnsureApplication: WPF Application created and dispatcher initialized.");
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
            if (Application.Current != null)
            {
                var dispatcher = Application.Current.Dispatcher;
                if (dispatcher.CheckAccess())
                {
                    SafeExecute(action);
                }
                else
                {
                    dispatcher.Invoke(() => SafeExecute(action));
                }
            }
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
