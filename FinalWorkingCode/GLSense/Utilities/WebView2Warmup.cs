using Microsoft.Web.WebView2.Core;
using Microsoft.Web.WebView2.Wpf;
using System;
using System.IO;
using System.Runtime.InteropServices;
using System.Threading.Tasks;

namespace GLSense.Utilities
{
    // CoreWebView2Environment.CreateAsync spins up a real Chromium process tree (browser +
    // GPU + network service + renderer + crashpad handler) - several seconds cold. GLLogin
    // and GLDrilldownCustomization both used to call CreateAsync independently against the
    // exact same AppPaths.LoginBrowserLogsPath/options, so whichever window opened first paid
    // that cost with zero visual feedback (confirmed via WindowDelay.gif: the WebView2 area
    // just sat blank for several seconds before GLLogin's login page appeared). This kicks the
    // same CreateAsync call off once, in the background, at ribbon load, and hands out the one
    // shared environment to both windows.
    public static class WebView2Warmup
    {
        private static readonly object _lock = new object();
        private static Task<CoreWebView2Environment> _environmentTask;

        public static void WarmUpInBackground()
        {
            lock (_lock)
            {
                if (_environmentTask == null)
                    _environmentTask = CreateEnvironmentAsync();
            }
        }

        public static Task<CoreWebView2Environment> GetEnvironmentAsync()
        {
            lock (_lock)
            {
                if (_environmentTask == null)
                    _environmentTask = CreateEnvironmentAsync();
                return _environmentTask;
            }
        }

        // 0x800700AA (ERROR_BUSY): the login browser profile folder is still held, e.g. by a WebView2
        // browser process that has not finished exiting. Seen in a customer log on every login
        // attempt for ~30s until Excel was restarted, and the window just closed silently.
        private const int ErrorBusy = unchecked((int)0x800700AA);
        private static readonly TimeSpan BusyRetryWindow = TimeSpan.FromSeconds(15);

        public const string InitFailedMessage =
            "Unable to start the browser component. Please wait a moment and try again. " +
            "If this keeps happening, close Excel completely and reopen it.";

        /// <summary>
        /// Initializes <paramref name="webView"/> with the shared environment, retrying for up to
        /// 15s while WebView2 reports ERROR_BUSY. Each retry drops the cached environment so a
        /// fresh one is created. <paramref name="onBusy"/> runs once, on the first busy result.
        /// Any other error, or still busy after 15s, is thrown to the caller.
        /// </summary>
        public static async Task EnsureCoreWebView2WithRetryAsync(WebView2 webView, Action onBusy = null)
        {
            var deadline = DateTime.UtcNow + BusyRetryWindow;
            for (int attempt = 1; ; attempt++)
            {
                var envTask = GetEnvironmentAsync();
                try
                {
                    await webView.EnsureCoreWebView2Async(await envTask);
                    if (attempt > 1)
                        LogUtility.LogInfo($"WebView2Warmup: browser initialized on attempt {attempt} after ERROR_BUSY.");
                    return;
                }
                catch (Exception ex)
                {
                    Forget(envTask);
                    if (!(ex is COMException com && com.HResult == ErrorBusy) || DateTime.UtcNow >= deadline)
                        throw;

                    LogUtility.LogWarn($"WebView2Warmup: browser profile folder in use (0x800700AA), attempt {attempt}; retrying.");
                    if (attempt == 1)
                        onBusy?.Invoke();
                    await Task.Delay(1000);
                }
            }
        }

        // Drops a failed environment so the next GetEnvironmentAsync creates a new one instead of
        // returning the same failed task for the rest of the Excel session.
        private static void Forget(Task<CoreWebView2Environment> envTask)
        {
            lock (_lock)
            {
                if (ReferenceEquals(_environmentTask, envTask))
                    _environmentTask = null;
            }
        }

        private static async Task<CoreWebView2Environment> CreateEnvironmentAsync()
        {
            try
            {
                string logDir = AppPaths.LoginBrowserLogsPath;
                var di = new DirectoryInfo(logDir);
                if (!di.Exists)
                    di.Create();

                var envOptions = new CoreWebView2EnvironmentOptions
                {
                    AllowSingleSignOnUsingOSPrimaryAccount = true
                };

                return await CoreWebView2Environment.CreateAsync(
                    browserExecutableFolder: null,
                    userDataFolder: di.FullName,
                    options: envOptions);
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex, "WebView2Warmup.CreateEnvironmentAsync");
                throw;
            }
        }
    }
}
