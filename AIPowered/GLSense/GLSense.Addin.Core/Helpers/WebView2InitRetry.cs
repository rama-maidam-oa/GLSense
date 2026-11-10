using GLSense.Addin.Core.Infrastructure;
using Microsoft.Web.WebView2.Core;
using Microsoft.Web.WebView2.Wpf;
using System;
using System.Runtime.InteropServices;
using System.Threading.Tasks;

namespace GLSense.Addin.Core.Helpers
{
    /// <summary>
    /// WebView2 initialization with a retry for 0x800700AA (ERROR_BUSY): the browser profile
    /// folder is still held, e.g. by a WebView2 browser process that has not finished exiting.
    /// Seen in a customer log on every login attempt for ~30s until Excel was restarted, while
    /// the login window just closed silently.
    /// </summary>
    internal static class WebView2InitRetry
    {
        private const int ErrorBusy = unchecked((int)0x800700AA);
        private static readonly TimeSpan BusyRetryWindow = TimeSpan.FromSeconds(15);

        public const string InitFailedMessage =
            "Unable to start the browser component. Please wait a moment and try again. " +
            "If this keeps happening, close Excel completely and reopen it.";

        /// <summary>
        /// Creates an environment with <paramref name="createEnvironment"/> and initializes
        /// <paramref name="webView"/> with it, retrying every second for up to 15s while WebView2
        /// reports ERROR_BUSY (a fresh environment each attempt). <paramref name="onBusy"/> runs once,
        /// on the first busy result. Returns the environment used. Any other error, or still busy
        /// after 15s, is thrown. The WPF control re-runs its init when the previous one faulted.
        /// </summary>
        public static async Task<CoreWebView2Environment> EnsureAsync(
            WebView2 webView, Func<Task<CoreWebView2Environment>> createEnvironment, Action? onBusy = null)
        {
            var deadline = DateTime.UtcNow + BusyRetryWindow;
            for (int attempt = 1; ; attempt++)
            {
                try
                {
                    var env = await createEnvironment();
                    await webView.EnsureCoreWebView2Async(env);
                    if (attempt > 1)
                        ServiceLocator.Logger?.LogInfo($"WebView2InitRetry: browser initialized on attempt {attempt} after ERROR_BUSY.");
                    return env;
                }
                catch (COMException ex) when (ex.HResult == ErrorBusy && DateTime.UtcNow < deadline)
                {
                    ServiceLocator.Logger?.LogWarn($"WebView2InitRetry: browser profile folder in use (0x800700AA), attempt {attempt}; retrying.");
                    if (attempt == 1)
                        onBusy?.Invoke();
                    await Task.Delay(1000);
                }
            }
        }
    }
}
