using Microsoft.Web.WebView2.Core;
using System;
using System.IO;
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
