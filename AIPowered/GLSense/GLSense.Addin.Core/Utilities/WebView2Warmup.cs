using Microsoft.Web.WebView2.Core;
using GLSense.Addin.Core.Infrastructure;
using System;
using System.IO;
using System.Threading.Tasks;

namespace GLSense.Addin.Core.Utilities
{
    // CoreWebView2Environment.CreateAsync spins up a real Chromium process tree (browser +
    // GPU + network service + renderer + crashpad handler) - several seconds cold. GLLogin
    // and GLDrilldownCustomization both used to call CreateAsync independently against the
    // exact same user data folder/options, so whichever window opened first paid that cost
    // with zero visual feedback. This kicks the same CreateAsync call off once, in the
    // background, at AppDomain init, and hands out the one shared environment to both
    // windows. Ported from FinalWorkingCode's Utilities\WebView2Warmup.cs.
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
                string logDir = ServiceLocator.Paths.LoginBrowserPath;
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
                ServiceLocator.Logger?.LogException(ex, "WebView2Warmup.CreateEnvironmentAsync");
                throw;
            }
        }
    }
}
