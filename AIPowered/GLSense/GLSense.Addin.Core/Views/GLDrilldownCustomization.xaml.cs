// GLDrilldownCustomization.xaml.cs in GLSense.Addin.Core
// Port of GLSense\Views\GLDrilldownCustomization.xaml.cs (FinalWorkingCode) - a
// WebView2-embedded window that navigates to
// "{LoginUrl}/web/public/cube/{cubeId}/drilldown/launcher?jwtParam={token}" and shows a
// busy overlay (AppOverlay) until the page finishes loading (NavigationCompleted).
//
// Adjustments made when porting into this project's architecture (mirrors GLLogin.xaml.cs
// - the only other WebView2 window in this project - and BaseWindow's established
// conventions; see those files' own header comments for the general rules referenced
// below):
//   - Base class DpiAwareWindow -> BaseWindow. MaxWidthCap/MaxHeightCap are still reset to
//     null in the constructor (same as the original) because this window intentionally
//     wants a much larger footprint (1250-1500px wide) than BaseWindow's default
//     work-area-clamped MaxWidthCap (1400d) would allow, and needs the ability to grow to
//     900px+ tall on ultrawide monitors - resetting both caps to null restores the
//     original DpiAwareWindow behavior of only being bound by SystemParameters.WorkArea.
//   - EnhancedDragDropHelper.EnableWindowDrag(this) / manual Header_MouseLeftButtonDown ->
//     the dedicated TitleBar_MouseLeftButtonDown handler (copied verbatim from
//     GLLogin.xaml.cs's own handler), wired to the title-bar Grid's MouseLeftButtonDown in
//     the XAML - matches every other window in this project.
//   - LogUtility.* (static) -> ServiceLocator.Logger?.* (Infrastructure.ServiceLocator),
//     same nullable-conditional call pattern GLLogin.xaml.cs uses throughout.
//   - WebView2 user-data folder: old code used AppPaths.LoginBrowserLogsPath (NOT
//     AppPaths.DrilldownBrowserLogsPath, even though that member already existed in the
//     old project's AppPaths.cs) - i.e. it deliberately (or by copy/paste from GLLogin)
//     reused the same profile folder as the Login window. This project's IPathProvider
//     exposes both an analogous LoginBrowserPath AND a dedicated DrilldownBrowserPath
//     (GLSense.Shared.PathProvider: "BrowserLogs\Login" vs "BrowserLogs\Drilldown").
//     Per the porting instructions for this file, the exact old AppPaths member used here
//     (LoginBrowserLogsPath) is preserved 1:1 as ServiceLocator.Paths.LoginBrowserPath -
//     i.e. the SAME property GLLogin.xaml.cs's own WebView_Loaded uses for its user data
//     folder - rather than switching to the seemingly-more-correct DrilldownBrowserPath,
//     since that would be a behavior change (a different on-disk profile / cookie jar),
//     not just a namespace/plumbing re-point. If GLLogin and this window are ever open at
//     the same time and that turns out to be a problem in practice, switch this one line to
//     ServiceLocator.Paths.DrilldownBrowserPath - that property exists precisely for this.
//   - AppOverlay control: already exists project-wide (same AppOverlayControl used by
//     GLLogin/GLCubeDetails) - declared in the XAML the same way, no new control needed.
//   - CancellationHelper: used exactly as the original did (a short-lived `using var`
//     local inside WebView_Loaded, not a field - this window has only one cancellable
//     operation and no navigation retries like GLLogin's).
using GLSense.Addin.Core.Helpers;
using GLSense.Addin.Core.Infrastructure;
using Microsoft.Web.WebView2.Core;
using System;
using System.IO;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Input;
using System.Windows.Threading;

#nullable enable
namespace GLSense.Addin.Core.Views
{
    /// <summary>
    /// Interaction logic for GLDrilldownCustomization.xaml
    /// </summary>
    public partial class GLDrilldownCustomization : BaseWindow
    {
        private Task? _webViewInitTask;

        public GLDrilldownCustomization()
        {
            InitializeComponent();
            ServiceLocator.Logger?.LogDebug("GLDrilldownCustomization constructor invoked");

            // This window wants a much larger footprint than BaseWindow's default
            // work-area-clamped caps allow (see header comment above).
            MaxWidthCap = null;
            MaxHeightCap = null;

            webView.Loaded += WebView_Loaded;

            webView.NavigationCompleted += WebView_NavigationCompleted;
        }

        // ---------- Title bar (drag / close) ----------

        private void TitleBar_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            try
            {
                if (e.LeftButton == MouseButtonState.Pressed)
                {
                    this.DragMove();
                }
            }
            catch (Exception ex)
            {
                ServiceLocator.Logger?.LogException(ex, "TitleBar_MouseLeftButtonDown error");
            }
        }

        private void BtnClose_Click(object sender, RoutedEventArgs e)
        {
            ServiceLocator.Logger?.LogDebug("GLDrilldownCustomization.BtnClose_Click invoked - closing window");
            Close();
        }

        // ---------- WebView2 lifecycle ----------

        private async void WebView_Loaded(object sender, RoutedEventArgs e)
        {
            ServiceLocator.Logger?.LogDebug("GLDrilldownCustomization.WebView_Loaded invoked");
            webView.Loaded -= WebView_Loaded; // prevent double init

            try
            {
                // 1) Ensure a writable user data folder (logs/profile)
                string logDir = ServiceLocator.Paths.LoginBrowserPath;
                DirectoryInfo di = new(logDir);
                if (!di.Exists)
                    di.Create();

                string webViewLogsPath = di.FullName;

                // 2) Create environment options FIRST
                var envOptions = new CoreWebView2EnvironmentOptions
                {
                    // Enable SSO if your scenario needs it
                    AllowSingleSignOnUsingOSPrimaryAccount = true
                };

                // 3) Create the environment with options + user data folder
                var env = await CoreWebView2Environment.CreateAsync(
                    browserExecutableFolder: null,
                    userDataFolder: webViewLogsPath,
                    options: envOptions);

                // 4) Initialize WebView2 with that environment
                _webViewInitTask = webView.EnsureCoreWebView2Async(env);
                await _webViewInitTask;

                // 5) Hook device permission handler and diagnostics after CoreWebView2 is ready
                webView.CoreWebView2.PermissionRequested += CoreWebView2_PermissionRequested;
                webView.CoreWebView2.ProcessFailed += CoreWebView2_ProcessFailed;

                // Optional: turn on DevTools during development
                webView.CoreWebView2.Settings.AreDevToolsEnabled = true;

                // 6) Diagnostics: log WebView2 runtime and SSO setting
                var version = webView.CoreWebView2.Environment.BrowserVersionString;
                ServiceLocator.Logger?.LogDebug($"WebView2 BrowserVersion={version}");
                ServiceLocator.Logger?.LogDebug($"AllowSingleSignOnUsingOSPrimaryAccount={envOptions.AllowSingleSignOnUsingOSPrimaryAccount}");

                if (AppState.Instance.SelectedCube != null)
                {
                    string drilldownUrl = $"{AppState.Instance.LoginUrl.TrimEnd('/')}/web/public/cube/{AppState.Instance.SelectedCube.CubeId}/drilldown/launcher?jwtParam=";
                    string finalUrl = $"{drilldownUrl}{AppState.Instance.LoginToken}";

                    ServiceLocator.Logger?.LogDebug($"GLDrilldownCustomization.WebView_Loaded: navigating to {drilldownUrl}<token redacted>");
                    using var cancellationHelper = new CancellationHelper();
                    webView.CoreWebView2.Navigate(finalUrl);
                    webView.Visibility = Visibility.Hidden;
                    await ShowBusyOverlayAsync(cancellationHelper, "Loading Drilldown Customization");
                }
                else
                {
                    ServiceLocator.Logger?.LogWarn("GLDrilldownCustomization.WebView_Loaded: no cube selected, skipping navigation");
                }
            }
            catch (Exception ex)
            {
                ServiceLocator.Logger?.LogException(ex, "WebView2 initialization failed in GLDrilldownCustomization");
            }
        }

        private async Task ShowBusyOverlayAsync(CancellationHelper helper, string message)
        {
            await Dispatcher.InvokeAsync(() =>
            {
                AppOverlayControl.ShowBusyasyn(
                    message: message + " (Click cancel to stop)",
                    cancelAction: async () =>
                    {
                        if (!helper.IsCancellationRequested)
                        {
                            helper.Cancel();
                            ServiceLocator.Logger?.LogWarn($"Operation cancelled by user: {message}");
                        }
                        await Task.Delay(80); // small feedback delay
                                              // HideBusyAsync() is already called in the overlay's handler
                    }
                );
                webView.Visibility = Visibility.Hidden; // ensure overlay is visible above WebView
            }, DispatcherPriority.Background);
        }

        private async void WebView_NavigationCompleted(object sender, CoreWebView2NavigationCompletedEventArgs e)
        {
            ServiceLocator.Logger?.LogDebug($"GLDrilldownCustomization.WebView_NavigationCompleted invoked: IsSuccess={e.IsSuccess}");
            if (!e.IsSuccess)
            {
                ServiceLocator.Logger?.LogWarn($"Navigation failed: {e.WebErrorStatus}");
                await AppOverlayControl.HideBusyAsync();
                webView.Visibility = Visibility.Visible;
                return;
            }

            try
            {
                // Wait until document.readyState == complete
                string script = "document.readyState";
                string result = await webView.CoreWebView2.ExecuteScriptAsync(script);

                // result comes with quotes like "\"complete\""
                if (result.Contains("complete"))
                {
                    ServiceLocator.Logger?.LogDebug("GLDrilldownCustomization.WebView_NavigationCompleted: document ready, hiding busy overlay");
                    // Small buffer for SPA rendering
                    await Task.Delay(500);
                    await AppOverlayControl.HideBusyAsync();
                    webView.Visibility = Visibility.Visible;
                }
            }
            catch (Exception ex)
            {
                ServiceLocator.Logger?.LogException(ex, "Error checking document.readyState");
                await AppOverlayControl.HideBusyAsync();
                webView.Visibility = Visibility.Visible;
            }
        }

        private void CoreWebView2_PermissionRequested(object? sender, CoreWebView2PermissionRequestedEventArgs e)
        {
            try
            {
                ServiceLocator.Logger?.LogDebug($"Permission requested: Kind={e.PermissionKind}, Uri={e.Uri}");

                // Decide what to allow
                switch (e.PermissionKind)
                {
                    case CoreWebView2PermissionKind.Microphone:
                    case CoreWebView2PermissionKind.Camera:
                    case CoreWebView2PermissionKind.Geolocation:
                    case CoreWebView2PermissionKind.MidiSystemExclusiveMessages:
                    case CoreWebView2PermissionKind.ClipboardRead:
                        {
                            e.State = CoreWebView2PermissionState.Allow;
                            e.Handled = true;
                            ServiceLocator.Logger?.LogDebug($"Permission allowed: {e.PermissionKind}");
                            break;
                        }

                    default:
                        e.State = CoreWebView2PermissionState.Deny;
                        e.Handled = true;
                        ServiceLocator.Logger?.LogWarn($"Permission denied: {e.PermissionKind}");
                        break;
                }
            }
            catch (Exception ex)
            {
                ServiceLocator.Logger?.LogException(ex, "PermissionRequested handler error");
                // Fail closed on error
                e.State = CoreWebView2PermissionState.Deny;
                e.Handled = true;
            }
        }

        private void CoreWebView2_ProcessFailed(object? sender, CoreWebView2ProcessFailedEventArgs e)
        {
            ServiceLocator.Logger?.LogWarn($"WebView2 process failed. Kind={e.ProcessFailedKind}");
        }
    }
}
#nullable disable
