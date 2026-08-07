using GLSense.Common;
using GLSense.Helpers;
using GLSense.Utilities;
using Microsoft.Web.WebView2.Core;
using System;
using System.IO;
using System.Text.Json;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Threading;

#nullable enable
namespace GLSense.Views
{
    /// <summary>
    /// Interaction logic for GLDrilldownCustomization.xaml
    /// </summary>
    public partial class GLDrilldownCustomization : DpiAwareWindow
    {
        private Task? _webViewInitTask;
        public GLDrilldownCustomization()
        {
            LogUtility.LogDebug("GLDrilldownCustomization.ctor invoked");
            InitializeComponent();
            EnhancedDragDropHelper.EnableWindowDrag(this);

            MaxWidthCap = null;
            MaxHeightCap = null;


            webView.Loaded += WebView_Loaded;

            webView.NavigationCompleted += WebView_NavigationCompleted;
        }
        private async void WebView_Loaded(object sender, RoutedEventArgs e)
        {
            LogUtility.LogDebug("GLDrilldownCustomization.WebView_Loaded invoked");
            webView.Loaded -= WebView_Loaded; // prevent double init

            try
            {
                // 1) Ensure a writable user data folder (logs/profile)
                string logDir = AppPaths.LoginBrowserLogsPath;
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
                LogUtility.LogDebug($"WebView2 BrowserVersion={version}");
                LogUtility.LogDebug($"AllowSingleSignOnUsingOSPrimaryAccount={envOptions.AllowSingleSignOnUsingOSPrimaryAccount}");


                if (AppState.Instance.SelectedCube != null)
                {
                    string drilldownUrl = $"{AppState.Instance.LoginUrl.TrimEnd('/')}/web/public/cube/{AppState.Instance.SelectedCube.CubeId}/drilldown/launcher?jwtParam=";
                    string finalUrl = $"{drilldownUrl}{AppState.Instance.LoginToken}";

                    LogUtility.LogDebug($"GLDrilldownCustomization.WebView_Loaded: navigating to drilldown launcher for cubeId={AppState.Instance.SelectedCube.CubeId}");
                    using var cancellationHelper = new CancellationHelper();
                    webView.CoreWebView2.Navigate(finalUrl);
                    webView.Visibility = Visibility.Hidden;
                    await ShowBusyOverlayAsync(cancellationHelper, "Loading Drilldown Customization");

                }
                else
                {
                    LogUtility.LogDebug("GLDrilldownCustomization.WebView_Loaded: no selected cube, skipping navigation");
                }

            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex, "WebView2 initialization failed in GLDrilldownCustomization");
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
                            LogUtility.LogWarn($"Operation cancelled by user: {message}");
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
            LogUtility.LogDebug($"GLDrilldownCustomization.WebView_NavigationCompleted invoked - IsSuccess={e.IsSuccess}");

            if (!e.IsSuccess)
            {
                LogUtility.LogWarn($"Navigation failed: {e.WebErrorStatus}");
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
                    LogUtility.LogDebug("GLDrilldownCustomization.WebView_NavigationCompleted: document ready, showing content");
                    // Small buffer for SPA rendering
                    await Task.Delay(500);
                    await AppOverlayControl.HideBusyAsync();
                    webView.Visibility = Visibility.Visible;
                }
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex, "Error checking document.readyState");
                await AppOverlayControl.HideBusyAsync();
                webView.Visibility = Visibility.Visible;
            }
        }
        private void CoreWebView2_PermissionRequested(object? sender, CoreWebView2PermissionRequestedEventArgs e)
        {
            try
            {
                LogUtility.LogDebug($"Permission requested: Kind={e.PermissionKind}, Uri={e.Uri}");

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
                            LogUtility.LogDebug($"Permission allowed: {e.PermissionKind}");
                            break;
                        }

                    default:
                        e.State = CoreWebView2PermissionState.Deny;
                        e.Handled = true;
                        LogUtility.LogWarn($"Permission denied: {e.PermissionKind}");
                        break;
                }
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex, "PermissionRequested handler error");
                // Fail closed on error
                e.State = CoreWebView2PermissionState.Deny;
                e.Handled = true;
            }
        }
        private void CoreWebView2_ProcessFailed(object? sender, CoreWebView2ProcessFailedEventArgs e)
        {
            LogUtility.LogWarn($"WebView2 process failed. Kind={e.ProcessFailedKind}");
        }
        private void BtnClose_Click(object sender, RoutedEventArgs e)
        {
            LogUtility.LogDebug("GLDrilldownCustomization.BtnClose_Click invoked");
            Close();
        }

        private async void BtnSaveLocally_Click(object sender, RoutedEventArgs e)
        {
            LogUtility.LogDebug("GLDrilldownCustomization.BtnSaveLocally_Click invoked");

            if (AppState.Instance.SelectedCube == null)
            {
                LogUtility.LogWarn("GLDrilldownCustomization.BtnSaveLocally_Click: no selected cube, aborting save.");
                // Hide webView for the same reason HideBusyAndShowWarnAsync does below - see its
                // comment. This path never starts the busy overlay, so webView is still visible
                // going in; without hiding it here too, this toast would show the same
                // WebView2-airspace symptom (only the header/footer dim, the WebView2 content
                // stays fully visible and un-blurred on top of the toast).
                webView.Visibility = Visibility.Hidden;
                await AppOverlayControl.ShowWarningAsync("No cube selected. Please select a cube first.");
                webView.Visibility = Visibility.Visible;
                return;
            }

            long cubeId = AppState.Instance.SelectedCube.CubeId;
            using var cancellationHelper = new CancellationHelper();
            string response = string.Empty;

            try
            {
                await ShowBusyOverlayAsync(cancellationHelper, "Saving customizations to this workbook");

                // drilldownType is left empty (not omitted) to get metadata for ALL drilldown
                // types in one call. Omitting the query parameter entirely returned a 404
                // ("No endpoint GET /reporting/rest/secure/finance/drilldown-metadata.") - the
                // server's routing requires the drilldownType key to be present, even with an
                // empty value, to match this endpoint.
                string apiUrl = $"{AppState.Instance.LoginUrl.TrimEnd('/')}{AppConstants.RestSecure}drilldown-metadata?cubeId={cubeId}&drilldownType=";

                LogUtility.LogDebug($"GLDrilldownCustomization.BtnSaveLocally_Click: calling API {apiUrl}");
                response = await ApiHelper.ServerAPI(apiUrl, "Form", "", "GET", cancellationHelper.GetToken())
                    ?? string.Empty;

                ValidateTransportResponse(response);

                var parsed = ApiResponseHelper.Parse<JsonElement>(response, JsonGlobals.Options);

                if (!parsed.IsSuccess)
                {
                    LogUtility.LogWarn($"GLDrilldownCustomization.BtnSaveLocally_Click: API returned failure - {parsed.ErrorMessage}");
                    await HideBusyAndShowErrorAsync(
                        string.IsNullOrWhiteSpace(parsed.ErrorMessage) ? "Failed to fetch drilldown metadata." : parsed.ErrorMessage);
                    return;
                }

                DrilldownMetadataXmlStore.Save(AppState.Instance.ExcelApp?.ActiveWorkbook, cubeId, response);

                LogUtility.LogDebug("GLDrilldownCustomization.BtnSaveLocally_Click: drilldown metadata saved locally successfully");
                await HideBusyAndShowSuccessAsync("Customizations saved to this workbook. They'll travel with it whenever it's shared or shipped.");
            }
            catch (OperationCanceledException)
            {
                LogUtility.LogWarn("GLDrilldownCustomization.BtnSaveLocally_Click: operation cancelled by user.");
                await HideBusyAndShowWarnAsync("Save cancelled.");
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex, "GLDrilldownCustomization.BtnSaveLocally_Click");
                LogUtility.LogRawJson("GLDrilldownCustomization.BtnSaveLocally_Click - error response", response);
                await HideBusyAndShowErrorAsync($"Failed to save customizations: {ex.Message}");
            }
        }

        // ShowBusyOverlayAsync hides the WebView2 control while the busy spinner is up (see
        // its Dispatcher.InvokeAsync block above). WebView2 hosts its own native child HWND,
        // which WPF's compositor always paints on top of everything in the WPF visual tree -
        // including AppOverlayControl's blur/dim/Panel.ZIndex - regardless of Z-order (the
        // classic WPF "airspace" limitation). So AppOverlayControl.ApplyBlurToSiblings() can
        // only ever dim its WPF-drawn siblings (the header/footer Borders); it can never touch
        // WebView2's own rendered pixels. The only real fix is what ShowBusyOverlayAsync
        // already does: hide webView outright while anything needs to visually cover that
        // area, and only show it again once nothing is covering it anymore.
        //
        // ShowErrorAsync/ShowSuccessAsync/ShowWarningAsync don't return until the toast itself
        // has actually been dismissed (auto-timeout or the user clicking its own close button)
        // - see AppOverlay.ShowToastAsync's TaskCompletionSource. So webView must stay hidden
        // across BOTH the busy spinner AND the toast, and only come back after the toast
        // await below completes - restoring it any earlier (e.g. right after HideBusyAsync,
        // before the toast) would let webView repaint over the content area while the toast is
        // still up, exactly reproducing the same partial-blur symptom.
        private async Task HideBusyAndShowErrorAsync(string errorMsg)
        {
            await AppOverlayControl.HideBusyAsync();
            if (!string.IsNullOrWhiteSpace(errorMsg))
            {
                await AppOverlayControl.ShowErrorAsync(errorMsg);
            }
            webView.Visibility = Visibility.Visible;
        }

        private async Task HideBusyAndShowSuccessAsync(string successMsg)
        {
            await AppOverlayControl.HideBusyAsync();
            if (!string.IsNullOrWhiteSpace(successMsg))
            {
                await AppOverlayControl.ShowSuccessAsync(successMsg);
            }
            webView.Visibility = Visibility.Visible;
        }

        private async Task HideBusyAndShowWarnAsync(string warnMsg)
        {
            await AppOverlayControl.HideBusyAsync();
            if (!string.IsNullOrWhiteSpace(warnMsg))
            {
                await AppOverlayControl.ShowWarningAsync(warnMsg);
            }
            webView.Visibility = Visibility.Visible;
        }

        // Mirrors GLUserConfig.ValidateTransportResponse - checks for empty/HTML/401/error
        // strings that ApiResponseHelper.Parse would otherwise fail to explain clearly.
        private static void ValidateTransportResponse(string response)
        {
            if (string.IsNullOrWhiteSpace(response))
            {
                throw new InvalidOperationException("Empty API response.");
            }

            if (response.IndexOf("(401)", StringComparison.OrdinalIgnoreCase) >= 0)
            {
                throw new UnauthorizedAccessException("Session expired.");
            }

            if (response.StartsWith("Error", StringComparison.OrdinalIgnoreCase) ||
                response.IndexOf("<!DOCTYPE", StringComparison.OrdinalIgnoreCase) >= 0)
            {
                throw new InvalidOperationException(response);
            }
        }

        private void Header_MouseLeftButtonDown(object sender, System.Windows.Input.MouseButtonEventArgs e)
        {
            if (e.ClickCount > 1 || WindowState == WindowState.Maximized)
                return;

            if (e.LeftButton == System.Windows.Input.MouseButtonState.Pressed)
            {
                try
                {
                    DragMove();
                }
                catch (Exception ex)
                {
                    // ignore drag failures (e.g. race with mouse-up), but log for diagnostics.
                    LogUtility.LogWarn($"GLDrilldownCustomization.Header_MouseLeftButtonDown: DragMove failed (ignored): {ex.Message}");
                }
            }
        }
    }
}
#nullable disable

