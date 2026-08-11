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
using GLSense.Addin.Core.Common;
using GLSense.Addin.Core.Helpers;
using GLSense.Addin.Core.Infrastructure;
using GLSense.Addin.Core.Utilities;
using Microsoft.Web.WebView2.Core;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text.Json;
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
        private WebView2NavigationResilience? _resilience;

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

            Closed += GLDrilldownCustomization_Closed;
        }

        // See GLLogin.GLLogin_Closed for the full reasoning: WebView2 spins up a real
        // Chromium browser-process tree per environment, and nothing else in this app's
        // lifecycle ever tears it down without this - confirmed via Task Manager (see
        // FinalWorkingCode's identical fix) showing dozens of orphaned msedgewebview2.exe
        // processes accumulating across sessions.
        private void GLDrilldownCustomization_Closed(object? sender, EventArgs e)
        {
            try
            {
                ServiceLocator.Logger?.LogDebug("GLDrilldownCustomization.GLDrilldownCustomization_Closed: disposing WebView2 control");

                webView.NavigationCompleted -= WebView_NavigationCompleted;

                if (webView.CoreWebView2 != null)
                {
                    _resilience?.Detach(webView.CoreWebView2);
                    webView.CoreWebView2.PermissionRequested -= CoreWebView2_PermissionRequested;
                }

                webView.Dispose();
            }
            catch (Exception ex)
            {
                ServiceLocator.Logger?.LogException(ex, "GLDrilldownCustomization.GLDrilldownCustomization_Closed: WebView2 dispose failed");
            }
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

                // Cert-error bypass (scoped to this window's own server) and retry-once
                // navigation. Ported from FinalWorkingCode's WebView2NavigationResilience
                // (popup-hosting piece excluded - see that class's header comment).
                _resilience = new WebView2NavigationResilience(nameof(GLDrilldownCustomization), this);
                _resilience.Attach(webView.CoreWebView2, GetTrustedHosts);

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
                    webView.Visibility = Visibility.Hidden;
                    await ShowBusyOverlayAsync(cancellationHelper, "Loading Drilldown Customization");

                    try
                    {
                        var result = await _resilience!.NavigateWithRetryAsync(webView.CoreWebView2, finalUrl, cancellationHelper.GetToken());
                        if (!result.IsSuccess)
                        {
                            ServiceLocator.Logger?.LogWarn($"GLDrilldownCustomization.WebView_Loaded: navigation to drilldown launcher failed after retry ({result.WebErrorStatus}).");
                            await AppOverlayControl.HideBusyAsync();
                            webView.Visibility = Visibility.Visible;
                            await AppOverlayControl.ShowErrorAsync("Unable to load the page. Please try again.");
                        }
                        // On success, the always-on WebView_NavigationCompleted handler below
                        // hides the busy overlay and restores visibility once
                        // document.readyState is complete.
                    }
                    catch (TaskCanceledException)
                    {
                        ServiceLocator.Logger?.LogWarn("GLDrilldownCustomization.WebView_Loaded: navigation was cancelled.");
                        await AppOverlayControl.HideBusyAsync();
                        webView.Visibility = Visibility.Visible;
                    }
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

        // Trusted-host set for WebView2NavigationResilience's certificate-error bypass -
        // this window always talks to the same, already-authenticated server
        // (AppState.Instance.LoginUrl).
        private IReadOnlyCollection<string> GetTrustedHosts()
        {
            try
            {
                var loginUrl = AppState.Instance.LoginUrl;
                if (string.IsNullOrWhiteSpace(loginUrl))
                    return Array.Empty<string>();

                return new[] { new Uri(loginUrl, UriKind.Absolute).Host };
            }
            catch (Exception ex)
            {
                ServiceLocator.Logger?.LogDebug($"GLDrilldownCustomization.GetTrustedHosts: could not resolve host from '{AppState.Instance.LoginUrl}': {ex.Message}");
                return Array.Empty<string>();
            }
        }

        // ---------- Save Locally ----------
        // Ported from GLSense\Views\GLDrilldownCustomization.xaml.cs (FinalWorkingCode)'s
        // BtnSaveLocally_Click/HideBusyAndShow*Async/ValidateTransportResponse. Fetches this
        // cube's drilldown metadata and stores it in a CustomXMLPart
        // (Common\DrilldownMetadataXmlStore.cs) so Drilldowns\DDDatatoWorksheet.cs's
        // ExtractMetadata can use it later when UserConfig.OverwriteDrilldownMetadata is enabled
        // (Views\GLUserConfig.xaml's "Overwrite drilldown metadata with locally saved" checkbox).
        private async void BtnSaveLocally_Click(object sender, RoutedEventArgs e)
        {
            ServiceLocator.Logger?.LogDebug("GLDrilldownCustomization.BtnSaveLocally_Click invoked");

            if (AppState.Instance.SelectedCube == null)
            {
                ServiceLocator.Logger?.LogWarn("GLDrilldownCustomization.BtnSaveLocally_Click: no selected cube, aborting save.");
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

                ServiceLocator.Logger?.LogDebug($"GLDrilldownCustomization.BtnSaveLocally_Click: calling API {apiUrl}");
                response = await ApiHelper.ServerAPI(apiUrl, "Form", "", "GET", cancellationHelper.GetToken())
                    ?? string.Empty;

                ValidateTransportResponse(response);

                var parsed = ApiResponseHelper.Parse<JsonElement>(response, JsonGlobals.Options);

                if (!parsed.IsSuccess)
                {
                    ServiceLocator.Logger?.LogWarn($"GLDrilldownCustomization.BtnSaveLocally_Click: API returned failure - {parsed.ErrorMessage}");
                    await HideBusyAndShowErrorAsync(
                        string.IsNullOrWhiteSpace(parsed.ErrorMessage) ? "Failed to fetch drilldown metadata." : parsed.ErrorMessage);
                    return;
                }

                DrilldownMetadataXmlStore.Save(ServiceLocator.ExcelApp?.ActiveWorkbook, cubeId, response);

                ServiceLocator.Logger?.LogDebug("GLDrilldownCustomization.BtnSaveLocally_Click: drilldown metadata saved locally successfully");
                await HideBusyAndShowSuccessAsync("Customizations saved to this workbook. They'll travel with it whenever it's shared or shipped.");
            }
            catch (OperationCanceledException)
            {
                ServiceLocator.Logger?.LogWarn("GLDrilldownCustomization.BtnSaveLocally_Click: operation cancelled by user.");
                await HideBusyAndShowWarnAsync("Save cancelled.");
            }
            catch (Exception ex)
            {
                ServiceLocator.Logger?.LogException(ex, "GLDrilldownCustomization.BtnSaveLocally_Click");
                ServiceLocator.Logger?.LogRawJson("GLDrilldownCustomization.BtnSaveLocally_Click - error response", response);
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
    }
}
#nullable disable
