using GLSense.Helpers;
using GLSense.Utilities;
using Microsoft.Web.WebView2.Core;
using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using System.Windows;

namespace GLSense.Views
{
#nullable enable
    /// <summary>
    /// Managed host for a WebView2 popup opened via CoreWebView2.NewWindowRequested (e.g. an
    /// identity provider's MFA/step-up-auth popup during an SSO/SAML/OIDC login flow) - see
    /// WebView2NavigationResilience and docs/superpowers/specs/2026-08-10-webview2-navigation-resilience-design.md.
    /// Without this, that popup would open as an unmanaged WebView2 window outside this app's
    /// own lifecycle/disposal handling.
    /// </summary>
    public partial class WebView2PopupWindow : DpiAwareWindow
    {
        private WebView2NavigationResilience? _resilience;

        public CoreWebView2? CoreWebView2 => webView.CoreWebView2;

        public WebView2PopupWindow()
        {
            InitializeComponent();
            EnhancedDragDropHelper.EnableWindowDrag(this);
            Closed += WebView2PopupWindow_Closed;
        }

        public async Task InitializeAsync(CoreWebView2Environment environment, Func<IReadOnlyCollection<string>> trustedHostsProvider)
        {
            await webView.EnsureCoreWebView2Async(environment);

            webView.CoreWebView2.PermissionRequested += CoreWebView2_PermissionRequested;
            webView.CoreWebView2.WindowCloseRequested += CoreWebView2_WindowCloseRequested;

            _resilience = new WebView2NavigationResilience("WebView2PopupWindow", this);
            _resilience.Attach(webView.CoreWebView2, trustedHostsProvider);
        }

        private void CoreWebView2_WindowCloseRequested(object? sender, object e)
        {
            // Fires when the popup's own JS calls window.close() (e.g. once the MFA/step-up
            // flow finishes) - close the hosting window the same way a user clicking the
            // close button would.
            Dispatcher.InvokeAsync(Close);
        }

        private void CoreWebView2_PermissionRequested(object? sender, CoreWebView2PermissionRequestedEventArgs e)
        {
            try
            {
                LogUtility.LogDebug($"WebView2PopupWindow permission requested: Kind={e.PermissionKind}, Uri={e.Uri}");

                switch (e.PermissionKind)
                {
                    case CoreWebView2PermissionKind.Microphone:
                    case CoreWebView2PermissionKind.Camera:
                    case CoreWebView2PermissionKind.Geolocation:
                    case CoreWebView2PermissionKind.MidiSystemExclusiveMessages:
                    case CoreWebView2PermissionKind.ClipboardRead:
                        e.State = CoreWebView2PermissionState.Allow;
                        e.Handled = true;
                        break;

                    default:
                        e.State = CoreWebView2PermissionState.Deny;
                        e.Handled = true;
                        LogUtility.LogWarn($"WebView2PopupWindow permission denied: {e.PermissionKind}");
                        break;
                }
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex, "WebView2PopupWindow.CoreWebView2_PermissionRequested");
                e.State = CoreWebView2PermissionState.Deny;
                e.Handled = true;
            }
        }

        // Mirrors GLLogin.GLLogin_Closed / GLDrilldownCustomization.GLDrilldownCustomization_Closed -
        // dispose the WebView2 control so this popup's Chromium process tree doesn't get orphaned.
        private void WebView2PopupWindow_Closed(object? sender, EventArgs e)
        {
            try
            {
                if (webView.CoreWebView2 != null)
                {
                    _resilience?.Detach(webView.CoreWebView2);
                    webView.CoreWebView2.PermissionRequested -= CoreWebView2_PermissionRequested;
                    webView.CoreWebView2.WindowCloseRequested -= CoreWebView2_WindowCloseRequested;
                }

                webView.Dispose();
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex, "WebView2PopupWindow.WebView2PopupWindow_Closed: WebView2 dispose failed");
            }

            // Owner reactivation on close is handled generically by
            // DpiAwareWindow.RestoreOwnerFocusOnClosed for every window in this app.
        }

        private void BtnClose_Click(object sender, RoutedEventArgs e)
        {
            Close();
        }
    }
#nullable disable
}
