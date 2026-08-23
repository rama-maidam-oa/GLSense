// WebView2PopupWindow.xaml.cs in GLSense.Addin.Core
// Port of GLSense\Views\WebView2PopupWindow.xaml.cs (FinalWorkingCode) - managed host for a
// WebView2 popup opened via CoreWebView2.NewWindowRequested (e.g. an identity provider's
// MFA/step-up-auth popup during an SSO/SAML/OIDC login flow) - see
// WebView2NavigationResilience.OnNewWindowRequested. Without this, that popup would open as
// an unmanaged WebView2 window outside this app's own lifecycle/disposal handling.
//
// Adjustments made when porting into this project's architecture (mirrors GLLogin.xaml.cs/
// GLDrilldownCustomization.xaml.cs conventions - see those files' own header comments for
// the general rules referenced below):
//   - Base class DpiAwareWindow -> DpiAwareWindow. Owner is still set directly via the WPF
//     Window.Owner property by WebView2NavigationResilience.OnNewWindowRequested (not
//     SetExcelOwner/ModalToExcel) - DpiAwareWindow.OnSourceInitialized only sets the native
//     Excel owner when WindowInteropHelper(this).Owner is still zero, so setting Owner to
//     another window (e.g. GLLogin) before Show() - exactly as OnNewWindowRequested does -
//     is left alone. See DpiAwareWindow.OnClosed for how this Owner is reactivated on close.
//   - EnhancedDragDropHelper.EnableWindowDrag(this) -> the dedicated
//     TitleBar_MouseLeftButtonDown handler (copied verbatim from GLLogin.xaml.cs's own
//     handler), matching every other window in this project.
//   - LogUtility.* (static) -> ServiceLocator.Logger?.*.
using GLSense.Addin.Core.Infrastructure;
using GLSense.Addin.Core.Utilities;
using Microsoft.Web.WebView2.Core;
using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Input;

namespace GLSense.Addin.Core.Views
{
#nullable enable
    public partial class WebView2PopupWindow : DpiAwareWindow
    {
        private WebView2NavigationResilience? _resilience;

        public CoreWebView2? CoreWebView2 => webView.CoreWebView2;

        public WebView2PopupWindow()
        {
            InitializeComponent();
            Closed += WebView2PopupWindow_Closed;
        }

        public async Task InitializeAsync(CoreWebView2Environment environment, Func<IReadOnlyCollection<string>> trustedHostsProvider)
        {
            await webView.EnsureCoreWebView2Async(environment);

            webView.CoreWebView2.PermissionRequested += CoreWebView2_PermissionRequested;
            webView.CoreWebView2.WindowCloseRequested += CoreWebView2_WindowCloseRequested;

            _resilience = new WebView2NavigationResilience(nameof(WebView2PopupWindow), this);
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
                ServiceLocator.Logger?.LogDebug($"WebView2PopupWindow permission requested: Kind={e.PermissionKind}, Uri={e.Uri}");

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
                        ServiceLocator.Logger?.LogWarn($"WebView2PopupWindow permission denied: {e.PermissionKind}");
                        break;
                }
            }
            catch (Exception ex)
            {
                ServiceLocator.Logger?.LogException(ex, "WebView2PopupWindow.CoreWebView2_PermissionRequested");
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
                ServiceLocator.Logger?.LogException(ex, "WebView2PopupWindow.WebView2PopupWindow_Closed: WebView2 dispose failed");
            }

            // Owner reactivation on close is handled generically by DpiAwareWindow.OnClosed for
            // every window in this app.
        }

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
            Close();
        }
    }
#nullable disable
}
