// WebView2NavigationResilience.cs in GLSense.Addin.Core
// Trimmed port of GLSense\Utilities\WebView2NavigationResilience.cs (FinalWorkingCode).
//
// Ported: ServerCertificateErrorDetected soft-fail (bypass cert errors only for the
// customer's own trusted/configured host), ProcessFailed recovery (reload on a
// recoverable renderer crash), and NavigateWithRetryAsync (retry-once navigation
// helper).
//
// NOT ported: NewWindowRequested/popup-hosting. FinalWorkingCode's version also hosts
// WebView2 popups (via WebView2PopupWindow) for identity-provider MFA/step-up-auth
// flows that call window.open() - this app has no such popup-hosting mechanism yet,
// and building it is new infrastructure, not a port. Until that exists, this class
// leaves NewWindowRequested unhooked, so WebView2's own default (unmanaged) popup
// handling applies exactly as it did before this port.
//
// Changes: LogUtility.* (static) -> ServiceLocator.Logger?.*. Constructor drops the
// ownerWindow parameter (only needed for popup hosting).
using GLSense.Addin.Core.Infrastructure;
using Microsoft.Web.WebView2.Core;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;

namespace GLSense.Addin.Core.Utilities
{
#nullable enable
    /// <summary>
    /// Makes WebView2 navigation in a login/SSO-style window tolerate non-fatal issues
    /// (a certificate error against the customer's own configured server, a transient
    /// connection failure) instead of leaving the page blank, while still respecting
    /// genuinely fatal failures. One instance per hosting window.
    /// </summary>
    public class WebView2NavigationResilience
    {
        private static readonly TimeSpan RetryDelay = TimeSpan.FromSeconds(1);

        private readonly string _windowName;
        private Func<IReadOnlyCollection<string>>? _trustedHostsProvider;
        private CoreWebView2? _core;

        public WebView2NavigationResilience(string windowName)
        {
            _windowName = windowName;
        }

        /// <summary>
        /// Hooks ServerCertificateErrorDetected and wraps ProcessFailed with a recovery
        /// step. Does not touch any pre-existing NavigationCompleted handler - those keep
        /// reacting to every completed navigation exactly as before.
        /// </summary>
        public void Attach(CoreWebView2 core, Func<IReadOnlyCollection<string>> trustedHostsProvider)
        {
            _core = core;
            _trustedHostsProvider = trustedHostsProvider;

            core.ServerCertificateErrorDetected += OnServerCertificateErrorDetected;
            core.ProcessFailed += OnProcessFailed;
        }

        public void Detach(CoreWebView2? core)
        {
            if (core == null)
                return;

            core.ServerCertificateErrorDetected -= OnServerCertificateErrorDetected;
            core.ProcessFailed -= OnProcessFailed;
        }

        private void OnServerCertificateErrorDetected(object? sender, CoreWebView2ServerCertificateErrorDetectedEventArgs e)
        {
            try
            {
                string host = TryGetHost(e.RequestUri);
                var trustedHosts = _trustedHostsProvider?.Invoke() ?? Array.Empty<string>();

                if (!string.IsNullOrEmpty(host) && trustedHosts.Contains(host, StringComparer.OrdinalIgnoreCase))
                {
                    e.Action = CoreWebView2ServerCertificateErrorAction.AlwaysAllow;
                    ServiceLocator.Logger?.LogWarn($"[{_windowName}] Bypassed certificate error ({e.ErrorStatus}) for trusted host '{host}' at {e.RequestUri}.");
                }
                else
                {
                    // Leave WebView2's default behavior (its own blocking interstitial) in
                    // place for anything outside the customer's own configured server - e.g.
                    // an identity-provider hop mid SSO/SAML/OIDC redirect. A cert error there
                    // is unusual enough that it should not be silently bypassed.
                    e.Action = CoreWebView2ServerCertificateErrorAction.Default;
                    ServiceLocator.Logger?.LogWarn($"[{_windowName}] Certificate error ({e.ErrorStatus}) for untrusted host '{host}' at {e.RequestUri} - not bypassed.");
                }
            }
            catch (Exception ex)
            {
                ServiceLocator.Logger?.LogException(ex, $"[{_windowName}] ServerCertificateErrorDetected handler error");
                e.Action = CoreWebView2ServerCertificateErrorAction.Default;
            }
        }

        private void OnProcessFailed(object? sender, CoreWebView2ProcessFailedEventArgs e)
        {
            ServiceLocator.Logger?.LogWarn($"[{_windowName}] WebView2 process failed. Kind={e.ProcessFailedKind}");

            switch (e.ProcessFailedKind)
            {
                case CoreWebView2ProcessFailedKind.RenderProcessExited:
                case CoreWebView2ProcessFailedKind.FrameRenderProcessExited:
                    try
                    {
                        _core?.Reload();
                        ServiceLocator.Logger?.LogWarn($"[{_windowName}] Reloaded after recoverable process failure ({e.ProcessFailedKind}).");
                    }
                    catch (Exception ex)
                    {
                        ServiceLocator.Logger?.LogException(ex, $"[{_windowName}] Reload after process failure failed");
                    }
                    break;

                default:
                    ServiceLocator.Logger?.LogWarn($"[{_windowName}] Unrecoverable WebView2 process failure ({e.ProcessFailedKind}) - not reloading.");
                    break;
            }
        }

        /// <summary>
        /// Navigates and awaits completion, retrying exactly once (after a short delay) if the
        /// first attempt failed. Returns the final CoreWebView2NavigationCompletedEventArgs
        /// (success or failure) to the caller, which decides what a final failure means for it
        /// (e.g. showing an overlay error). Cancellation via <paramref name="ct"/> propagates as
        /// a TaskCanceledException, matching this codebase's existing navigation call sites.
        /// </summary>
        public async Task<CoreWebView2NavigationCompletedEventArgs> NavigateWithRetryAsync(
            CoreWebView2 core, string url, CancellationToken ct)
        {
            var firstAttempt = await NavigateOnceAsync(core, url, ct);
            if (firstAttempt.IsSuccess)
                return firstAttempt;

            ServiceLocator.Logger?.LogWarn($"[{_windowName}] Navigation to {url} failed ({firstAttempt.WebErrorStatus}); retrying once.");

            await Task.Delay(RetryDelay, ct);

            return await NavigateOnceAsync(core, url, ct);
        }

        private static Task<CoreWebView2NavigationCompletedEventArgs> NavigateOnceAsync(
            CoreWebView2 core, string url, CancellationToken ct)
        {
            var tcs = new TaskCompletionSource<CoreWebView2NavigationCompletedEventArgs>();
            CancellationTokenRegistration ctRegistration = default;

            void Handler(object? sender, CoreWebView2NavigationCompletedEventArgs e)
            {
                core.NavigationCompleted -= Handler;
                ctRegistration.Dispose();
                tcs.TrySetResult(e);
            }

            core.NavigationCompleted += Handler;
            ctRegistration = ct.Register(() =>
            {
                core.NavigationCompleted -= Handler;
                tcs.TrySetCanceled();
            });

            core.Navigate(url);
            return tcs.Task;
        }

        private static string TryGetHost(string? uri)
        {
            try
            {
                return string.IsNullOrEmpty(uri) ? string.Empty : new Uri(uri).Host;
            }
            catch
            {
                // Malformed URI - can be ignored as expected; callers treat an empty host as no match.
                return string.Empty;
            }
        }
    }
#nullable disable
}
