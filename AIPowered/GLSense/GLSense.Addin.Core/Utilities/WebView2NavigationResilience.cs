// WebView2NavigationResilience.cs in GLSense.Addin.Core
// Port of GLSense\Utilities\WebView2NavigationResilience.cs (FinalWorkingCode).
//
// Ported: ServerCertificateErrorDetected soft-fail (bypass cert errors only for the
// customer's own trusted/configured host), ProcessFailed recovery (reload on a
// recoverable renderer crash), NavigateWithRetryAsync (retry-once navigation helper),
// and NewWindowRequested/popup-hosting via Views.WebView2PopupWindow (added once that
// class existed in this project - see its own header comment). OnNewWindowRequested
// shows the popup BEFORE awaiting InitializeAsync (EnsureCoreWebView2Async), matching
// FinalWorkingCode's fix for the deadlock that reordering avoids: the WPF WebView2
// control's native handle is only realized once the window's visual tree actually
// loads, which requires the window to already be shown - awaiting
// EnsureCoreWebView2Async first hangs forever waiting on a precondition nothing will
// ever satisfy, and holds the NewWindowRequested deferral open indefinitely, which can
// make the *parent* window's renderer look unresponsive too.
//
// Changes: LogUtility.* (static) -> ServiceLocator.Logger?.*.
using GLSense.Addin.Core.Infrastructure;
using GLSense.Addin.Core.Views;
using Microsoft.Web.WebView2.Core;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;

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
        private readonly Window? _ownerWindow;
        private Func<IReadOnlyCollection<string>>? _trustedHostsProvider;
        private CoreWebView2? _core;

        public WebView2NavigationResilience(string windowName, Window? ownerWindow = null)
        {
            _windowName = windowName;
            _ownerWindow = ownerWindow;
        }

        /// <summary>
        /// Hooks ServerCertificateErrorDetected, NewWindowRequested, and wraps ProcessFailed
        /// with a recovery step. Does not touch any pre-existing NavigationCompleted handler -
        /// those keep reacting to every completed navigation exactly as before.
        /// </summary>
        public void Attach(CoreWebView2 core, Func<IReadOnlyCollection<string>> trustedHostsProvider)
        {
            _core = core;
            _trustedHostsProvider = trustedHostsProvider;

            core.ServerCertificateErrorDetected += OnServerCertificateErrorDetected;
            core.NewWindowRequested += OnNewWindowRequested;
            core.ProcessFailed += OnProcessFailed;
        }

        public void Detach(CoreWebView2? core)
        {
            if (core == null)
                return;

            core.ServerCertificateErrorDetected -= OnServerCertificateErrorDetected;
            core.NewWindowRequested -= OnNewWindowRequested;
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

        private async void OnNewWindowRequested(object? sender, CoreWebView2NewWindowRequestedEventArgs e)
        {
            var deferral = e.GetDeferral();
            WebView2PopupWindow? popup = null;
            try
            {
                popup = new WebView2PopupWindow();
                if (_ownerWindow != null)
                    popup.Owner = _ownerWindow;

                // Show BEFORE EnsureCoreWebView2Async (inside InitializeAsync) - see this
                // class's own header comment for the full reasoning. The popup briefly
                // shows blank chrome until its content finishes initializing - expected
                // and harmless.
                popup.Show();

                await popup.InitializeAsync(_core!.Environment, _trustedHostsProvider!);

                e.NewWindow = popup.CoreWebView2;
                e.Handled = true;

                ServiceLocator.Logger?.LogDebug($"[{_windowName}] Hosted popup window for {e.Uri}");
            }
            catch (Exception ex)
            {
                ServiceLocator.Logger?.LogException(ex, $"[{_windowName}] Failed to host popup window for {e.Uri}");
                // Fall back to WebView2's own default (unmanaged) popup rather than leaving
                // the new-window navigation stuck - only stop if it truly has to. Close our
                // own (already-shown, now-broken) popup first so it doesn't linger alongside
                // WebView2's default one.
                try { popup?.Close(); } catch { /* best-effort cleanup - can be ignored as expected */ }
                e.Handled = false;
            }
            finally
            {
                deferral.Complete();
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
