# WebView2 Navigation Resilience — Design

**Date:** 2026-08-10
**Scope:** `GLSense/Views/GLLogin.xaml.cs`, `GLSense/Views/GLDrilldownCustomization.xaml.cs`

## Problem

Both windows navigate a `Microsoft.Web.WebView2.Wpf.WebView2` control to a customer-configured
server address (`GLLogin`, for SSO/SAML/OIDC login) or an already-authenticated drilldown URL
(`GLDrilldownCustomization`). At some customer sites the page fails to display or never finishes
navigating, due to things like firewall interference, TLS/certificate warnings, or other
WebView2-level navigation/process failures. Today:

- Neither window hooks `ServerCertificateErrorDetected` — a certificate error hits WebView2's
  default behavior (its own blocking interstitial), with no path to continue automatically.
- Neither window hooks `NewWindowRequested` — an identity provider that pops up a separate
  window (e.g. for MFA/step-up auth) would get an unmanaged, untracked WebView2 popup outside
  this app's own lifecycle/disposal handling.
- `NavigationCompleted` failures are logged via `LogUtility.LogWarn` but never retried, and the
  window is simply left in whatever state it was in (often a blank `WebView2` area) — no
  user-facing feedback.
- `ProcessFailed` is logged but never triggers a recovery attempt (e.g. `Reload()`).

## Goals

1. Certificate errors against the customer's own configured server should not block navigation.
2. Genuinely fatal failures (host truly unreachable, cert error against an untrusted/third-party
   host, unrecoverable process failure) must still stop navigation — nothing here should mask a
   real problem.
3. A transient failure gets one automatic retry before being treated as fatal.
4. Identity-provider popups (SSO/SAML/OIDC step-up/MFA windows) are hosted in a managed window,
   not left to WebView2's unmanaged default.
5. Every time navigation ultimately can't proceed, it's recorded as a `LogWarn` — never silently
   swallowed — and the user sees a friendly error instead of a blank page.
6. Both windows behave identically.

## Non-goals

- No changes to `Utilities/StrictCertificateValidator.cs` or the app's REST API TLS validation
  policy — this design only affects WebView2 page rendering, not the app's HTTP API calls.
- No change to WebView2's handling of sub-resource (image/script/XHR) failures — `NavigationCompleted`
  is main-frame-only in the WebView2 API, which is the right granularity here; adding
  `WebResourceError` handling for every sub-resource would be noisy without adding real signal.
- No new automated test project. This repository has none today (confirmed via `GLSense.sln`),
  and every prior fix this session was verified by building + exercising the real app + reading
  the NLog output — this design follows that same convention (see Verification, below).

## Architecture

Two new files, one shared helper class used by both windows:

### `GLSense/Utilities/WebView2NavigationResilience.cs`

One instance per window (not static), created in each window's existing `WebView_Loaded`
right after `CoreWebView2` becomes available.

```csharp
public class WebView2NavigationResilience
{
    public WebView2NavigationResilience(string windowName);

    // Hooks ServerCertificateErrorDetected, NewWindowRequested, and wraps the
    // existing ProcessFailed handling with a recovery step. Does not touch any
    // pre-existing NavigationCompleted handler.
    public void Attach(CoreWebView2 core, Func<IReadOnlyCollection<string>> trustedHostsProvider);
    public void Detach(CoreWebView2 core);

    // Replaces the raw `core.Navigate(url)` + one-shot-handler pattern at each
    // call site. Retries exactly once on failure, then returns the final result.
    public Task<CoreWebView2NavigationCompletedEventArgs> NavigateWithRetryAsync(
        CoreWebView2 core, string url, CancellationToken ct);
}
```

`trustedHostsProvider` is a callback, not a fixed list, because `GLLogin`'s trusted host changes
at runtime (user can switch servers via the dropdown); `GLDrilldownCustomization` returns a
single fixed host (from `AppState.Instance.LoginUrl`).

### `GLSense/Views/WebView2PopupWindow.xaml` / `.xaml.cs`

A minimal borderless `DpiAwareWindow` subclass with a single `WebView2` control, created only
when `NewWindowRequested` fires. Shares the parent's `CoreWebView2Environment` (so the SSO
session/cookies carry over), gets the same cert-bypass wiring via
`WebView2NavigationResilience.Attach` (same trusted hosts as the parent), and disposes itself
the same way `GLLogin`/`GLDrilldownCustomization` already do:

- `CoreWebView2.WindowCloseRequested` (JS `window.close()`) → closes the window.
- `Closed` → unhooks its `CoreWebView2` events and disposes the control.

No retry logic here — it's a short-lived auth popup; if it fails, the user retries from the main
flow.

## Mechanics

### Certificate-error bypass

```csharp
core.ServerCertificateErrorDetected += (s, e) =>
{
    var host = e.RequestUri != null ? new Uri(e.RequestUri).Host : "";
    if (trustedHostsProvider().Contains(host, StringComparer.OrdinalIgnoreCase))
    {
        e.Action = CoreWebView2ServerCertificateErrorAction.AlwaysAllow;
        LogUtility.LogWarn($"[{windowName}] Bypassed cert error ({e.ErrorStatus}) for trusted host '{host}'.");
    }
    else
    {
        e.Action = CoreWebView2ServerCertificateErrorAction.Default; // unchanged from today
        LogUtility.LogWarn($"[{windowName}] Cert error ({e.ErrorStatus}) for untrusted host '{host}' - not bypassed.");
    }
};
```

Trusted hosts:
- `GLLogin`: host portion of the currently-selected server's `Address` (re-read at bypass time,
  so switching the dropdown selection updates it without re-attaching anything).
- `GLDrilldownCustomization`: host portion of `AppState.Instance.LoginUrl` (fixed for the
  window's lifetime).

Hosts hit mid-redirect that are *not* in this set (identity-provider domains, etc.) are
deliberately left on WebView2's default cert-error behavior — the design does not add any new
blocking behavior beyond what exists today for those hosts, only removes it for the customer's
own server.

### Retry-once flow

`NavigateWithRetryAsync(core, url, ct)`:
1. Subscribe a one-shot `NavigationCompleted` handler, call `core.Navigate(url)`, await the
   result (registering `ct` for cancellation).
2. If it succeeded, return the result immediately.
3. If it failed and this was the first attempt: `LogWarn` the failure (URL + `WebErrorStatus`),
   wait ~1 second, and repeat from step 1 once.
4. If it failed on the second attempt, return that (final) failed result to the caller.

Cancellation (e.g. the user clicks Cancel on the busy overlay) aborts immediately with no retry,
matching the existing `CancellationHelper`/`TaskCanceledException` handling already present in
`GLLogin`.

This replaces the manual `Handler`/`TaskCompletionSource` block at the `Navigate()` call sites in
`NavigateToServerAsync` and `NavigateToLoginAsync` (`GLLogin`), and the direct `Navigate()` call
in `WebView_Loaded` (`GLDrilldownCustomization`). The always-attached `NavigationCompleted`
handlers used for login-success detection (`GLLogin`) and document-ready checking
(`GLDrilldownCustomization`) are untouched — they keep firing for every navigation completion as
they do today.

### Popup / new-window handling

```csharp
core.NewWindowRequested += async (s, e) =>
{
    var deferral = e.GetDeferral();
    try
    {
        var popup = new WebView2PopupWindow();
        await popup.InitializeAsync(core.Environment, trustedHostsProvider);
        e.NewWindow = popup.CoreWebView2;
        e.Handled = true;
        popup.Show();
    }
    catch (Exception ex)
    {
        LogUtility.LogException(ex, $"[{windowName}] Failed to host popup window for {e.Uri}");
        e.Handled = false; // fall back to WebView2's own default popup rather than blocking
    }
    finally
    {
        deferral.Complete();
    }
};
```

If our own popup setup fails, we deliberately fall back to letting WebView2 open its default,
unmanaged popup rather than leaving the new-window navigation stuck — consistent with "only stop
if it truly has to."

### Process-failure recovery

Enhances the existing `ProcessFailed` handler (currently log-only in both windows):

- `RenderProcessExited` / `FrameRenderProcessExited` (recoverable): call `core.Reload()` once,
  `LogWarn` the recovery attempt.
- `BrowserProcessExited` (unrecoverable): skip reload, route into the same fatal-failure path
  (log + overlay error) as a failed navigation.

## Error handling & logging

**Removing duplicate logging.** Both windows' always-attached `NavigationCompleted` handlers
currently have their own `LogUtility.LogWarn("Navigation failed: ...")` on `!e.IsSuccess`
(`GLLogin.WebView_NavigationCompleted`, `GLDrilldownCustomization.WebView_NavigationCompleted`).
Once `NavigateWithRetryAsync` owns retry/final-outcome logging, these two lines are **removed** —
otherwise the same failed attempt is logged twice, with different wording, from two different
places. The always-on handlers keep their existing early-return on failure; they simply stop
being a second logging path for the same event.

**Fatal-failure UX**, added at each of the three navigate call sites: on a final (post-retry)
failure, hide the busy overlay, restore `webView.Visibility`, and call
`AppOverlayControl.ShowErrorAsync("Unable to load the page. Please try again.")` — a generic
user-facing message. The real `CoreWebView2WebErrorStatus` goes only into the log line, never the
UI.

## Verification

No automated test project exists in this solution. Verification is manual, following this
session's established pattern (build + exercise the real app + inspect
`%LOCALAPPDATA%\ORBIT\Excel_Logs\GLSense_Logs\Logs\GLSense_Logs_<date>.log`):

1. **Normal login** against a known-good server — confirm no behavior change (one clean
   navigation, login succeeds).
2. **Self-signed/internal-CA cert** against a trusted (configured) server — confirm the page
   loads without an interstitial, and a `LogWarn` records the bypass.
3. **Cert error against an untrusted host** — confirm WebView2's default interstitial still
   appears (behavior unchanged for non-trusted hosts).
4. **Unreachable server** — confirm one retry occurs (~1s gap visible in log timestamps), then
   the overlay error shows and a final `LogWarn` is recorded.
5. **IdP popup** (if a test identity provider with a step-up/MFA popup is available) — confirm it
   opens inside the new managed popup window and closes cleanly on completion. If no such IdP is
   available for testing, this is flagged as unverified rather than assumed working.

## Open risks / follow-ups

- `ServerCertificateErrorDetected` and the `Action`/`Deferral` shapes used here require a
  reasonably current WebView2 SDK; the implementation step should confirm the referenced
  `Microsoft.Web.WebView2` package version exposes them before writing code against them.
- The generic overlay error message intentionally omits the raw `WebErrorStatus` from the UI;
  if support ends up needing that detail from users directly (rather than pulling the log), the
  message could later include a short code, but that's out of scope unless requested.
