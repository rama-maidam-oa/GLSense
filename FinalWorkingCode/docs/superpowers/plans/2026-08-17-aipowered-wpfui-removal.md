# AIPowered WPF-UI Removal + Window-Flash Fix Port Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** On a new branch cut from `11.1.0`, remove the WPF-UI (`Wpf.Ui`/`Wpf.Ui.Abstractions`) dependency from `AIPowered\GLSense\GLSense.Addin.Core` entirely, and port every behavioral bug fix from the `11.1.0-window-flash-redo` branch's `FinalWorkingCode\GLSense` changes into AIPowered's architecture, with zero regressions to AIPowered's existing functionality.

**Architecture:** `BaseWindow.cs` currently derives from `Wpf.Ui.Controls.FluentWindow` and already independently re-implements most of `DpiAwareWindow`'s DPI/centering/resettle logic (with its own timing model: resettle in `OnLoaded` + again in `OnContentRendered`). This plan (a) swaps `BaseWindow`'s parent to plain `Window` and replaces its internals with a line-for-line adaptation of `FinalWorkingCode\GLSense\Utilities\DpiAwareWindow.cs`'s newer timing model (resettle in `OnSourceInitialized`, before first paint — this is the actual fix for the "blank window" flash bug, and AIPowered doesn't have it yet), (b) removes the `Wpf.Ui`/`Wpf.Ui.Abstractions` NuGet dependency and rewrites the theme bootstrapper to drop its now-uncompilable `Wpf.Ui.Appearance` calls, (c) ports the two brand-new utility classes (`WebView2Warmup`, `WindowLoadingPlaceholder`) into AIPowered, wired through the new `BaseWindow`, and (d) ports the 6 specific window/file behavior fixes plus 2 unrelated data-correctness fixes, verbatim in effect, adapted to AIPowered's `ServiceLocator`/AppDomain-split conventions per `AIPowered\GLSense\PORTING_GUIDE.md`.

**Tech Stack:** .NET Framework 4.8.1, WPF, VSTO/AddinExpress (`AddinModule.cs` host + AppDomain-isolated `GLSense.Addin.Core`), WebView2, SQLite.

**Spec:** No separate spec doc — this plan is driven directly by (1) the diff between `11.1.0` and `11.1.0-window-flash-redo` in `FinalWorkingCode\GLSense` (19 files, ~1000 lines — window-flash/DPI fixes + 2 unrelated data fixes), and (2) `AIPowered\GLSense\PORTING_GUIDE.md` (porting conventions) and `AIPowered\GLSense\CLAUDE.md` (AIPowered's own fix history — not fully read; consult its DPI/resettle sections, e.g. "section 1.4e", if a step in this plan behaves unexpectedly during execution).

## Global Constraints

- No automated test project exists for either codebase — every task's "verify" step is a `msbuild`/`dotnet build` compile check plus a manual smoke-test instruction (open the affected window(s), confirm behavior), not a unit test. This is the established verification pattern for this repo (see `PORTING_GUIDE.md` section 7, which describes build + manual verification, not unit tests).
- Every static reference must be re-pointed per `PORTING_GUIDE.md`: `AppState.Instance.X` stays `AppState.Instance.X` (AIPowered has its own smaller `AppState`, already has `SelectedCube`/`SelectedLedger`/`LoginToken`/etc. — verified in this plan's research, do not add new AppState members unless a task below says to), `LogUtility.*` → `ServiceLocator.Logger?.*`, `AppPaths.*` → `ServiceLocator.Paths.*`.
- New `.cs`/`.xaml` files must get explicit `<Compile>`/`<Page>` entries in `GLSense.Addin.Core.csproj` — this project type does not glob automatically.
- Every method invoked fire-and-forget across the AppDomain boundary (anything reached from `AddinEntry.OnRibbonAction`) must catch everything internally and never rethrow (`PORTING_GUIDE.md` section 2).
- Do not touch `11.1.0`, `main`, or `11.1.0-window-flash-redo` — all work happens on the new branch created in Task 1.
- FinalWorkingCode's `AppConstants.cs` `DefaultCommitDate` change (`"13-Aug-2026"` → `"17-Aug-2026"`) is a version-stamp bump, not a behavior fix, and is intentionally **not** ported to AIPowered's `AppConstants.cs` by this plan — it isn't "a change" in the sense the rest of this plan ports; update it manually only if/when AIPowered's own build-date convention calls for it.
- **Out of scope for this plan (flagged, not silently dropped):** AIPowered's window *chrome* (title-bar row height/margins/style-key names) is structurally different from FinalWorkingCode's (`Themes/GlobalStyles.xaml` comparison confirmed: AIPowered uses a 2-row Grid + `TitleBarGridStyle`/`ExtendsContentIntoTitleBar`; FinalWorkingCode uses a 3-row Grid + `HeaderBar`/`CustomWindowCloseButtonStyle` + `Margin="10"` around everything, `WindowStyle="None"`). This plan makes AIPowered's windows **behave** identically (same fixes, same timing, no WPF-UI dependency, `WindowStyle="None"`, no regressions) but does **not** rebuild all 26 windows' visual layout pixel-for-pixel to match FinalWorkingCode's specific chrome styling — that is a separate, much larger per-window visual-parity pass (touches all 26 `BaseWindow`-derived windows' XAML, not just the 6 with behavior fixes). Flag to the user for an explicit go/no-go before starting that pass; see the note at the end of this plan for the recipe if it's requested.

---

## Task 1: Create the new branch

**Files:** none (git operation only)

- [ ] **Step 1: Confirm working tree is clean and fetch latest**

```bash
git status
git fetch origin
```

Expected: clean working tree (no uncommitted changes reported).

- [ ] **Step 2: Create the new branch from `11.1.0`**

```bash
git checkout 11.1.0
git pull origin 11.1.0
git checkout -b 11.1.0-aipowered-wpfui-removal
```

- [ ] **Step 3: Bring the FinalWorkingCode window-flash fixes onto the new branch**

The 15 commits below (already verified, build-verified per their own commit messages) contain the FinalWorkingCode-side changes this plan's AIPowered work is a port of. Merge them in so both codebases end up consistent on this branch:

```bash
git merge 11.1.0-window-flash-redo
```

Expected: fast-forward or clean merge (both branches share `11.1.0` as their merge-base, and no FinalWorkingCode file has been touched independently on `11.1.0` since). If there's a conflict, stop and resolve manually rather than guessing — do not force.

- [ ] **Step 4: Verify branch state**

```bash
git log --oneline -5
git status
```

Expected: `git log` shows the window-flash-fix commits at the tip; `git status` is clean.

---

## Task 2: Rewrite `BaseWindow.cs` — remove FluentWindow, port DpiAwareWindow's timing model

**Files:**
- Modify: `AIPowered/GLSense/GLSense.Addin.Core/Views/BaseWindow.cs`

**Interfaces:**
- Consumes: `ServiceLocator.Logger`, `ServiceLocator.ExcelHandle`, `MouseWheelFocusHelper.EnableHoverToScroll(Window)`, `ExcelWindowHelper.ActivateExcelMainWindow()`, `DpiAwarenessHelper.SetPerMonitorAware()`/`GetWindowDpi(Window)` (all pre-existing in AIPowered, unchanged).
- Produces (unchanged public surface — every one of the 26 derived windows keeps compiling against these): `EnableAutoLayoutRefresh`, `EnableExcelCentering`, `EnableEscapeToClose`, `AutoClampToWorkArea`, `WorkAreaMargin`, `WindowCaption`, `IconSymbol`, `CenterInExcel`, `ModalToExcel`, `MaxWidthCap`, `MaxHeightCap`, `ShowDialog()`, `SafeShowDialog()`. New: `DisableAutoSizing` (bool, default false — ported from `DpiAwareWindow`, needed by dialogs that should ignore auto-sizing entirely), `SetExcelOwner(IntPtr)`, `ShowDialogWithOwner(IntPtr)`, `ShowWithOwner(IntPtr)` (ported from `DpiAwareWindow` — not currently on `BaseWindow`, but harmless additions; nothing currently calls the old `ModalToExcel`-based owner path in a way these conflict with). `ForceSizeToContentResettle()`/`PumpDispatcherFrame()` stay `protected` (GLLOVs's `DataLoadedAction` in Task 10 still calls them).

This task keeps the class name `BaseWindow` and its constructor-based `ServiceLocator.ExcelHandle` owner-setting flow (AIPowered's own convention, not present in `DpiAwareWindow`) — only the FluentWindow inheritance, WPF-UI bootstrap calls, and the resettle timing model are replaced.

- [ ] **Step 1: Replace the full file content**

```csharp
using GLSense.Addin.Core.Helpers;
using GLSense.Addin.Core.Infrastructure;
using GLSense.Addin.Core.Utilities;
using System;
using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Input;
using System.Windows.Interop;
using System.Windows.Media;

namespace GLSense.Addin.Core.Views
{
    public abstract class BaseWindow : Window
    {
        [DllImport("user32.dll")]
        private static extern bool GetWindowRect(IntPtr hWnd, out RECT lpRect);

        [DllImport("user32.dll")]
        private static extern bool SetForegroundWindow(IntPtr hWnd);

        [DllImport("user32.dll", SetLastError = true)]
        private static extern bool SetWindowPos(IntPtr hWnd, IntPtr hWndInsertAfter,
            int X, int Y, int cx, int cy, uint uFlags);

        [DllImport("user32.dll")]
        private static extern bool RedrawWindow(IntPtr hWnd, IntPtr lprcUpdate, IntPtr hrgnUpdate, uint flags);

        private const uint SWP_NOMOVE = 0x0002;
        private const uint SWP_NOSIZE = 0x0001;
        private const uint SWP_NOZORDER = 0x0004;
        private const uint SWP_NOACTIVATE = 0x0010;
        private const uint SWP_FRAMECHANGED = 0x0020;

        private const uint RDW_INVALIDATE = 0x0001;
        private const uint RDW_ERASE = 0x0004;
        private const uint RDW_FRAME = 0x0400;
        private const uint RDW_ALLCHILDREN = 0x0080;
        private const uint RDW_UPDATENOW = 0x0100;

        [StructLayout(LayoutKind.Sequential)]
        private struct RECT
        {
            public int Left;
            public int Top;
            public int Right;
            public int Bottom;
            public readonly int Width => Right - Left;
            public readonly int Height => Bottom - Top;
        }

        private HwndSource _hwndSource;
        private double _currentScaleFactor = 1.0;
        private readonly ScaleTransform _dpiScaleTransform = new ScaleTransform(1.0, 1.0);
        private readonly string _windowName;
        private IntPtr _excelHandle;
        private bool _ownerSet;
        private bool _initialLayoutApplied;
        private double _initialMaxWidth = double.NaN;
        private double _initialMaxHeight = double.NaN;
        private double _initialMinHeight = double.NaN;
        private System.Windows.Threading.DispatcherTimer _resizeSettleTimer;

        public bool EnableAutoLayoutRefresh { get; set; } = true;
        public bool EnableExcelCentering { get; set; } = true;
        public bool EnableEscapeToClose { get; set; } = true;
        public bool AutoClampToWorkArea { get; set; } = true;
        public double WorkAreaMargin { get; set; } = 24d;
        public double? MaxWidthCap { get; set; } = 1400d;
        public double? MaxHeightCap { get; set; } = null;
        public double MinContentScale { get; set; } = 0.85;

        // Escape hatch for dialogs that should ignore all auto-sizing/clamping (ported
        // from FinalWorkingCode's DpiAwareWindow - not previously exposed on BaseWindow).
        public bool DisableAutoSizing { get; set; } = false;

        public string WindowCaption
        {
            get => Title;
            set => Title = value;
        }

        // FontAwesome PackIconFontAwesomeKind name bound to each window's title-bar
        // iconPacks:PackIconFontAwesome via TitleBarIconStyle - kept as a plain string
        // (not the enum type) since WPF's binding engine coerces a string source into an
        // enum-typed target property automatically. Unchanged from the pre-WPF-UI-removal
        // BaseWindow - every derived window's XAML already sets IconSymbol="SomeKind".
        public string IconSymbol { get; set; } = "KeySolid";
        public bool CenterInExcel { get; set; } = true;
        public bool ModalToExcel { get; set; } = true;

        protected BaseWindow()
        {
            try
            {
                _windowName = GetType().Name;
                ServiceLocator.Logger?.LogDebug($"[{_windowName}] constructing window");

                using (DpiAwarenessHelper.SetPerMonitorAware())
                {
                    this.UseLayoutRounding = true;
                    this.SnapsToDevicePixels = true;
                    TextOptions.SetTextFormattingMode(this, TextFormattingMode.Display);
                    RenderOptions.SetBitmapScalingMode(this, BitmapScalingMode.HighQuality);
                }

                try
                {
                    _excelHandle = ServiceLocator.ExcelHandle;
                    if (_excelHandle != IntPtr.Zero)
                        ServiceLocator.Logger?.LogDebug($"[{_windowName}] Excel handle obtained from ServiceLocator: {_excelHandle}");
                    else
                        ServiceLocator.Logger?.LogWarn($"[{_windowName}] Excel handle is IntPtr.Zero from ServiceLocator");
                }
                catch (Exception ex)
                {
                    ServiceLocator.Logger?.LogException(ex, $"[{_windowName}] Failed to get Excel handle from ServiceLocator");
                }

                this.SourceInitialized += OnSourceInitialized;
                this.KeyDown += BaseWindow_KeyDown;
                this.Closed += OnClosedRestoreFocus;

                MouseWheelFocusHelper.EnableHoverToScroll(this);

                ServiceLocator.Logger?.LogDebug($"[{_windowName}] constructor completed with DPI awareness");
            }
            catch (Exception ex)
            {
                ServiceLocator.Logger?.LogException(ex, $"[{_windowName}] constructor error");
            }
        }

        public void SetExcelOwner(IntPtr excelHwnd)
        {
            try
            {
                _excelHandle = excelHwnd;
                var helper = new WindowInteropHelper(this);
                helper.Owner = excelHwnd;
                _ownerSet = true;
            }
            catch (Exception ex)
            {
                ServiceLocator.Logger?.LogException(ex, $"[{_windowName}] SetExcelOwner error");
            }
        }

        public bool? ShowDialogWithOwner(IntPtr excelHwnd)
        {
            SetExcelOwner(excelHwnd);
            return this.ShowDialog();
        }

        public void ShowWithOwner(IntPtr excelHwnd)
        {
            SetExcelOwner(excelHwnd);
            this.Show();
        }

        // ShowDialog with proper Excel ownership using ServiceLocator - unchanged from
        // pre-removal BaseWindow (this is AIPowered's own convention, not present in
        // FinalWorkingCode's DpiAwareWindow, which uses SetExcelOwner explicitly instead).
        public new bool? ShowDialog()
        {
            try
            {
                ServiceLocator.Logger?.LogDebug($"[{_windowName}] ShowDialog called");

                if (!this.Dispatcher.CheckAccess())
                    return this.Dispatcher.Invoke(() => ShowDialog());

                if (ModalToExcel && !_ownerSet)
                {
                    try
                    {
                        _excelHandle = ServiceLocator.ExcelHandle;
                    }
                    catch (Exception ex)
                    {
                        ServiceLocator.Logger?.LogDebug($"[{_windowName}] Could not get Excel handle from ServiceLocator: {ex.Message}");
                    }

                    if (_excelHandle != IntPtr.Zero)
                    {
                        try
                        {
                            var helper = new WindowInteropHelper(this);
                            if (helper.Owner == IntPtr.Zero)
                            {
                                helper.Owner = _excelHandle;
                                _ownerSet = true;
                                ServiceLocator.Logger?.LogDebug($"[{_windowName}] Excel owner set in ShowDialog: {_excelHandle}");
                            }
                        }
                        catch (Exception ex)
                        {
                            ServiceLocator.Logger?.LogDebug($"[{_windowName}] Could not set Excel owner in ShowDialog: {ex.Message}");
                        }
                    }
                    else
                    {
                        ServiceLocator.Logger?.LogWarn($"[{_windowName}] Excel handle is IntPtr.Zero in ShowDialog");
                    }
                }

                return base.ShowDialog();
            }
            catch (Exception ex)
            {
                ServiceLocator.Logger?.LogException(ex, $"[{_windowName}] ShowDialog error");
                throw;
            }
        }

        public void SafeShowDialog()
        {
            try
            {
                ServiceLocator.Logger?.LogDebug($"[{_windowName}] SafeShowDialog called");
                if (this.Dispatcher.CheckAccess())
                    this.ShowDialog();
                else
                    this.Dispatcher.Invoke(() => this.ShowDialog());
            }
            catch (Exception ex)
            {
                ServiceLocator.Logger?.LogException(ex, $"[{_windowName}] SafeShowDialog error");
            }
        }

        // Ported from FinalWorkingCode's DpiAwareWindow.OnSourceInitialized: runs the
        // DPI/fit/center pass synchronously while the window has no on-screen presence
        // yet (SourceInitialized fires once the HWND exists, strictly before
        // Show()/ShowDialog() calls ShowWindow) - not later via a deferred OnLoaded pass,
        // which is what caused the visible resize/reposition "pop" AIPowered's previous
        // FluentWindow-based BaseWindow still had (its OnLoaded ran this after the window
        // was already shown at its placeholder CenterOwner-computed position/size). Doing
        // the same math here means Show()/ShowDialog() paints the correct final
        // size/position on the very first frame.
        private void OnSourceInitialized(object sender, EventArgs e)
        {
            try
            {
                ServiceLocator.Logger?.LogDebug($"[{_windowName}] source initialized");
                _hwndSource = PresentationSource.FromVisual(this) as HwndSource;
                _hwndSource?.AddHook(WndProc);

                if (ModalToExcel && _excelHandle != IntPtr.Zero)
                {
                    try
                    {
                        var helper = new WindowInteropHelper(this);
                        if (helper.Owner == IntPtr.Zero)
                        {
                            helper.Owner = _excelHandle;
                            _ownerSet = true;
                            ServiceLocator.Logger?.LogDebug($"[{_windowName}] Excel owner set successfully: {_excelHandle}");
                        }
                    }
                    catch (Exception ex)
                    {
                        ServiceLocator.Logger?.LogDebug($"[{_windowName}] Could not set Excel owner: {ex.Message}");
                    }
                }

                _currentScaleFactor = GetCurrentScaleFactor();

                if (!DisableAutoSizing)
                {
                    CaptureInitialWindowConstraints();
                    ApplyLayoutRefresh();
                }

                int placeholderGen = WindowLoadingPlaceholder.ShowMatching(
                    Left, Top, ResolveExpectedWidth(), ResolveExpectedHeight(), _excelHandle);
                HookPlaceholderDismissal(placeholderGen);
            }
            catch (Exception ex)
            {
                ServiceLocator.Logger?.LogException(ex, $"[{_windowName}] OnSourceInitialized error");
            }
        }

        private void HookPlaceholderDismissal(int generation)
        {
            if (generation < 0)
                return;

            void Dismiss(object sender, EventArgs e)
            {
                this.ContentRendered -= Dismiss;
                this.Closed -= Dismiss;
                WindowLoadingPlaceholder.Hide(generation);
            }

            this.ContentRendered += Dismiss;
            this.Closed += Dismiss;
        }

        private double ResolveExpectedWidth()
        {
            if (double.IsNaN(Width) || double.IsInfinity(Width))
                return double.NaN;

            double w = Width;
            if (!double.IsNaN(MinWidth)) w = Math.Max(w, MinWidth);
            if (!double.IsPositiveInfinity(MaxWidth)) w = Math.Min(w, MaxWidth);
            return w;
        }

        private double ResolveExpectedHeight()
        {
            if (double.IsNaN(Height) || double.IsInfinity(Height))
                return double.NaN;

            double h = Height;
            if (!double.IsNaN(MinHeight)) h = Math.Max(h, MinHeight);
            if (!double.IsPositiveInfinity(MaxHeight)) h = Math.Min(h, MaxHeight);
            return h;
        }

        private void ApplyLayoutRefresh()
        {
            try
            {
                if (!EnableAutoLayoutRefresh || DisableAutoSizing)
                    return;

                FitToAvailableWorkArea();

                if (EnableExcelCentering && !_initialLayoutApplied)
                {
                    _initialLayoutApplied = true;
                    CenterOverOwnerOnce();
                }
            }
            catch (Exception ex)
            {
                ServiceLocator.Logger?.LogException(ex, $"[{_windowName}] ApplyLayoutRefresh error");
            }
        }

        // See CLAUDE.md section 1.4e for AIPowered's own prior history of this mechanism.
        // Retained here (called from GLLOVs's DataLoadedAction, per Task 10) for windows
        // whose real content settles asynchronously after first paint.
        protected void ForceSizeToContentResettle()
        {
            try
            {
                double previousLeft = this.Left;
                double previousTop = this.Top;
                double previousWidth = this.ActualWidth > 0 ? this.ActualWidth : this.Width;
                double previousHeight = this.ActualHeight > 0 ? this.ActualHeight : this.Height;

                var mode = this.SizeToContent;
                this.SizeToContent = SizeToContent.Manual;
                this.UpdateLayout();

                if (this.ActualWidth > 0 && this.ActualHeight > 0)
                {
                    this.Width = this.ActualWidth + 1.0;
                    this.Height = this.ActualHeight + 1.0;
                    this.UpdateLayout();
                }

                this.SizeToContent = mode;
                this.UpdateLayout();

                ForceFrameRedraw();

                if (CenterInExcel)
                {
                    double newWidth = this.ActualWidth > 0 ? this.ActualWidth : this.Width;
                    double newHeight = this.ActualHeight > 0 ? this.ActualHeight : this.Height;

                    if (Math.Abs(newWidth - previousWidth) > 0.5 || Math.Abs(newHeight - previousHeight) > 0.5)
                        RecenterAfterSizeChange(previousLeft, previousTop, previousWidth, previousHeight);
                }

                ServiceLocator.Logger?.LogDebug($"[{_windowName}] SizeToContent resettled ({mode})");
            }
            catch (Exception ex)
            {
                ServiceLocator.Logger?.LogException(ex, $"[{_windowName}] ForceSizeToContentResettle error");
            }
        }

        protected void PumpDispatcherFrame()
        {
            try
            {
                var frame = new System.Windows.Threading.DispatcherFrame();
                this.Dispatcher.BeginInvoke(new Action(() => frame.Continue = false),
                    System.Windows.Threading.DispatcherPriority.Background);
                System.Windows.Threading.Dispatcher.PushFrame(frame);
            }
            catch (Exception ex)
            {
                ServiceLocator.Logger?.LogException(ex, $"[{_windowName}] PumpDispatcherFrame error");
            }
        }

        // WindowStyle="None" windows still get a DWM-drawn drop shadow around their real
        // client area, and a programmatic resize can leave stale rendering behind at the
        // old edge. SWP_FRAMECHANGED alone only recomputes the non-client frame/shadow,
        // not the client area - RedrawWindow with INVALIDATE|ERASE|FRAME|ALLCHILDREN|
        // UPDATENOW forces an immediate full erase-and-repaint. Ported from
        // FinalWorkingCode's DpiAwareWindow.ForceFrameRedraw.
        private void ForceFrameRedraw()
        {
            try
            {
                var hwnd = _hwndSource?.Handle ?? IntPtr.Zero;
                if (hwnd == IntPtr.Zero)
                    return;

                SetWindowPos(hwnd, IntPtr.Zero, 0, 0, 0, 0,
                    SWP_NOMOVE | SWP_NOSIZE | SWP_NOZORDER | SWP_NOACTIVATE | SWP_FRAMECHANGED);
                RedrawWindow(hwnd, IntPtr.Zero, IntPtr.Zero,
                    RDW_INVALIDATE | RDW_ERASE | RDW_FRAME | RDW_ALLCHILDREN | RDW_UPDATENOW);
            }
            catch (Exception ex)
            {
                ServiceLocator.Logger?.LogException(ex, $"[{_windowName}] ForceFrameRedraw error");
            }
        }

        private void RecenterAfterSizeChange(double previousLeft, double previousTop, double previousWidth, double previousHeight)
        {
            try
            {
                if (double.IsNaN(previousLeft) || double.IsNaN(previousTop) ||
                    double.IsNaN(previousWidth) || double.IsNaN(previousHeight) ||
                    previousWidth <= 0 || previousHeight <= 0)
                    return;

                double centerX = previousLeft + (previousWidth / 2.0);
                double centerY = previousTop + (previousHeight / 2.0);
                PositionAroundCenter(centerX, centerY);
            }
            catch (Exception ex)
            {
                ServiceLocator.Logger?.LogException(ex, $"[{_windowName}] RecenterAfterSizeChange error");
            }
        }

        private void PositionAroundCenter(double centerX, double centerY)
        {
            double effectiveWidth = ActualWidth > 0 ? ActualWidth : Width;
            double effectiveHeight = ActualHeight > 0 ? ActualHeight : Height;

            if (double.IsNaN(effectiveWidth) || double.IsNaN(effectiveHeight) || effectiveWidth <= 0 || effectiveHeight <= 0)
                return;

            double newLeft = centerX - (effectiveWidth / 2.0);
            double newTop = centerY - (effectiveHeight / 2.0);

            var workArea = SystemParameters.WorkArea;
            if (effectiveWidth < workArea.Width)
                newLeft = Math.Max(workArea.Left, Math.Min(newLeft, workArea.Right - effectiveWidth));
            if (effectiveHeight < workArea.Height)
                newTop = Math.Max(workArea.Top, Math.Min(newTop, workArea.Bottom - effectiveHeight));

            Left = newLeft;
            Top = newTop;
        }

        private void CenterOverOwnerOnce()
        {
            try
            {
                UpdateLayout();

                double centerX, centerY;

                if (_excelHandle != IntPtr.Zero && GetWindowRect(_excelHandle, out RECT ownerRectPx) &&
                    ownerRectPx.Width > 0 && ownerRectPx.Height > 0)
                {
                    double scale = _currentScaleFactor > 0 ? _currentScaleFactor : 1.0;
                    double ownerLeft = ownerRectPx.Left / scale;
                    double ownerTop = ownerRectPx.Top / scale;
                    double ownerWidth = (ownerRectPx.Right - ownerRectPx.Left) / scale;
                    double ownerHeight = (ownerRectPx.Bottom - ownerRectPx.Top) / scale;

                    centerX = ownerLeft + (ownerWidth / 2.0);
                    centerY = ownerTop + (ownerHeight / 2.0);
                }
                else
                {
                    var workArea = SystemParameters.WorkArea;
                    centerX = workArea.Left + (workArea.Width / 2.0);
                    centerY = workArea.Top + (workArea.Height / 2.0);
                }

                PositionAroundCenter(centerX, centerY);
            }
            catch (Exception ex)
            {
                ServiceLocator.Logger?.LogException(ex, $"[{_windowName}] CenterOverOwnerOnce error");
            }
        }

        private double GetCurrentScaleFactor()
        {
            try
            {
                if (_hwndSource?.CompositionTarget != null)
                {
                    var scale = _hwndSource.CompositionTarget.TransformToDevice.M11;
                    if (scale > 0) return scale;
                }
            }
            catch { /* fall through */ }

            try
            {
                var dpi = DpiAwarenessHelper.GetWindowDpi(this);
                if (dpi > 0) return dpi / 96.0;
            }
            catch { /* ignore */ }

            return 1.0;
        }

        private void FitToAvailableWorkArea()
        {
            if (!AutoClampToWorkArea || DisableAutoSizing)
                return;

            try
            {
                if (Content is not FrameworkElement root)
                    return;

                var workArea = SystemParameters.WorkArea;
                var availableWidth = Math.Max(0, workArea.Width - (WorkAreaMargin * 2));
                var availableHeight = Math.Max(0, workArea.Height - (WorkAreaMargin * 2));

                var requestedMaxWidth = GetEffectiveRequestedMaxWidth();
                if (!double.IsPositiveInfinity(requestedMaxWidth))
                    availableWidth = Math.Min(availableWidth, requestedMaxWidth);

                if (MaxWidthCap.HasValue)
                    availableWidth = Math.Min(availableWidth, MaxWidthCap.Value);
                if (MaxHeightCap.HasValue)
                    availableHeight = Math.Min(availableHeight, MaxHeightCap.Value);

                root.Measure(new Size(double.PositiveInfinity, double.PositiveInfinity));
                var desiredWidth = root.DesiredSize.Width;
                var desiredHeight = root.DesiredSize.Height;

                if (!double.IsNaN(_initialMinHeight) && _initialMinHeight > 0)
                    desiredHeight = Math.Max(desiredHeight, _initialMinHeight);

                if (desiredWidth <= 0 || desiredHeight <= 0)
                    return;

                var rawScale = Math.Min(availableWidth / desiredWidth, availableHeight / desiredHeight);
                var fitScale = Math.Min(1.0, rawScale);
                if (MinContentScale > 0 && fitScale < MinContentScale)
                    fitScale = MinContentScale;

                ApplyScaleTransform(fitScale);

                var targetWidth = Math.Min(desiredWidth * fitScale, availableWidth);
                var targetHeight = Math.Min(desiredHeight * fitScale, availableHeight);

                double previousLeft = Left;
                double previousTop = Top;
                double previousWidth = Width;
                double previousHeight = Height;
                bool sizeChanged = false;

                if (targetWidth > 0 && (double.IsNaN(previousWidth) || Math.Abs(targetWidth - previousWidth) > 0.5))
                {
                    Width = targetWidth;
                    sizeChanged = true;
                }
                if (targetHeight > 0 && (double.IsNaN(previousHeight) || Math.Abs(targetHeight - previousHeight) > 0.5))
                {
                    Height = targetHeight;
                    sizeChanged = true;
                }

                MaxWidth = availableWidth;
                MaxHeight = availableHeight;

                if (sizeChanged)
                {
                    RecenterAfterSizeChange(previousLeft, previousTop, previousWidth, previousHeight);
                    ForceFrameRedraw();
                }
            }
            catch (Exception ex)
            {
                ServiceLocator.Logger?.LogException(ex, $"[{_windowName}] FitToAvailableWorkArea error");
            }
        }

        private void ApplyScaleTransform(double scaleFactor)
        {
            if (Content is not FrameworkElement element)
                return;

            if (Math.Abs(scaleFactor - 1.0) < 0.001)
            {
                element.LayoutTransform = Transform.Identity;
                return;
            }

            if (Math.Abs(scaleFactor - _currentScaleFactor) < 0.001)
                return;

            try
            {
                _dpiScaleTransform.ScaleX = scaleFactor;
                _dpiScaleTransform.ScaleY = scaleFactor;
                element.LayoutTransform = _dpiScaleTransform;
                element.InvalidateMeasure();
            }
            catch (Exception ex)
            {
                ServiceLocator.Logger?.LogException(ex, $"[{_windowName}] ApplyScaleTransform error");
            }
        }

        private void CaptureInitialWindowConstraints()
        {
            if (double.IsNaN(_initialMaxWidth)) _initialMaxWidth = MaxWidth;
            if (double.IsNaN(_initialMaxHeight)) _initialMaxHeight = MaxHeight;
            if (double.IsNaN(_initialMinHeight)) _initialMinHeight = MinHeight;
        }

        private double GetEffectiveRequestedMaxWidth()
        {
            var maxWidth = double.IsNaN(_initialMaxWidth) ? MaxWidth : _initialMaxWidth;
            if (double.IsPositiveInfinity(maxWidth))
                return double.PositiveInfinity;
            return maxWidth + 200;
        }

        protected override void OnRenderSizeChanged(SizeChangedInfo sizeInfo)
        {
            base.OnRenderSizeChanged(sizeInfo);

            if (DisableAutoSizing || !AutoClampToWorkArea)
                return;

            // Debounce: content (e.g. a DataGrid) can grow across several back-to-back
            // RenderSizeChanged events as rows populate - reacting to every one made an
            // already-visible window visibly resize/reposition more than once in quick
            // succession ("dancing"). Only run EnsureFitsWorkArea once rendering has been
            // quiet for 120ms. Ported from FinalWorkingCode's DpiAwareWindow.
            _resizeSettleTimer?.Stop();
            _resizeSettleTimer = new System.Windows.Threading.DispatcherTimer
            {
                Interval = TimeSpan.FromMilliseconds(120)
            };
            _resizeSettleTimer.Tick += (s, e) =>
            {
                _resizeSettleTimer.Stop();
                try { EnsureFitsWorkArea(); }
                catch (Exception ex) { ServiceLocator.Logger?.LogException(ex, $"[{_windowName}] OnRenderSizeChanged (debounced clamp)"); }
            };
            _resizeSettleTimer.Start();
        }

        protected void EnsureFitsWorkArea(double? marginOverride = null)
        {
            if (DisableAutoSizing)
                return;

            var margin = marginOverride ?? WorkAreaMargin;
            try
            {
                double previousLeft = Left;
                double previousTop = Top;
                double previousWidth = Width;
                double previousHeight = Height;
                bool sizeChanged = false;

                var workArea = SystemParameters.WorkArea;
                var baseMaxWidth = Math.Max(0, workArea.Width - margin);
                var baseMaxHeight = Math.Max(0, workArea.Height - margin);

                var requestedMaxWidth = GetEffectiveRequestedMaxWidth();
                if (!double.IsPositiveInfinity(requestedMaxWidth))
                    baseMaxWidth = Math.Min(baseMaxWidth, requestedMaxWidth);
                if (MaxWidthCap.HasValue)
                    baseMaxWidth = Math.Min(baseMaxWidth, MaxWidthCap.Value);
                if (MaxHeightCap.HasValue)
                    baseMaxHeight = Math.Min(baseMaxHeight, MaxHeightCap.Value);

                var effectiveMaxWidth = double.IsPositiveInfinity(MaxWidth) ? baseMaxWidth : Math.Min(MaxWidth, baseMaxWidth);
                var effectiveMaxHeight = double.IsPositiveInfinity(MaxHeight) ? baseMaxHeight : Math.Min(MaxHeight, baseMaxHeight);

                MaxWidth = effectiveMaxWidth;
                MaxHeight = effectiveMaxHeight;

                if (MinWidth > effectiveMaxWidth) MinWidth = effectiveMaxWidth;
                if (MinHeight > effectiveMaxHeight) MinHeight = effectiveMaxHeight;

                if (Width > effectiveMaxWidth) { Width = effectiveMaxWidth; sizeChanged = true; }
                else if (Width < MinWidth) { Width = MinWidth; sizeChanged = true; }

                if (Height > effectiveMaxHeight) { Height = effectiveMaxHeight; sizeChanged = true; }
                else if (Height < MinHeight) { Height = MinHeight; sizeChanged = true; }

                if (sizeChanged)
                {
                    RecenterAfterSizeChange(previousLeft, previousTop, previousWidth, previousHeight);
                    ForceFrameRedraw();
                }
            }
            catch (Exception ex)
            {
                ServiceLocator.Logger?.LogException(ex, $"[{_windowName}] EnsureFitsWorkArea error");
            }
        }

        private void BaseWindow_KeyDown(object sender, KeyEventArgs e)
        {
            try
            {
                if (!e.Handled && e.Key == Key.Escape && EnableEscapeToClose)
                {
                    e.Handled = true;
                    ServiceLocator.Logger?.LogDebug($"[{_windowName}] Escape pressed - closing window");
                    this.Close();
                }
            }
            catch (Exception ex)
            {
                ServiceLocator.Logger?.LogException(ex, $"[{_windowName}] BaseWindow_KeyDown error");
            }
        }

        private IntPtr WndProc(IntPtr hwnd, int msg, IntPtr wParam, IntPtr lParam, ref bool handled)
        {
            try
            {
                const int WM_DPICHANGED = 0x02E0;
                const int WM_ACTIVATE = 0x0006;
                const int WA_ACTIVE = 1;
                const int WA_CLICKACTIVE = 2;

                switch (msg)
                {
                    case WM_DPICHANGED when !DisableAutoSizing:
                        AdjustForDpiChange((uint)wParam, lParam);
                        handled = true;
                        break;

                    case WM_ACTIVATE:
                        if (((int)wParam == WA_ACTIVE || (int)wParam == WA_CLICKACTIVE) && _ownerSet && _excelHandle != IntPtr.Zero)
                        {
                            this.Dispatcher.BeginInvoke(new Action(() =>
                            {
                                try { if (this.IsVisible && !this.IsActive) this.Focus(); }
                                catch (Exception ex) { ServiceLocator.Logger?.LogDebug($"WM_ACTIVATE handler error: {ex.Message}"); }
                            }), System.Windows.Threading.DispatcherPriority.Background);
                        }
                        break;
                }
            }
            catch (Exception ex)
            {
                ServiceLocator.Logger?.LogException(ex, $"[{_windowName}] WndProc error for msg {msg}");
            }

            return IntPtr.Zero;
        }

        private void AdjustForDpiChange(uint newDpi, IntPtr lParam)
        {
            try
            {
                var scaleFactor = newDpi / 96.0;
                _currentScaleFactor = scaleFactor;

                ApplyScaleTransform(scaleFactor);

                if (lParam != IntPtr.Zero)
                {
                    var rect = Marshal.PtrToStructure<RECT>(lParam);
                    bool autoSized = this.SizeToContent != SizeToContent.Manual;

                    this.Dispatcher.BeginInvoke(new Action(() =>
                    {
                        try
                        {
                            this.Left = rect.Left / scaleFactor;
                            this.Top = rect.Top / scaleFactor;

                            if (!autoSized)
                            {
                                this.Width = rect.Width / scaleFactor;
                                this.Height = rect.Height / scaleFactor;
                                FitToAvailableWorkArea();
                            }
                            else
                            {
                                ForceSizeToContentResettle();
                                PumpDispatcherFrame();
                            }
                        }
                        catch (Exception ex)
                        {
                            ServiceLocator.Logger?.LogException(ex, $"[{_windowName}] AdjustForDpiChange (resize) error");
                        }
                    }));
                }
            }
            catch (Exception ex)
            {
                ServiceLocator.Logger?.LogException(ex, $"[{_windowName}] AdjustForDpiChange error");
            }
        }

        protected override void OnClosed(EventArgs e)
        {
            try
            {
                if (_hwndSource != null)
                {
                    _hwndSource.RemoveHook(WndProc);
                    _hwndSource = null;
                }
            }
            catch (Exception ex)
            {
                ServiceLocator.Logger?.LogException(ex, $"[{_windowName}] OnClosed error");
            }

            base.OnClosed(e);
        }

        // Closing a window does not reliably return focus to its owner - the OS can just
        // as easily hand it to whatever unrelated application is next in its own
        // activation history. Unchanged from the pre-removal BaseWindow.
        private void OnClosedRestoreFocus(object sender, EventArgs e)
        {
            try
            {
                if (Owner != null)
                    Owner.Activate();
                else if (_ownerSet && _excelHandle != IntPtr.Zero)
                    ExcelWindowHelper.ActivateExcelMainWindow();
            }
            catch (Exception ex)
            {
                ServiceLocator.Logger?.LogException(ex, $"[{_windowName}] OnClosedRestoreFocus error");
            }
        }
    }
}
```

- [ ] **Step 2: Remove the now-orphaned `Wpf.Ui.Controls` using and confirm no other file references `BaseWindow`'s old drag-override behavior**

Run:
```bash
grep -rn "OnMouseLeftButtonDown\|OnMouseLeftButtonUp\|OnMouseMove" AIPowered/GLSense/GLSense.Addin.Core/Views/*.xaml.cs
```
Expected: no derived window overrides these (confirmed during research - every window already uses its own dedicated `TitleBar_MouseLeftButtonDown` handler calling `DragMove()`, e.g. `GLLogin.xaml.cs:151`). The window-wide drag override removed from `BaseWindow` in Step 1 was flagged in `PORTING_GUIDE.md` as a known bug (intercepts clicks meant for buttons/inputs) - removing it is a bug fix, not a regression.

- [ ] **Step 3: Build and fix compile errors**

```bash
msbuild AIPowered/GLSense/GLSense.sln /t:GLSense_Addin_Core /p:Configuration=Debug
```
Expected at this point: compile errors only in `WpfUiBootstrapper.cs`/`AddinEntry.cs` (still calling `WpfUiBootstrapper.Initialize()`/`SetLightTheme()` and referencing `Wpf.Ui.Appearance`) and any `.xaml` files still using `ExtendsContentIntoTitleBar="True"` (a `FluentWindow`-only DP) - both fixed in Tasks 3 and 6. Do not attempt to fix those errors here; just confirm no *other* unexpected errors appear (e.g. a window's code-behind directly calling a WPF-UI type would show up as a new, unexpected error - none were found during research, but this build catches it if research missed one).

---

## Task 3: Rewrite the theme bootstrapper — drop `Wpf.Ui.Appearance`, rename off the WPF-UI name

**Files:**
- Modify: `AIPowered/GLSense/GLSense.Addin.Core/Utilities/WpfUiBootstrapper.cs` → rename to `AppThemeBootstrapper.cs`
- Modify: `AIPowered/GLSense/GLSense.Addin.Core/Views/BaseWindow.cs` (no change needed - Task 2's rewrite already dropped the `WpfUiBootstrapper.Initialize()`/`SetLightTheme()` calls from the constructor entirely, since AIPowered's original bootstrap-on-every-window-construction pattern was itself a WPF-UI-era workaround; theming now happens once, at ribbon load, same as FinalWorkingCode's `MahAppsBootstrapper.Init`)
- Modify: `AIPowered/GLSense/GLSense.Addin.Core/AddinEntry.cs:137-139`
- Modify: `AIPowered/GLSense/GLSense.Addin.Core/GLSense.Addin.Core.csproj` (rename `<Compile Include="Utilities\WpfUiBootstrapper.cs" />` to `AppThemeBootstrapper.cs`)

`AppThemeBootstrapper` keeps steps 1+3 of the original (`AddRequiredResources`/`AddFallbackResources` - these already guarantee every resource key exists as a plain hand-authored `SolidColorBrush`, with zero WPF-UI involvement, confirmed during research) and drops step 2 (`LoadWpfUiFromPackUris`, which loads the real `Wpf.Ui.dll` theme dictionaries - the only step that requires the package) and the `SetDarkTheme`/`SetLightTheme` methods' `Wpf.Ui.Appearance.ApplicationThemeManager.Apply(...)` calls (the one call in this file with no fallback - it cannot be ported as-is since the type is gone).

- [ ] **Step 1: Create the renamed file with `LoadWpfUiFromPackUris` and the `ApplicationThemeManager` calls removed**

```csharp
using GLSense.Addin.Core.Infrastructure;
using System;
using System.Linq;
using System.Windows;
using System.Windows.Media;

namespace GLSense.Addin.Core.Utilities
{
    public static class AppThemeBootstrapper
    {
        private static bool _initialized;
        private static readonly object _lock = new object();

        public static void Initialize()
        {
            if (_initialized)
                return;

            lock (_lock)
            {
                if (_initialized)
                    return;

                try
                {
                    ServiceLocator.Logger?.LogDebug("AppThemeBootstrapper: Starting initialization...");

                    EnsureApplication();
                    LoadAllResourcesManually();

                    _initialized = true;
                    ServiceLocator.Logger?.LogDebug("AppThemeBootstrapper: Initialization completed successfully.");
                }
                catch (Exception ex)
                {
                    ServiceLocator.Logger?.LogException(ex, "AppThemeBootstrapper.Initialize: Initialization failed");
                    throw;
                }
            }
        }

        private static void EnsureApplication()
        {
            if (Application.Current == null)
            {
                ServiceLocator.Logger?.LogDebug("AppThemeBootstrapper: Creating new Application instance...");
                var app = new Application();

                app.DispatcherUnhandledException += (s, e) =>
                {
                    ServiceLocator.Logger?.LogException(e.Exception, "AppThemeBootstrapper: Unhandled WPF Dispatcher exception (suppressed, UI kept alive)");
                    e.Handled = true;
                };
            }
            else
            {
                ServiceLocator.Logger?.LogDebug("AppThemeBootstrapper: Using existing Application instance.");
            }
        }

        private static void LoadAllResourcesManually()
        {
            if (Application.Current == null)
                throw new InvalidOperationException("Application.Current is null");

            var app = Application.Current;
            var mergedDictionaries = app.Resources.MergedDictionaries;

            ServiceLocator.Logger?.LogDebug("AppThemeBootstrapper: Loading all resources manually...");

            var existingDicts = mergedDictionaries.ToList();
            foreach (var dict in existingDicts)
            {
                mergedDictionaries.Remove(dict);
                ServiceLocator.Logger?.LogDebug($"AppThemeBootstrapper: Removed existing dictionary: {dict.Source}");
            }

            AddRequiredResources(app);
            AddFallbackResources(app);

            ServiceLocator.Logger?.LogDebug("AppThemeBootstrapper: Resource loading complete.");
        }

        private static void AddRequiredResources(Application app)
        {
            ServiceLocator.Logger?.LogDebug("AppThemeBootstrapper: Adding required resources...");

            var accentColor = (Color)ColorConverter.ConvertFromString("#0078D7");

            var accentBrush = new SolidColorBrush(accentColor);
            var subtleBrush = new SolidColorBrush(Color.FromArgb(255, 240, 240, 240));
            var subtleSecondaryBrush = new SolidColorBrush(Color.FromArgb(255, 225, 225, 225));
            var solidAccentBrush = new SolidColorBrush(accentColor);

            accentBrush.Freeze();
            subtleBrush.Freeze();
            subtleSecondaryBrush.Freeze();
            solidAccentBrush.Freeze();

            var resources = app.Resources;

            AddResourceIfMissing(resources, "SystemAccentColor", accentColor);
            AddResourceIfMissing(resources, "SystemAccentColorPrimary", accentColor);
            AddResourceIfMissing(resources, "SystemAccentColorPrimaryBrush", accentBrush);
            AddResourceIfMissing(resources, "SystemAccentColorSecondary", accentColor);
            AddResourceIfMissing(resources, "SystemAccentColorTertiary", accentColor);

            AddResourceIfMissing(resources, "ControlBackgroundBrush", new SolidColorBrush(Color.FromArgb(255, 255, 255, 255)));
            AddResourceIfMissing(resources, "ControlSubtleBackgroundBrush", subtleBrush);
            AddResourceIfMissing(resources, "ControlSubtleSecondaryBrush", subtleSecondaryBrush);
            AddResourceIfMissing(resources, "ControlSolidAccentBrush", solidAccentBrush);
            AddResourceIfMissing(resources, "ControlTextBrush", new SolidColorBrush(Color.FromArgb(255, 0, 0, 0)));
            AddResourceIfMissing(resources, "ControlBorderBrush", new SolidColorBrush(Color.FromArgb(255, 200, 200, 200)));

            AddResourceIfMissing(resources, "CardBackgroundFillColorDefaultBrush", new SolidColorBrush(Color.FromArgb(255, 255, 255, 255)));
            AddResourceIfMissing(resources, "CardStrokeColorDefaultBrush", new SolidColorBrush(Color.FromArgb(255, 220, 220, 220)));

            AddResourceIfMissing(resources, "TextFillColorPrimaryBrush", new SolidColorBrush(Color.FromArgb(255, 0, 0, 0)));
            AddResourceIfMissing(resources, "TextFillColorSecondaryBrush", new SolidColorBrush(Color.FromArgb(255, 100, 100, 100)));

            AddResourceIfMissing(resources, "ApplicationBackgroundBrush", new SolidColorBrush(Color.FromArgb(255, 240, 240, 240)));

            ServiceLocator.Logger?.LogDebug("AppThemeBootstrapper: Required resources added.");
        }

        private static void AddResourceIfMissing(ResourceDictionary resources, string key, object value)
        {
            if (!resources.Contains(key))
            {
                resources[key] = value;
                ServiceLocator.Logger?.LogDebug($"AppThemeBootstrapper: Added resource '{key}'");
            }
        }

        private static void AddFallbackResources(Application app)
        {
            ServiceLocator.Logger?.LogDebug("AppThemeBootstrapper: Adding fallback resources...");

            var mergedDictionaries = app.Resources.MergedDictionaries;
            var fallbackDict = new ResourceDictionary();

            var keysToCheck = new[]
            {
                "ControlSubtleSecondaryBrush", "ControlSolidAccentBrush", "SystemAccentColorPrimary",
                "SystemAccentColorPrimaryBrush", "ControlBackgroundBrush", "ControlTextBrush",
                "ControlBorderBrush", "CardBackgroundFillColorDefaultBrush", "CardStrokeColorDefaultBrush",
                "TextFillColorPrimaryBrush", "TextFillColorSecondaryBrush", "ApplicationBackgroundBrush"
            };

            foreach (var key in keysToCheck)
            {
                bool exists = app.Resources.Contains(key) || mergedDictionaries.Any(d => d.Contains(key));
                if (exists)
                    continue;

                object value = key switch
                {
                    "ControlSubtleSecondaryBrush" => new SolidColorBrush(Color.FromArgb(255, 225, 225, 225)),
                    "ControlSolidAccentBrush" => new SolidColorBrush((Color)ColorConverter.ConvertFromString("#0078D7")),
                    "SystemAccentColorPrimary" => (Color)ColorConverter.ConvertFromString("#0078D7"),
                    "SystemAccentColorPrimaryBrush" => new SolidColorBrush((Color)ColorConverter.ConvertFromString("#0078D7")),
                    "ControlBackgroundBrush" => new SolidColorBrush(Color.FromArgb(255, 255, 255, 255)),
                    "ControlTextBrush" => new SolidColorBrush(Color.FromArgb(255, 0, 0, 0)),
                    "ControlBorderBrush" => new SolidColorBrush(Color.FromArgb(255, 200, 200, 200)),
                    "CardBackgroundFillColorDefaultBrush" => new SolidColorBrush(Color.FromArgb(255, 255, 255, 255)),
                    "CardStrokeColorDefaultBrush" => new SolidColorBrush(Color.FromArgb(255, 220, 220, 220)),
                    "TextFillColorPrimaryBrush" => new SolidColorBrush(Color.FromArgb(255, 0, 0, 0)),
                    "TextFillColorSecondaryBrush" => new SolidColorBrush(Color.FromArgb(255, 100, 100, 100)),
                    "ApplicationBackgroundBrush" => new SolidColorBrush(Color.FromArgb(255, 240, 240, 240)),
                    _ => null
                };

                if (value != null)
                {
                    fallbackDict[key] = value;
                    ServiceLocator.Logger?.LogDebug($"AppThemeBootstrapper: Added fallback resource '{key}'");
                }
            }

            if (fallbackDict.Count > 0)
            {
                mergedDictionaries.Add(fallbackDict);
                ServiceLocator.Logger?.LogDebug($"AppThemeBootstrapper: Fallback dictionary added with {fallbackDict.Count} resources");
            }
        }

        public static bool IsInitialized => _initialized;
    }
}
```

- [ ] **Step 2: Delete the old file and add the new one to the csproj**

```bash
git mv AIPowered/GLSense/GLSense.Addin.Core/Utilities/WpfUiBootstrapper.cs AIPowered/GLSense/GLSense.Addin.Core/Utilities/AppThemeBootstrapper.cs
```
Then in `GLSense.Addin.Core.csproj`, change:
```xml
<Compile Include="Utilities\WpfUiBootstrapper.cs" />
```
to:
```xml
<Compile Include="Utilities\AppThemeBootstrapper.cs" />
```

- [ ] **Step 3: Update the one call site in `AddinEntry.cs`**

In `AddinEntry.cs`, replace lines 135-139:
```csharp
                // Ensure WPF Application exists - ONLY CALL THIS ONCE
                WpfAppManager.EnsureApplication();
                //Implementing WPF-UI Bootstrapper
                WpfUiBootstrapper.Initialize();
                WpfUiBootstrapper.SetLightTheme();
```
with:
```csharp
                // Ensure WPF Application exists - ONLY CALL THIS ONCE
                WpfAppManager.EnsureApplication();
                AppThemeBootstrapper.Initialize();
```

- [ ] **Step 4: Build and confirm no remaining `Wpf.Ui` references in code**

```bash
grep -rn "Wpf\.Ui\|WpfUiBootstrapper" AIPowered/GLSense/GLSense.Addin.Core --include=*.cs
msbuild AIPowered/GLSense/GLSense.sln /t:GLSense_Addin_Core /p:Configuration=Debug
```
Expected: grep returns nothing; build still fails only on the `ExtendsContentIntoTitleBar` XAML attributes (fixed in Task 6) and the `Wpf.Ui`/`Wpf.Ui.Abstractions` package references still being present in the csproj (removed in Task 4).

---

## Task 4: Remove `Wpf.Ui`/`Wpf.Ui.Abstractions` package references

**Files:**
- Modify: `AIPowered/GLSense/GLSense.Addin.Core/GLSense.Addin.Core.csproj`
- Modify: `AIPowered/GLSense/GLSense.Addin.Core/packages.config`

- [ ] **Step 1: Remove the two `<Reference>` entries from the csproj**

Delete these two entries (verified at approximately lines 113-119 during research):
```xml
    <Reference Include="Wpf.Ui, Version=4.3.0.0, Culture=neutral, PublicKeyToken=11f9f5cc97b3ffd6, processorArchitecture=MSIL">
      <HintPath>..\packages\WPF-UI.4.3.0\lib\net481\Wpf.Ui.dll</HintPath>
    </Reference>
    <Reference Include="Wpf.Ui.Abstractions, Version=4.3.0.0, Culture=neutral, PublicKeyToken=11f9f5cc97b3ffd6, processorArchitecture=MSIL">
      <HintPath>..\packages\WPF-UI.Abstractions.4.3.0\lib\net481\Wpf.Ui.Abstractions.dll</HintPath>
    </Reference>
```

- [ ] **Step 2: Remove the two `<package>` entries from `packages.config`**

Delete the `WPF-UI` and `WPF-UI.Abstractions` (both version `4.3.0`) entries. Leave every other package (`MahApps.Metro.IconPacks.Core`/`FontAwesome`, SQLite, WebView2, etc.) untouched.

- [ ] **Step 3: Check sibling projects don't also reference these packages before physically deleting the NuGet package folders**

```bash
grep -rln "Wpf.Ui" AIPowered/GLSense/*/*.csproj AIPowered/GLSense/*/packages.config 2>/dev/null
```
Expected: only `GLSense.Addin.Core.csproj`/`packages.config` (now edited) match. If any other project (`GLSense`, `GLSense.Contracts`, `GLSense.Loader.Core`, `GLSense.Shared`, `GLSense.LocalUpdateHost`) also references them, stop and handle that project too before proceeding - do not delete a package still in use elsewhere.

- [ ] **Step 4: Delete the on-disk package folders (only after Step 3 confirms nothing else references them)**

```bash
rm -rf "AIPowered/GLSense/packages/WPF-UI.4.3.0" "AIPowered/GLSense/packages/WPF-UI.Abstractions.4.3.0"
```

- [ ] **Step 5: Build and confirm the package-reference errors are gone**

```bash
msbuild AIPowered/GLSense/GLSense.sln /t:GLSense_Addin_Core /p:Configuration=Debug
```
Expected: remaining errors are only the `ExtendsContentIntoTitleBar`/`WindowStyle="SingleBorderWindow"` XAML attributes (Task 6) and the two new utility files not yet added (Tasks 5 and would show as "type not found" in Task 7's edits, not before).

---

## Task 5: Port `WebView2Warmup.cs`

**Files:**
- Create: `AIPowered/GLSense/GLSense.Addin.Core/Utilities/WebView2Warmup.cs`
- Modify: `AIPowered/GLSense/GLSense.Addin.Core/GLSense.Addin.Core.csproj`

**Interfaces:**
- Produces: `WebView2Warmup.WarmUpInBackground()` (void), `WebView2Warmup.GetEnvironmentAsync()` (`Task<CoreWebView2Environment>`) - consumed by Task 7 (wiring), Task 8 (`GLLogin.xaml.cs`), Task 9 (`GLDrilldownCustomization.xaml.cs`).

- [ ] **Step 1: Create the file**

```csharp
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
```

- [ ] **Step 2: Add to the csproj**

Add next to the other `Utilities\*.cs` entries:
```xml
    <Compile Include="Utilities\WebView2Warmup.cs" />
```

- [ ] **Step 3: Build**

```bash
msbuild AIPowered/GLSense/GLSense.sln /t:GLSense_Addin_Core /p:Configuration=Debug
```
Expected: no new errors from this file (`ServiceLocator.Paths.LoginBrowserPath` already exists - confirmed in use by the current `GLLogin.xaml.cs`/`GLDrilldownCustomization.xaml.cs`).

---

## Task 6: Port `WindowLoadingPlaceholder.cs`

**Files:**
- Create: `AIPowered/GLSense/GLSense.Addin.Core/Utilities/WindowLoadingPlaceholder.cs`
- Modify: `AIPowered/GLSense/GLSense.Addin.Core/GLSense.Addin.Core.csproj`

**Interfaces:**
- Consumes: `AppConstants.GLAccentHex` (already exists in AIPowered's `AppConstants.cs:40` = `"#2E86AB"`).
- Produces: `WindowLoadingPlaceholder.WarmUpInBackground()`, `WindowLoadingPlaceholder.ShowMatching(double, double, double, double, IntPtr)` (returns `int` generation), `WindowLoadingPlaceholder.Hide(int)` - all consumed by `BaseWindow.cs` (already wired in Task 2's rewrite) and Task 7 (warm-up call).

- [ ] **Step 1: Create the file**

```csharp
using GLSense.Addin.Core.Infrastructure;
using System;
using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using System.Windows.Media.Animation;
using System.Windows.Shapes;
using System.Windows.Threading;

namespace GLSense.Addin.Core.Utilities
{
    // Shows one small, reused loading-indicator window sized/positioned to match where the
    // real window will appear, and hides it once the real window's own ContentRendered
    // fires. Reusing a single instance (Hide(), never Close()) pays its own one-time
    // per-instance HWND-creation/first-paint cost once, off-screen, at AppDomain init -
    // every later ShowMatching() call just toggles visibility (and resizes/repositions) an
    // HWND that already exists and has already painted. Ported verbatim from
    // FinalWorkingCode's Utilities\WindowLoadingPlaceholder.cs - see that file's header
    // comment for the full investigation history behind this approach.
    public static class WindowLoadingPlaceholder
    {
        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool GetWindowRect(IntPtr hWnd, out RECT lpRect);

        [DllImport("user32.dll")]
        private static extern uint GetDpiForWindow(IntPtr hwnd);

        [StructLayout(LayoutKind.Sequential)]
        private struct RECT
        {
            public int Left;
            public int Top;
            public int Right;
            public int Bottom;
        }

        private static readonly object _lock = new object();
        private static Window _window;
        private static int _generation;
        private static DispatcherTimer _safetyTimer;

        public static void WarmUpInBackground()
        {
            var app = Application.Current;
            if (app == null)
            {
                ServiceLocator.Logger?.LogWarn("WindowLoadingPlaceholder.WarmUpInBackground: no Application.Current yet, skipping.");
                return;
            }

            try
            {
                app.Dispatcher.BeginInvoke(new Action(() =>
                {
                    try
                    {
                        EnsureCreated();
                        ServiceLocator.Logger?.LogDebug("WindowLoadingPlaceholder: warmed up (paid its own first-paint cost off-screen).");
                    }
                    catch (Exception ex)
                    {
                        ServiceLocator.Logger?.LogException(ex, "WindowLoadingPlaceholder.WarmUpInBackground (inner)");
                    }
                }), DispatcherPriority.ApplicationIdle);
            }
            catch (Exception ex)
            {
                ServiceLocator.Logger?.LogException(ex, "WindowLoadingPlaceholder.WarmUpInBackground");
            }
        }

        public static int ShowMatching(double left, double top, double width, double height, IntPtr excelOwnerHwnd)
        {
            int myGeneration;
            try
            {
                EnsureCreated();

                bool hasTarget = !double.IsNaN(left) && !double.IsNaN(top) &&
                                 !double.IsNaN(width) && !double.IsNaN(height) &&
                                 !double.IsInfinity(width) && !double.IsInfinity(height) &&
                                 width > 0 && height > 0;

                if (hasTarget)
                {
                    _window.SizeToContent = SizeToContent.Manual;
                    _window.Width = width;
                    _window.Height = height;
                    _window.Left = left;
                    _window.Top = top;
                }
                else
                {
                    _window.SizeToContent = SizeToContent.WidthAndHeight;
                    _window.UpdateLayout();
                    PositionGenericNear(excelOwnerHwnd);
                }

                lock (_lock)
                {
                    myGeneration = ++_generation;
                }

                _window.Show();

                _safetyTimer?.Stop();
                _safetyTimer = new DispatcherTimer { Interval = TimeSpan.FromMilliseconds(3000) };
                _safetyTimer.Tick += (s, e) => Hide(myGeneration);
                _safetyTimer.Start();
            }
            catch (Exception ex)
            {
                ServiceLocator.Logger?.LogException(ex, "WindowLoadingPlaceholder.ShowMatching");
                myGeneration = -1;
            }

            return myGeneration;
        }

        public static void Hide(int generation)
        {
            try
            {
                lock (_lock)
                {
                    if (generation != _generation)
                        return;
                }

                _safetyTimer?.Stop();
                _safetyTimer = null;
                _window?.Hide();
            }
            catch (Exception ex)
            {
                ServiceLocator.Logger?.LogException(ex, "WindowLoadingPlaceholder.Hide");
            }
        }

        private static void EnsureCreated()
        {
            if (_window != null)
                return;

            var ring = new Ellipse
            {
                Width = 48,
                Height = 48,
                Stroke = new SolidColorBrush((Color)ColorConverter.ConvertFromString(AppConstants.GLAccentHex)),
                StrokeThickness = 5,
                StrokeDashArray = new DoubleCollection { 2.4, 1.6 },
                RenderTransformOrigin = new Point(0.5, 0.5),
                RenderTransform = new RotateTransform(0)
            };

            var spin = new DoubleAnimation(0, 360, TimeSpan.FromSeconds(1))
            {
                RepeatBehavior = RepeatBehavior.Forever
            };
            ring.RenderTransform.BeginAnimation(RotateTransform.AngleProperty, spin);

            var text = new TextBlock
            {
                Text = "Loading...",
                Margin = new Thickness(0, 16, 0, 0),
                HorizontalAlignment = HorizontalAlignment.Center,
                FontSize = 18,
                Foreground = new SolidColorBrush(Color.FromRgb(0x21, 0x25, 0x29))
            };

            var panel = new StackPanel
            {
                Orientation = Orientation.Vertical,
                HorizontalAlignment = HorizontalAlignment.Center,
                VerticalAlignment = VerticalAlignment.Center
            };
            panel.Children.Add(ring);
            panel.Children.Add(text);

            var border = new Border
            {
                Background = Brushes.White,
                BorderBrush = new SolidColorBrush(Color.FromRgb(0xD1, 0xD5, 0xDB)),
                BorderThickness = new Thickness(1),
                CornerRadius = new CornerRadius(10),
                Padding = new Thickness(28),
                Child = panel
            };

            _window = new Window
            {
                Content = border,
                SizeToContent = SizeToContent.WidthAndHeight,
                WindowStyle = WindowStyle.None,
                AllowsTransparency = false,
                ResizeMode = ResizeMode.NoResize,
                ShowInTaskbar = false,
                ShowActivated = false,
                Topmost = true,
                WindowStartupLocation = WindowStartupLocation.Manual,
                Left = -32000,
                Top = -32000
            };

            _window.Show();
            _window.Hide();
        }

        private static void PositionGenericNear(IntPtr excelOwnerHwnd)
        {
            try
            {
                double centerX, centerY;

                if (excelOwnerHwnd != IntPtr.Zero && GetWindowRect(excelOwnerHwnd, out RECT rect) &&
                    rect.Right > rect.Left && rect.Bottom > rect.Top)
                {
                    double scale = 1.0;
                    try
                    {
                        uint dpi = GetDpiForWindow(excelOwnerHwnd);
                        if (dpi > 0) scale = dpi / 96.0;
                    }
                    catch { /* fall back to 1.0 */ }

                    centerX = ((rect.Left + rect.Right) / 2.0) / scale;
                    centerY = ((rect.Top + rect.Bottom) / 2.0) / scale;
                }
                else
                {
                    var workArea = SystemParameters.WorkArea;
                    centerX = workArea.Left + (workArea.Width / 2.0);
                    centerY = workArea.Top + (workArea.Height / 2.0);
                }

                _window.Left = centerX - (_window.ActualWidth / 2.0);
                _window.Top = centerY - (_window.ActualHeight / 2.0);
            }
            catch (Exception ex)
            {
                ServiceLocator.Logger?.LogException(ex, "WindowLoadingPlaceholder.PositionGenericNear");
            }
        }
    }
}
```

- [ ] **Step 2: Add to the csproj**

```xml
    <Compile Include="Utilities\WindowLoadingPlaceholder.cs" />
```

- [ ] **Step 3: Build**

```bash
msbuild AIPowered/GLSense/GLSense.sln /t:GLSense_Addin_Core /p:Configuration=Debug
```
Expected: `BaseWindow.cs`'s reference to `WindowLoadingPlaceholder.ShowMatching`/`.Hide` (added in Task 2) now resolves. Remaining errors: only the `ExtendsContentIntoTitleBar` XAML attributes (Task 7).

---

## Task 7: Wire warm-up calls + remove `ExtendsContentIntoTitleBar`/switch to `WindowStyle="None"` across all 26 windows

**Files:**
- Modify: `AIPowered/GLSense/GLSense.Addin.Core/AddinEntry.cs:135-139`
- Modify all 26 files below (mechanical attribute change, one line each)

**Interfaces:** none new.

- [ ] **Step 1: Add the warm-up calls to `AddinEntry.Initialize`**

Immediately after the `AppThemeBootstrapper.Initialize();` line from Task 3 Step 3, add:
```csharp
                WebView2Warmup.WarmUpInBackground();
                WindowLoadingPlaceholder.WarmUpInBackground();
```
So the block reads:
```csharp
                // Ensure WPF Application exists - ONLY CALL THIS ONCE
                WpfAppManager.EnsureApplication();
                AppThemeBootstrapper.Initialize();
                WebView2Warmup.WarmUpInBackground();
                WindowLoadingPlaceholder.WarmUpInBackground();
```
Add `using GLSense.Addin.Core.Utilities;` to `AddinEntry.cs` if not already present (it is - `AddinEntry.cs:9` already has this using per the file already read during research).

- [ ] **Step 2: For each of the 26 files below, remove the `ExtendsContentIntoTitleBar="True"` attribute and change `WindowStyle="SingleBorderWindow"` to `WindowStyle="None"`**

Files (all under `AIPowered/GLSense/GLSense.Addin.Core/Views/`):
`GLUserConfig.xaml`, `GLServerConfiguration.xaml`, `GLSegmentValues.xaml`, `GLSegmentRef.xaml`, `GLSegmentManager.xaml`, `GLSegmentFunctions.xaml`, `GLSegmentDiscovery.xaml`, `GLRollerGroups.xaml`, `GLLoginDetails.xaml`, `GLLogin.xaml`, `GLLOVs.xaml`, `GLJobsMonitor.xaml`, `GLGetPeriodStartEnd.xaml`, `GLGetPeriodDetails.xaml`, `GLGetPeriodByYear.xaml`, `GLGetPeriodByDate.xaml`, `GLGetPeriod.xaml`, `GLExpandOptions.xaml`, `GLDrilldownCustomization.xaml`, `GLCubeDetails.xaml`, `GLAbout.xaml`, `AttachmentsDialog.xaml`, `WebView2PopupWindow.xaml`, `GLWaitWindow.xaml`, `GLDailyRates.xaml`, `GLMessageWindow.xaml`.

In each file, change (exact text confirmed present in this form via `GLCubeDetails.xaml`/`GLUserConfig.xaml`/`GLLOVs.xaml` during research; other 23 files use the identical two-attribute pair per the earlier grep sweep):
```xml
                  WindowStyle="SingleBorderWindow"
                  ExtendsContentIntoTitleBar="True"
```
to:
```xml
                  WindowStyle="None"
```
(i.e. delete the `ExtendsContentIntoTitleBar` line entirely and change `SingleBorderWindow` to `None` on the remaining line - exact surrounding whitespace/attribute order may differ slightly per file; preserve each file's own formatting style, only these two attributes change.)

Verify after editing all 26:
```bash
grep -rln "ExtendsContentIntoTitleBar" AIPowered/GLSense/GLSense.Addin.Core/Views/*.xaml
grep -rln "WindowStyle=\"SingleBorderWindow\"" AIPowered/GLSense/GLSense.Addin.Core/Views/*.xaml
```
Expected: both commands return no files.

- [ ] **Step 3: Build the full solution**

```bash
msbuild AIPowered/GLSense/GLSense.sln /p:Configuration=Debug
```
Expected: clean build, zero errors. If any window still fails to compile, it means that file has an additional WPF-UI-specific usage not caught during research (e.g. a `ui:SymbolIcon` somewhere) - fix it by replacing with the equivalent `iconPacks:PackIconFontAwesome` (see `PORTING_GUIDE.md` section 6 for the icon-name-mapping caveat) rather than reintroducing the package.

- [ ] **Step 4: Manual smoke test - open every one of the 26 windows once**

Launch Excel with the add-in loaded (Debug build) and trigger every ribbon action that opens a `BaseWindow`-derived window (Login, Cube Details, User Config, LOVs, Roller Groups, Segment Values/Ref/Manager/Functions/Discovery, Daily Rates, Get Period family, Jobs Monitor, About, Attachments, Drilldown Customization, Server Configuration, Wait/Message/Expand-Options/WebView2Popup as they appear during normal flows, Login Details). Confirm for each: window opens without a WPF exception, has no native OS title bar, drag-by-title-bar still works, Escape still closes it (where applicable), Close button still works. This is the regression check for the base-class swap.

---

## Task 8: Port GLLogin.xaml.cs WebView2 warm-up + busy overlay

**Files:**
- Modify: `AIPowered/GLSense/GLSense.Addin.Core/Views/GLLogin.xaml.cs`

**Interfaces:**
- Consumes: `WebView2Warmup.GetEnvironmentAsync()` (Task 5).

- [ ] **Step 1: Replace the `WebView_Loaded` method**

Replace the current method (lines 222-277) with:
```csharp
        private async void WebView_Loaded(object sender, RoutedEventArgs e)
        {
            webView.Loaded -= WebView_Loaded; // prevent double init

            // CoreWebView2Environment creation spins up a real Chromium process tree and
            // can take several seconds cold; previously nothing was shown while it ran, so
            // this whole area just sat blank. The environment is now shared/pre-warmed via
            // WebView2Warmup (kicked off at AddinEntry.Initialize), so this mostly just
            // awaits an already-in-flight or already-completed task. Ported from
            // FinalWorkingCode's identical GLLogin.xaml.cs fix.
            webView.Visibility = Visibility.Collapsed;
            AppOverlayControl.ShowBusyasyn("Initializing browser component...");
            try
            {
                var env = await WebView2Warmup.GetEnvironmentAsync();

                // Initialize WebView2 with that environment
                _webViewInitTask = webView.EnsureCoreWebView2Async(env);
                await _webViewInitTask;

                // Hook device permission handler and diagnostics after CoreWebView2 is ready
                webView.CoreWebView2.PermissionRequested += CoreWebView2_PermissionRequested;
                webView.CoreWebView2.NavigationCompleted += WebView_NavigationCompleted;

                // Cert-error bypass (scoped to the customer's own configured server) and
                // retry-once navigation for SSO/SAML/OIDC redirects. Ported from
                // FinalWorkingCode's WebView2NavigationResilience (popup-hosting piece
                // excluded - see that class's header comment).
                _resilience = new WebView2NavigationResilience(nameof(GLLogin), this);
                _resilience.Attach(webView.CoreWebView2, GetTrustedHosts);

                // Optional: turn on DevTools during development
                webView.CoreWebView2.Settings.AreDevToolsEnabled = true;

                // Diagnostics: log WebView2 runtime
                var version = webView.CoreWebView2.Environment.BrowserVersionString;
                ServiceLocator.Logger?.LogDebug($"WebView2 BrowserVersion={version}");
            }
            catch (Exception ex)
            {
                ServiceLocator.Logger?.LogException(ex, "WebView2 initialization failed in GLLogin");
                Close();
            }
            finally
            {
                await AppOverlayControl.HideBusyAsync();
                webView.Visibility = Visibility.Visible;
            }
        }
```

Add `using GLSense.Addin.Core.Utilities;` if not already present (it already is - `GLLogin.xaml.cs:40`).

- [ ] **Step 2: Build**

```bash
msbuild AIPowered/GLSense/GLSense.sln /t:GLSense_Addin_Core /p:Configuration=Debug
```

- [ ] **Step 3: Manual smoke test**

Open GLLogin from the ribbon. Confirm: a busy overlay with "Initializing browser component..." shows immediately (instead of a blank WebView2 area), then the login page loads and the overlay hides. Confirm login flow (server select → navigate → cookie extraction → cube fetch → close) still works end to end.

---

## Task 9: Port GLDrilldownCustomization.xaml.cs WebView2 warm-up + busy overlay

**Files:**
- Modify: `AIPowered/GLSense/GLSense.Addin.Core/Views/GLDrilldownCustomization.xaml.cs`

**Interfaces:**
- Consumes: `WebView2Warmup.GetEnvironmentAsync()` (Task 5), `ShowBusyOverlayAsync` (already exists in this file, unchanged).

- [ ] **Step 1: Replace the `WebView_Loaded` method**

Replace the current method (lines 136-225) with:
```csharp
        private async void WebView_Loaded(object sender, RoutedEventArgs e)
        {
            ServiceLocator.Logger?.LogDebug("GLDrilldownCustomization.WebView_Loaded invoked");
            webView.Loaded -= WebView_Loaded; // prevent double init

            // CoreWebView2Environment creation spins up a real Chromium process tree and can
            // take several seconds cold; previously nothing was shown while it ran here
            // either (same gap as GLLogin). The environment is shared/pre-warmed via
            // WebView2Warmup (kicked off at AddinEntry.Initialize), and both this window and
            // GLLogin already point at the exact same ServiceLocator.Paths.LoginBrowserPath
            // user data folder, so sharing one instance is also strictly less wasteful than
            // each window spinning up its own. Ported from FinalWorkingCode's identical fix.
            using var initCancellation = new CancellationHelper();
            await ShowBusyOverlayAsync(initCancellation, "Initializing browser component...");
            try
            {
                var env = await WebView2Warmup.GetEnvironmentAsync();

                // Initialize WebView2 with that environment
                _webViewInitTask = webView.EnsureCoreWebView2Async(env);
                await _webViewInitTask;

                // Hook device permission handler and diagnostics after CoreWebView2 is ready
                webView.CoreWebView2.PermissionRequested += CoreWebView2_PermissionRequested;

                // Cert-error bypass (scoped to this window's own server) and retry-once
                // navigation. Ported from FinalWorkingCode's WebView2NavigationResilience
                // (popup-hosting piece excluded - see that class's header comment).
                _resilience = new WebView2NavigationResilience(nameof(GLDrilldownCustomization), this);
                _resilience.Attach(webView.CoreWebView2, GetTrustedHosts);

                // Optional: turn on DevTools during development
                webView.CoreWebView2.Settings.AreDevToolsEnabled = true;

                // Diagnostics: log WebView2 runtime
                var version = webView.CoreWebView2.Environment.BrowserVersionString;
                ServiceLocator.Logger?.LogDebug($"WebView2 BrowserVersion={version}");

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
                    await AppOverlayControl.HideBusyAsync();
                    webView.Visibility = Visibility.Visible;
                }
            }
            catch (Exception ex)
            {
                ServiceLocator.Logger?.LogException(ex, "WebView2 initialization failed in GLDrilldownCustomization");
                await AppOverlayControl.HideBusyAsync();
                webView.Visibility = Visibility.Visible;
            }
        }
```

Add `using GLSense.Addin.Core.Utilities;` if not already present (it already is - `GLDrilldownCustomization.xaml.cs:45`).

- [ ] **Step 2: Build**

```bash
msbuild AIPowered/GLSense/GLSense.sln /t:GLSense_Addin_Core /p:Configuration=Debug
```

- [ ] **Step 3: Manual smoke test**

Open GLDrilldownCustomization from the ribbon with a cube selected. Confirm the busy overlay shows "Initializing browser component..." immediately, then "Loading Drilldown Customization" once environment init completes, then the drilldown page renders and the overlay hides. Confirm the "no cube selected" path (open without a selected cube) still shows the WebView area (blank, no stuck overlay) instead of hanging.

---

## Task 10: Port GLLOVs prepare-before-show

**Files:**
- Modify: `AIPowered/GLSense/GLSense.Addin.Core/Views/GLLOVs.xaml` (remove `Loaded="Window_Loaded"`)
- Modify: `AIPowered/GLSense/GLSense.Addin.Core/Views/GLLOVs.xaml.cs` (`Window_Loaded` → `PrepareAsync`)
- Modify: `AIPowered/GLSense/GLSense.Addin.Core/ViewModels/GLLovViewModel.cs` (`LoadDataAsync` awaits `LoadLovRowsAsync` explicitly)
- Modify: `AIPowered/GLSense/GLSense.Addin.Core/AddinEntry.cs` (new `ShowLovsWindow` method + updated `case "ShowLOVs"`)

**Interfaces:**
- Produces: `GLLOVs.PrepareAsync()` (public `Task`, replaces the private `Window_Loaded` event handler) - consumed by the new `AddinEntry.ShowLovsWindow()`.

- [ ] **Step 1: `GLLOVs.xaml` - remove the `Loaded` wiring**

Change:
```xml
                  WindowStyle="None"
                  Loaded="Window_Loaded">
```
(after Task 7 has already changed `WindowStyle` to `"None"`) to:
```xml
                  WindowStyle="None">
```

- [ ] **Step 2: `GLLOVs.xaml.cs` - rename `Window_Loaded` to public `PrepareAsync`**

Replace:
```csharp
        private async void Window_Loaded(object sender, RoutedEventArgs e)
        {
            ServiceLocator.Logger?.LogDebug("GLLOVs.Window_Loaded invoked");
            try
            {
                Excel.Range rng = ServiceLocator.ExcelApp.ActiveCell;
                string sheetName = ((Excel.Worksheet)rng.Parent).Name;
                string cellAddress = rng.Address[true, true, Excel.XlReferenceStyle.xlA1, false];
                string addr = $"'{sheetName}'!{cellAddress}";

                GlobalStateViewModel.Instance.ReferenceText = addr;

                if (AppState.Instance.SelectedCube != null && AppState.Instance.SelectedLedger != null)
                {
                    ServiceLocator.Logger?.LogDebug($"GLLOVs.Window_Loaded: loading data for cubeId={AppState.Instance.SelectedCube.CubeId}, ledgerId={AppState.Instance.SelectedLedger.LedgerId}");
                    await vm.LoadDataAsync(AppState.Instance.SelectedCube.CubeId, AppState.Instance.SelectedLedger.LedgerId);

                    await Dispatcher.InvokeAsync(() =>
                    {
                        cmbLedgers.Text = vm.LOV_SelectedLedger.LedgerName;
                    });
                }
            }
            catch (Exception ex)
            {
                ServiceLocator.Logger?.LogException(ex, "GLLOVs.Window_Loaded");
            }
        }
```
with:
```csharp
        /// <summary>
        /// Loads the active-cell reference and the LOV grid data. Called by
        /// AddinEntry.ShowLovsWindow and awaited *before* ShowDialog() - not wired to the
        /// Loaded event like most other windows - so the window's very first frame already
        /// has the real content in it instead of appearing blank and filling in a moment
        /// later. Ported from FinalWorkingCode's identical GLLOVs.xaml.cs fix.
        /// </summary>
        public async Task PrepareAsync()
        {
            ServiceLocator.Logger?.LogDebug("GLLOVs.PrepareAsync invoked");
            try
            {
                Excel.Range rng = ServiceLocator.ExcelApp.ActiveCell;
                string sheetName = ((Excel.Worksheet)rng.Parent).Name;
                string cellAddress = rng.Address[true, true, Excel.XlReferenceStyle.xlA1, false];
                string addr = $"'{sheetName}'!{cellAddress}";

                GlobalStateViewModel.Instance.ReferenceText = addr;

                if (AppState.Instance.SelectedCube != null && AppState.Instance.SelectedLedger != null)
                {
                    ServiceLocator.Logger?.LogDebug($"GLLOVs.PrepareAsync: loading data for cubeId={AppState.Instance.SelectedCube.CubeId}, ledgerId={AppState.Instance.SelectedLedger.LedgerId}");
                    await vm.LoadDataAsync(AppState.Instance.SelectedCube.CubeId, AppState.Instance.SelectedLedger.LedgerId);

                    await Dispatcher.InvokeAsync(() =>
                    {
                        cmbLedgers.Text = vm.LOV_SelectedLedger.LedgerName;
                    });
                }
            }
            catch (Exception ex)
            {
                ServiceLocator.Logger?.LogException(ex, "GLLOVs.PrepareAsync");
            }
        }
```

- [ ] **Step 3: `GLLovViewModel.cs` - set the backing field directly and await `LoadLovRowsAsync` explicitly**

In `LoadDataAsync` (lines 118-160), replace the `_dispatcher.InvokeAsync` block's ledger-selection logic and add the explicit await after it:
```csharp
                await _dispatcher.InvokeAsync(() =>
                {
                    Ledgers.Clear();
                    foreach (var l in allLedgers)
                    {
                        Ledgers.Add(l);
                    }

                    // Sets the backing field directly (not the LOV_SelectedLedger property)
                    // specifically to skip its setter's LoadLovRows() call - that call is
                    // fire-and-forget (Task.Run, never awaited by its caller), which meant
                    // this method previously returned as soon as the ledger dropdown was
                    // populated while the actual grid content kept loading in the background
                    // *after* PrepareAsync had already gone on to call ShowDialog. Explicitly
                    // awaiting LoadLovRowsAsync() below instead makes this method genuinely
                    // block until the grid has real data. A later user-driven ledger change
                    // (window already open) still goes through the normal property setter
                    // further down in this class, so that fire-and-forget-with-busy-overlay
                    // UX is unaffected. Ported from FinalWorkingCode's identical fix.
                    if (AppState.Instance.SelectedLedger != null)
                    {
                        _lOV_SelectedLedger = AppState.Instance.SelectedLedger;
                        ServiceLocator.Logger?.LogDebug($"GLLovViewModel.LoadDataAsync: using AppState.SelectedLedger \"{LOV_SelectedLedger?.LedgerName}\" (LedgerId={LOV_SelectedLedger?.LedgerId}).");
                    }
                    else
                    {
                        _lOV_SelectedLedger = Ledgers.FirstOrDefault(x => x.LedgerId == defaultLedgerId);
                        if (LOV_SelectedLedger == null)
                            ServiceLocator.Logger?.LogWarn($"GLLovViewModel.LoadDataAsync: no ledger found matching defaultLedgerId={defaultLedgerId} among {Ledgers.Count} loaded ledger(s).");
                        else
                            ServiceLocator.Logger?.LogDebug($"GLLovViewModel.LoadDataAsync: defaulted to ledger \"{LOV_SelectedLedger.LedgerName}\" (LedgerId={LOV_SelectedLedger.LedgerId}).");
                    }
                    OnPropertyChanged(nameof(LOV_SelectedLedger));
                });

                await LoadLovRowsAsync();

                ServiceLocator.Logger?.LogDebug("GLLovViewModel.LoadDataAsync: completed.");
```
(This replaces everything from `await _dispatcher.InvokeAsync(() =>` through the pre-existing `ServiceLocator.Logger?.LogDebug("GLLovViewModel.LoadDataAsync: completed.");` line - the `try`/`catch` wrapper around it is unchanged.)

- [ ] **Step 4: `AddinEntry.cs` - dedicated show-method for LOVs instead of the shared `ShowGroupCWindow`**

Replace:
```csharp
                case "ShowLOVs":
                    ShowGroupCWindow("ShowLOVs", () => new GLLOVs());
                    break;
```
with:
```csharp
                case "ShowLOVs":
                    ShowLovsWindow();
                    break;
```

Add a new method next to `ShowGroupCWindow` (after it, same class):
```csharp
        /// <summary>
        /// GLLOVs needs its own show-method (not the shared ShowGroupCWindow) because its
        /// data load must be awaited BEFORE ShowDialog() - see GLLOVs.PrepareAsync's own
        /// comment for why - unlike every other Group C/H window, which loads its data from
        /// its own Loaded event after the window is already shown. Ported from
        /// FinalWorkingCode's AddinModule.RibLOVs_OnClick.
        /// </summary>
        private void ShowLovsWindow()
        {
            const string actionLabel = "ShowLOVs";
            try
            {
                ServiceLocator.Logger?.LogDebug($"{actionLabel}: Opening window...");

                WpfAppManager.InvokeOnWpfThread(async () =>
                {
                    try
                    {
                        var win = new GLLOVs();
                        win.CenterInExcel = true;
                        win.ModalToExcel = true;
                        win.ShowInTaskbar = false;

                        // Awaited before ShowDialog (not on the Loaded event like most other
                        // windows) so the window's first frame already has the LOV grid data
                        // in it - see GLLOVs.PrepareAsync for why.
                        await win.PrepareAsync();

                        // PrepareAsync's internal awaits (LoadDataAsync -> repository/SQLite
                        // calls) can resume on a different thread than the one that
                        // constructed win. ShowDialog touches win's HWND (a DispatcherObject),
                        // so it must run back on win's own Dispatcher thread explicitly.
                        win.Dispatcher.Invoke(() => win.ShowDialog());

                        ServiceLocator.Logger?.LogDebug($"{actionLabel}: Dialog closed.");
                    }
                    catch (Exception ex)
                    {
                        ServiceLocator.Logger?.LogException(ex, $"{actionLabel}: ShowDialog error");
                    }
                });
            }
            catch (Exception ex)
            {
                ServiceLocator.Logger?.LogException(ex, $"{actionLabel}: Error");
            }
        }
```

Check `WpfAppManager.InvokeOnWpfThread`'s signature accepts an `async ()` lambda (i.e. takes `Action`, and C# allows an `async void`-shaped lambda to satisfy an `Action` parameter) - confirm by reading `Utilities/WpfAppManager.cs`'s `InvokeOnWpfThread` signature before this step; if it only accepts `Action` and internally does something that doesn't tolerate the lambda completing before the async body finishes (e.g. a synchronization construct that assumes synchronous completion), wrap using `Dispatcher.InvokeAsync` directly against the WPF dispatcher instead - do not silently drop the `await win.PrepareAsync()` ordering requirement.

- [ ] **Step 5: Build**

```bash
msbuild AIPowered/GLSense/GLSense.sln /t:GLSense_Addin_Core /p:Configuration=Debug
```

- [ ] **Step 6: Manual smoke test**

Open GLLOVs from the ribbon (RibLOVs) with a cube+ledger selected. Confirm: the window's first visible frame already shows the ledger combo populated and the LOV grid with real rows (not an empty grid that fills in after a visible delay). Confirm the "no cube/ledger selected" path still opens the window without throwing (grid stays empty, no exception). Confirm changing the ledger dropdown after the window is open still reloads the grid with a busy overlay (this path goes through the unchanged `LOV_SelectedLedger` property setter, not `PrepareAsync`).

---

## Task 11: Port GLCubeDetails.xaml sizing fix

**Files:**
- Modify: `AIPowered/GLSense/GLSense.Addin.Core/Views/GLCubeDetails.xaml`

- [ ] **Step 1: Change the size bounds**

Change:
```xml
                  MinWidth="680" MaxWidth="750"
```
to:
```xml
                  MinWidth="800" MaxWidth="950"
```
(`MinHeight="500" MaxHeight="650"` is unchanged - FinalWorkingCode's fix only widened this window, matching its own `Views\GLCubeDetails.xaml` diff.)

- [ ] **Step 2: Build and smoke-test**

```bash
msbuild AIPowered/GLSense/GLSense.sln /t:GLSense_Addin_Core /p:Configuration=Debug
```
Open GLCubeDetails from the ribbon. Confirm the window is noticeably wider (fits the Ledger Name/Last Refreshed Date/Time Zone columns without truncation for typical ledger names) and no longer requires the DataGrid to horizontal-scroll for common content widths.

---

## Task 12: Port GLUserConfig.xaml sizing fix

**Files:**
- Modify: `AIPowered/GLSense/GLSense.Addin.Core/Views/GLUserConfig.xaml`

- [ ] **Step 1: Change the size bounds**

Change:
```xml
                  MinWidth="500" MaxWidth="600"
```
to:
```xml
                  MinWidth="550" MaxWidth="630"
```
(`MinHeight="400" MaxHeight="560"` is unchanged.)

- [ ] **Step 2: Build and smoke-test**

```bash
msbuild AIPowered/GLSense/GLSense.sln /t:GLSense_Addin_Core /p:Configuration=Debug
```
Open GLUserConfig from the ribbon. Confirm the window is slightly wider and the DrillDowns/Options tabs both still render correctly at the new size.

---

## Task 13: Port AppOverlay minimum busy-overlay duration fix

**Files:**
- Modify: `AIPowered/GLSense/GLSense.Addin.Core/Views/AppOverlay.xaml.cs`

**Interfaces:** no signature change to `HideBusyAsync()` - callers (all 26 windows) are unaffected.

- [ ] **Step 1: Add the minimum-duration constant and check**

Add above `HideBusyAsync`:
```csharp
        // If a genuinely fast operation (e.g. a SQLite query that now finishes in under
        // 100ms thanks to the warm-up/load-before-show fixes) shows the busy overlay and
        // hides it again almost immediately, that reads as a flicker/splash rather than a
        // smooth load - reported as the window looking like it's "dancing" together with an
        // "inner blink." Holding the overlay up to this minimum is imperceptible for any
        // operation that's actually slow enough to need a loading indicator in the first
        // place, but turns a sub-100ms flash into a calm, readable state. Ported from
        // FinalWorkingCode's identical AppOverlay.xaml.cs fix.
        private const int MinBusyDurationMs = 400;

        public async Task HideBusyAsync()
        {
            ServiceLocator.Logger?.LogDebug("AppOverlay.HideBusyAsync invoked");

            if (_busyStart.HasValue)
            {
                var remaining = TimeSpan.FromMilliseconds(MinBusyDurationMs) - (DateTime.UtcNow - _busyStart.Value);
                if (remaining > TimeSpan.Zero)
                    await Task.Delay(remaining);
            }

            var tcs = new TaskCompletionSource<bool>();
```
(i.e. insert the `MinBusyDurationMs` constant declaration and the `if (_busyStart.HasValue) { ... }` block as the very first lines of the existing `HideBusyAsync` method body, before its existing `var tcs = new TaskCompletionSource<bool>();` line - everything after that line in the method is unchanged. `_busyStart` already exists as a field in this file, set in `ShowBusyasyn`.)

- [ ] **Step 2: Build**

```bash
msbuild AIPowered/GLSense/GLSense.sln /t:GLSense_Addin_Core /p:Configuration=Debug
```

- [ ] **Step 3: Manual smoke test**

Trigger a fast operation that shows/hides the busy overlay quickly (e.g. reopen GLLOVs for an already-cached ledger). Confirm the overlay is visibly readable for at least ~400ms instead of flashing.

---

## Task 14: Port Excel-formula-injection guard fix

**Files:**
- Modify: `AIPowered/GLSense/GLSense.Addin.Core/Drilldowns/DDDatatoWorksheet.cs:1648-1683`

- [ ] **Step 1: Replace the method**

Replace:
```csharp
        private static object ProtectExcelFormulaLikeText(object rawValue)
        {
            var text = SafeToString(rawValue);
            var equalsCount = CountEqualsSafely(text, rawValue);
            if (equalsCount > 1)
            {
                // Ensure we write as text for Excel
                return "'" + text;
            }
            return rawValue;
        }
        private static string SafeToString(object value) => value?.ToString() ?? string.Empty;
        private static int CountEqualsSafely(string text, object rawValue)
        {
            if (string.IsNullOrWhiteSpace(text) || !text.Contains("=")) return 0;
            try { return text.Count(c => c == '='); }
            catch (Exception ex)
            {
                ServiceLocator.Logger.LogError($"There are multiple '=' symbols in string value: {rawValue}. Exception: {ex.Message}");
                return 0;
            }
        }
```
with:
```csharp
        private static object ProtectExcelFormulaLikeText(object rawValue)
        {
            var text = SafeToString(rawValue);

            if (StartsWithEqualsSign(text))
            {
                // A leading "=" is what makes Excel interpret a pasted value as a live
                // formula (the actual injection vector) - prefix with an apostrophe so it's
                // written as text instead. Ported from FinalWorkingCode's identical fix.
                return "'" + text;
            }

            return rawValue;
        }
        private static string SafeToString(object value) => value?.ToString() ?? string.Empty;
        private static bool StartsWithEqualsSign(string text)
        {
            return !string.IsNullOrEmpty(text) && text.TrimStart().StartsWith("=", StringComparison.Ordinal);
        }
```

- [ ] **Step 2: Build**

```bash
msbuild AIPowered/GLSense/GLSense.sln /t:GLSense_Addin_Core /p:Configuration=Debug
```

- [ ] **Step 3: Manual smoke test**

Run a drilldown-to-worksheet that includes a cell value containing a single `=` sign somewhere other than the start (e.g. `A=B`) - confirm it is now written as a plain value (not prefixed with `'`), since the old logic (count > 1) incorrectly triggered for any string with 2+ `=` characters anywhere, while the fix only guards genuine formula-injection (leading `=`). Also confirm a value that actually starts with `=` (e.g. `=SUM(A1)`) is written prefixed with `'` (i.e. as literal text, not a live formula).

---

## Task 15: Port multi-ledger currency matching fix

**Files:**
- Modify: `AIPowered/GLSense/GLSense.Addin.Core/Models/BalanceDtoModel.cs:307-340`

- [ ] **Step 1: Replace the block**

Replace:
```csharp
            var matchedLedgerForCurrency = default(LedgerRecord);
            ...
            else
            {
                string ldgerNameNormalized = NormalizeStrings(ledgerName);
                var ledgerNames = ldgerNameNormalized.ToString().Split(';').Select(name => name.Trim());
                var matchingLedgers = ledgerRecord.Where(l => ledgerNames.Contains(l.LedgerName));
                if (!matchingLedgers.Any())
                    ServiceLocator.Logger?.LogWarn($"BalanceDto.CreateFromXllParameters: no matching ledger(s) found for LedgerName='{ledgerName}'. CellRef={cellRef}");
                balance.ledgerIdList = matchingLedgers.Any() ? matchingLedgers.Select(l => (object)l.LedgerId).ToArray() : null;
                balance.coaid = matchingLedgers.FirstOrDefault()?.Coaid.ToString();
                ledgerId = matchingLedgers.FirstOrDefault()?.LedgerId ?? 0;
                matchedLedgerForCurrency = matchingLedgers.FirstOrDefault();
            }
            if (matchedLedgerForCurrency != null)
                balance.isFunctionalCurrency = matchedLedgerForCurrency.CurrencyCode == balance.currencyCode;
            else
                balance.isFunctionalCurrency = true; // Default to true if no matching ledger found
```
with:
```csharp
            List<LedgerRecord> matchedLedgersForCurrency = null;
            ...
            else
            {
                string ldgerNameNormalized = NormalizeStrings(ledgerName);
                var ledgerNames = ldgerNameNormalized.ToString().Split(';').Select(name => name.Trim());
                var matchingLedgers = ledgerRecord.Where(l => ledgerNames.Contains(l.LedgerName)).ToList();
                if (!matchingLedgers.Any())
                    ServiceLocator.Logger?.LogWarn($"BalanceDto.CreateFromXllParameters: no matching ledger(s) found for LedgerName='{ledgerName}'. CellRef={cellRef}");
                balance.ledgerIdList = matchingLedgers.Any() ? matchingLedgers.Select(l => (object)l.LedgerId).ToArray() : null;
                balance.coaid = matchingLedgers.FirstOrDefault()?.Coaid.ToString();
                ledgerId = matchingLedgers.FirstOrDefault()?.LedgerId ?? 0;
                // Any one of the formula's own matched ledger(s) works here for
                // ledgerId/coaid - ledgers named together in the same formula call share the
                // same currency-code comparison outcome for THIS check, so every matched
                // ledger must be checked individually below: true if ANY of them has a
                // functional currency equal to this balance's currency code, false only if
                // none do. Ported from FinalWorkingCode's identical fix.
                matchedLedgersForCurrency = matchingLedgers;
            }
            if (matchedLedgersForCurrency != null && matchedLedgersForCurrency.Count > 0)
                balance.isFunctionalCurrency = matchedLedgersForCurrency.Any(l => l.CurrencyCode == balance.currencyCode);
            else
                balance.isFunctionalCurrency = true; // Default to true if no matching ledger found
```
(The `...` in both blocks represents the unchanged lines between the variable declaration and the `else` branch - e.g. the `if (ledgerRecord == null || ledgerRecord.Count == 0) { ... }` branch above it - copy this task's before/after exactly around the code actually present in the file at the read line numbers 307-340, do not guess at the omitted lines; re-read the file at edit time to confirm exact surrounding context before applying.)

- [ ] **Step 2: Build**

```bash
msbuild AIPowered/GLSense/GLSense.sln /t:GLSense_Addin_Core /p:Configuration=Debug
```

- [ ] **Step 3: Manual smoke test**

Run a balance formula referencing 2+ ledgers with different functional currencies via a `;`-separated `LedgerName` argument, where only one of them matches the balance's currency code. Confirm `isFunctionalCurrency` now evaluates `true` (any-match), whereas before the fix it depended on which ledger happened to be `FirstOrDefault()` in the matched set (non-deterministic-looking, order-dependent result).

---

## Task 16: Full-solution build + end-to-end regression pass

**Files:** none (verification only)

- [ ] **Step 1: Clean full-solution build**

```bash
msbuild AIPowered/GLSense/GLSense.sln /t:Clean,Build /p:Configuration=Debug
msbuild AIPowered/GLSense/GLSense.sln /t:Clean,Build /p:Configuration=Release
```
Expected: zero errors, zero new warnings introduced by this plan's changes (pre-existing warnings unrelated to this work are out of scope).

- [ ] **Step 2: Confirm zero remaining WPF-UI footprint**

```bash
grep -rn "Wpf.Ui\|WPF-UI\|WpfUiBootstrapper\|FluentWindow\|ExtendsContentIntoTitleBar" AIPowered/GLSense --include=*.cs --include=*.xaml --include=*.csproj --include=packages.config
```
Expected: no matches anywhere in `AIPowered/GLSense`.

- [ ] **Step 3: Full manual regression pass**

Using the running add-in (Debug build): log in (GLLogin), select a cube (GLCubeDetails), open every ribbon-triggered window at least once (per Task 7 Step 4's list), run at least one drilldown that exercises `DDDatatoWorksheet` (Task 14), run at least one balance formula referencing multiple ledgers (Task 15), open GLLOVs and confirm load-before-show (Task 10), open GLDrilldownCustomization with and without a selected cube (Task 9). Confirm no window shows a blank/flash frame on open, no window fails to close/reopen cleanly, and the add-in does not crash or log an unhandled exception during any of this.

- [ ] **Step 4: Commit**

```bash
git add -A
git status
```
Review the full diff list before committing - confirm no unintended files (build artifacts, `obj/`/`bin/` output, deleted package folders that shouldn't be tracked) are staged.
```bash
git commit -m "Remove WPF-UI dependency from AIPowered BaseWindow; port window-flash/DPI fixes and 2 data-correctness fixes from FinalWorkingCode"
```

---

## Note: chrome visual-parity (not in this plan)

This plan makes AIPowered's 26 `BaseWindow`-derived windows behave like FinalWorkingCode's (no WPF-UI, same DPI/resettle/placeholder timing, same specific bug fixes) but leaves AIPowered's own custom title-bar chrome layout in place (`TitleBarGridStyle`, 2-row Grid, `Themes/GlobalStyles.xaml`) rather than rebuilding it to match FinalWorkingCode's `HeaderBar`/3-row-Grid/`Margin="10"` structure pixel-for-pixel. If true pixel-level chrome parity is wanted, that's a separate, larger follow-up plan (touches all 26 windows' XAML layout, not just the 6 with behavior fixes, and risks visual regressions across windows that have zero other changes in this plan) - flag it explicitly and get a separate go-ahead before starting it, rather than folding it silently into this one.
