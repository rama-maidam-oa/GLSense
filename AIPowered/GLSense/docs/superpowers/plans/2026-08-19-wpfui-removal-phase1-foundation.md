# Phase 1: WPF-UI Removal Foundation + 3 Pilot Windows Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Remove AIPowered's `BaseWindow : Wpf.Ui.Controls.FluentWindow` dependency on the WPF-UI package,
replacing its sizing/DPI/centering engine with FinalWorkingCode's proven `DpiAwareWindow` model, and prove the
new foundation against 3 pilot windows before the remaining ~23 windows are ported in later batches.

**Architecture:** `BaseWindow` becomes a plain `Window` subclass carrying `DpiAwareWindow`'s full
`SourceInitialized`-time (pre-paint) fit/scale/center engine plus a ported `WindowLoadingPlaceholder` overlay,
while keeping AIPowered-specific integration points (`ServiceLocator`-based logging/Excel-handle ownership,
`ModalToExcel`/`WindowCaption`/`IconSymbol` properties — kept because dozens of not-yet-ported windows still
reference them). `GlobalStyles.xaml`/`Generic.xaml` gain every FinalWorkingCode-only style/brush they're
currently missing (additive — AIPowered-only keys are NOT deleted yet, since ~26 not-yet-ported windows still
depend on them; each later batch deletes exactly the keys its own newly-ported windows stop needing).

**Tech Stack:** .NET Framework 4.8.1, WPF, AddinExpress. MSBuild via
`C:\Program Files\Microsoft Visual Studio\2022\Community\MSBuild\Current\Bin\MSBuild.exe`.

**Spec:** `docs/superpowers/specs/2026-08-19-wpfui-removal-finalworkingcode-sync-design.md`

## Global Constraints

- Build verification command for every task: `msbuild GLSense.Addin.Core\GLSense.Addin.Core.csproj /t:Build
  /p:Configuration=Debug /p:SignAssembly=false /v:normal /nologo` — must report `0 Error(s)`. The
  `SignAssembly=false` override is verification-only (this machine lacks the real `GLSense.Contracts.pfx`);
  never use it for a real deliverable build.
- Do NOT delete any AIPowered-only `GlobalStyles.xaml`/`Generic.xaml` style key in this plan unless this
  plan's own tasks are its last consumer (verified by grep — see Task 8). Every other AIPowered-only key stays
  until its own later batch removes it.
- Do NOT touch `AppOverlay.xaml`/`.xaml.cs` or `ExcelRefEditControl.xaml`/`.xaml.cs` beyond the audit in Task 3
  — `AppOverlay.xaml.cs` in particular is already ahead of FinalWorkingCode (newer dim/slide-toast system) and
  must not be replaced.
- Every C#/XAML file this plan touches uses the standard namespace mapping already established throughout this
  codebase: `GLSense.Views` → `GLSense.Addin.Core.Views`, `GLSense.Utilities` → `GLSense.Addin.Core.Utilities`,
  `GLSense.Controls` → `GLSense.Addin.Core.Controls`, `GLSense.Converters` → `GLSense.Addin.Core.Converters`,
  `GLSense;component` → `GLSense.Addin.Core;component`, `LogUtility.*` → `ServiceLocator.Logger?.*`.

---

### Task 1: Port `WindowLoadingPlaceholder`

**Files:**
- Create: `GLSense.Addin.Core\Utilities\WindowLoadingPlaceholder.cs`

**Interfaces:**
- Produces: `WindowLoadingPlaceholder.ShowMatching(double left, double top, double width, double height,
  IntPtr excelOwnerHwnd) : int` and `WindowLoadingPlaceholder.Hide(int generation) : void` — consumed by
  Task 2's `BaseWindow.OnSourceInitialized`/`HookPlaceholderDismissal`.

- [ ] **Step 1: Write the file**

Source: `FinalWorkingCode\GLSense\Utilities\WindowLoadingPlaceholder.cs`, adapted only as follows (no other
line changes — the sizing/positioning/generation-token logic is unchanged):
- Namespace `GLSense.Utilities` → `GLSense.Addin.Core.Utilities`.
- Every `LogUtility.LogDebug(...)` / `LogUtility.LogWarn(...)` / `LogUtility.LogException(...)` →
  `ServiceLocator.Logger?.LogDebug(...)` / `ServiceLocator.Logger?.LogWarn(...)` /
  `ServiceLocator.Logger?.LogException(...)`, and add `using GLSense.Addin.Core.Infrastructure;` for
  `ServiceLocator`.
- `AppConstants.GLAccentHex` stays as-is — already exists at
  `GLSense.Addin.Core\AppConstants.cs:40` (`"#2E86AB"`), confirmed by grep; add
  `using GLSense.Addin.Core;` (or the actual namespace `AppConstants` lives in in this project) if not already
  implied by an existing using.
- Drop the `WarmUpInBackground()` method entirely — it exists in FinalWorkingCode to pre-warm the placeholder
  at ribbon-load time from a call site (`AddinModule_OnRibbonLoaded` or similar) that does not exist in
  AIPowered's architecture. Confirm via `grep -rn "WindowLoadingPlaceholder.WarmUpInBackground"` across
  `GLSense.Addin.Core` and `GLSense` that nothing calls it (nothing will, since this is a new file) — no call
  site needs to be added for this port; `EnsureCreated()` already lazily creates the placeholder on its first
  real `ShowMatching()` call, so pre-warming is a pure optimization, not a correctness requirement.

Full adapted content:

```csharp
using System;
using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using System.Windows.Media.Animation;
using System.Windows.Shapes;
using System.Windows.Threading;
using GLSense.Addin.Core.Infrastructure;

namespace GLSense.Addin.Core.Utilities
{
    // Targets the confirmed root cause behind the "blank window before loading" flash:
    // showing a *new* real window (GLCubeDetails, GLLOVs, etc.) always pays a genuine
    // per-instance cost - a fresh native HWND, a fresh visual tree, a fresh first layout/
    // paint pass - between the moment it's shown and ContentRendered. This is per-instance,
    // not per-type, so no amount of pre-warming a window *type* can pre-pay it.
    //
    // Every earlier attempt to hide this gap by manipulating the *real* window itself
    // (Opacity=0 until ready, parking it off-screen until ContentRendered) caused a worse,
    // independently-reproduced DWM black-flash regression. This takes a different approach
    // entirely: never touch the real window's Opacity/Position/Visibility. Instead, show one
    // small, deliberately trivial, reused loading indicator window sized/positioned to match
    // where the real window will appear, and hide it once the real window's own
    // ContentRendered fires. Reusing a single instance (Hide(), never Close()) means its own
    // one-time per-instance cost is paid once, off-screen - every later ShowMatching() call
    // just toggles visibility (and resizes/repositions) an HWND that already exists and has
    // already painted, which is fast.
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

        /// <summary>
        /// Shows the reused loading indicator matching the real window's own resolved
        /// Left/Top/Width/Height, so the transition reads as "the window was already
        /// there, its content just finished loading" instead of a small, unrelated box
        /// appearing somewhere else and then jumping to a differently-sized/positioned
        /// real window. Falls back to a small generic box centered near excelOwnerHwnd
        /// (or the primary screen's work area) when width/height aren't usable (NaN,
        /// infinite, or &lt;= 0) - e.g. a DisableAutoSizing dialog whose size was never
        /// resolved by BaseWindow's own fit-to-content pass. Returns a generation
        /// token - pass it to Hide(generation) so a stale dismissal (e.g. from a window
        /// that's still loading after a newer one already took over the placeholder)
        /// can't hide a placeholder a different, later call is now responsible for.
        /// </summary>
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
                        return; // superseded by a newer ShowMatching() - not ours to hide
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

            // Pay the one-time per-instance HWND-creation/first-paint cost right here,
            // off-screen, instead of on whichever real window first triggers ShowMatching().
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
                        if (dpi > 0)
                            scale = dpi / 96.0;
                    }
                    catch
                    {
                        // fall back to 1.0
                    }

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

- [ ] **Step 2: Add the file to the project**

Edit `GLSense.Addin.Core.csproj`: add a `<Compile Include="Utilities\WindowLoadingPlaceholder.cs" />` entry
next to the existing `<Compile Include="Utilities\WpfAppManager.cs" />` entry (same `<ItemGroup>`).

- [ ] **Step 3: Build verification**

Run: `msbuild GLSense.Addin.Core\GLSense.Addin.Core.csproj /t:Build /p:Configuration=Debug
/p:SignAssembly=false /v:normal /nologo`
Expected: `Build succeeded. 0 Error(s)` (this file isn't referenced by anything yet, so it just needs to compile
standalone).

- [ ] **Step 4: Commit**

```bash
git add GLSense.Addin.Core/Utilities/WindowLoadingPlaceholder.cs GLSense.Addin.Core/GLSense.Addin.Core.csproj
git commit -m "Port WindowLoadingPlaceholder from FinalWorkingCode"
```

---

### Task 2: Rewrite `BaseWindow.cs` to drop WPF-UI and adopt DpiAwareWindow's engine

**Files:**
- Modify: `GLSense.Addin.Core\Views\BaseWindow.cs` (full rewrite)

**Interfaces:**
- Consumes: `WindowLoadingPlaceholder.ShowMatching(...)`/`.Hide(...)` (Task 1).
- Produces: `BaseWindow : Window` with public members every existing window still needs, unchanged in
  signature: `EnableAutoLayoutRefresh`, `EnableExcelCentering`, `EnableEscapeToClose`, `AutoClampToWorkArea`,
  `WorkAreaMargin`, `ModalToExcel`, `CenterInExcel`, `WindowCaption`, `IconSymbol`, `MaxWidthCap`,
  `MaxHeightCap` (all bool/double/string properties, same names/types as before — 47 `WindowCaption` call
  sites and 27 `IconSymbol` call sites across the codebase must keep compiling unchanged). New:
  `DisableAutoSizing`, `MinContentScale` (ported from `DpiAwareWindow`, ungated defaults preserve current
  behavior — `DisableAutoSizing = false`, `MinContentScale = 0.85`). Protected: `ForceSizeToContentResettle()`
  and `PumpDispatcherFrame()` are REMOVED (no longer the sizing mechanism) — grep for any derived window still
  calling them before this task is considered done (Step 4 covers this).

- [ ] **Step 1: Confirm no other file calls the two methods being removed**

Run: `grep -rn "ForceSizeToContentResettle\|PumpDispatcherFrame" GLSense.Addin.Core --include="*.cs" | grep -v
"/obj/\|/bin/\|Views/BaseWindow.cs"`

Expected right now (before this task's rewrite): these calls exist in `GLCubeDetails.xaml.cs`,
`GLServerConfiguration` callers, etc. per `CLAUDE.md` section 1. **This step is a discovery step, not a
pass/fail gate** — record every file that calls either method; Step 5 below removes each call site (they
become no-ops once the new sizing model no longer needs a manual resettle, since every window either already
uses `SizeToContent="Manual"` — unaffected by this removal — or gets its own pre-paint fit-to-content pass from
the new engine on every relevant layout change).

- [ ] **Step 2: Write the new `BaseWindow.cs`**

Full replacement content — `DpiAwareWindow`'s complete sizing/DPI/centering engine
(`FinalWorkingCode\GLSense\Utilities\DpiAwareWindow.cs`), adapted to keep AIPowered's own
`ServiceLocator`-based Excel-ownership model (instead of `DpiAwareWindow`'s generic `SetExcelOwner(IntPtr)` /
`ShowDialogWithOwner(IntPtr)` pattern, which would require touching every one of the ~30 call sites that
currently just call `.ShowDialog()`/`.Show()` directly and rely on `BaseWindow` auto-setting the owner), and
keeping `WindowCaption`/`IconSymbol`/`ModalToExcel` (still referenced by not-yet-ported windows), plus
`DpiAwareWindow`'s toast-dismiss-on-any-input behavior (`DismissActiveToast`/`IsInteractionOverlayVisible`) —
a genuine feature AIPowered's current `BaseWindow` lacks entirely:

```csharp
using GLSense.Addin.Core.Helpers;
using GLSense.Addin.Core.Infrastructure;
using GLSense.Addin.Core.Utilities;
using System;
using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Interop;
using System.Windows.Media;
using System.Windows.Threading;

namespace GLSense.Addin.Core.Views
{
    public abstract class BaseWindow : Window
    {
        private HwndSource _hwndSource;
        private IntPtr _excelHandle;
        private bool _ownerSet;
        private double _currentScaleFactor = 1.0;
        private readonly ScaleTransform _dpiScaleTransform = new ScaleTransform(1.0, 1.0);
        private readonly string _windowName;
        private bool _isDragging;
        private Point _dragStartPoint;
        private bool _layoutRefreshPending;
        private bool _initialLayoutApplied;
        private double _initialMaxWidth = double.NaN;
        private double _initialMaxHeight = double.NaN;
        private double _initialMinHeight = double.NaN;
        private DispatcherTimer _resizeSettleTimer;

        public bool EnableAutoLayoutRefresh { get; set; } = true;
        public bool EnableExcelCentering { get; set; } = true;
        public bool EnableEscapeToClose { get; set; } = true;
        public bool AutoClampToWorkArea { get; set; } = true;
        public double WorkAreaMargin { get; set; } = 24d;
        public bool ModalToExcel { get; set; } = true;

        // Compat no-op: WPF-UI's FluentWindow declared this DP to extend window content
        // into its own custom title-bar chrome. BaseWindow no longer derives from
        // FluentWindow and never rendered that chrome to begin with, so this does
        // nothing - it exists purely because ~26 not-yet-ported windows' XAML sets
        // ExtendsContentIntoTitleBar="True" directly on their <views:BaseWindow> root
        // tag, and XAML markup compilation (MC3072) fails for any attribute that isn't a
        // real property on the tag's type. Remove this once every window has been ported
        // away from setting it (grep for "ExtendsContentIntoTitleBar" across Views/*.xaml
        // before deleting - same removal pattern as IconSymbol/WindowCaption).
        public bool ExtendsContentIntoTitleBar { get; set; }
        public bool CenterInExcel { get; set; } = true;
        public double? MaxWidthCap { get; set; } = 1400d;
        public double? MaxHeightCap { get; set; } = null;
        public double MinContentScale { get; set; } = 0.85;

        /// <summary>
        /// When true, completely disables all auto-sizing/clamping/centering logic in this
        /// class. Use for message boxes/dialogs that must respect user resizing exactly as
        /// authored, with no fit-to-content or work-area clamping at all.
        /// </summary>
        public bool DisableAutoSizing { get; set; } = false;

        public string WindowCaption
        {
            get => Title;
            set => Title = value;
        }

        // FontAwesome PackIconFontAwesomeKind name (e.g. "KeySolid"), still bound by every
        // not-yet-ported window's title-bar Grid (Style="{StaticResource TitleBarIconStyle}",
        // Kind="{Binding IconSymbol, RelativeSource={RelativeSource AncestorType=BaseWindow}}").
        // Ported windows stop using this (they hardcode their own header icon per
        // FinalWorkingCode's own-header-per-window pattern instead) - kept here only because
        // ~23 windows still reference it as of this task; remove once every window has been
        // ported and nothing binds to it any more (grep for "AncestorType=views:BaseWindow"
        // and "IconSymbol=" across Views/*.xaml before deleting).
        public string IconSymbol { get; set; } = "KeySolid";

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
                    if (_excelHandle == IntPtr.Zero)
                        ServiceLocator.Logger?.LogWarn($"[{_windowName}] Excel handle is IntPtr.Zero from ServiceLocator");
                }
                catch (Exception ex)
                {
                    ServiceLocator.Logger?.LogException(ex, $"[{_windowName}] Failed to get Excel handle from ServiceLocator");
                }

                AddHandler(UIElement.PreviewMouseDownEvent, new MouseButtonEventHandler(OnWindowPreviewMouseDown), true);
                AddHandler(UIElement.PreviewKeyDownEvent, new KeyEventHandler(OnWindowPreviewKeyDown), true);
                AddHandler(UIElement.PreviewTextInputEvent, new TextCompositionEventHandler(OnWindowPreviewTextInput), true);

                this.SourceInitialized += OnSourceInitialized;
                this.Loaded += OnLoaded;
                this.Closed += RestoreOwnerFocusOnClosed;

                MouseWheelFocusHelper.EnableHoverToScroll(this);

                ServiceLocator.Logger?.LogDebug($"[{_windowName}] constructor completed");
            }
            catch (Exception ex)
            {
                ServiceLocator.Logger?.LogException(ex, $"[{_windowName}] constructor error");
            }
        }

        // Dismisses the shared toast overlay (if any AppOverlay is visible on this window)
        // on any click/keypress/text input anywhere in the window - ported from
        // DpiAwareWindow, a genuine missing feature vs. AIPowered's previous BaseWindow.
        private void OnWindowPreviewMouseDown(object sender, MouseButtonEventArgs e) => DismissActiveToast();
        private void OnWindowPreviewTextInput(object sender, TextCompositionEventArgs e) => DismissActiveToast();

        private void OnWindowPreviewKeyDown(object sender, KeyEventArgs e)
        {
            DismissActiveToast();

            if (e.Key == Key.Escape && IsInteractionOverlayVisible())
            {
                e.Handled = true;
                return;
            }

            if (!e.Handled && e.Key == Key.Escape && EnableEscapeToClose)
            {
                e.Handled = true;
                ServiceLocator.Logger?.LogDebug($"[{_windowName}] Escape pressed - closing window");
                this.Close();
            }
        }

        private void DismissActiveToast()
        {
            try
            {
                if (FindName("AppOverlayControl") is FrameworkElement overlay)
                {
                    var dismissMethod = overlay.GetType().GetMethod("DismissToast", Type.EmptyTypes);
                    dismissMethod?.Invoke(overlay, null);
                }
            }
            catch (Exception ex)
            {
                ServiceLocator.Logger?.LogDebug($"[{_windowName}] toast dismiss ignored: {ex.Message}");
            }
        }

        private bool IsInteractionOverlayVisible()
        {
            try
            {
                if (FindName("AppOverlayControl") is FrameworkElement overlay)
                {
                    var isBusy = overlay.GetType().GetProperty("IsBusyVisible")?.GetValue(overlay) as bool? ?? false;
                    var isConfirm = overlay.GetType().GetProperty("IsConfirmVisible")?.GetValue(overlay) as bool? ?? false;
                    return isBusy || isConfirm;
                }
            }
            catch (Exception ex)
            {
                ServiceLocator.Logger?.LogDebug($"[{_windowName}] overlay interaction check ignored: {ex.Message}");
            }

            return false;
        }

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
                        }
                    }
                    catch (Exception ex)
                    {
                        ServiceLocator.Logger?.LogDebug($"[{_windowName}] Could not set Excel owner: {ex.Message}");
                    }
                }
                else if (ModalToExcel && _excelHandle == IntPtr.Zero)
                {
                    ServiceLocator.Logger?.LogWarn($"[{_windowName}] Cannot set Excel owner - handle is IntPtr.Zero");
                }

                _currentScaleFactor = DpiAwarenessHelper.GetWindowDpi(this) / 96.0;

                // Run the DPI/fit/center pass now, synchronously, while the window still has
                // no on-screen presence at all (SourceInitialized fires once the HWND exists
                // but strictly before Show()/ShowDialog() makes it visible) - not later via a
                // deferred Dispatcher callback, which used to run after the window was
                // already visible at a stale placeholder size/position, producing a visible
                // resize/reposition "pop" right on top of the window's first frame.
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

        private void OnLoaded(object sender, RoutedEventArgs e)
        {
            ServiceLocator.Logger?.LogDebug($"[{_windowName}] loaded");
        }

        private void QueueLayoutRefresh(DispatcherPriority priority)
        {
            if (_layoutRefreshPending)
                return;

            _layoutRefreshPending = true;
            Dispatcher.BeginInvoke(new Action(() =>
            {
                _layoutRefreshPending = false;
                ApplyLayoutRefresh();
            }), priority);
        }

        private void ApplyLayoutRefresh()
        {
            try
            {
                if (!EnableAutoLayoutRefresh || DisableAutoSizing)
                    return;

                AdjustForCurrentDpi();
                FitToAvailableWorkArea();

                if (EnableExcelCentering && CenterInExcel && !_initialLayoutApplied)
                {
                    _initialLayoutApplied = true;
                    CenterOverExcelOnce();
                }
            }
            catch (Exception ex)
            {
                ServiceLocator.Logger?.LogException(ex, $"[{_windowName}] ApplyLayoutRefresh error");
            }
        }

        public void RefreshWindowLayout()
        {
            try
            {
                if (DisableAutoSizing)
                    return;

                if (!Dispatcher.CheckAccess())
                {
                    QueueLayoutRefresh(DispatcherPriority.Background);
                    return;
                }

                if (Content is FrameworkElement root)
                {
                    root.InvalidateMeasure();
                    root.InvalidateArrange();
                }

                QueueLayoutRefresh(DispatcherPriority.Render);
            }
            catch (Exception ex)
            {
                ServiceLocator.Logger?.LogException(ex, $"[{_windowName}] RefreshWindowLayout error");
            }
        }

        private IntPtr WndProc(IntPtr hwnd, int msg, IntPtr wParam, IntPtr lParam, ref bool handled)
        {
            try
            {
                const int WM_DPICHANGED = 0x02E0;

                if (msg == WM_DPICHANGED && !DisableAutoSizing)
                {
                    AdjustForDpiChange((uint)wParam, lParam);
                    handled = true;
                }
            }
            catch (Exception ex)
            {
                ServiceLocator.Logger?.LogException(ex, $"[{_windowName}] WndProc error for msg {msg}");
            }

            return IntPtr.Zero;
        }

        private void AdjustForCurrentDpi()
        {
            try
            {
                _currentScaleFactor = DpiAwarenessHelper.GetWindowDpi(this) / 96.0;
            }
            catch (Exception ex)
            {
                ServiceLocator.Logger?.LogException(ex, $"[{_windowName}] AdjustForCurrentDpi error");
            }
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

                    this.Dispatcher.BeginInvoke(new Action(() =>
                    {
                        try
                        {
                            this.Left = rect.Left / scaleFactor;
                            this.Top = rect.Top / scaleFactor;
                            this.Width = rect.Width / scaleFactor;
                            this.Height = rect.Height / scaleFactor;
                            FitToAvailableWorkArea();
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

                MaxWidth = Math.Min(MaxWidth, availableWidth);
                MaxHeight = Math.Min(MaxHeight, availableHeight);

                if (sizeChanged && EnableExcelCentering && CenterInExcel)
                    RecenterAfterSizeChange(previousLeft, previousTop, previousWidth, previousHeight);
            }
            catch (Exception ex)
            {
                ServiceLocator.Logger?.LogException(ex, $"[{_windowName}] FitToAvailableWorkArea error");
            }
        }

        private void RecenterAfterSizeChange(double previousLeft, double previousTop, double previousWidth, double previousHeight)
        {
            try
            {
                if (double.IsNaN(previousLeft) || double.IsNaN(previousTop) ||
                    double.IsNaN(previousWidth) || double.IsNaN(previousHeight) ||
                    previousWidth <= 0 || previousHeight <= 0)
                {
                    return;
                }

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

        // Centers over Excel's own main window (via ServiceLocator.ExcelHandle), using this
        // window's real, post-layout size - called exactly once, right after the first
        // FitToAvailableWorkArea pass has resolved the window's true size.
        private void CenterOverExcelOnce()
        {
            try
            {
                UpdateLayout();

                double centerX, centerY;

                if (_excelHandle != IntPtr.Zero && GetWindowRect(_excelHandle, out RECT excelRect) &&
                    excelRect.Width > 0 && excelRect.Height > 0)
                {
                    double scale = _currentScaleFactor > 0 ? _currentScaleFactor : 1.0;
                    double excelLeft = excelRect.Left / scale;
                    double excelTop = excelRect.Top / scale;
                    double excelWidth = excelRect.Width / scale;
                    double excelHeight = excelRect.Height / scale;

                    centerX = excelLeft + (excelWidth / 2.0);
                    centerY = excelTop + (excelHeight / 2.0);
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
                ServiceLocator.Logger?.LogException(ex, $"[{_windowName}] CenterOverExcelOnce error");
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

        // Content that grows/shrinks AFTER the window is already shown (e.g. a DataGrid
        // populating asynchronously well past initial load) fires RenderSizeChanged, not
        // SourceInitialized/WM_DPICHANGED - neither of which OnSourceInitialized/
        // AdjustForDpiChange's own FitToAvailableWorkArea calls are reached by. Without
        // this, a window whose content grows post-show would render past its intended
        // Max bounds/off-center until the user manually resized it or a DPI change
        // happened to fire. Debounced (120ms) since a DataGrid can raise many
        // back-to-back RenderSizeChanged events while its rows populate - reacting to
        // every single one would make the window visibly "dance" (resize/reposition more
        // than once in quick succession) instead of settling once, quietly, after
        // rendering goes quiet.
        protected override void OnRenderSizeChanged(SizeChangedInfo sizeInfo)
        {
            base.OnRenderSizeChanged(sizeInfo);

            if (DisableAutoSizing || !AutoClampToWorkArea)
                return;

            _resizeSettleTimer?.Stop();
            _resizeSettleTimer = new DispatcherTimer { Interval = TimeSpan.FromMilliseconds(120) };
            _resizeSettleTimer.Tick += (s, e) =>
            {
                _resizeSettleTimer.Stop();
                try
                {
                    FitToAvailableWorkArea();
                }
                catch (Exception ex)
                {
                    ServiceLocator.Logger?.LogException(ex, $"[{_windowName}] OnRenderSizeChanged (debounced clamp) error");
                }
            };
            _resizeSettleTimer.Start();
        }

        protected override void OnMouseLeftButtonDown(MouseButtonEventArgs e)
        {
            try
            {
                if (e.LeftButton == MouseButtonState.Pressed)
                {
                    _isDragging = true;
                    _dragStartPoint = e.GetPosition(this);
                    this.DragMove();
                }
            }
            catch (Exception ex)
            {
                ServiceLocator.Logger?.LogException(ex, $"[{_windowName}] OnMouseLeftButtonDown error");
            }
            base.OnMouseLeftButtonDown(e);
        }

        protected override void OnMouseLeftButtonUp(MouseButtonEventArgs e)
        {
            _isDragging = false;
            base.OnMouseLeftButtonUp(e);
        }

        protected override void OnMouseMove(MouseEventArgs e)
        {
            if (_isDragging && e.LeftButton == MouseButtonState.Pressed)
            {
                try
                {
                    var currentPos = e.GetPosition(this);
                    var diff = currentPos - _dragStartPoint;
                    var workingArea = SystemParameters.WorkArea;

                    var newLeft = Math.Max(workingArea.Left, Math.Min(workingArea.Right - this.Width, this.Left + diff.X));
                    var newTop = Math.Max(workingArea.Top, Math.Min(workingArea.Bottom - this.Height, this.Top + diff.Y));

                    this.Left = newLeft;
                    this.Top = newTop;
                    _dragStartPoint = currentPos;
                }
                catch (Exception ex)
                {
                    ServiceLocator.Logger?.LogException(ex, $"[{_windowName}] OnMouseMove error");
                }
            }
            base.OnMouseMove(e);
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

        private void RestoreOwnerFocusOnClosed(object sender, EventArgs e)
        {
            try
            {
                if (Owner != null)
                {
                    Owner.Activate();
                }
                else if (_ownerSet)
                {
                    ExcelWindowHelper.ActivateExcelMainWindow();
                }
            }
            catch (Exception ex)
            {
                ServiceLocator.Logger?.LogException(ex, $"[{_windowName}] RestoreOwnerFocusOnClosed error");
            }
        }

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
                            }
                        }
                        catch (Exception ex)
                        {
                            ServiceLocator.Logger?.LogDebug($"[{_windowName}] Could not set Excel owner in ShowDialog: {ex.Message}");
                        }
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

        [DllImport("user32.dll")]
        private static extern bool GetWindowRect(IntPtr hWnd, out RECT lpRect);

        [StructLayout(LayoutKind.Sequential)]
        private struct RECT
        {
            public int Left;
            public int Top;
            public int Right;
            public int Bottom;
            public int Width => Right - Left;
            public int Height => Bottom - Top;
        }
    }
}
```

**Design notes for the reviewer:**
- `ExcelWindowHelper.ActivateExcelMainWindow()` is assumed to already exist (it's referenced in the CURRENT
  `BaseWindow.cs` at the line `ExcelWindowHelper.ActivateExcelMainWindow();` inside `OnClosed`) — confirm via
  `grep -rn "class ExcelWindowHelper"` before this task is done; if it doesn't exist under that exact name,
  find its actual location/name and fix the `using`/call accordingly rather than inventing a new helper.
- `GetDpiForWindow`/`IsWindow`/`SetParent`/`GetParent`/`GetWindowThreadProcessId`/`SetForegroundWindow`/
  `BringWindowToTop` P/Invoke declarations present in the OLD `BaseWindow.cs` are dropped — grep confirms (Step
  1 above) whether anything outside this file called them; if the grep in Step 1 turns up a caller depending
  on one of these P/Invokes specifically being declared in `BaseWindow` (unlikely — they're private), restore
  just that declaration.

- [ ] **Step 3: Build verification (compile-only check)**

Run: `msbuild GLSense.Addin.Core\GLSense.Addin.Core.csproj /t:Build /p:Configuration=Debug
/p:SignAssembly=false /v:normal /nologo`

Expected: errors, not success — every window still referencing `ForceSizeToContentResettle()`/
`PumpDispatcherFrame()` (protected members that no longer exist) will fail to compile. This is the exact,
complete list of call sites Step 4 must fix; do not proceed to "0 errors" yet.

- [ ] **Step 4: Fix every call site the build just flagged**

For each error `'BaseWindow' does not contain a definition for 'ForceSizeToContentResettle'` (or
`PumpDispatcherFrame`) in a specific file:
- If the call is inside `OnContentRendered`/`OnLoaded`-style "resettle again once real data loads" code (per
  `CLAUDE.md` sections 1.4b/1.4c/24.1/26.3.5) — delete the call (and its surrounding `if
  (this.SizeToContent != SizeToContent.Manual)` guard, if the resulting block becomes empty) since the new
  `BaseWindow` re-runs `FitToAvailableWorkArea` from `AdjustForDpiChange`/its own `OnSourceInitialized` pass
  instead of relying on a manual per-window resettle call. Do NOT delete the surrounding method itself if it
  does other work (e.g. `DataGridColumnFillHelper.Refresh(...)` calls in the same method stay).
  Do NOT touch this file's un-related business logic.
- Record the full list of files fixed this way — Task 8 uses this list to confirm none of them still
  reference the two removed methods.

- [ ] **Step 5: Build verification**

Run: `msbuild GLSense.Addin.Core\GLSense.Addin.Core.csproj /t:Build /p:Configuration=Debug
/p:SignAssembly=false /v:normal /nologo`
Expected: `Build succeeded. 0 Error(s)`.

- [ ] **Step 6: Commit**

```bash
git add -A
git commit -m "Replace BaseWindow's WPF-UI/FluentWindow engine with DpiAwareWindow's sizing model"
```

---

### Task 3: Add FinalWorkingCode-only theme keys into `GlobalStyles.xaml`/`Generic.xaml` (additive only)

**Files:**
- Modify: `GLSense.Addin.Core\Themes\GlobalStyles.xaml`
- Modify: `GLSense.Addin.Core\Themes\Generic.xaml`
- Modify (audit only, patch if needed): `GLSense.Addin.Core\Views\GLBalanceConfigurator.xaml`

**Interfaces:**
- Produces: every `StaticResource` key FinalWorkingCode's `GlobalStyles.xaml`/`Generic.xaml` defines is now
  also resolvable in AIPowered. No key currently in AIPowered's files is removed by this task.

- [ ] **Step 1: Identify every FinalWorkingCode-only key**

Run (from repo root `D:\SQLLite_Test\GLSense`):
```
diff -u FinalWorkingCode/GLSense/Themes/GlobalStyles.xaml AIPowered/GLSense/GLSense.Addin.Core/Themes/GlobalStyles.xaml
diff -u FinalWorkingCode/GLSense/Themes/Generic.xaml AIPowered/GLSense/GLSense.Addin.Core/Themes/Generic.xaml
```
Confirmed-present FinalWorkingCode-only keys as of this plan's authoring (re-run the diff — this list must
match; if it doesn't, the two files changed since this plan was written and the diff's actual output is
authoritative): `SurfaceBrush`, `SurfaceColor`, `ProgressButton`, `TitleHeaderTextBlock`, `HeaderBar`,
`HeaderLabel`, `PaginationButton`, `PaginationInfoText`, `CompactBorderStyle`, `CompactCardStyle`,
`CompactCombo`, `CompactRefEditControl`, `CustomWindowCloseButtonStyle`, `EditableDataGridText`, `LabelStyle`,
`ModernDataGridCheckBoxColumn`, `ModernDataGridRowWithBorder`, `ModernDataGridTextBlock`, `SectionHeader`,
`SectionHeaderStyle`, `SectionHeaderStyle1`, `SingleClickCheckBoxStyle`, `SmallTextBox`, `SpinnerButtonStyle`,
`TabItemStyle`, `TextBlockLabelStyle`, `WarningToolTipStyle`.

- [ ] **Step 2: Copy each key's full definition block into AIPowered's files**

For `GlobalStyles.xaml`: read `FinalWorkingCode\GLSense\Themes\GlobalStyles.xaml` in full. For each key in the
Step 1 list, copy its complete XAML element (the whole `<Style x:Key="...">...</Style>` /
`<SolidColorBrush x:Key="...">` / `<Color x:Key="...">` / `<ControlTemplate x:Key="...">` block, including any
`<Style.Triggers>`/`<Style.Resources>` children) verbatim. Append all copied blocks into AIPowered's
`GlobalStyles.xaml`, inserted as a new section immediately before the closing `</ResourceDictionary>` tag,
under a comment header:
```xml
    <!-- ==================== Ported from FinalWorkingCode (2026-08-19) ==================== -->
```
Apply exactly these text substitutions to the copied blocks (nothing else):
- `clr-namespace:GLSense.Views` → `clr-namespace:GLSense.Addin.Core.Views`
- `clr-namespace:GLSense.Controls` → `clr-namespace:GLSense.Addin.Core.Controls`
- `pack://application:,,,/GLSense;component/` → `pack://application:,,,/GLSense.Addin.Core;component/`
- `/GLSense;component/` → `/GLSense.Addin.Core;component/`

If a copied block's `TargetType`/`BasedOn` references another key that turns out to ALSO be FinalWorkingCode-
only and isn't in the Step 1 list (i.e. Step 1's diff missed a transitive dependency), add that key's block
too, applying the same substitutions — do not leave a dangling reference.

Repeat the identical process for `Generic.xaml` against its own FinalWorkingCode-only diff content.

- [ ] **Step 3: Audit `AppOverlay.xaml`, `ExcelRefEditControl.xaml`, `GLBalanceConfigurator.xaml` for anything this step might affect**

Run:
```
grep -o 'StaticResource [A-Za-z0-9_]*' GLSense.Addin.Core/Views/AppOverlay.xaml | sort -u
grep -o 'StaticResource [A-Za-z0-9_]*' GLSense.Addin.Core/Views/ExcelRefEditControl.xaml | sort -u
grep -o 'StaticResource [A-Za-z0-9_]*' GLSense.Addin.Core/Views/GLBalanceConfigurator.xaml | sort -u
```
Since this task is purely additive (nothing is removed), every key these 3 files already reference remains
resolvable — this step is a confirmation, not expected to require any fix. If any key referenced by these 3
files is missing from AIPowered's `GlobalStyles.xaml`/`Generic.xaml` even after Step 2 (i.e. it's neither an
existing AIPowered key nor one of the Step 1 FinalWorkingCode-only keys — meaning it doesn't exist in EITHER
file, a genuine typo/bug in the referencing XAML unrelated to this migration), log it in this task's commit
message rather than inventing a new style to satisfy it.

- [ ] **Step 4: Build verification**

Run: `msbuild GLSense.Addin.Core\GLSense.Addin.Core.csproj /t:Build /p:Configuration=Debug
/p:SignAssembly=false /v:normal /nologo`
Expected: `Build succeeded. 0 Error(s)` — adding resource dictionary entries doesn't affect compilation, but
this confirms the XAML is still well-formed (a malformed addition would fail markup compilation).

- [ ] **Step 5: Commit**

```bash
git add GLSense.Addin.Core/Themes/GlobalStyles.xaml GLSense.Addin.Core/Themes/Generic.xaml
git commit -m "Add FinalWorkingCode-only theme styles/brushes to GlobalStyles.xaml/Generic.xaml (additive)"
```

---

### Task 4: Remove the WPF-UI package

**Files:**
- Modify: `GLSense.Addin.Core\GLSense.Addin.Core.csproj`
- Modify: `GLSense.Addin.Core\packages.config`
- Modify: `GLSense.Addin.Core\Utilities\WpfUiBootstrapper.cs` (delete)
- Modify: every `.xaml` file under `GLSense.Addin.Core\Views\` and `GLSense.Addin.Core\Themes\` containing
  `xmlns:ui="http://schemas.lepo.co/wpfui/2022/xaml"` (remove the declaration line only — confirmed via the
  earlier grep that no `<ui:...>` element exists anywhere, so no element removal is needed)

- [ ] **Step 1: Confirm zero remaining `<ui:` element usage**

Run: `grep -rn "<ui:" GLSense.Addin.Core --include="*.xaml" | grep -v "/obj/\|/bin/"`
Expected: no output. (Already confirmed during this plan's authoring — re-verify before deleting anything,
since new code may have been added since.)

- [ ] **Step 2: Delete `WpfUiBootstrapper.cs` and its call site**

Delete `GLSense.Addin.Core\Utilities\WpfUiBootstrapper.cs`. In the OLD `BaseWindow.cs` this call site was:
```csharp
if (!WpfUiBootstrapper.IsInitialized)
{
    WpfUiBootstrapper.Initialize();
    WpfUiBootstrapper.SetLightTheme();
}
```
This is already absent from Task 2's new `BaseWindow.cs` (it never had this block) — confirm via `grep -rn
"WpfUiBootstrapper" GLSense.Addin.Core --include="*.cs" | grep -v "/obj/\|/bin/"` that no other file
references it before deleting the `.csproj`'s `<Compile Include="Utilities\WpfUiBootstrapper.cs" />` entry.

- [ ] **Step 3: Remove the WPF-UI `<Reference>` entries from the csproj**

Edit `GLSense.Addin.Core.csproj`: remove both:
```xml
    <Reference Include="Wpf.Ui, Version=4.3.0.0, Culture=neutral, PublicKeyToken=11f9f5cc97b3ffd6, processorArchitecture=MSIL">
      <HintPath>..\packages\WPF-UI.4.3.0\lib\net481\Wpf.Ui.dll</HintPath>
    </Reference>
    <Reference Include="Wpf.Ui.Abstractions, Version=4.3.0.0, Culture=neutral, PublicKeyToken=11f9f5cc97b3ffd6, processorArchitecture=MSIL">
      <HintPath>..\packages\WPF-UI.Abstractions.4.3.0\lib\net481\Wpf.Ui.Abstractions.dll</HintPath>
    </Reference>
```
and the `<Compile Include="Utilities\WpfUiBootstrapper.cs" />` entry.

- [ ] **Step 4: Remove the packages.config entry**

Edit `GLSense.Addin.Core\packages.config`: remove the `WPF-UI` (and `WPF-UI.Abstractions` if listed
separately) `<package .../>` line(s).

- [ ] **Step 5: Remove `xmlns:ui` declarations from every XAML file**

Run this PowerShell to do the removal mechanically and precisely (one exact line pattern, applied file by
file, no other content touched):
```powershell
Get-ChildItem "GLSense.Addin.Core" -Recurse -Filter *.xaml | Where-Object { $_.FullName -notmatch '\\obj\\|\\bin\\' } | ForEach-Object {
    $content = Get-Content $_.FullName -Raw
    if ($content -match 'xmlns:ui="http://schemas\.lepo\.co/wpfui/2022/xaml"\s*\r?\n?') {
        $updated = $content -replace 'xmlns:ui="http://schemas\.lepo\.co/wpfui/2022/xaml"\s*\r?\n?', ''
        Set-Content -Path $_.FullName -Value $updated -NoNewline -Encoding UTF8
        Write-Output "Updated: $($_.FullName)"
    }
}
```
Manually re-open each file this prints and confirm the XAML is still well-formed (no leftover trailing
whitespace issue on the previous attribute's line) — a `>` or the next attribute must immediately follow
cleanly.

- [ ] **Step 6: Build verification**

Run: `msbuild GLSense.Addin.Core\GLSense.Addin.Core.csproj /t:Build /p:Configuration=Debug
/p:SignAssembly=false /v:normal /nologo`
Expected: `Build succeeded. 0 Error(s)`.

Also confirm the WPF-UI DLLs are no longer copied to the output:
```powershell
Test-Path "GLSense.Addin.Core\bin\Debug\Wpf.Ui.dll"
```
Expected: `False`.

- [ ] **Step 7: Commit**

```bash
git add -A
git commit -m "Remove WPF-UI package dependency entirely"
```

---

### Task 5: Port pilot window 1 — `GLWaitWindow`

**Files:**
- Modify: `GLSense.Addin.Core\Views\GLWaitWindow.xaml`
- Modify: `GLSense.Addin.Core\Views\GLWaitWindow.xaml.cs`

- [ ] **Step 1: Replace the XAML wholesale**

Source: `FinalWorkingCode\GLSense\Views\GLWaitWindow.xaml`. Apply the standard adaptation (already established
by this file's own current header comment, and used identically for every other ported window in this
codebase):
- Root element `<utils:DpiAwareWindow x:Class="GLSense.Views.GLWaitWindow"` → `<views:BaseWindow
  x:Class="GLSense.Addin.Core.Views.GLWaitWindow"`, closing tag `</utils:DpiAwareWindow>` →
  `</views:BaseWindow>`.
- `xmlns:utils="clr-namespace:GLSense.Utilities"` line removed (no longer needed — `BaseWindow` lives in
  `Views`); add `xmlns:views="clr-namespace:GLSense.Addin.Core.Views"` and
  `xmlns:local="clr-namespace:GLSense.Addin.Core.Views"` (matching every other ported window's namespace
  block) if not already present from the base FinalWorkingCode file (it uses `local:AppOverlay` for the
  overlay control).
- `<ResourceDictionary Source="pack://application:,,,/GLSense;component/Themes/GlobalStyles.xaml"/>` →
  `<ResourceDictionary Source="/GLSense.Addin.Core;component/Themes/GlobalStyles.xaml" />`.
- Keep FinalWorkingCode's own header (the inline `<Border Grid.Row="0" Background="{StaticResource
  PrimaryBrush}">` with `iconPacks:PackIconFontAwesome Kind="RotateSolid"` + `txtTitle` `TextBlock`) exactly as
  authored — do NOT reintroduce AIPowered's previous `TitleBarGridStyle`/`IconSymbol`-bound header; this pilot
  is the first window in the codebase to adopt FinalWorkingCode's own-header-per-window pattern.
- Keep FinalWorkingCode's `SizeToContent="Height" Width="420" MinHeight="180" MaxHeight="350"` sizing exactly
  (this is a real, deliberate difference from AIPowered's current `SizeToContent="WidthAndHeight" MinWidth="420"
  MaxWidth="420"` — FinalWorkingCode's is correct per the golden reference).
- The `AppOverlay` element stays `x:Name="AppOverlayControl"` (matches `BaseWindow`'s `DismissActiveToast`/
  `IsInteractionOverlayVisible` lookup-by-name from Task 2) — confirm `local:AppOverlay` resolves to
  `GLSense.Addin.Core.Views.AppOverlay` via the `xmlns:local` mapping above.

- [ ] **Step 2: Replace the code-behind, preserving AIPowered-specific members**

Keep every field/method already in `GLSense.Addin.Core\Views\GLWaitWindow.xaml.cs` (the class body is already
adapted — `ServiceLocator.Logger?.*`, `BaseWindow` base class, `TitleBar_MouseLeftButtonDown` — none of this
changes). The only edits this task makes:
- `TitleBar_MouseLeftButtonDown`'s XAML hookup: since Step 1's header now uses FinalWorkingCode's own `Grid`
  structure (not `Style="{StaticResource TitleBarGridStyle}" MouseLeftButtonDown="TitleBar_MouseLeftButtonDown"`),
  confirm the copied header `Border`/`Grid` still wires `MouseLeftButtonDown="TitleBar_MouseLeftButtonDown"` on
  itself (FinalWorkingCode's own XAML wires this on the `DockPanel`/`Grid` header element already — keep that
  same wiring, just renamed to match this project's existing handler name if FinalWorkingCode used a
  differently-named handler for this specific window).
- No other code-behind logic changes — `StartMonitoring`, `SetProcessTitle`, `SetProcessMessage`,
  `RequestClose`, `BtnCancel_Click`, `ShowConfirmToastAsync`, the `IDisposable` pattern all stay exactly as they
  are today.

- [ ] **Step 3: Build verification**

Run: `msbuild GLSense.Addin.Core\GLSense.Addin.Core.csproj /t:Build /p:Configuration=Debug
/p:SignAssembly=false /v:normal /nologo`
Expected: `Build succeeded. 0 Error(s)`.

- [ ] **Step 4: Commit**

```bash
git add GLSense.Addin.Core/Views/GLWaitWindow.xaml GLSense.Addin.Core/Views/GLWaitWindow.xaml.cs
git commit -m "Port GLWaitWindow XAML wholesale from FinalWorkingCode (pilot 1)"
```

---

### Task 6: Port pilot window 2 — `GLMessageWindow`

**Files:**
- Modify: `GLSense.Addin.Core\Views\GLMessageWindow.xaml`
- Modify: `GLSense.Addin.Core\Views\GLMessageWindow.xaml.cs`

- [ ] **Step 1: Replace the XAML wholesale**

Source: `FinalWorkingCode\GLSense\Views\GLMessageWindow.xaml`. Same adaptation rules as Task 5 Step 1 (root
element/class name, namespace mapping, resource dictionary source path). Specific differences to preserve
exactly from FinalWorkingCode (do NOT keep AIPowered's current `TitleBarGridStyle` header or
`SizeToContent="WidthAndHeight" MinHeight="165" MaxHeight="500" MinWidth="400" MaxWidth="620"
WindowStyle="SingleBorderWindow" ExtendsContentIntoTitleBar="True"` — replace with FinalWorkingCode's own):
- `WindowStyle="None" Background="#F8F9FA" AllowsTransparency="false" ShowInTaskbar="False" Topmost="False"
  ShowActivated="True"`, same `MinHeight="165" MaxHeight="500" MinWidth="400" MaxWidth="620"`.
- FinalWorkingCode's own `DockPanel`-based header (`HeaderPanel`, `MouseLeftButtonDown="HeaderPanel_MouseLeftButtonDown"`,
  `CustomWindowCloseButtonStyle` close button) replaces AIPowered's `TitleBarGridStyle`/`TitleBarIconStyle`/
  `TitleBarTextStyle`/`TitleBarCloseButtonStyle` header Grid.
- The local `<Style TargetType="ScrollBar">`/`<Style TargetType="Thumb">` overrides in FinalWorkingCode's own
  `<utils:DpiAwareWindow.Resources>` come along as part of the wholesale copy (window-scoped, harmless).

- [ ] **Step 2: Replace the code-behind, preserving AIPowered-specific members**

FinalWorkingCode's `GLMessageWindow.xaml.cs` constructor signature is `GLMessageWindow(string message,
MessageBoxIcon icon, MessageBoxButtons buttons = MessageBoxButtons.OK)` — confirm this matches AIPowered's
current constructor exactly (it does, per the file already read during this plan's authoring) so no call site
elsewhere in the codebase (`CommonFunctions.GLSenseMessage`) needs updating.
- Rename the event handler `HeaderPanel_MouseLeftButtonDown` to match whatever Step 1's copied XAML actually
  wires (keep FinalWorkingCode's own name — no reason to rename it to `TitleBar_MouseLeftButtonDown` here,
  since this window's own convention differs from `GLWaitWindow`'s and both are valid, already-established
  patterns in this codebase).
- `EnhancedDragDropHelper.EnableWindowDrag(this)` (used in FinalWorkingCode's constructor) — grep
  `GLSense.Addin.Core\Helpers\EnhancedDragDropHelper.cs` to confirm this helper already exists in AIPowered
  (`GLWaitWindow.xaml.cs`'s own header comment says it does, via `GLSense.Addin.Core.Helpers`). If present,
  keep this line as FinalWorkingCode has it (no need for AIPowered's current per-handler
  `TitleBar_MouseLeftButtonDown` if FinalWorkingCode already solves it via the shared drag helper — use
  whichever this specific file's own FinalWorkingCode source actually does, verbatim).
- All `AddDialogButton`/`SetMessageIcon`/`SetupButtons`/`Result` logic stays exactly as FinalWorkingCode has it
  (already what AIPowered has today, confirmed identical during this plan's authoring) —
  `ServiceLocator.Logger?.*` replaces `LogUtility.*` per the standard mapping.

- [ ] **Step 3: Build verification**

Run: `msbuild GLSense.Addin.Core\GLSense.Addin.Core.csproj /t:Build /p:Configuration=Debug
/p:SignAssembly=false /v:normal /nologo`
Expected: `Build succeeded. 0 Error(s)`.

- [ ] **Step 4: Commit**

```bash
git add GLSense.Addin.Core/Views/GLMessageWindow.xaml GLSense.Addin.Core/Views/GLMessageWindow.xaml.cs
git commit -m "Port GLMessageWindow XAML wholesale from FinalWorkingCode (pilot 2)"
```

---

### Task 7: Port pilot window 3 — `GLSegmentValues`

**Files:**
- Modify: `GLSense.Addin.Core\Views\GLSegmentValues.xaml`
- Modify: `GLSense.Addin.Core\Views\GLSegmentValues.xaml.cs`

- [ ] **Step 1: Read both sides in full before editing**

Read `FinalWorkingCode\GLSense\Views\GLSegmentValues.xaml` (621 lines) and
`FinalWorkingCode\GLSense\Views\GLSegmentValues.xaml.cs` (436 lines) in full, alongside AIPowered's current
507-line `.xaml.cs`. This file carries real, AIPowered-specific business logic (per `CLAUDE.md` sections 7,
7.0b, 7.1, 7.1b, 8.1, 9.2, 35) that must be individually verified present or correctly superseded — this task
is NOT a blind copy the way Tasks 5/6 are close to being.

- [ ] **Step 2: Replace the XAML, keeping AIPowered's `SegmentSelectorViewModel`-driven bindings intact**

Source: `FinalWorkingCode\GLSense\Views\GLSegmentValues.xaml`. Apply the standard namespace/root-element/
resource-path adaptation (same as Tasks 5/6). Specific known-correct target state (already independently
achieved once in AIPowered per `CLAUDE.md` section 8.1, then must survive re-porting from FinalWorkingCode
without regressing):
- `SizeToContent="Manual"` with explicit `Width="740" Height="700"` (matching section 7's fix).
- `dgLeft`'s "Description" column and `dgRight`'s "Segment" column as native `Width="*"` — NOT wrapped in
  `DataGridColumnFillHelper.EnableFillColumn(...)` calls (per section 8.1's finding that the helper is actively
  harmful for a `SizeToContent="Manual"` window). If FinalWorkingCode's own XAML/constructor never had
  `DataGridColumnFillHelper` calls at all (it doesn't — that class is AIPowered-only, built to work around a
  WPF-UI/`SizeToContent="WidthAndHeight"` measurement quirk that doesn't exist in FinalWorkingCode's
  architecture), this is automatically satisfied by a faithful wholesale copy — do not add
  `DataGridColumnFillHelper` calls back in.
- The Segment/Hierarchy row and Search row as direct, non-scrolling children of the outer `Grid` (per section
  9.2's fix — verify FinalWorkingCode's own XAML already has this structure, since it's the reference this
  fix was modeled on; if FinalWorkingCode's actual current XAML differs from what section 9.2 assumed, prefer
  FinalWorkingCode's real, current file content over the historical CLAUDE.md description).
- Is-Summary/Value/Description column tooltips (section 7.1) and the "Showing:"-labeled `PageRangeText` +
  paging footer icons (sections 7.1, 8.6) — keep whatever FinalWorkingCode's own file currently has for these;
  do not re-derive them from the CLAUDE.md prose description if the live FinalWorkingCode file has since
  evolved further.

- [ ] **Step 3: Reconcile the code-behind — do not blindly overwrite**

`GLSegmentValues.xaml.cs`'s constructor wires a `DataLoadedAction` callback on the shared
`SegmentSelectorViewModel` with a `_hasResettledAfterInitialLoad` guard (per sections 1.4b/24.1/26.3). Since
Task 2's new `BaseWindow` no longer has `ForceSizeToContentResettle()`/`PumpDispatcherFrame()` (removed) and
this window is `SizeToContent="Manual"` (never needed that resettle mechanism in the first place — confirmed
by `CLAUDE.md` section 7's own note that `SizeToContent="Manual"` windows are gated out of that machinery
entirely), remove this `DataLoadedAction`-triggered resettle call and its guard field entirely; do NOT replace
it with a call to anything on the new `BaseWindow`, since nothing is needed here.

Preserve exactly, verifying each is still present after the edit (cross-reference against `CLAUDE.md` sections
35 and this file's own current AIPowered content):
- The `SelectedHierarchy` backing-field-clear-on-segment-change fix (section 35).
- Any `SegmentSelectorViewModel` wiring specific to this window (`windowName="val"` construction parameter or
  equivalent) — do not change how this window constructs/shares the ViewModel.
- `ServiceLocator.Logger?.*`/`ServiceLocator.ExcelApp` usage (already correct, keep as-is).

- [ ] **Step 4: Build verification**

Run: `msbuild GLSense.Addin.Core\GLSense.Addin.Core.csproj /t:Build /p:Configuration=Debug
/p:SignAssembly=false /v:normal /nologo`
Expected: `Build succeeded. 0 Error(s)`.

- [ ] **Step 5: Commit**

```bash
git add GLSense.Addin.Core/Views/GLSegmentValues.xaml GLSense.Addin.Core/Views/GLSegmentValues.xaml.cs
git commit -m "Port GLSegmentValues XAML wholesale from FinalWorkingCode (pilot 3)"
```

---

### Task 8: Progressive prune of now-dead AIPowered-only theme keys + final verification

**Files:**
- Modify: `GLSense.Addin.Core\Themes\GlobalStyles.xaml`

- [ ] **Step 1: Find AIPowered-only keys with zero remaining consumers**

For each AIPowered-only key identified during this plan's authoring (`ModernButton`, `ModernTextBox`,
`SafeCheckBoxTemplate`, `PlainButtonTemplate`, `ModernScrollViewer`, `ModernSeparator`, `ModernTabControl`,
`ModernTabItem`, `ModernListBoxItem`, `ModernDataGridCell`, `ModernDataGridColumnHeader`, `ModernDataGridRow`,
`ReadOnlyTextBox`, `IndeterminateProgressBar`, `WhiteBrush`, `WhiteColor`, `DataGridSelectedBrush`,
`ChromeStyleToolTip`, `ComboBoxFocusVisualStyle`, `CardBorderStyle`, `HeaderBorderStyle`, `LabelIconStyle`,
`LabelIconBorderStyle`, `LabelWithIconStyle`, `LoginCardStyle`, `ValidationErrorStyle`, `ToastCloseButtonStyle`,
`FirstPageButtonStyle`, `LastPageButtonStyle`, `NextPageButtonStyle`, `PreviousPageButtonStyle`, `SmallIcon`,
`MediumIcon`, `LargeIcon`, `VerySmallIcon`, `SuccessBrush`, `SuccessColor`, `WarningBrush`, `WarningColor`,
`InfoBrush`, `InfoColor`, `DangerDarkBrush`, `DangerDarkColor`, `DangerDarkerBrush`, `DangerDarkerColor`,
`SuccessMessage`, `WarningMessage`, `InfoMessage`, `ErrorMessage`, `WindowTitleBarBrush`, `WindowTitleBarColor`,
`BackgroundBrush`, `BackgroundColor`, `HeaderTextBlock`, `ContentBorderStyle`, `TitleBarGridStyle`,
`TitleBarIconStyle`, `TitleBarTextStyle`, `TitleBarButtonStyle`, `TitleBarCloseButtonStyle`, `LabelTextStyle`,
`ModernDatePicker`), run:

```
grep -rl "StaticResource <KEY>}" GLSense.Addin.Core/Views/*.xaml GLSense.Addin.Core/Themes/Generic.xaml 2>/dev/null
```

Any key with **zero** matches now (after Tasks 5-7 ported `GLWaitWindow`/`GLMessageWindow`/`GLSegmentValues`
away from it) is safe to delete. Based on this plan's own authoring-time audit, expect at minimum
`IndeterminateProgressBar` (was `GLWaitWindow`'s sole consumer) to qualify — re-run the grep for the real,
current answer rather than trusting this list, since Tasks 5-7's actual edits may have kept or dropped
different keys than assumed here.

- [ ] **Step 2: Delete each zero-consumer key's definition block from `GlobalStyles.xaml`**

For each key confirmed dead in Step 1, remove its complete `<Style x:Key="...">`/`<SolidColorBrush
x:Key="...">`/`<Color x:Key="...">` block from `GlobalStyles.xaml`. Leave every other AIPowered-only key
(still consumed by not-yet-ported windows) untouched — this is a partial, conservative prune, not a full
cleanup; the remaining ~20 windows still ported in later batches will each trigger their own round of this
same prune step.

- [ ] **Step 3: Full solution component build + pilot regression check**

Run: `msbuild GLSense.Addin.Core\GLSense.Addin.Core.csproj /t:Build /p:Configuration=Debug
/p:SignAssembly=false /v:normal /nologo`
Expected: `Build succeeded. 0 Error(s)`.

Run: `grep -rn "ForceSizeToContentResettle\|PumpDispatcherFrame\|WpfUiBootstrapper" GLSense.Addin.Core
--include="*.cs" | grep -v "/obj/\|/bin/"`
Expected: no output (confirms Task 2 Step 4's fixes were complete and Task 4's deletion left no stragglers).

- [ ] **Step 4: Commit**

```bash
git add GLSense.Addin.Core/Themes/GlobalStyles.xaml
git commit -m "Prune GlobalStyles.xaml keys with zero remaining consumers after Phase 1 pilots"
```

---

## Handoff to the user

After Task 8, tell the user Phase 1 is ready for their own rebuild + fresh Excel relaunch (per the standing
deployment note: a running Excel session won't pick up a rebuilt DLL, and the versioned deployment folder only
refreshes via `post_build.cmd`). Ask them to specifically check:
1. `GLWaitWindow`, `GLMessageWindow`, `GLSegmentValues` open without a blank gap, center correctly over Excel,
   and (for `GLSegmentValues`) the dual DataGrids scroll/select correctly and the window doesn't resize when
   switching segments.
2. Every other window (not yet ported) still opens — if the additive-only theme approach in Task 3 worked as
   designed, they should be visually unchanged from before this plan.
3. `GLMessageWindow` still appears correctly for every `CommonFunctions.GLSenseMessage` call site across the
   app (it's used everywhere as the app's message-box replacement — a good smoke test for broad coverage).
