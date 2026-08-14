using System;
using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Input;
using System.Windows.Interop;
using System.Windows.Media;

namespace GLSense.Utilities
{
    public class DpiAwareWindow : Window
    {
        [DllImport("user32.dll")]
        private static extern bool SetForegroundWindow(IntPtr hWnd);

        private HwndSource _hwndSource;
        private IntPtr _excelOwnerHwnd = IntPtr.Zero;
        private double _currentScaleFactor = 1.0;
        private readonly ScaleTransform _dpiScaleTransform = new ScaleTransform(1.0, 1.0);
        private readonly string _windowName;
        private bool _layoutRefreshPending;
        private bool _initialLayoutApplied;
        private double _initialMaxWidth = double.NaN;
        private double _initialMaxHeight = double.NaN;
        private double _initialMinHeight = double.NaN;

        public bool EnableAutoLayoutRefresh { get; set; } = true;
        public bool EnableExcelCentering { get; set; } = true;
        public bool EnableEscapeToClose { get; set; } = true;

        public double CurrentScaleFactor => _currentScaleFactor;
        public bool AutoClampToWorkArea { get; set; } = true;
        public double WorkAreaMargin { get; set; } = 24d;
        public double? MaxWidthCap { get; set; } = 1400d;
        public double? MaxHeightCap { get; set; } = null;

        /// <summary>
        /// When set to true, completely disables all auto-sizing and clamping logic.
        /// Use this for message boxes and dialogs that should respect user resizing.
        /// </summary>
        public bool DisableAutoSizing { get; set; } = false;

        public double MinContentScale { get; set; } = 0.85;

        public DpiAwareWindow()
        {
            try
            {
                _windowName = GetType().Name;
                LogUtility.LogDebug($"[{_windowName}] constructing window");

                AddHandler(UIElement.PreviewMouseDownEvent, new MouseButtonEventHandler(OnWindowPreviewMouseDown), true);
                AddHandler(UIElement.PreviewKeyDownEvent, new KeyEventHandler(OnWindowPreviewKeyDown), true);
                AddHandler(UIElement.PreviewTextInputEvent, new TextCompositionEventHandler(OnWindowPreviewTextInput), true);

                // Ensures mouse wheel scrolling works on hover for every window derived
                // from this base class - see MouseWheelFocusHelper for the root cause
                // harmless no-op here since a shown top-level Window already has focus
                // the fix matters for content embedded via ElementHost/HWND-reparenting,
                // e.g. the Balance Configurator task pane).
                MouseWheelFocusHelper.EnableHoverToScroll(this);

                WpfAppManager.EnsureApplication();

                WpfAppManager.InvokeOnWpfThread(() =>
                {
                    try
                    {
                        using (DpiAwarenessHelper.SetPerMonitorAware())
                        {
                            this.UseLayoutRounding = true;
                            this.SnapsToDevicePixels = true;
                            TextOptions.SetTextFormattingMode(this, TextFormattingMode.Display);
                            RenderOptions.SetBitmapScalingMode(this, BitmapScalingMode.HighQuality);

                            this.SourceInitialized += OnSourceInitialized;
                            this.Loaded += OnLoaded;
                            this.ContentRendered += OnContentRenderedDebug;
                            this.Closed += OnClosedDebug;
                            this.Closed += RestoreOwnerFocusOnClosed;
                            this.Unloaded += OnUnloadedDebug;
                        }
                    }
                    catch (Exception ex)
                    {
                        LogUtility.LogException(ex, $"DpiAwareWindow constructor ({_windowName})");
                    }
                });
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex, "DpiAwareWindow constructor (fatal)");
            }
        }

        protected override void OnInitialized(EventArgs e)
        {
            try
            {
                LogUtility.LogDebug($"[{_windowName}] initialized");
                base.OnInitialized(e);
            }
            catch (System.IO.IOException ex)
            {
                LogUtility.LogException(ex, $"DpiAwareWindow.OnInitialized ({_windowName})");
            }
        }

        public void SetExcelOwner(IntPtr excelHwnd)
        {
            try
            {
                var helper = new WindowInteropHelper(this);
                helper.Owner = excelHwnd;
                _excelOwnerHwnd = excelHwnd;
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex, "DpiAwareWindow.SetExcelOwner");
            }
        }

        public bool? ShowDialogWithOwner(IntPtr excelHwnd)
        {
            int placeholderGen = WindowLoadingPlaceholder.ShowNear(excelHwnd);
            HookPlaceholderDismissal(placeholderGen);
            try
            {
                SetExcelOwner(excelHwnd);
                return this.ShowDialog();
            }
            catch (System.IO.IOException ex)
            {
                LogUtility.LogException(ex, "DpiAwareWindow.ShowDialogWithOwner (IO, retrying)");
                System.Threading.Thread.Sleep(100);
                SetExcelOwner(excelHwnd);
                return this.ShowDialog();
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex, "DpiAwareWindow.ShowDialogWithOwner (retrying)");
                try
                {
                    return this.ShowDialog();
                }
                catch (Exception innerEx)
                {
                    LogUtility.LogException(innerEx, "DpiAwareWindow.ShowDialogWithOwner (critical, retry failed)");
                    return null;
                }
            }
        }

        public void ShowWithOwner(IntPtr excelHwnd)
        {
            int placeholderGen = WindowLoadingPlaceholder.ShowNear(excelHwnd);
            HookPlaceholderDismissal(placeholderGen);
            try
            {
                SetExcelOwner(excelHwnd);
                this.Show();
            }
            catch (System.IO.IOException ex)
            {
                LogUtility.LogException(ex, "DpiAwareWindow.ShowWithOwner (IO, retrying)");
                System.Threading.Thread.Sleep(100);
                SetExcelOwner(excelHwnd);
                this.Show();
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex, "DpiAwareWindow.ShowWithOwner (retrying)");
                try
                {
                    this.Show();
                }
                catch (Exception innerEx)
                {
                    LogUtility.LogException(innerEx, "DpiAwareWindow.ShowWithOwner (critical, retry failed)");
                }
            }
        }

        // Dismisses the shared WindowLoadingPlaceholder once this window's own first real
        // frame is ready (ContentRendered), or immediately if it closes before that ever
        // happens (e.g. an exception during load) - the placeholder's own 3-second safety
        // timer covers any path that hits neither. Hooked once per Show/ShowDialog call
        // rather than in the constructor, since the placeholder should only be up for the
        // span between "user asked to see this window" and "this window is actually ready."
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

        private void OnWindowPreviewMouseDown(object sender, MouseButtonEventArgs e)
        {
            DismissActiveToast();
        }

        private void OnWindowPreviewKeyDown(object sender, KeyEventArgs e)
        {
            DismissActiveToast();

            if (e.Key == Key.Escape && IsInteractionOverlayVisible())
            {
                e.Handled = true;
                return;
            }

            if (EnableEscapeToClose && e.Key == Key.Escape)
            {
                e.Handled = true;
                Close();
            }
        }

        private void OnWindowPreviewTextInput(object sender, TextCompositionEventArgs e)
        {
            DismissActiveToast();
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
                LogUtility.LogDebug($"[{_windowName}] toast dismiss ignored: {ex.Message}");
            }
        }

        private bool IsInteractionOverlayVisible()
        {
            try
            {
                if (FindName("AppOverlayControl") is GLSense.Views.AppOverlay overlay)
                    return overlay.IsBusyVisible || overlay.IsConfirmVisible;
            }
            catch (Exception ex)
            {
                LogUtility.LogDebug($"[{_windowName}] overlay interaction check ignored: {ex.Message}");
            }

            return false;
        }

        private void OnSourceInitialized(object sender, EventArgs e)
        {
            try
            {
                LogUtility.LogDebug($"[{_windowName}] source initialized");
                _hwndSource = PresentationSource.FromVisual(this) as HwndSource;
                _hwndSource?.AddHook(WndProc);

                // Run the DPI/fit/center pass now, synchronously, while the window still
                // has no on-screen presence at all (SourceInitialized fires once the HWND
                // exists but strictly before Show()/ShowDialog() calls ShowWindow) - not
                // later via OnLoaded's deferred Dispatcher callback, which used to run
                // after the window was already visible at its placeholder
                // WindowStartupLocation="CenterOwner" position/size (computed before
                // layout resolved the window's real content size), producing a visible
                // resize/reposition "pop" right on top of the window's first frame. Doing
                // the exact same math here instead means Show() paints the correct final
                // size/position on the very first frame - there is nothing left to
                // visibly correct afterward. This does NOT touch Opacity/Visibility/
                // Position of an already-visible window - it only sets these properties
                // before the window has ever been shown, which is the same thing any WPF
                // app does when it sizes/positions a window up front.
                //
                // OnLoaded's own QueueLayoutRefresh call still runs afterward as a
                // harmless, idempotent safety re-check (e.g. in case DPI/content weren't
                // fully resolved yet at this earlier point) - it should no-op in the
                // common case since FitToAvailableWorkArea/CenterOverOwnerOnce already
                // did the real work here.
                if (!DisableAutoSizing)
                {
                    CaptureInitialWindowConstraints();
                    ApplyLayoutRefresh();
                }
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex, $"DpiAwareWindow.OnSourceInitialized ({_windowName})");
            }
        }

        private void OnLoaded(object sender, RoutedEventArgs e)
        {
            try
            {
                LogUtility.LogDebug($"[{_windowName}] loaded - applying DPI adjustments");

                if (!DisableAutoSizing)
                {
                    CaptureInitialWindowConstraints();
                    QueueLayoutRefresh(System.Windows.Threading.DispatcherPriority.Loaded);
                }

                LogUtility.LogDebug($"[{_windowName}] load complete");
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex, $"DpiAwareWindow.OnLoaded ({_windowName})");
            }
        }

        private void QueueLayoutRefresh(System.Windows.Threading.DispatcherPriority priority)
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

                if (EnableExcelCentering && !_initialLayoutApplied)
                {
                    _initialLayoutApplied = true;
                    CenterOverOwnerOnce();
                }
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex, "DpiAwareWindow.ApplyLayoutRefresh");
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
                    QueueLayoutRefresh(System.Windows.Threading.DispatcherPriority.Background);
                    return;
                }

                if (Content is FrameworkElement root)
                {
                    root.InvalidateMeasure();
                    root.InvalidateArrange();
                }

                QueueLayoutRefresh(System.Windows.Threading.DispatcherPriority.Render);
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex, "DpiAwareWindow.RefreshWindowLayout");
            }
        }

        private IntPtr WndProc(IntPtr hwnd, int msg, IntPtr wParam, IntPtr lParam, ref bool handled)
        {
            try
            {
                const int WM_DPICHANGED = 0x02E0;

                if (msg == WM_DPICHANGED && !DisableAutoSizing)
                {
                    uint newDpi = (uint)wParam;
                    AdjustForDpiChange(newDpi, lParam);
                    handled = true;
                }
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex, "DpiAwareWindow.WndProc");
            }

            return IntPtr.Zero;
        }

        private void AdjustForCurrentDpi()
        {
            try
            {
                _currentScaleFactor = GetCurrentScaleFactor();
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex, "DpiAwareWindow.AdjustForCurrentDpi");
            }
        }

        private void AdjustForDpiChange(uint newDpi, IntPtr lParam)
        {
            try
            {
                var scaleFactor = GetScaleFactorFromDpi(newDpi);

                ApplyScaleTransform(scaleFactor);

                if (lParam != IntPtr.Zero)
                {
                    var rect = Marshal.PtrToStructure<Rect>(lParam);

                    this.Dispatcher.BeginInvoke(new Action(() =>
                    {
                        try
                        {
                            this.Left = rect.Left / scaleFactor;
                            this.Top = rect.Top / scaleFactor;
                            this.Width = rect.Width / scaleFactor;
                            this.Height = rect.Height / scaleFactor;

                            // Windows' suggested rect (above) only keeps the window under the
                            // cursor/at the same relative position during a DPI change - it has
                            // no idea about our own MaxWidthCap/content-fit rules, so a window
                            // dragged from a large, high-res monitor onto a smaller/lower-res one
                            // at a different scale can land larger than the new monitor's work
                            // area. Re-run the same clamp+recenter pass OnLoaded already does, so
                            // a live cross-monitor drag ends up exactly as constrained as a fresh
                            // open on that same monitor would be.
                            FitToAvailableWorkArea();
                        }
                        catch (Exception ex)
                        {
                            LogUtility.LogException(ex, "DpiAwareWindow.AdjustForDpiChange (resize)");
                        }
                    }));
                }
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex, "DpiAwareWindow.AdjustForDpiChange");
            }
        }

        private void ApplyScaleTransform(double scaleFactor)
        {
            if (Content is not FrameworkElement element)
            {
                return;
            }

            if (Math.Abs(scaleFactor - 1.0) < 0.001)
            {
                element.LayoutTransform = Transform.Identity;
                return;
            }

            if (Math.Abs(scaleFactor - _currentScaleFactor) < 0.001)
            {
                return;
            }

            try
            {
                _dpiScaleTransform.ScaleX = scaleFactor;
                _dpiScaleTransform.ScaleY = scaleFactor;
                element.LayoutTransform = _dpiScaleTransform;
                element.InvalidateMeasure();
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex, "DpiAwareWindow.ApplyScaleTransform");
            }
        }

        private void FitToAvailableWorkArea()
        {
            if (!AutoClampToWorkArea || DisableAutoSizing)
                return;

            try
            {
                if (Content is not FrameworkElement root)
                {
                    return;
                }

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

                // If an initial MinHeight was specified on the window, treat it as
                // a logical minimum content height when calculating fit and scale.
                // This prevents the window from being rendered too-small when no
                // explicit content requires space, while still allowing the
                // clamping logic below to reduce the window if the work area is
                // smaller (so it will not prevent clamping to the taskbar).
                if (!double.IsNaN(_initialMinHeight) && _initialMinHeight > 0)
                {
                    desiredHeight = Math.Max(desiredHeight, _initialMinHeight);
                }

                if (desiredWidth <= 0 || desiredHeight <= 0)
                    return;

                var rawScale = Math.Min(availableWidth / desiredWidth, availableHeight / desiredHeight);
                var fitScale = Math.Min(1.0, rawScale);

                if (MinContentScale > 0 && fitScale < MinContentScale)
                {
                    fitScale = MinContentScale;
                }

                ApplyScaleTransform(fitScale);

                var targetWidth = Math.Min(desiredWidth * fitScale, availableWidth);
                var targetHeight = Math.Min(desiredHeight * fitScale, availableHeight);

                // Capture the size/position as they stood before this method changes them,
                // so we can recenter around the same center point afterward (see
                // RecenterAfterSizeChange for why this matters - WindowStartupLocation only
                // centers the window once, and any resize after that anchors at the current
                // Left/Top, silently drifting the window off-center as it grows/shrinks).
                double previousLeft = Left;
                double previousTop = Top;
                double previousWidth = Width;
                double previousHeight = Height;
                bool sizeChanged = false;

                if (targetWidth > 0 && Math.Abs(targetWidth - previousWidth) > 0.5)
                {
                    Width = targetWidth;
                    sizeChanged = true;
                }

                if (targetHeight > 0 && Math.Abs(targetHeight - previousHeight) > 0.5)
                {
                    Height = targetHeight;
                    sizeChanged = true;
                }

                MaxWidth = availableWidth;
                MaxHeight = availableHeight;

                if (sizeChanged)
                {
                    RecenterAfterSizeChange(previousLeft, previousTop, previousWidth, previousHeight);
                }
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex, "DpiAwareWindow.FitToAvailableWorkArea");
            }
        }

        /// <summary>
        /// Recenters the window around the same center point it had before Width/Height
        /// were just changed by FitToAvailableWorkArea/EnsureFitsWorkArea, instead of
        /// leaving Left/Top untouched. A resize always grows/shrinks anchored at the
        /// window's current top-left corner, so without this, any post-centering resize
        /// (e.g. content growing once async data finishes loading, or the safety clamp in
        /// EnsureFitsWorkArea kicking in) silently drifts the window's true center away
        /// from wherever WindowStartupLocation originally centered it (typically
        /// CenterOwner against the Excel window) - this was the root cause of windows
        /// appearing off-center in the shipped MSI. Only called when this class itself
        /// changed Width/Height; never touches Left/Top for a plain user-initiated
        /// drag-resize (ResizeMode="CanResize"), since that doesn't go through either of
        /// those two methods unless the drag actually violates Min/MaxWidth/Height, in
        /// which case re-centering after the forced clamp is the correct behavior anyway.
        /// </summary>
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
                LogUtility.LogException(ex, "DpiAwareWindow.RecenterAfterSizeChange");
            }
        }

        /// <summary>
        /// Positions the window so its own center lands on (centerX, centerY) - i.e.
        /// Left/Top = center minus half of this window's own Width/Height - clamped so it
        /// can't be pushed off the visible work area. Shared by RecenterAfterSizeChange
        /// (recentering around the window's own previous center after a resize) and
        /// CenterOverOwnerOnce (centering around the owner/work-area's center on first
        /// layout). Centering must always subtract half of *this* window's size from the
        /// target center point - using the center point directly as Left/Top (as
        /// WindowStartupLocation effectively did here, since it ran before SizeToContent=
        /// "Manual" plus MinWidth/MinHeight had resolved this window's real size) leaves
        /// the window's left/top edge sitting at the center instead of the window's own
        /// center landing there.
        /// </summary>
        private void PositionAroundCenter(double centerX, double centerY)
        {
            if (double.IsNaN(Width) || double.IsNaN(Height) || Width <= 0 || Height <= 0)
                return;

            double newLeft = centerX - (Width / 2.0);
            double newTop = centerY - (Height / 2.0);

            var workArea = SystemParameters.WorkArea;
            if (Width < workArea.Width)
                newLeft = Math.Max(workArea.Left, Math.Min(newLeft, workArea.Right - Width));
            if (Height < workArea.Height)
                newTop = Math.Max(workArea.Top, Math.Min(newTop, workArea.Bottom - Height));

            Left = newLeft;
            Top = newTop;
        }

        /// <summary>
        /// Explicitly centers the window over its owner (or the work area, if there is no
        /// owner) using this window's real, post-layout Width/Height - called exactly once,
        /// right after the first FitToAvailableWorkArea pass has resolved the window's true
        /// size. This is the fix for windows appearing off-center: WindowStartupLocation=
        /// "CenterOwner" (set in XAML) already ran once by this point, but for a window
        /// using SizeToContent="Manual" with only MinWidth/MinHeight/MaxWidth constraints
        /// (no explicit Width/Height), WPF performs that positioning before layout has
        /// resolved the window's real MinWidth/MinHeight-driven size - so it centers using
        /// a placeholder width, leaving the window's left edge sitting near the owner's
        /// center instead of the window's own center landing there. Re-centering here with
        /// the now-final Width/Height corrects that, independent of whatever
        /// WindowStartupLocation computed beforehand.
        /// </summary>
        private void CenterOverOwnerOnce()
        {
            try
            {
                double centerX;
                double centerY;

                IntPtr ownerHwnd = new WindowInteropHelper(this).Owner;
                if (ownerHwnd != IntPtr.Zero && GetWindowRect(ownerHwnd, out Rect ownerRectPx) &&
                    ownerRectPx.Width > 0 && ownerRectPx.Height > 0)
                {
                    double scale = GetCurrentScaleFactor();
                    double ownerLeft = ownerRectPx.Left / scale;
                    double ownerTop = ownerRectPx.Top / scale;
                    double ownerWidth = ownerRectPx.Width / scale;
                    double ownerHeight = ownerRectPx.Height / scale;

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
                LogUtility.LogException(ex, "DpiAwareWindow.CenterOverOwnerOnce");
            }
        }

        private double GetCurrentScaleFactor()
        {
            try
            {
                if (_hwndSource?.CompositionTarget != null)
                {
                    var scale = _hwndSource.CompositionTarget.TransformToDevice.M11;
                    if (scale > 0)
                        return scale;
                }
            }
            catch
            {
                // fall through to fallback logic
            }

            try
            {
                var dpi = DpiAwarenessHelper.GetWindowDpi(this);
                if (dpi > 0)
                    return dpi / 96.0;
            }
            catch
            {
                // ignore and fall back to 1.0
            }

            return 1.0;
        }

        private static double GetScaleFactorFromDpi(uint dpi)
        {
            var scale = dpi / 96.0;
            return scale > 0 ? scale : 1.0;
        }

        protected double DipToPixels(double dip)
        {
            return dip * _currentScaleFactor;
        }

        protected double PixelsToDip(double pixels)
        {
            return pixels / _currentScaleFactor;
        }

        [DllImport("user32.dll")]
        private static extern uint GetDpiForWindow(IntPtr hwnd);

        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool GetWindowRect(IntPtr hWnd, out Rect lpRect);

        [StructLayout(LayoutKind.Sequential)]
        private struct Rect
        {
            public int Left;
            public int Top;
            public int Right;
            public int Bottom;
            public readonly int Width => Right - Left;
            public readonly int Height => Bottom - Top;
        }

        protected override void OnRenderSizeChanged(SizeChangedInfo sizeInfo)
        {
            base.OnRenderSizeChanged(sizeInfo);

            // Skip clamping if auto-sizing is disabled
            if (DisableAutoSizing || !AutoClampToWorkArea)
                return;

            try
            {
                EnsureFitsWorkArea();
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex, "DpiAwareWindow.OnRenderSizeChanged (clamp)");
            }
        }

        protected void EnsureFitsWorkArea(double? marginOverride = null)
        {
            // Skip if auto-sizing is disabled
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
                {
                    baseMaxWidth = Math.Min(baseMaxWidth, requestedMaxWidth);
                }

                if (MaxWidthCap.HasValue)
                {
                    baseMaxWidth = Math.Min(baseMaxWidth, MaxWidthCap.Value);
                }

                if (MaxHeightCap.HasValue)
                {
                    baseMaxHeight = Math.Min(baseMaxHeight, MaxHeightCap.Value);
                }

                var effectiveMaxWidth = double.IsPositiveInfinity(MaxWidth)
                    ? baseMaxWidth
                    : Math.Min(MaxWidth, baseMaxWidth);

                var effectiveMaxHeight = double.IsPositiveInfinity(MaxHeight)
                    ? baseMaxHeight
                    : Math.Min(MaxHeight, baseMaxHeight);

                MaxWidth = effectiveMaxWidth;
                MaxHeight = effectiveMaxHeight;

                if (MinWidth > effectiveMaxWidth)
                {
                    MinWidth = effectiveMaxWidth;
                }

                if (MinHeight > effectiveMaxHeight)
                {
                    MinHeight = effectiveMaxHeight;
                }

                if (Width > effectiveMaxWidth)
                {
                    Width = effectiveMaxWidth;
                    sizeChanged = true;
                }
                else if (Width < MinWidth)
                {
                    Width = MinWidth;
                    sizeChanged = true;
                }

                if (Height > effectiveMaxHeight)
                {
                    Height = effectiveMaxHeight;
                    sizeChanged = true;
                }
                else if (Height < MinHeight)
                {
                    Height = MinHeight;
                    sizeChanged = true;
                }

                // Only recenter when this method itself just forced a clamp - a plain
                // user drag-resize (ResizeMode="CanResize") stays within Min/MaxWidth/Height
                // and never reaches here, so ordinary manual resizing is left untouched.
                if (sizeChanged)
                    RecenterAfterSizeChange(previousLeft, previousTop, previousWidth, previousHeight);
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex, "DpiAwareWindow.EnsureFitsWorkArea");
            }
        }

        private void OnContentRenderedDebug(object sender, EventArgs e)
        {
            try
            {
                LogUtility.LogDebug($"[{_windowName}] content rendered");
            }
            catch
            {
                // swallow
            }
        }

        private void OnClosedDebug(object sender, EventArgs e)
        {
            try
            {
                LogUtility.LogDebug($"[{_windowName}] closed");
            }
            catch
            {
                // swallow
            }
        }

        // Show() (non-modal), unlike ShowDialog(), does not automatically reactivate
        // the owner when a window closes - the OS just falls through to whatever
        // window is next in its own activation history, which can be a completely
        // unrelated application (confirmed via a live test with WebView2PopupWindow:
        // focus landed on a background terminal window, not Excel). Centralized here
        // so every DpiAwareWindow-derived window (dialogs, popups, message/wait
        // windows) gets this for free on any close - normal, programmatic, or forced -
        // rather than each window needing its own copy of this fix.
        private void RestoreOwnerFocusOnClosed(object sender, EventArgs e)
        {
            try
            {
                if (Owner != null)
                {
                    // Owned by another WPF window (e.g. a popup owned by GLLogin) -
                    // reactivating it is enough; that window's own native ownership
                    // chain (set via SetExcelOwner) takes care of Excel in turn.
                    Owner.Activate();
                }
                else if (_excelOwnerHwnd != IntPtr.Zero)
                {
                    // Owned directly by Excel's native HWND (SetExcelOwner /
                    // ShowWithOwner / ShowDialogWithOwner) - Window.Owner is never set
                    // for this path since Excel isn't a WPF Window, so force Excel's
                    // own window back to the foreground explicitly.
                    SetForegroundWindow(_excelOwnerHwnd);
                }
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex, $"DpiAwareWindow.RestoreOwnerFocusOnClosed ({_windowName})");
            }
        }

        private void OnUnloadedDebug(object sender, RoutedEventArgs e)
        {
            try
            {
                LogUtility.LogDebug($"[{_windowName}] unloaded");
            }
            catch
            {
                // swallow
            }
        }

        private void CaptureInitialWindowConstraints()
        {
            if (double.IsNaN(_initialMaxWidth))
            {
                _initialMaxWidth = MaxWidth;
            }

            if (double.IsNaN(_initialMaxHeight))
            {
                _initialMaxHeight = MaxHeight;
            }

            if (double.IsNaN(_initialMinHeight))
            {
                _initialMinHeight = MinHeight;
            }
        }

        private double GetEffectiveRequestedMaxWidth()
        {
            var maxWidth = double.IsNaN(_initialMaxWidth) ? MaxWidth : _initialMaxWidth;

            if (double.IsPositiveInfinity(maxWidth))
                return double.PositiveInfinity;

            return maxWidth + 200;
        }
    }
}
