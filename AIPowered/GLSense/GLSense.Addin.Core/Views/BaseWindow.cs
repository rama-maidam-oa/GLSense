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
        private double _initialMinWidth = double.NaN;
        private double _initialMinHeight = double.NaN;
        private DispatcherTimer _resizeSettleTimer;

        public bool EnableAutoLayoutRefresh { get; set; } = true;
        public bool EnableExcelCentering { get; set; } = true;
        public bool EnableEscapeToClose { get; set; } = true;
        public bool AutoClampToWorkArea { get; set; } = true;
        public double WorkAreaMargin { get; set; } = 24d;
        public bool ModalToExcel { get; set; } = true;

        // Compat shim: under WPF-UI's old FluentWindow base class, this DP actively
        // suppressed the native Windows title bar via WindowChrome so the window's own
        // custom header (TitleBarGridStyle) could be the only visible caption/chrome.
        // BaseWindow no longer derives from FluentWindow, and plain Window has no
        // equivalent built-in mechanism - so without OnInitialized's handling below, every
        // one of the ~23 not-yet-ported windows (which all set
        // WindowStyle="SingleBorderWindow" + ExtendsContentIntoTitleBar="True" + their own
        // custom header row) would show a NATIVE Windows caption bar stacked directly above
        // their own custom header, both showing the title. See OnInitialized for the actual
        // fix. Remove this (and OnInitialized's handling of it) once every window has been
        // individually ported to set WindowStyle="None" itself in its own XAML (grep for
        // "ExtendsContentIntoTitleBar" across Views/*.xaml before deleting - same removal
        // pattern as IconSymbol/WindowCaption).
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
                this.KeyDown += BaseWindow_KeyDown;

                MouseWheelFocusHelper.EnableHoverToScroll(this);

                ServiceLocator.Logger?.LogDebug($"[{_windowName}] constructor completed");
            }
            catch (Exception ex)
            {
                ServiceLocator.Logger?.LogException(ex, $"[{_windowName}] constructor error");
            }
        }

        protected override void OnInitialized(EventArgs e)
        {
            base.OnInitialized(e);

            // Compat shim for not-yet-ported windows: under WPF-UI's old FluentWindow base
            // class, ExtendsContentIntoTitleBar="True" actively suppressed the native title
            // bar via WindowChrome. Under plain Window there's no equivalent mechanism, so
            // without this, every window still setting this flag (all ~23 not yet ported to
            // their own WindowStyle="None" header) would show a native Windows caption bar
            // stacked above their own custom TitleBarGridStyle header. Switching WindowStyle
            // to None here matches exactly what the already-ported pilot windows
            // (GLWaitWindow, GLMessageWindow, GLSegmentValues) do natively in their own XAML
            // - this is the real target end-state, not a workaround. Remove this override
            // once every window has been individually ported to set WindowStyle="None"
            // itself (grep for "ExtendsContentIntoTitleBar" across Views/*.xaml before
            // deleting).
            if (ExtendsContentIntoTitleBar && WindowStyle == WindowStyle.SingleBorderWindow)
            {
                WindowStyle = WindowStyle.None;
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
            }
        }

        // Bubbling (not tunneling) - fires only after every child control has had a chance
        // to mark the key handled for its own purpose (closing a dropdown, canceling a cell edit).
        private void BaseWindow_KeyDown(object sender, KeyEventArgs e)
        {
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
                // AppOverlay lives in this same assembly/namespace, so a direct typed check
                // is used instead of reflection-by-property-name (which DismissActiveToast
                // above still needs, since AppOverlay's DismissToast is only reachable via
                // a method lookup there for unrelated reasons - left as-is).
                if (FindName("AppOverlayControl") is AppOverlay overlay)
                    return overlay.IsBusyVisible || overlay.IsConfirmVisible;
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

                ApplyScaleTransform(scaleFactor);
                _currentScaleFactor = scaleFactor;

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

                if (!double.IsNaN(_initialMinWidth) && _initialMinWidth > 0)
                    desiredWidth = Math.Max(desiredWidth, _initialMinWidth);

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

                // Recenter unconditionally whenever this method itself just changed
                // Width/Height - NOT gated on EnableExcelCentering/CenterInExcel. Those
                // flags only govern the one-time CenterOverExcelOnce() initial-placement
                // pass; they say nothing about whether a window that was already
                // positioned some other way (e.g. WindowStartupLocation="CenterOwner",
                // used by GLSegmentValues/GLSegmentRef/GLSegmentManager/
                // GLBalanceConfigurator, all of which set EnableExcelCentering=false)
                // should be allowed to silently drift off-center when THIS method resizes
                // it. A resize always grows/shrinks anchored at the current top-left
                // corner, so without recentering here, any of those 4 windows would
                // reintroduce the exact off-center bug CLAUDE.md section 33 already fixed.
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

        // Lighter-weight counterpart to FitToAvailableWorkArea, ported verbatim (aside from
        // the logging call) from the reference DpiAwareWindow.EnsureFitsWorkArea. Unlike
        // FitToAvailableWorkArea (meant for initial load - remeasures Content at infinite
        // available size and rescales via LayoutTransform), this only clamps the window's
        // already-resolved Width/Height against its own Min/Max bounds - no content
        // remeasure, no rescale. This is the correct method for OnRenderSizeChanged's
        // debounced post-show reclamp: a window whose content changes size well after it's
        // already visible (e.g. AppOverlay.EnsureOwnerWindowRoomForConfirm deliberately
        // growing Height to fit a confirm popup) should have that new size clamped to fit
        // the work area, not have its entire content remeasured from scratch and rescaled -
        // which previously caused the just-grown Height to be measured back down again,
        // shrinking a popup whose own MaxHeight is bound to a percentage of the window's
        // Height (see OnRenderSizeChanged's own comment for the full bug history).
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
            if (double.IsNaN(_initialMinWidth)) _initialMinWidth = MinWidth;
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
                    // EnsureFitsWorkArea, NOT FitToAvailableWorkArea - the latter remeasures
                    // Content at infinite available size and rescales via LayoutTransform,
                    // which is only correct for the INITIAL fit-to-content pass. Calling it
                    // here (a post-show debounced reclamp) re-measured content from scratch
                    // and could shrink a size a caller had just deliberately grown (e.g.
                    // AppOverlay.EnsureOwnerWindowRoomForConfirm growing Height to fit a
                    // confirm popup) right back down. EnsureFitsWorkArea only clamps the
                    // already-resolved Width/Height against Min/Max - see its own comment.
                    EnsureFitsWorkArea();
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

                // Stop the debounced resize-settle timer (OnRenderSizeChanged) so it can't
                // tick once more and call EnsureFitsWorkArea/ForceFrameRedraw against an
                // already-closed window.
                _resizeSettleTimer?.Stop();
                _resizeSettleTimer = null;
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

        [DllImport("user32.dll")]
        private static extern bool SetWindowPos(IntPtr hWnd, IntPtr hWndInsertAfter, int X, int Y, int cx, int cy, uint uFlags);

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

        // WindowStyle="None" windows still get a DWM-drawn drop shadow around their real
        // client area, and resizing programmatically (Width/Height set from code, not a
        // user drag) can leave stale rendering behind at the old edge - see CLAUDE.md
        // section 1.4d and the reference DpiAwareWindow.ForceFrameRedraw for the original,
        // screenshot-confirmed defect this fixes. SWP_FRAMECHANGED alone only recomputes
        // the non-client frame/shadow, not the client area, so RedrawWindow with
        // INVALIDATE|ERASE|FRAME|ALLCHILDREN|UPDATENOW forces an immediate full
        // erase-and-repaint regardless of the exact cause. NOMOVE/NOSIZE/NOZORDER/
        // NOACTIVATE mean this call itself never moves/resizes/activates anything -
        // Width/Height/Left/Top are already set separately by the caller.
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
