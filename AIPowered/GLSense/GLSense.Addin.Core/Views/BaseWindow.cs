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

                ApplyScaleTransform(scaleFactor);
                _currentScaleFactor = scaleFactor;

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
