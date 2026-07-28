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
        private HwndSource _hwndSource;
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
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex, "DpiAwareWindow.SetExcelOwner");
            }
        }

        public bool? ShowDialogWithOwner(IntPtr excelHwnd)
        {
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

                if (targetWidth > 0)
                    Width = targetWidth;

                if (targetHeight > 0)
                    Height = targetHeight;

                MaxWidth = availableWidth;
                MaxHeight = availableHeight;
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex, "DpiAwareWindow.FitToAvailableWorkArea");
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
                }
                else if (Width < MinWidth)
                {
                    Width = MinWidth;
                }

                if (Height > effectiveMaxHeight)
                {
                    Height = effectiveMaxHeight;
                }
                else if (Height < MinHeight)
                {
                    Height = MinHeight;
                }
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