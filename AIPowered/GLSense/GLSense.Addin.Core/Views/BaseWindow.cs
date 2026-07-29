using GLSense.Addin.Core.Helpers;
using GLSense.Addin.Core.Infrastructure;
using GLSense.Addin.Core.Utilities;
using System;
using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Input;
using System.Windows.Interop;
using System.Windows.Media;
using Wpf.Ui.Controls;

namespace GLSense.Addin.Core.Views
{
    public abstract class BaseWindow : FluentWindow
    {
        private HwndSource _hwndSource;
        private double _currentScaleFactor = 1.0;
        private readonly ScaleTransform _dpiScaleTransform = new ScaleTransform(1.0, 1.0);
        private readonly string _windowName;
        private bool _isDragging;
        private Point _dragStartPoint;
        private IntPtr _excelHandle;
        private bool _ownerSet;

        public bool EnableAutoLayoutRefresh { get; set; } = true;
        public bool EnableExcelCentering { get; set; } = true;
        public bool EnableEscapeToClose { get; set; } = true;
        public bool AutoClampToWorkArea { get; set; } = true;
        public double WorkAreaMargin { get; set; } = 24d;

        public string WindowCaption
        {
            get => Title;
            set => Title = value;
        }

        // FontAwesome PackIconFontAwesomeKind name (e.g. "KeySolid") bound to each
        // window's title-bar iconPacks:PackIconFontAwesome via TitleBarIconStyle - kept
        // as a plain string (not the enum type) since WPF's binding engine coerces a
        // string source into an enum-typed target property automatically via the
        // enum's default TypeConverter, so every derived window can keep just setting
        // IconSymbol="SomeKind" in XAML like before. Switched from WPF-UI's Symbol
        // names to FontAwesome Kind names project-wide so every icon in the app comes
        // from the same, already-proven icon set FinalWorkingCode uses everywhere.
        public string IconSymbol { get; set; } = "KeySolid";
        public bool CenterInExcel { get; set; } = true;
        public bool ModalToExcel { get; set; } = true;

        protected BaseWindow()
        {
            try
            {
                _windowName = GetType().Name;
                ServiceLocator.Logger?.LogDebug($"[{_windowName}] constructing window");

                // Initialize WPF-UI
                if (!WpfUiBootstrapper.IsInitialized)
                {
                    WpfUiBootstrapper.Initialize();
                    WpfUiBootstrapper.SetLightTheme();
                }

                // Set up DPI awareness with PerMonitorV2
                using (DpiAwarenessHelper.SetPerMonitorAware())
                {
                    this.UseLayoutRounding = true;
                    this.SnapsToDevicePixels = true;
                    TextOptions.SetTextFormattingMode(this, TextFormattingMode.Display);
                    RenderOptions.SetBitmapScalingMode(this, BitmapScalingMode.HighQuality);
                }

                // Get Excel handle from ServiceLocator
                try
                {
                    _excelHandle = ServiceLocator.ExcelHandle;
                    if (_excelHandle != IntPtr.Zero)
                    {
                        ServiceLocator.Logger?.LogDebug($"[{_windowName}] Excel handle obtained from ServiceLocator: {_excelHandle}");
                    }
                    else
                    {
                        ServiceLocator.Logger?.LogWarn($"[{_windowName}] Excel handle is IntPtr.Zero from ServiceLocator");
                    }
                }
                catch (Exception ex)
                {
                    ServiceLocator.Logger?.LogException(ex, $"[{_windowName}] Failed to get Excel handle from ServiceLocator");
                }

                // Set up events
                this.SourceInitialized += OnSourceInitialized;
                this.Loaded += OnLoaded;

                // See CLAUDE.md section 1.4e. OnLoaded's resettle runs synchronously
                // inside the Loaded event - i.e. BEFORE ShowDialog()'s nested message
                // loop has ever painted this window on screen. ForceSizeToContentResettle's
                // SetWindowPos(SWP_FRAMECHANGED) call (1.4d) only has something to force
                // Windows/DWM to recompose once the window actually has a first frame on
                // screen to begin with - calling it pre-paint is a no-op, same as it would
                // be for a user "clicking the resize border" on a window that isn't even
                // visible yet. Windows whose own Loaded handler chains an async
                // continuation (GLAbout/GLJobsMonitor/GLRollerGroups) or fires a
                // DataLoadedAction callback (GLLOVs/GLSegmentRef/GLSegmentManager/
                // GLSegmentValues) happen to get a second resettle call for free, always
                // reached well after first paint - that's an accident of their timing, not
                // something GLCubeDetails/GLServerConfiguration/GLMessageWindow have.
                // ContentRendered is the one WPF event guaranteed to fire only after the
                // window's content has actually been rendered/painted for the first time,
                // so hooking the SAME resettle there gives every BaseWindow-derived window
                // a guaranteed post-paint pass without needing a bespoke per-window hook.
                this.ContentRendered += OnContentRendered;

                // Escape-to-close for every BaseWindow-derived dialog, opt-out via
                // EnableEscapeToClose (GLWaitWindow sets this false so Cancel is the only
                // way out). Wired on the bubbling KeyDown (not PreviewKeyDown) and gated on
                // !e.Handled so a DataGridCell/ComboBox that already consumed Escape for its
                // own purpose (cancel cell edit, close dropdown) is left alone instead of
                // also closing the whole window.
                this.KeyDown += BaseWindow_KeyDown;

                // Ensures mouse wheel scrolling works on hover for every window derived
                // from this base class - see MouseWheelFocusHelper for the root cause
                // (harmless no-op here since a shown top-level Window already has focus;
                // the fix matters for content that's HWND-reparented into a host task
                // pane, e.g. the Balance Configurator).
                MouseWheelFocusHelper.EnableHoverToScroll(this);

                ServiceLocator.Logger?.LogDebug($"[{_windowName}] constructor completed with DPI awareness");
            }
            catch (Exception ex)
            {
                ServiceLocator.Logger?.LogException(ex, $"[{_windowName}] constructor error");
            }
        }

        private void OnSourceInitialized(object sender, EventArgs e)
        {
            try
            {
                ServiceLocator.Logger?.LogDebug($"[{_windowName}] source initialized");
                _hwndSource = PresentationSource.FromVisual(this) as HwndSource;

                if (_hwndSource != null)
                {
                    _hwndSource.AddHook(WndProc);
                }

                // Set Excel as owner using ServiceLocator.ExcelHandle
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
                else if (ModalToExcel && _excelHandle == IntPtr.Zero)
                {
                    ServiceLocator.Logger?.LogWarn($"[{_windowName}] Cannot set Excel owner - handle is IntPtr.Zero");
                }

                // Get initial DPI
                _currentScaleFactor = DpiAwarenessHelper.GetWindowDpi(this) / 96.0;
                ServiceLocator.Logger?.LogDebug($"[{_windowName}] Initial DPI scale: {_currentScaleFactor}");
            }
            catch (Exception ex)
            {
                ServiceLocator.Logger?.LogException(ex, $"[{_windowName}] OnSourceInitialized error");
            }
        }

        private void OnLoaded(object sender, RoutedEventArgs e)
        {
            try
            {
                ServiceLocator.Logger?.LogDebug($"[{_windowName}] loaded - applying DPI adjustments");

                // Center in Excel if enabled
                if (CenterInExcel)
                {
                    CenterWindowInExcel();
                }

                // Apply DPI adjustments
                AdjustForCurrentDpi();
                FitToAvailableWorkArea();

                // Ensure window is visible and on top of Excel
                if (_ownerSet && _excelHandle != IntPtr.Zero)
                {
                    this.Focus();
                }

                // SizeToContent windows can settle on a stale/undersized first measurement
                // (e.g. a gap after the last DataGrid column, or blank space below the
                // footer) that only corrects itself once the user manually resizes the
                // window - which forces WPF to fully redo its layout from scratch. Toggling
                // SizeToContent off and back on has the same effect, so do that once
                // automatically instead of relying on the user to nudge the window.
                //
                // This used to be deferred via Dispatcher.BeginInvoke(..., ContextIdle),
                // gambling that the dispatcher queue would still be busy long enough (e.g.
                // from JIT/resource-loading on a "cold" first open of a given window type)
                // for the ContextIdle callback to fire and fix the layout before the first
                // frame was actually painted. That worked by luck on a cold first open, but
                // on a "warm" reopen of the same window (JIT already done, styles already
                // loaded, far less dispatcher traffic) the window's stale/gappy first layout
                // pass would win the race and get painted before ContextIdle ever fired -
                // exactly matching the reported "looks right the first time, distorted/gappy
                // on close+reopen" symptom, and affecting every window using this base class.
                //
                // Fix: run the resettle synchronously right now, then pump a nested
                // dispatcher frame (WPF's "DoEvents" equivalent - PumpDispatcherFrame below)
                // so any pending native WM_SIZE/layout work queued as part of it is fully
                // flushed before this handler returns and the window is actually presented
                // on screen. This removes the timing race entirely instead of depending on
                // how busy the dispatcher happens to be on any given open.
                if (this.SizeToContent != SizeToContent.Manual)
                {
                    ForceSizeToContentResettle();
                    PumpDispatcherFrame();
                }

                ServiceLocator.Logger?.LogDebug($"[{_windowName}] load complete");
            }
            catch (Exception ex)
            {
                ServiceLocator.Logger?.LogException(ex, $"[{_windowName}] OnLoaded error");
            }
        }

        // See CLAUDE.md section 1.4e. Guaranteed to fire only after this window's
        // content has actually been rendered/painted on screen for the first time -
        // unlike OnLoaded's own resettle (which runs before that first paint), this one
        // can actually make SetWindowPos(SWP_FRAMECHANGED) do something, the same way a
        // user's resize-border click only works on a window that's already visible.
        // Gives every BaseWindow-derived window a guaranteed post-paint resettle pass
        // without depending on that window happening to have its own async Loaded
        // continuation or DataLoadedAction callback landing after paint by accident.
        private void OnContentRendered(object sender, EventArgs e)
        {
            try
            {
                if (this.SizeToContent != SizeToContent.Manual)
                {
                    ForceSizeToContentResettle();
                    PumpDispatcherFrame();
                }
            }
            catch (Exception ex)
            {
                ServiceLocator.Logger?.LogException(ex, $"[{_windowName}] OnContentRendered error");
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

        // Protected (not private): windows whose content genuinely loads asynchronously
        // after Loaded (e.g. GLCubeDetails' DataGrid, populated many awaits deep inside
        // its own Loaded handler, well after OnLoaded's own synchronous resettle below
        // has already run against an empty grid) need to trigger this same resettle
        // again once their real content is actually in place - see CLAUDE.md section 1
        // for the full history of this mechanism.
        protected void ForceSizeToContentResettle()
        {
            try
            {
                // Capture the position/size the window had before this resettle runs, so
                // that if the resettle actually changes its rendered size (the whole point
                // of this method - forcing SizeToContent to regrow/reflow to real content),
                // we can recenter it afterward instead of leaving Left/Top untouched. A
                // resize always grows/shrinks anchored at the current top-left corner, so
                // without this, every resettle call (there are several: once in OnLoaded,
                // again in OnContentRendered, again from the DPI-change handler for
                // auto-sized windows) silently drifts the window away from wherever
                // CenterWindowInExcel originally centered it - the reported "windows not
                // centered" MSI bug. ActualWidth/ActualHeight (not Width/Height) are used
                // here since SizeToContent windows don't reliably keep the Width/Height DPs
                // in sync with their true rendered size.
                double previousLeft = this.Left;
                double previousTop = this.Top;
                double previousWidth = this.ActualWidth > 0 ? this.ActualWidth : this.Width;
                double previousHeight = this.ActualHeight > 0 ? this.ActualHeight : this.Height;

                var mode = this.SizeToContent;
                this.SizeToContent = SizeToContent.Manual;
                this.UpdateLayout();

                // The UpdateLayout() calls here only flush WPF's own logical
                // measure/arrange tree - they don't reliably force the underlying Win32
                // HWND to actually resize to match on every machine/timing (WPF's
                // SizeToContent -> native window resize hookup isn't guaranteed to be
                // synchronously flushed by InvalidateMeasure/UpdateLayout alone). That's
                // exactly why windows like GLMessageWindow kept rendering with stale,
                // oversized chrome (extra blank space beyond the button/content, or a
                // gap to the right of the close button) even after this method's two
                // SizeToContent toggles above - until the user's own manual drag-resize
                // forced a genuine native WM_SIZE round-trip. Nudging Width/Height by a
                // full pixel (not the sub-pixel 0.1 this used to be - see below) while
                // SizeToContent is Manual (so the sets actually take effect instead of
                // being immediately overwritten) forces that same round-trip
                // programmatically instead of relying on the user to do it.
                if (this.ActualWidth > 0 && this.ActualHeight > 0)
                {
                    // This used to nudge by only +0.1 logical units. At every DPI scale
                    // factor WPF actually renders at (100%, 125%, 150%...), 0.1 logical
                    // units rounds to LESS THAN ONE physical device pixel - meaning the
                    // native HWND's actual on-screen size never changed at all, so this
                    // "nudge" was likely a near-total no-op at the Win32 level the entire
                    // time. That lines up exactly with continuing reports that the gap
                    // persists on windows still relying on this resettle (GLCubeDetails,
                    // GLServerConfiguration) even after 1.4b/1.4c wired the resettle to
                    // re-run once real content loads - re-running a no-op still does
                    // nothing. A full pixel guarantees an actual, measurable device-pixel
                    // size change regardless of DPI scale.
                    this.Width = this.ActualWidth + 1.0;
                    this.Height = this.ActualHeight + 1.0;
                    this.UpdateLayout();
                }

                this.SizeToContent = mode;
                this.UpdateLayout();

                // Belt-and-braces: explicitly force Windows to recompute this window's
                // non-client frame and let DWM recompose it, WITHOUT moving or resizing
                // it (SWP_NOMOVE/SWP_NOSIZE) - this is the actual mechanism a user
                // clicking (not even dragging) the resize border triggers internally via
                // the modal sizing loop, and is a well-known, more direct way to force a
                // stale-looking window frame to redraw than hoping a logical WPF
                // Width/Height nudge produces a large enough physical pixel delta.
                if (_hwndSource?.Handle != null && _hwndSource.Handle != IntPtr.Zero)
                {
                    SetWindowPos(_hwndSource.Handle, IntPtr.Zero, 0, 0, 0, 0,
                        SWP_NOMOVE | SWP_NOSIZE | SWP_NOZORDER | SWP_NOACTIVATE | SWP_FRAMECHANGED);
                }

                if (CenterInExcel)
                {
                    double newWidth = this.ActualWidth > 0 ? this.ActualWidth : this.Width;
                    double newHeight = this.ActualHeight > 0 ? this.ActualHeight : this.Height;

                    if (Math.Abs(newWidth - previousWidth) > 0.5 || Math.Abs(newHeight - previousHeight) > 0.5)
                    {
                        RecenterAfterSizeChange(previousLeft, previousTop, previousWidth, previousHeight, newWidth, newHeight);
                    }
                }

                ServiceLocator.Logger?.LogDebug($"[{_windowName}] SizeToContent resettled ({mode})");
            }
            catch (Exception ex)
            {
                ServiceLocator.Logger?.LogException(ex, $"[{_windowName}] ForceSizeToContentResettle error");
            }
        }

        // WPF's classic "DoEvents" equivalent: pushes a nested dispatcher frame that
        // processes every pending operation at Background priority and above (which
        // covers Send/Normal/DataBind/Render/Loaded/Background - i.e. essentially
        // everything short of the Idle bands) until the frame is told to stop. Used
        // right after ForceSizeToContentResettle() so the native HWND resize it
        // triggers is guaranteed to be fully flushed - synchronously, deterministically -
        // before OnLoaded returns and the window is actually painted on screen, instead
        // of hoping a deferred ContextIdle callback happens to win a race against the
        // first paint.
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

        private void CenterWindowInExcel()
        {
            try
            {
                if (_excelHandle == IntPtr.Zero)
                {
                    ServiceLocator.Logger?.LogWarn("CenterWindowInExcel: Excel handle is IntPtr.Zero");
                    return;
                }

                if (!GetWindowRect(_excelHandle, out RECT excelRect))
                {
                    ServiceLocator.Logger?.LogWarn("CenterWindowInExcel: Failed to get Excel window rect");
                    return;
                }

                var excelWidth = excelRect.Right - excelRect.Left;
                var excelHeight = excelRect.Bottom - excelRect.Top;
                var windowWidth = this.ActualWidth > 0 ? this.ActualWidth : this.Width;
                var windowHeight = this.ActualHeight > 0 ? this.ActualHeight : this.Height;

                var left = excelRect.Left + (excelWidth - windowWidth) / 2;
                var top = excelRect.Top + (excelHeight - windowHeight) / 2;

                left = Math.Max(excelRect.Left + 10, Math.Min(excelRect.Right - windowWidth - 10, left));
                top = Math.Max(excelRect.Top + 10, Math.Min(excelRect.Bottom - windowHeight - 10, top));

                this.Left = left;
                this.Top = top;

                ServiceLocator.Logger?.LogDebug($"Centered at Left={left}, Top={top}");
            }
            catch (Exception ex)
            {
                ServiceLocator.Logger?.LogException(ex, "CenterWindowInExcel error");
            }
        }

        /// <summary>
        /// Recenters the window around the same center point it had before Width/Height
        /// just changed (in ForceSizeToContentResettle or FitToAvailableWorkArea), instead
        /// of leaving Left/Top untouched. See ForceSizeToContentResettle's comment for the
        /// full reasoning - this is the fix for the reported "windows not centered on
        /// screen" bug in the shipped MSI, ported from FinalWorkingCode's identical fix in
        /// DpiAwareWindow.cs. Takes explicit before/after width/height (rather than reading
        /// Width/Height itself) because callers need different sources depending on
        /// whether SizeToContent is involved: FitToAvailableWorkArea assigns the Width/Height
        /// DPs directly, while ForceSizeToContentResettle's real size change is only
        /// reliably observable via ActualWidth/ActualHeight once UpdateLayout() has run.
        /// </summary>
        private void RecenterAfterSizeChange(double previousLeft, double previousTop,
            double previousWidth, double previousHeight, double newWidth, double newHeight)
        {
            try
            {
                if (double.IsNaN(previousLeft) || double.IsNaN(previousTop) ||
                    double.IsNaN(previousWidth) || double.IsNaN(previousHeight) ||
                    previousWidth <= 0 || previousHeight <= 0 ||
                    double.IsNaN(newWidth) || double.IsNaN(newHeight) ||
                    newWidth <= 0 || newHeight <= 0)
                {
                    return;
                }

                double centerX = previousLeft + (previousWidth / 2.0);
                double centerY = previousTop + (previousHeight / 2.0);

                double newLeft = centerX - (newWidth / 2.0);
                double newTop = centerY - (newHeight / 2.0);

                // Clamp so recentering never pushes the window off the visible work area
                // (e.g. if the old center point was near a screen edge).
                var workArea = SystemParameters.WorkArea;
                if (newWidth < workArea.Width)
                    newLeft = Math.Max(workArea.Left, Math.Min(newLeft, workArea.Right - newWidth));
                if (newHeight < workArea.Height)
                    newTop = Math.Max(workArea.Top, Math.Min(newTop, workArea.Bottom - newHeight));

                this.Left = newLeft;
                this.Top = newTop;

                ServiceLocator.Logger?.LogDebug($"[{_windowName}] Recentered after size change: Left={newLeft}, Top={newTop}");
            }
            catch (Exception ex)
            {
                ServiceLocator.Logger?.LogException(ex, $"[{_windowName}] RecenterAfterSizeChange error");
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
                    case WM_DPICHANGED:
                        var newDpi = (uint)wParam;
                        AdjustForDpiChange(newDpi, lParam);
                        handled = true;
                        break;

                    case WM_ACTIVATE:
                        if ((int)wParam == WA_ACTIVE || (int)wParam == WA_CLICKACTIVE)
                        {
                            if (_ownerSet && _excelHandle != IntPtr.Zero)
                            {
                                this.Dispatcher.BeginInvoke(new Action(() =>
                                {
                                    try
                                    {
                                        if (this.IsVisible && !this.IsActive)
                                        {
                                            this.Focus();
                                        }
                                    }
                                    catch (Exception ex)
                                    {
                                        ServiceLocator.Logger?.LogDebug($"WM_ACTIVATE handler error: {ex.Message}");
                                    }
                                }), System.Windows.Threading.DispatcherPriority.Background);
                            }
                        }
                        break;
                }
            }
            catch (Exception ex)
            {
                ServiceLocator.Logger?.LogException(ex, $"WndProc error for msg {msg}");
            }

            return IntPtr.Zero;
        }

        private void AdjustForCurrentDpi()
        {
            try
            {
                _currentScaleFactor = DpiAwarenessHelper.GetWindowDpi(this) / 96.0;
                ServiceLocator.Logger?.LogDebug($"[{_windowName}] Current DPI scale: {_currentScaleFactor}");
            }
            catch (Exception ex)
            {
                ServiceLocator.Logger?.LogException(ex, "AdjustForCurrentDpi error");
            }
        }

        private void AdjustForDpiChange(uint newDpi, IntPtr lParam)
        {
            try
            {
                var scaleFactor = newDpi / 96.0;
                _currentScaleFactor = scaleFactor;

                if (Content is FrameworkElement element)
                {
                    if (Math.Abs(scaleFactor - 1.0) > 0.001)
                    {
                        _dpiScaleTransform.ScaleX = scaleFactor;
                        _dpiScaleTransform.ScaleY = scaleFactor;
                        element.LayoutTransform = _dpiScaleTransform;
                        element.InvalidateMeasure();
                    }
                    else
                    {
                        element.LayoutTransform = Transform.Identity;
                    }
                }

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

                            // For SizeToContent windows, forcing Width/Height from Windows'
                            // suggested DPI rect here "freezes" the window at whatever
                            // intermediate size it happened to be measured at when this
                            // (very common, near-immediate-on-open) DPI notification arrived -
                            // overriding SizeToContent's own natural, content-fit measurement.
                            // That is exactly what caused windows to open with a stale/wrong
                            // size (a gap after the last DataGrid column, or blank space below
                            // the footer) that only self-corrected once the user manually
                            // resized the window and forced a fresh layout pass. Leaving
                            // Width/Height alone here lets SizeToContent re-measure at the new
                            // DPI scale on its own (the LayoutTransform update above already
                            // invalidates measure), which produces the correct size the first
                            // time instead of only after a later, unrelated resize.
                            if (!autoSized)
                            {
                                this.Width = rect.Width / scaleFactor;
                                this.Height = rect.Height / scaleFactor;
                            }
                            else
                            {
                                // The LayoutTransform + InvalidateMeasure() applied above only
                                // flush WPF's own logical measure/arrange tree - exactly the
                                // same category of problem ForceSizeToContentResettle()/
                                // PumpDispatcherFrame() were built to fix for the initial-open
                                // case in OnLoaded (a logical invalidate that never reliably
                                // forces a genuine native HWND resize on its own). This add-in
                                // only applies Per-Monitor-V2 DPI awareness to a scoped thread
                                // context (see DpiAwarenessHelper.SetPerMonitorAware(), since
                                // Excel's own process isn't PMv2-manifested), so on any monitor
                                // running above 100% scaling (125%/150% - the norm on most
                                // business laptops) WM_DPICHANGED can fire moments after
                                // OnLoaded's own resettle already completed, silently
                                // re-collapsing the window back to a stale measurement and
                                // reproducing the exact "gap until a manual resize" symptom -
                                // this time driven by the DPI change notification instead of
                                // the original "*" row measurement bug. Re-running the same
                                // resettle here (which OnLoaded doesn't repeat, since
                                // WM_DPICHANGED can arrive well after Loaded has already fired)
                                // closes that gap instead of waiting for the user to resize.
                                ForceSizeToContentResettle();
                                PumpDispatcherFrame();
                            }
                        }
                        catch (Exception ex)
                        {
                            ServiceLocator.Logger?.LogException(ex, "AdjustForDpiChange resize error");
                        }
                    }));
                }

                ServiceLocator.Logger?.LogDebug($"[{_windowName}] DPI changed to {newDpi}, scale: {scaleFactor}");
            }
            catch (Exception ex)
            {
                ServiceLocator.Logger?.LogException(ex, "AdjustForDpiChange error");
            }
        }

        private void FitToAvailableWorkArea()
        {
            if (!AutoClampToWorkArea)
                return;

            try
            {
                var workArea = SystemParameters.WorkArea;
                var availableWidth = Math.Max(0, workArea.Width - (WorkAreaMargin * 2));
                var availableHeight = Math.Max(0, workArea.Height - (WorkAreaMargin * 2));

                if (this.MaxWidthCap.HasValue)
                    availableWidth = Math.Min(availableWidth, this.MaxWidthCap.Value);
                if (this.MaxHeightCap.HasValue)
                    availableHeight = Math.Min(availableHeight, this.MaxHeightCap.Value);

                // Take the smaller of the screen-derived limit and whatever MaxWidth/
                // MaxHeight the window itself already declares in XAML (e.g.
                // GLCubeDetails' MaxWidth="750" MaxHeight="650"). Unconditionally
                // assigning here used to blow those tighter, author-chosen caps away
                // with the much larger screen work-area size (or the 1400 default
                // MaxWidthCap), which is why SizeToContent windows appeared to ignore
                // their declared Min/Max and grow past the intended bounds - e.g.
                // GLCubeDetails expanding off-screen once its DataGrid populated with
                // long ledger names, since nothing smaller than ~1400px/screen-height
                // was left to stop it.
                this.MaxWidth = Math.Min(this.MaxWidth, availableWidth);
                this.MaxHeight = Math.Min(this.MaxHeight, availableHeight);

                double previousLeft = this.Left;
                double previousTop = this.Top;
                double previousWidth = this.Width;
                double previousHeight = this.Height;
                bool sizeChanged = false;

                if (this.Width > availableWidth)
                {
                    this.Width = availableWidth;
                    sizeChanged = true;
                }
                if (this.Height > availableHeight)
                {
                    this.Height = availableHeight;
                    sizeChanged = true;
                }

                if (sizeChanged && CenterInExcel)
                    RecenterAfterSizeChange(previousLeft, previousTop, previousWidth, previousHeight, this.Width, this.Height);
            }
            catch (Exception ex)
            {
                ServiceLocator.Logger?.LogException(ex, "FitToAvailableWorkArea error");
            }
        }

        // Win32 API imports
        [DllImport("user32.dll")]
        private static extern bool GetWindowRect(IntPtr hWnd, out RECT lpRect);

        [DllImport("user32.dll")]
        private static extern uint GetDpiForWindow(IntPtr hwnd);

        [DllImport("user32.dll")]
        private static extern bool IsWindow(IntPtr hWnd);

        [DllImport("user32.dll")]
        private static extern bool SetForegroundWindow(IntPtr hWnd);

        [DllImport("user32.dll")]
        private static extern bool BringWindowToTop(IntPtr hWnd);

        [DllImport("user32.dll")]
        private static extern IntPtr SetParent(IntPtr hWndChild, IntPtr hWndNewParent);

        [DllImport("user32.dll")]
        private static extern IntPtr GetParent(IntPtr hWnd);

        [DllImport("user32.dll")]
        private static extern int GetWindowThreadProcessId(IntPtr hWnd, out int lpdwProcessId);

        [DllImport("user32.dll", SetLastError = true)]
        private static extern bool SetWindowPos(IntPtr hWnd, IntPtr hWndInsertAfter,
            int X, int Y, int cx, int cy, uint uFlags);

        private const uint SWP_NOMOVE = 0x0002;
        private const uint SWP_NOSIZE = 0x0001;
        private const uint SWP_NOZORDER = 0x0004;
        private const uint SWP_NOACTIVATE = 0x0010;
        private const uint SWP_FRAMECHANGED = 0x0020;

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

        public double? MaxWidthCap { get; set; } = 1400d;
        public double? MaxHeightCap { get; set; } = null;

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

        // ShowDialog with proper Excel ownership using ServiceLocator
        public new bool? ShowDialog()
        {
            try
            {
                ServiceLocator.Logger?.LogDebug($"[{_windowName}] ShowDialog called");

                if (!this.Dispatcher.CheckAccess())
                {
                    return this.Dispatcher.Invoke(() => ShowDialog());
                }

                // Ensure Excel owner is set before showing
                if (ModalToExcel && !_ownerSet)
                {
                    // Get fresh Excel handle from ServiceLocator
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
                {
                    this.ShowDialog();
                }
                else
                {
                    this.Dispatcher.Invoke(() => this.ShowDialog());
                }
            }
            catch (Exception ex)
            {
                ServiceLocator.Logger?.LogException(ex, $"[{_windowName}] SafeShowDialog error");
            }
        }
    }
}