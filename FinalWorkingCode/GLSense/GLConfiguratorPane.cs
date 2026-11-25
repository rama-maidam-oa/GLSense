using AddinExpress.XL;
using GLSense.Utilities;
using GLSense.Views;
using System;
using System.Drawing;
using System.Runtime.InteropServices;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Windows.Forms.Integration;

namespace GLSense
{
    public partial class GLConfiguratorPane : AddinExpress.XL.ADXExcelTaskPane
    {
        private GLBalanceConfigurator _wpfControl;
        private ElementHost _host;
        // Absolute floor/ceiling (DIP) - the hard safety rails GetWidthBoundsPx() clamps
        // the percentage-derived bounds into, and the fallback used outright when Excel's
        // own window width isn't available yet (e.g. early startup). _minWidthDip must
        // stay >= GLBalanceConfigurator.MinimumConfiguratorWidth (595), which the header
        // content actually needs to fit without overflowing.
        //
        // 520 (the originally proposed floor) was NOT enough: at that width, the
        // header's icon + "Balance Configurator" title (Grid.Column="0", Auto-sized)
        // plus the close button (Grid.Column="2", Auto-sized) together needed more room
        // than a 520-wide container provides - WPF's Grid does not compress Auto
        // columns to force-fit. Confirmed via direct measurement: the close button's
        // own computed screen position sat a consistent 36px past the pane's actual
        // right edge, independent of any layout/repaint invalidation - genuine content
        // overflow, not a timing bug. 75 DIP of headroom over that measured shortfall
        // leaves a safety margin for font/ClearType rendering variance across machines.
        private readonly int _minWidthDip = 635;
        private readonly int _maxWidthDip = 940;
        // OISR-23017: what the WPF content itself needs - GLBalanceConfigurator.xaml's inner Grid MinWidth
        // (640; its ScrollViewer has horizontal scrolling disabled) plus the vertical scrollbar (~18). The
        // WPF host never gets narrower than this; when Add-in Express's docking area is narrower (it
        // decides that width, not us - e.g. 493 px at 100% was seen), the pane scrolls horizontally
        // (AutoScroll) instead of forcing itself wider than the docking area and running off the screen.
        private const int ContentMinWidthDip = 640 + 18;
        // The pane's width target as a fraction of Excel's own current window width
        // (AppState.Instance.ExcelApp.Width) - see GetWidthBoundsPx(). Keeps the pane a
        // consistent proportion of the available space across any monitor/resolution/DPI,
        // instead of a fixed DIP size that ate up to ~60% of the window at 175% scale on
        // a modest-resolution screen (see commit 738fbed's fixed-DIP version).
        private const float MinWidthPercent = 0.25f;
        private const float MaxWidthPercent = 0.35f;
        // Guarantees at least this much room between the computed min and max (DIP,
        // DPI-scaled in GetWidthBoundsPx()) even when the percentage-derived max would
        // otherwise fall below the absolute floor - without this, minPx and maxPx could
        // collapse to the exact same value on any non-maximized/narrower Excel window
        // (confirmed via debug logging: minWidthPx=520, maxWidthPx=520 and similar at
        // other DPIs), locking the splitter dead with zero room to drag.
        private readonly int _minRangeSlackDip = 100;
        // Width the pane opens at (GLConfiguratorPane_ADXBeforeTaskPaneShow) - distinct
        // from _minWidthDip so the pane can launch wider than its floor.
        private readonly int _defaultWidthDip = 650;
        private readonly int _minHeightDip = 300;
        private const int DefaultDpi = 96;
        private const int WM_WINDOWPOSCHANGING = 0x0046;
        private const int WM_SIZING = 0x0214;
        private const int WMSZ_LEFT = 1;
        private const int WMSZ_RIGHT = 2;
        private const int WMSZ_TOP = 3;
        private const int WMSZ_TOPLEFT = 4;
        private const int WMSZ_TOPRIGHT = 5;
        private const int WMSZ_BOTTOM = 6;
        private const int WMSZ_BOTTOMLEFT = 7;
        private const int WMSZ_BOTTOMRIGHT = 8;

        [StructLayout(LayoutKind.Sequential)]
        private struct WINDOWPOS
        {
            public IntPtr hwnd;
            public IntPtr hwndInsertAfter;
            public int x;
            public int y;
            public int cx;
            public int cy;
            public uint flags;
        }

        [StructLayout(LayoutKind.Sequential)]
        private struct RECT
        {
            public int Left;
            public int Top;
            public int Right;
            public int Bottom;
        }

        [DllImport("user32.dll")]
        private static extern uint GetDpiForWindow(IntPtr hWnd);

        [DllImport("user32.dll")]
        private static extern bool GetWindowRect(IntPtr hWnd, out RECT lpRect);

        // Control.Width can go stale relative to the pane's real on-screen size - during
        // a live DPI/resolution change, Add-in Express's own internal reflow can move
        // the real native window without going through whatever path keeps WinForms'
        // cached Control.Width in sync (confirmed via direct GetWindowRect comparison).
        // Always read the real rect here instead of trusting this.Width for clamp math.
        private int GetActualWidthPx()
        {
            if (this.IsHandleCreated && GetWindowRect(this.Handle, out RECT rect))
            {
                return rect.Right - rect.Left;
            }

            return this.Width;
        }

        private int GetEffectiveDpi()
        {
            try
            {
                if (this.IsHandleCreated)
                {
                    // OISR-23017: the monitor's real DPI first - GetDpiForWindow keeps reporting
                    // Excel's start-up DPI for this pane after a display-scale change with Excel open.
                    uint monitorDpi = DpiAwarenessHelper.GetMonitorDpi(this.Handle);
                    if (monitorDpi > 0)
                        return (int)monitorDpi;
                    return (int)GetDpiForWindow(this.Handle);
                }
            }
            catch
            {
                // ignore and fall back
            }

            return this.DeviceDpi > 0 ? this.DeviceDpi : DefaultDpi;
        }

        // Excel.Application.Width is documented as POINTS (72/inch), but doesn't
        // reliably reflect the window's actual current-DPI physical size (confirmed via
        // live testing - a points/dpi conversion came out ~1.6x too large at 175%
        // scale, a known rough edge of these Office COM automation properties not being
        // fully per-monitor-DPI-aware). Application.Hwnd + GetWindowRect instead reads
        // the window's real physical pixel rect directly via Win32 - the same unit
        // space this.Width/e.NewRegionSize.Width already use.
        private int GetExcelWindowWidthPx()
        {
            try
            {
                var excelApp = AppState.Instance.ExcelApp;
                if (excelApp == null)
                {
                    return 0;
                }

                IntPtr hwnd = (IntPtr)excelApp.Hwnd;
                if (hwnd == IntPtr.Zero || !GetWindowRect(hwnd, out RECT rect))
                {
                    return 0;
                }

                return rect.Right - rect.Left;
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex, "GLConfiguratorPane.GetExcelWindowWidthPx");
                return 0;
            }
        }

        // Percentage-of-Excel-window bounds, clamped into the absolute [_minWidthDip,
        // _maxWidthDip] safety rails (DPI-scaled). Falls back to those absolute bounds
        // outright when Excel's own window width isn't available yet (e.g. early
        // startup, GetExcelWindowWidthPx() returning 0).
        private (int minPx, int maxPx) GetWidthBoundsPx()
        {
            int dpi = GetEffectiveDpi();
            int absoluteMinPx = (int)Math.Round(_minWidthDip * dpi / (float)DefaultDpi);
            int absoluteMaxPx = (int)Math.Round(_maxWidthDip * dpi / (float)DefaultDpi);

            int excelWindowWidthPx = GetExcelWindowWidthPx();
            if (excelWindowWidthPx <= 0)
            {
                return (absoluteMinPx, absoluteMaxPx);
            }

            int minPx = (int)Math.Round(excelWindowWidthPx * MinWidthPercent);
            int maxPx = (int)Math.Round(excelWindowWidthPx * MaxWidthPercent);
            int minRangeSlackPx = (int)Math.Round(_minRangeSlackDip * dpi / (float)DefaultDpi);

            minPx = Math.Max(minPx, absoluteMinPx);
            // Math.Max(maxPx, minPx) alone would let maxPx collapse to exactly minPx
            // whenever the percentage-derived max falls below the floor - guarantee at
            // least minRangeSlackPx of room instead, then clamp to the ceiling (and, in
            // case the ceiling itself is tighter than min+slack, back down to at least
            // minPx so max never ends up below min).
            maxPx = Math.Max(maxPx, minPx + minRangeSlackPx);
            maxPx = Math.Min(maxPx, absoluteMaxPx);
            maxPx = Math.Max(maxPx, minPx);

            return (minPx, maxPx);
        }

        // Called from AddinModule.adxExcelAppEvents1_WindowResize whenever Excel's OWN
        // window is resized/moved/maximized - the pane can't detect that on its own
        // (GLConfiguratorPane_Resize only fires for the pane's own size changes), but a
        // change in Excel's window width shifts the percentage-derived bounds even when
        // the pane's own width hasn't changed yet, so it needs an explicit re-clamp.
        public void RecomputeWidthBounds()
        {
            var (minPx, maxPx) = GetWidthBoundsPx();
            int actualWidthPx = GetActualWidthPx();
            // OISR-23017: cap at the max only - never push the pane wider than the docking area Add-in
            // Express gave it (that ran its right side off the screen); a narrower area scrolls (AutoScroll).
            int clampedWidth = Math.Min(actualWidthPx, maxPx);
            if (clampedWidth != this.Width)
            {
                LogUtility.LogDebug($"GLConfiguratorPane.RecomputeWidthBounds: this.Width={this.Width} (actual={actualWidthPx}) -> {clampedWidth} (minPx={minPx}, maxPx={maxPx})");
                this.Width = clampedWidth;
            }

            // Defensive layout refresh after a programmatic resize (see the matching
            // comment in GLConfiguratorPane_Resize for why this alone did NOT fix the
            // close-button overflow bug - kept anyway as cheap, harmless insurance).
            _wpfControl?.InvalidateMeasure();
            _wpfControl?.InvalidateArrange();
            _wpfControl?.UpdateLayout();
            if (_host != null)
            {
                _host.Invalidate();
            }
            _wpfControl?.InvalidateVisual();
        }

        public GLConfiguratorPane()
        {
            InitializeComponent();

            // Enable DPI-aware sizing for the WinForms host
            this.AutoScaleMode = AutoScaleMode.Dpi;

            ApplyDpiAwareSizing(GetEffectiveDpi());

            // ---- REVERT NOTE (fix applied for "Balance Configurator appears zoomed in
            // for some users" - see chat/CLAUDE.md history) -------------------------------
            // Original code before this fix (kept here so this can be reverted exactly if
            // the fix below ever needs to be backed out):
            //
            //     using (DpiAwarenessHelper.SetPerMonitorAware())
            //     {
            //         _wpfControl = new GLBalanceConfigurator(this);
            //     }
            //
            //     _wpfControl.OnCloseRequested += () => this.Visible = false;
            //
            //     _host = new ElementHost
            //     {
            //         Dock = DockStyle.Fill,
            //         MinimumSize = this.MinimumSize,
            //         Child = _wpfControl
            //     };
            //
            //     this.Controls.Add(_host);
            //
            // Why that was wrong: the WPF content's real native window (ElementHost's
            // HwndSource) is NOT created when _wpfControl/_host are constructed (building
            // a WPF object creates no HWND at all) - WinForms creates it lazily, only once
            // the handle is actually needed (typically when this task pane's own handle is
            // realized by Excel/ADX and the control tree cascades handle creation down to
            // its children). The old "using" block only covered the WPF object's managed
            // construction, so by the time the real ElementHost handle got created later,
            // the thread's DPI context had already been reverted back to whatever it was
            // before - the WPF content ended up rendering under whatever DPI awareness was
            // ambient at that later, untimed moment instead of Per-Monitor-V2. That race is
            // what caused the pane to intermittently render "zoomed in"/blurry - Windows
            // falls back to bitmap-stretching the content to the monitor's actual DPI
            // instead of it rendering natively - most visible on >100% display scaling
            // and/or when Excel isn't on the primary monitor when the pane first loads.
            // Note this is NOT the same fix as WpfAppManager.cs's "never revert" pattern:
            // that dedicated WPF dispatcher thread only ever hosts WPF windows, so leaving
            // it permanently Per-Monitor-V2 is safe. This task pane instead runs on
            // Excel's own main UI thread (shared with the rest of Excel's UI), so the
            // context here is always explicitly reverted afterward via "using" - we widen
            // the scope to cover ElementHost creation, and also reapply it in
            // HandleCreated below to cover the (normal, ADX-driven) case where this task
            // pane's own handle - and so the ElementHost's cascade-created handle - isn't
            // realized until after this constructor has already returned.
            using (DpiAwarenessHelper.SetPerMonitorAware())
            {
                _wpfControl = new GLBalanceConfigurator(this);

                // Host WPF control inside WinForms
                _host = new ElementHost
                {
                    Dock = DockStyle.Fill,
                    MinimumSize = this.MinimumSize,
                    Child = _wpfControl
                };

                this.Controls.Add(_host);
            }

            // OISR-23017: scroll horizontally when the docking area is narrower than the content needs
            // (the host's MinimumSize / AutoScrollMinSize, set in ApplyDpiAwareSizing) instead of the pane
            // forcing itself wider than the docking area. Dock=Fill lays the host out in the scrollable area.
            this.AutoScroll = true;
            ApplyDpiAwareSizing(GetEffectiveDpi());

            _wpfControl.OnCloseRequested += () => this.Visible = false;

            // Covers the case where this task pane's own native handle - and therefore
            // the ElementHost's cascade-created handle - is realized after this
            // constructor returns (the normal case for an ADX-hosted task pane), so the
            // WPF content's HwndSource still ends up created under Per-Monitor-V2.
            this.HandleCreated += GLConfiguratorPane_HandleCreated;

            // Handle resize events
            this.Resize += GLConfiguratorPane_Resize;

            // Freezes the pane's width between MinimumSize.Width and a DPI-scaled
            // _maxWidthDip during a live splitter drag. This is Add-in Express's own
            // dedicated resize-constraint event - fired on every mouse-move while the
            // user drags the splitter (ADXContainerPane.OnMouseMove -> DoMouseMove ->
            // VerifyConstrains -> ADXForm.VerifyConstraints -> OnADXSplitterMove),
            // distinct from the WM_SIZING/WM_WINDOWPOSCHANGING overrides below (which
            // independently enforce only the minimum, and are left unchanged).
            // e.NewRegionSize is read-only - there is no way to substitute a clamped
            // size, only accept or reject the exact proposed one - but since this fires
            // continuously during the drag, rejecting anything outside the [Min, Max]
            // range stops the pane from tracking the mouse any further once it hits
            // either bound.
            //
            // Deliberately does NOT touch this.MaximumSize/ApplyDpiAwareSizing - an
            // earlier version set MaximumSize and re-clamped this.Size from there, which
            // ran during DPI-change/HandleCreated handling too, not just live drags.
            // That caused a resize feedback loop (continuous flicker, Excel going
            // unresponsive - confirmed via Windows Event Log as an Application Hang,
            // AppHangB1) when switching display resolution/scale (1360x768 @125%) while
            // the pane was open - most likely fighting either Add-in Express's own
            // docking-layout recalculation or Windows' WM_DPICHANGED resize protocol,
            // both of which are far more sensitive to a size change happening during
            // DPI-change handling than to one only ever happening from a user's own live
            // mouse drag. Computing maxWidthPx fresh here, scoped only to this
            // user-driven event, avoids running any of that during DPI/resolution
            // changes.
            this.ADXSplitterMove += GLConfiguratorPane_ADXSplitterMove;
            this.ADXAfterTaskPaneShow += GLConfiguratorPane_ADXAfterTaskPaneShow;

        }

        // OISR-23017: Add-in Express re-lays its docking area out after the show and can leave it
        // narrower than the content (480 px at 100% here), so widen it once it has settled.
        private void GLConfiguratorPane_ADXAfterTaskPaneShow(object sender, ADXAfterTaskPaneShowEventArgs e)
        {
            try
            {
                this.BeginInvoke(new Action(() =>
                {
                    try
                    {
                        int actualWidthPx = GetActualWidthPx();
                        int contentMinWidthPx = (int)Math.Round(ContentMinWidthDip * GetEffectiveDpi() / (float)DefaultDpi);
                        if (this.Visible && actualWidthPx < contentMinWidthPx)
                        {
                            var (_, maxWidthPx) = GetWidthBoundsPx();
                            int defaultWidthPx = (int)Math.Round(_defaultWidthDip * GetEffectiveDpi() / (float)DefaultDpi);
                            int targetWidthPx = Math.Min(Math.Max(defaultWidthPx, contentMinWidthPx), Math.Max(maxWidthPx, contentMinWidthPx));
                            LogUtility.LogDebug($"GLConfiguratorPane_ADXAfterTaskPaneShow: widening {actualWidthPx} -> {targetWidthPx}");
                            this.Width = targetWidthPx;
                        }
                    }
                    catch (Exception ex)
                    {
                        LogUtility.LogException(ex, "GLConfiguratorPane_ADXAfterTaskPaneShow (deferred)");
                    }
                }));
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex, "GLConfiguratorPane_ADXAfterTaskPaneShow");
            }
        }

        private void GLConfiguratorPane_ADXSplitterMove(object sender, ADXSplitterMoveEventArgs e)
        {
            // Both bounds MUST come from the same GetWidthBoundsPx() call - comparing a
            // freshly-computed max against a stale cached min (this.MinimumSize.Width,
            // from a prior WinForms-only computation) previously let the two fall out of
            // sync after a live DPI change and invert the allowed range, blocking every
            // resize (confirmed reproduction: cursor showed resize, dragging did nothing,
            // after switching 125% -> 100% with the pane already open).
            var (minWidthPx, maxWidthPx) = GetWidthBoundsPx();
            if (e.NewRegionSize.Width < minWidthPx || e.NewRegionSize.Width > maxWidthPx)
            {
                e.Cancel = true;
            }
        }

        private void GLConfiguratorPane_HandleCreated(object sender, EventArgs e)
        {
            using (DpiAwarenessHelper.SetPerMonitorAware())
            {
                // Touching Handle forces WinForms to realize the ElementHost's native
                // window now, while the per-monitor context is active, if it has not
                // already been created by this point.
                if (_host != null)
                {
                    _ = _host.Handle;
                }
            }

            // GetEffectiveDpi() in the constructor ran before this.IsHandleCreated was
            // true, so it always fell back to the stale Control.DeviceDpi (96) instead
            // of the monitor's real DPI - permanently locking MinimumSize/_host.MinimumSize
            // to the un-scaled 600x300px floor regardless of actual display scaling. That
            // let a user shrink the pane (and the ElementHost/WPF content inside it) well
            // below the size GLBalanceConfigurator's own MinWidth actually needs at the
            // real DPI, causing a layout squeeze/ghosting artifact while dragging the
            // pane's resize corner. Now that this pane's own handle exists, GetEffectiveDpi()
            // can return the real per-monitor DPI, so recompute the sizing here to correct
            // that floor.
            ApplyDpiAwareSizing(GetEffectiveDpi());
            SyncWpfContentScale();
        }

        // OISR-23017: the WPF content keeps the DPI it was created with (Excel's start-up DPI). A docked
        // task pane gets no WM_DPICHANGED (see below), so after a display-scale change with Excel open the
        // content was drawn too small (100% -> 150%) or too big and clipped (150% -> 100%). Scale it by
        // monitor DPI / WPF's DPI; 1.0 (no transform) whenever the two agree, e.g. Excel started at this scale.
        private double _appliedContentScale = 1.0;

        private bool _reshowPending;

        private void SyncWpfContentScale(bool reshowIfVisible = false)
        {
            try
            {
                if (_wpfControl == null || !this.IsHandleCreated)
                    return;

                uint monitorDpi = DpiAwarenessHelper.GetMonitorDpi(this.Handle);
                DpiAwarenessHelper.LogIfDpiChanged(monitorDpi, "Balance Configurator pane");
                double wpfDpi = System.Windows.Media.VisualTreeHelper.GetDpi(_wpfControl).PixelsPerInchX;
                if (monitorDpi == 0 || wpfDpi <= 0)
                    return;

                double scale = monitorDpi / wpfDpi;
                if (Math.Abs(scale - _appliedContentScale) < 0.01)
                    return;

                _appliedContentScale = scale;
                _wpfControl.LayoutTransform = Math.Abs(scale - 1.0) < 0.01
                    ? System.Windows.Media.Transform.Identity
                    : new System.Windows.Media.ScaleTransform(scale, scale);

                // The DPI changed while the pane exists: MinimumSize is still in the old DPI's pixels, and
                // WM_WINDOWPOSCHANGING/SetBoundsCore enforce it - after 150% -> 100% that kept the pane at
                // its 150% width (content clipped on the right) whatever the Resize clamp asked for.
                // Re-derive it for the new DPI; the Resize clamp then brings the width into range.
                ApplyDpiAwareSizing(monitorDpi);

                // A live DPI change while the pane is open: the pane resizes, but Add-in Express does not
                // re-lay out the docking area around it, so after 150% -> 100% the content sat shifted right
                // with the close button and right-hand controls clipped until the user resized it. Re-show
                // it once (the same path as opening it - ADXBeforeTaskPaneShow sets the width for the
                // current DPI); its content is kept. Not done when the pane is merely being created/shown.
                if (reshowIfVisible && this.Visible && !_reshowPending)
                {
                    _reshowPending = true;
                    this.BeginInvoke(new Action(() =>
                    {
                        // ADXBeforeTaskPaneShow keeps the pane hidden unless displayConfigurator is set (only
                        // the ribbon's Balance click sets it, so Add-in Express can't pop the pane up on its own);
                        // set it around the re-show exactly as RibFSG_OnClick does, or the pane stays closed.
                        AppState.Instance.displayConfigurator = true;
                        try
                        {
                            this.Visible = false;
                            this.Visible = true;
                        }
                        catch (Exception ex)
                        {
                            LogUtility.LogException(ex, "GLConfiguratorPane.SyncWpfContentScale (re-show)");
                        }
                        finally
                        {
                            AppState.Instance.displayConfigurator = false;
                            _reshowPending = false;
                        }
                    }));
                }
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex, "GLConfiguratorPane.SyncWpfContentScale");
            }
        }

        // Only ever called from the constructor and HandleCreated (both effectively
        // initial-setup) - Form.DpiChanged/WM_DPICHANGED was tried as a live-update
        // trigger too, but confirmed via debug logging (dpi=96/120/168 transitions
        // visible in GLConfiguratorPane_Resize's own log line, with zero corresponding
        // "DpiChanged FIRED" lines) to never fire at all for this docked/Add-in-Express-
        // subclassed pane - Windows only sends WM_DPICHANGED to genuine top-level
        // windows. GLConfiguratorPane_Resize is what actually keeps the pane sized
        // correctly across a live DPI/resolution change instead.
        private void ApplyDpiAwareSizing(float dpiX)
        {
            var scale = dpiX / 96f;
            int minHeightPx = (int)Math.Round(_minHeightDip * scale);
            int contentMinWidthPx = (int)Math.Round(ContentMinWidthDip * scale);

            // OISR-23017: no minimum WIDTH on the pane form itself any more - Add-in Express decides the
            // docking area's width, and a form minimum wider than that made the pane run off the screen
            // (WndProc/SetBoundsCore enforce MinimumSize). The WPF host keeps the content's real minimum
            // and the form scrolls horizontally (AutoScroll) when the docking area is narrower. The user's
            // splitter drags are still kept within [min, max] by ADXSplitterMove.
            this.MinimumSize = new Size(0, minHeightPx);
            this.AutoScrollMinSize = new Size(contentMinWidthPx, minHeightPx);
            if (_host != null)
            {
                _host.MinimumSize = new Size(contentMinWidthPx, minHeightPx);
            }

            this.Size = new Size(this.Width, Math.Max(this.Height, minHeightPx));
        }
        private void GLConfiguratorPane_Resize(object sender, EventArgs e)
        {
            SyncWpfContentScale(reshowIfVisible: true); // OISR-23017: Resize is what fires on a live DPI change for this pane

            // Resize (backed by WM_SIZE) fires for ANY actual size change to this
            // control, regardless of cause - unlike WM_DPICHANGED/Form.DpiChanged, which
            // Windows only sends to genuine top-level windows and, confirmed via debug
            // logging, never reaches this docked/Add-in-Express-subclassed pane at all
            // (a DpiChanged handler was wired up and logged here temporarily; despite the
            // dpi value below visibly transitioning 96/120/168 across a live display
            // scale change, its "FIRED" log line never once appeared - removed once
            // confirmed). Re-clamping into [min, max] here, computed fresh every time, is
            // what actually keeps the pane correctly sized across a live DPI/resolution
            // change. (Excel's own window resizing/moving is handled separately by
            // RecomputeWidthBounds, called from AddinModule's WindowResize handler.)
            var (minWidthPx, maxWidthPx) = GetWidthBoundsPx();
            int actualWidthPx = GetActualWidthPx();

            // OISR-23017: cap at the max only (see RecomputeWidthBounds) - below the min the pane scrolls.
            int clampedWidth = Math.Min(actualWidthPx, maxWidthPx);
            if (clampedWidth != this.Width)
            {
                this.Width = clampedWidth; // the host follows via Dock=Fill (never below its MinimumSize)
            }

            // Defensive layout refresh after a programmatic resize. NOTE: this was
            // originally added as a fix attempt for the header close button rendering
            // past the pane's right edge, but direct PointToScreen measurement of the
            // button proved this had ZERO effect (byte-identical screen position
            // before/after) - the real root cause was WPF Grid Auto-column content
            // overflow (the header's icon + title + close button needed more width
            // than the container provided; Grid does not compress Auto columns to
            // force-fit), fixed by widening _minWidthDip to 595 and adding
            // ClipToBounds="True" in GLBalanceConfigurator.xaml. Kept here anyway as
            // cheap, harmless insurance after a programmatic resize.
            _wpfControl?.InvalidateMeasure();
            _wpfControl?.InvalidateArrange();
            _wpfControl?.UpdateLayout();
            if (_host != null)
            {
                _host.Invalidate();
            }
            _wpfControl?.InvalidateVisual();
        }
        public async Task RelaunchPane(bool showBusyOverlay = true)
        {
            try
            {
                LogUtility.LogDebug($"GLConfiguratorPane.RelaunchPane invoked (showBusyOverlay={showBusyOverlay}).");
                if (_wpfControl != null && _wpfControl.Dispatcher != null)
                {
                    // Dispatcher.InvokeAsync(async () => ...) doesn't wait for the inner Task -
                    // the DispatcherOperation completes as soon as the delegate hits its first
                    // await. Use a non-async delegate + Task.Unwrap() so the real completion is
                    // awaited (see GLWaitWindow.ShowConfirmToastAsync for the same pattern).
                    await _wpfControl.Dispatcher.InvokeAsync(() => _wpfControl.ReLoadConfigurator(showBusyOverlay)).Task.Unwrap();
                }
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex, "GLConfiguratorPane.RelaunchPane");
            }
        }
        /// <summary>
        /// Whether the hosted GLBalanceConfigurator currently has an explicit saved
        /// configuration selected. Read directly (not marshaled) - this is called from
        /// AddinModule.cs's Excel SheetSelectionChange handler, which already runs on the
        /// same STA/UI thread the WPF control lives on.
        /// </summary>
        public bool HasSavedConfigurationSelected() => _wpfControl?.HasSavedConfigurationSelected ?? false;

        public async Task ResetPaneReference()
        {
            try
            {
                LogUtility.LogDebug("GLConfiguratorPane.ResetPaneReference invoked.");
                if (_wpfControl != null && _wpfControl.Dispatcher != null)
                {
                    await _wpfControl.Dispatcher.InvokeAsync(() =>
                    {
                        GLBalanceConfigurator.ResetCellReference();
                    });
                }
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex, "GLConfiguratorPane.ResetPaneReference");
            }
        }
        private void GLConfiguratorPane_ADXBeforeTaskPaneShow(object sender, ADXBeforeTaskPaneShowEventArgs e)
        {
            try
            {
                GLConfiguratorPane pane = sender as GLConfiguratorPane;
                if (pane != null)
                {
                    pane.Visible = AppState.Instance.displayConfigurator;
                    LogUtility.LogDebug($"GLConfiguratorPane_ADXBeforeTaskPaneShow fired. Visible={pane.Visible}");

                    // Set size when showing - launches at _defaultWidthDip (550), not the
                    // MinimumSize floor (520), so the two can differ.
                    if (pane.Visible)
                    {
                        pane.SyncWpfContentScale(); // OISR-23017
                        int defaultWidthPx = (int)Math.Round(_defaultWidthDip * pane.GetEffectiveDpi() / (float)DefaultDpi);
                        pane.Width = Math.Max(defaultWidthPx, pane.MinimumSize.Width);
                    }
                }
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex, "GLConfiguratorPane_ADXBeforeTaskPaneShow");
            }

        }

        protected override void WndProc(ref Message m)
        {
            // NOTE: a WM_MOUSEWHEEL branch was tried here (forwarding directly into
            // SuggestAppendComboBox) as part of the ADXTaskPane mouse-wheel fix saga
            // (CLAUDE.md section 24.3). It was removed: disabling the global low-level
            // mouse hook (SuggestAppendComboBox's other fix attempt) proved this WndProc
            // override never actually received WM_MOUSEWHEEL for the popup in the first
            // place - the message simply never reaches this control's window procedure via
            // normal routing in this hosting context, only via a raw, pre-routing
            // low-level hook. See SuggestAppendComboBox.cs for the current fix (a
            // low-level hook on its own dedicated, non-blocking thread).
            if (m.Msg == WM_SIZING && m.LParam != IntPtr.Zero)
            {
                var rc = (RECT)Marshal.PtrToStructure(m.LParam, typeof(RECT));
                int width = rc.Right - rc.Left;
                if (width < this.MinimumSize.Width)
                {
                    int minWidth = this.MinimumSize.Width;
                    switch ((int)m.WParam)
                    {
                        case WMSZ_LEFT:
                        case WMSZ_TOPLEFT:
                        case WMSZ_BOTTOMLEFT:
                            rc.Left = rc.Right - minWidth;
                            break;
                        default:
                            rc.Right = rc.Left + minWidth;
                            break;
                    }

                    Marshal.StructureToPtr(rc, m.LParam, true);
                }
            }
            else if (m.Msg == WM_WINDOWPOSCHANGING && m.LParam != IntPtr.Zero)
            {
                var pos = (WINDOWPOS)Marshal.PtrToStructure(m.LParam, typeof(WINDOWPOS));
                if (pos.cx < this.MinimumSize.Width)
                {
                    pos.cx = this.MinimumSize.Width;
                    Marshal.StructureToPtr(pos, m.LParam, true);
                }
            }

            base.WndProc(ref m);
        }

        protected override void SetBoundsCore(int x, int y, int width, int height, BoundsSpecified specified)
        {
            var dpi = GetEffectiveDpi();
            int minHeightPx = (int)Math.Round(_minHeightDip * dpi / (float)DefaultDpi);

            // OISR-23017: width is no longer floored here - Add-in Express sizes the pane to its docking
            // area, and forcing it wider ran its right side off the screen. Narrower areas scroll instead.

            if ((specified & BoundsSpecified.Height) != 0 && height < minHeightPx)
            {
                height = minHeightPx;
            }

            base.SetBoundsCore(x, y, width, height, specified);
        }
    }
}
