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
        private readonly int _minWidthDip = 520;
        private readonly int _maxWidthDip = 625;
        // Width the pane opens at (GLConfiguratorPane_ADXBeforeTaskPaneShow) - distinct
        // from _minWidthDip so the pane can launch wider than its floor. Must stay >=
        // GLBalanceConfigurator.MinimumConfiguratorWidth (520) - that WPF-side floor
        // actively pushes the pane back up to it if this pane is ever narrower.
        private readonly int _defaultWidthDip = 550;
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

        private int GetEffectiveDpi()
        {
            try
            {
                if (this.IsHandleCreated)
                {
                    return (int)GetDpiForWindow(this.Handle);
                }
            }
            catch
            {
                // ignore and fall back
            }

            return this.DeviceDpi > 0 ? this.DeviceDpi : DefaultDpi;
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

        }

        private void GLConfiguratorPane_ADXSplitterMove(object sender, ADXSplitterMoveEventArgs e)
        {
            // Both bounds MUST be computed fresh from the same live DPI reading here -
            // comparing a freshly-computed max against the cached this.MinimumSize.Width
            // (only ever updated via ApplyDpiAwareSizing, itself only reachable through
            // HandleCreated/the constructor - see GLConfiguratorPane_Resize's own comment
            // on why DpiChanged is not part of that list) let the two fall out of sync
            // after a live DPI change: MinimumSize.Width stayed pinned to whatever DPI
            // was active when the pane was first shown (e.g. 750px at 125%) while this
            // method's max recomputed against the NEW dpi (e.g. 700px at 100%) - with
            // Min(750) > Max(700), every single drag position failed both comparisons at
            // once, so the splitter never accepted any resize at all (confirmed
            // reproduction: cursor showed resize, dragging did nothing, after switching
            // 125% -> 100% with the pane already open).
            var dpi = GetEffectiveDpi();
            int minWidthPx = (int)Math.Round(_minWidthDip * dpi / (float)DefaultDpi);
            int maxWidthPx = (int)Math.Round(_maxWidthDip * dpi / (float)DefaultDpi);
            LogUtility.LogDebug($"GLConfiguratorPane_ADXSplitterMove: dpi={dpi}, NewRegionSize.Width={e.NewRegionSize.Width}, minWidthPx={minWidthPx}, maxWidthPx={maxWidthPx}, cachedMinimumSize.Width={this.MinimumSize.Width}");
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
            int minWidthPx = (int)Math.Round(_minWidthDip * scale);
            int minHeightPx = (int)Math.Round(_minHeightDip * scale);

            this.MinimumSize = new Size(minWidthPx, minHeightPx);
            if (_host != null)
            {
                _host.MinimumSize = this.MinimumSize;
            }

            this.Size = new Size(Math.Max(this.Width, minWidthPx), Math.Max(this.Height, minHeightPx));
        }
        private void GLConfiguratorPane_Resize(object sender, EventArgs e)
        {
            // Resize (backed by WM_SIZE) fires for ANY actual size change to this
            // control, regardless of cause - unlike WM_DPICHANGED/Form.DpiChanged, which
            // Windows only sends to genuine top-level windows and, confirmed via debug
            // logging, never reaches this docked/Add-in-Express-subclassed pane at all
            // (a DpiChanged handler was wired up and logged here temporarily; despite the
            // dpi value below visibly transitioning 96/120/168 across a live display
            // scale change, its "FIRED" log line never once appeared - removed once
            // confirmed). Re-clamping into [min, max] here, computed fresh from the live
            // DPI every time, is what actually keeps the pane correctly sized across a
            // live DPI/resolution change.
            var dpi = GetEffectiveDpi();
            int minWidthPx = (int)Math.Round(_minWidthDip * dpi / (float)DefaultDpi);
            int maxWidthPx = (int)Math.Round(_maxWidthDip * dpi / (float)DefaultDpi);
            LogUtility.LogDebug($"GLConfiguratorPane_Resize: dpi={dpi}, this.Width={this.Width}, minWidthPx={minWidthPx}, maxWidthPx={maxWidthPx}");

            int clampedWidth = Math.Min(Math.Max(this.Width, minWidthPx), maxWidthPx);
            if (clampedWidth != this.Width)
            {
                this.Width = clampedWidth;
                if (_host != null)
                {
                    _host.Width = clampedWidth;
                }
            }

            // Update WPF control if needed
            _wpfControl?.UpdateLayout();
        }
        public async Task RelaunchPane()
        {
            try
            {
                LogUtility.LogDebug("GLConfiguratorPane.RelaunchPane invoked.");
                if (_wpfControl != null && _wpfControl.Dispatcher != null)
                {
                    // Dispatcher.InvokeAsync(async () => ...) doesn't wait for the inner Task -
                    // the DispatcherOperation completes as soon as the delegate hits its first
                    // await. Use a non-async delegate + Task.Unwrap() so the real completion is
                    // awaited (see GLWaitWindow.ShowConfirmToastAsync for the same pattern).
                    await _wpfControl.Dispatcher.InvokeAsync(() => _wpfControl.ReLoadConfigurator()).Task.Unwrap();
                }
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex, "GLConfiguratorPane.RelaunchPane");
            }
        }
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
            int minWidthPx = (int)Math.Round(_minWidthDip * dpi / (float)DefaultDpi);
            int minHeightPx = (int)Math.Round(_minHeightDip * dpi / (float)DefaultDpi);

            if ((specified & BoundsSpecified.Width) != 0 && width < minWidthPx)
            {
                width = minWidthPx;
            }

            if ((specified & BoundsSpecified.Height) != 0 && height < minHeightPx)
            {
                height = minHeightPx;
            }

            base.SetBoundsCore(x, y, width, height, specified);
        }
    }
}
