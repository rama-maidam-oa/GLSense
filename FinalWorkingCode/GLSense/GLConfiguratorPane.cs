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
        private readonly int _minWidthDip = 600;
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
            LogUtility.LogInfo("[DPI-DIAG] GLConfiguratorPane ctor: start");
            InitializeComponent();
            LogUtility.LogInfo("[DPI-DIAG] GLConfiguratorPane ctor: InitializeComponent done");

            // Enable DPI-aware sizing for the WinForms host
            this.AutoScaleMode = AutoScaleMode.Dpi;

            LogUtility.LogInfo($"[DPI-DIAG] GLConfiguratorPane ctor: before ApplyDpiAwareSizing, GetEffectiveDpi={GetEffectiveDpi()}");
            ApplyDpiAwareSizing(GetEffectiveDpi());
            LogUtility.LogInfo("[DPI-DIAG] GLConfiguratorPane ctor: ApplyDpiAwareSizing done");
            this.DpiChanged += GLConfiguratorPane_DpiChanged;

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
            LogUtility.LogInfo("[DPI-DIAG] GLConfiguratorPane ctor: before SetPerMonitorAware/new GLBalanceConfigurator/ElementHost block");
            using (DpiAwarenessHelper.SetPerMonitorAware())
            {
                _wpfControl = new GLBalanceConfigurator(this);
                LogUtility.LogInfo("[DPI-DIAG] GLConfiguratorPane ctor: new GLBalanceConfigurator(this) done");

                // Host WPF control inside WinForms
                _host = new ElementHost
                {
                    Dock = DockStyle.Fill,
                    MinimumSize = this.MinimumSize,
                    Child = _wpfControl
                };
                LogUtility.LogInfo("[DPI-DIAG] GLConfiguratorPane ctor: ElementHost created, before Controls.Add");

                this.Controls.Add(_host);
                LogUtility.LogInfo("[DPI-DIAG] GLConfiguratorPane ctor: Controls.Add(_host) done");
            }
            LogUtility.LogInfo("[DPI-DIAG] GLConfiguratorPane ctor: SetPerMonitorAware/ElementHost block complete");

            _wpfControl.OnCloseRequested += () => this.Visible = false;

            // Covers the case where this task pane's own native handle - and therefore
            // the ElementHost's cascade-created handle - is realized after this
            // constructor returns (the normal case for an ADX-hosted task pane), so the
            // WPF content's HwndSource still ends up created under Per-Monitor-V2.
            this.HandleCreated += GLConfiguratorPane_HandleCreated;

            // Handle resize events
            this.Resize += GLConfiguratorPane_Resize;

        }

        private void GLConfiguratorPane_HandleCreated(object sender, EventArgs e)
        {
            LogUtility.LogInfo("[DPI-DIAG] GLConfiguratorPane.HandleCreated: start");
            using (DpiAwarenessHelper.SetPerMonitorAware())
            {
                // Touching Handle forces WinForms to realize the ElementHost's native
                // window now, while the per-monitor context is active, if it has not
                // already been created by this point.
                if (_host != null)
                {
                    LogUtility.LogInfo("[DPI-DIAG] GLConfiguratorPane.HandleCreated: before touching _host.Handle");
                    _ = _host.Handle;
                    LogUtility.LogInfo("[DPI-DIAG] GLConfiguratorPane.HandleCreated: _host.Handle touched OK");
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
            LogUtility.LogInfo($"[DPI-DIAG] GLConfiguratorPane.HandleCreated: before re-applying ApplyDpiAwareSizing, GetEffectiveDpi={GetEffectiveDpi()}");
            ApplyDpiAwareSizing(GetEffectiveDpi());
            LogUtility.LogInfo("[DPI-DIAG] GLConfiguratorPane.HandleCreated: end");
        }
        private void GLConfiguratorPane_DpiChanged(object sender, DpiChangedEventArgs e)
        {
            LogUtility.LogWarn($"[DPI-DIAG] GLConfiguratorPane.DpiChanged: start, DeviceDpiOld={e.DeviceDpiOld}, DeviceDpiNew={e.DeviceDpiNew}");
            ApplyDpiAwareSizing(e.DeviceDpiNew);
            LogUtility.LogWarn("[DPI-DIAG] GLConfiguratorPane.DpiChanged: ApplyDpiAwareSizing returned OK, end");
        }

        private void ApplyDpiAwareSizing(float dpiX)
        {
            LogUtility.LogInfo($"[DPI-DIAG] GLConfiguratorPane.ApplyDpiAwareSizing: start, dpiX={dpiX}");
            var scale = dpiX / 96f;
            int minWidthPx = (int)Math.Round(_minWidthDip * scale);
            int minHeightPx = (int)Math.Round(_minHeightDip * scale);

            this.MinimumSize = new Size(minWidthPx, minHeightPx);
            if (_host != null)
            {
                _host.MinimumSize = this.MinimumSize;
            }

            this.Size = new Size(Math.Max(this.Width, minWidthPx), Math.Max(this.Height, minHeightPx));
            LogUtility.LogInfo($"[DPI-DIAG] GLConfiguratorPane.ApplyDpiAwareSizing: end, MinimumSize={this.MinimumSize}, Size={this.Size}");
        }
        private void GLConfiguratorPane_Resize(object sender, EventArgs e)
        {
            var dpi = GetEffectiveDpi();
            var dipWidth = this.Width * DefaultDpi / (float)dpi;

            // Ensure minimum size in DIPs
            if (dipWidth < _minWidthDip)
            {
                int minWidthPx = (int)Math.Round(_minWidthDip * dpi / (float)DefaultDpi);
                this.Width = minWidthPx;
                if (_host != null)
                {
                    _host.Width = minWidthPx;
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
                    await _wpfControl.Dispatcher.InvokeAsync(async () =>
                    {
                        await _wpfControl.ReLoadConfigurator();
                    });
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

                    // Set size when showing
                    if (pane.Visible)
                    {
                        pane.Width = pane.MinimumSize.Width;  // Ensure proper width when shown
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
