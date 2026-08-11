// GLConfiguratorPane.cs in GLSense (host)
//
// Group H (Balance Configurator pane) - port of GLConfiguratorPane.cs (FinalWorkingCode).
// All DPI-aware minimum-size enforcement (ApplyDpiAwareSizing/WndProc/SetBoundsCore) is
// pure host-side WinForms/Win32 code, unchanged from the original.
//
// What DID change: the original directly did
// `_wpfControl = new GLBalanceConfigurator(this); _host = new ElementHost { Child =
// _wpfControl }; this.Controls.Add(_host);` - trivial, because everything ran in one
// AppDomain. GLBalanceConfigurator (+ its ~3000-line GLConfiguratorViewModel) now live in
// GLSense.Addin.Core's separate, hot-reloadable AppDomain, and a WPF FrameworkElement
// isn't MarshalByRefObject - it can't cross an AppDomain boundary by reference, so this
// class can no longer `new` it up directly.
//
// Instead this uses an HWND-reparenting bridge (see GLSense.Addin.Core.Views.
// ConfiguratorPaneHost and IGLSenseAddin.CreateConfiguratorPaneContent's doc comment for
// the full rationale): Addin.Core creates the WPF content as a real, self-contained,
// borderless top-level Window on its own WPF thread and hands back only its native window
// handle (an IntPtr - safe to cross an AppDomain boundary within the same process). This
// class then:
//   1. Win32 SetParent's that handle into itself.
//   2. Rewrites the child's window style bits (WS_POPUP -> WS_CHILD, drops
//      WS_CAPTION/WS_THICKFRAME) so it renders as embedded content.
//   3. Keeps it sized to match via MoveWindow on HandleCreated/Resize.
//   4. AttachThreadInput's the two threads so keyboard focus/Tab-navigation flows into
//      the reparented window (it's still pumped by Addin.Core's own WPF thread, not this
//      pane's/Excel's main STA thread).
using AddinExpress.XL;
using System;
using System.Drawing;
using System.Runtime.InteropServices;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace GLSense
{
    public partial class GLConfiguratorPane : AddinExpress.XL.ADXExcelTaskPane
    {
        private readonly int _minWidthDip = 600;
        private readonly int _minHeightDip = 300;
        private const int DefaultDpi = 96;
        private const int WM_WINDOWPOSCHANGING = 0x0046;
        private const int WM_SIZING = 0x0214;
        private const int WMSZ_LEFT = 1;
        private const int WMSZ_TOPLEFT = 4;
        private const int WMSZ_BOTTOMLEFT = 7;

        // Win32 interop for the HWND-reparenting bridge.
        [DllImport("user32.dll")] private static extern IntPtr SetParent(IntPtr hWndChild, IntPtr hWndNewParent);
        [DllImport("user32.dll")] private static extern IntPtr GetWindowLongPtr(IntPtr hWnd, int nIndex);
        [DllImport("user32.dll")] private static extern IntPtr SetWindowLongPtr(IntPtr hWnd, int nIndex, IntPtr dwNewLong);
        [DllImport("user32.dll")] private static extern bool MoveWindow(IntPtr hWnd, int X, int Y, int nWidth, int nHeight, bool bRepaint);
        [DllImport("user32.dll")] private static extern bool AttachThreadInput(uint idAttach, uint idAttachTo, bool fAttach);
        [DllImport("user32.dll")] private static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint lpdwProcessId);
        [DllImport("kernel32.dll")] private static extern uint GetCurrentThreadId();
        [DllImport("user32.dll")] private static extern uint GetDpiForWindow(IntPtr hWnd);

        private const int GWL_STYLE = -16;
        private const long WS_CHILD = 0x40000000L;
        private const long WS_POPUP = unchecked((long)0x80000000L);
        private const long WS_CAPTION = 0x00C00000L;
        private const long WS_THICKFRAME = 0x00040000L;
        private const long WS_SYSMENU = 0x00080000L;
        private const long WS_MAXIMIZEBOX = 0x00010000L;
        private const long WS_MINIMIZEBOX = 0x00020000L;

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

        /// <summary>
        /// Host-only reentrancy guard (old monolith's AppState.Instance.displayConfigurator).
        /// AppState now lives in Addin.Core and nothing there ever read this flag - it only
        /// ever gated ADXBeforeTaskPaneShow's decision to allow/suppress an unsolicited
        /// pane show (e.g. Excel auto-restoring a previously-visible pane on workbook
        /// reopen), so it's kept entirely host-side rather than crossing the AppDomain
        /// boundary for no reason. Set true by RibFSG_OnClick for the duration of its own
        /// toggle, false otherwise.
        /// </summary>
        public static bool DisplayConfigurator { get; set; }

        private IntPtr _contentHwnd = IntPtr.Zero;

        private int GetEffectiveDpi()
        {
            try
            {
                if (this.IsHandleCreated)
                {
                    return (int)GetDpiForWindow(this.Handle);
                }
            }
            catch (Exception ex)
            {
                // ignore and fall back
                GlobalsEx.Context?.Logger?.LogDebug($"GLConfiguratorPane.GetEffectiveDpi: GetDpiForWindow failed, falling back ({ex.Message})");
            }

            return this.DeviceDpi > 0 ? this.DeviceDpi : DefaultDpi;
        }

        public GLConfiguratorPane()
        {
            InitializeComponent();

            this.AutoScaleMode = AutoScaleMode.Dpi;
            ApplyDpiAwareSizing(GetEffectiveDpi());
            this.DpiChanged += GLConfiguratorPane_DpiChanged;
            this.HandleCreated += GLConfiguratorPane_HandleCreated;
            this.HandleDestroyed += GLConfiguratorPane_HandleDestroyed;
        }

        private void GLConfiguratorPane_HandleCreated(object sender, EventArgs e)
        {
            GlobalsEx.Context?.Logger?.LogDebug("GLConfiguratorPane_HandleCreated fired - embedding content.");
            EmbedContent();

            // GetEffectiveDpi() in the constructor ran before this.IsHandleCreated was
            // true, so it always fell back to the stale DeviceDpi (96) instead of the
            // monitor's real DPI - permanently locking MinimumSize to the un-scaled
            // 600x300px floor regardless of actual display scaling, letting the pane (and
            // its embedded content) be shrunk below the size it actually needs at the real
            // DPI. Now that this pane's own handle exists, GetEffectiveDpi() can return the
            // real per-monitor DPI, so recompute the sizing here to correct that floor.
            // Ported from FinalWorkingCode's identical fix in GLConfiguratorPane.cs.
            ApplyDpiAwareSizing(GetEffectiveDpi());
        }

        private void GLConfiguratorPane_HandleDestroyed(object sender, EventArgs e)
        {
            // Best-effort: don't tear down the Addin.Core-side Window here - it's reused
            // across pane show/hide cycles (see ConfiguratorPaneHost's idempotent
            // CreateContent). Only Shutdown/hot-reload teardown closes it for real.
            GlobalsEx.Context?.Logger?.LogDebug("GLConfiguratorPane_HandleDestroyed fired - clearing cached content HWND.");
            _contentHwnd = IntPtr.Zero;
        }

        private void EmbedContent()
        {
            try
            {
                GlobalsEx.Context?.Logger?.LogDebug("GLConfiguratorPane.EmbedContent: requesting content HWND from Addin.Core.");
                _contentHwnd = GlobalsEx.Addin?.CreateConfiguratorPaneContent() ?? IntPtr.Zero;
                if (_contentHwnd == IntPtr.Zero)
                {
                    GlobalsEx.Context?.Logger?.LogError("GLConfiguratorPane: CreateConfiguratorPaneContent returned a null handle.");
                    return;
                }

                // Restyle from a top-level popup window into embedded child content.
                long style = GetWindowLongPtr(_contentHwnd, GWL_STYLE).ToInt64();
                style &= ~(WS_POPUP | WS_CAPTION | WS_THICKFRAME | WS_SYSMENU | WS_MAXIMIZEBOX | WS_MINIMIZEBOX);
                style |= WS_CHILD;
                SetWindowLongPtr(_contentHwnd, GWL_STYLE, new IntPtr(style));

                SetParent(_contentHwnd, this.Handle);
                ResizeContent();

                // The reparented window is still pumped by Addin.Core's own WPF thread,
                // not this pane's (Excel's main STA thread) - without attaching thread
                // input, keyboard focus/Tab-navigation into the embedded content is
                // unreliable. Safe to call repeatedly/idempotently.
                uint childThreadId = GetWindowThreadProcessId(_contentHwnd, out _);
                uint thisThreadId = GetCurrentThreadId();
                if (childThreadId != 0 && childThreadId != thisThreadId)
                {
                    AttachThreadInput(thisThreadId, childThreadId, true);
                }
            }
            catch (Exception ex)
            {
                GlobalsEx.Context?.Logger?.LogException(ex, "GLConfiguratorPane.EmbedContent");
            }
        }

        private void GLConfiguratorPane_DpiChanged(object sender, DpiChangedEventArgs e)
        {
            ApplyDpiAwareSizing(e.DeviceDpiNew);
        }

        private void ApplyDpiAwareSizing(float dpiX)
        {
            var scale = dpiX / 96f;
            int minWidthPx = (int)Math.Round(_minWidthDip * scale);
            int minHeightPx = (int)Math.Round(_minHeightDip * scale);

            this.MinimumSize = new Size(minWidthPx, minHeightPx);
            this.Size = new Size(Math.Max(this.Width, minWidthPx), Math.Max(this.Height, minHeightPx));
        }

        private void GLConfiguratorPane_Resize(object sender, EventArgs e)
        {
            var dpi = GetEffectiveDpi();
            var dipWidth = this.Width * DefaultDpi / (float)dpi;

            if (dipWidth < _minWidthDip)
            {
                int minWidthPx = (int)Math.Round(_minWidthDip * dpi / (float)DefaultDpi);
                this.Width = minWidthPx;
            }

            ResizeContent();
        }

        private void ResizeContent()
        {
            if (_contentHwnd != IntPtr.Zero)
            {
                MoveWindow(_contentHwnd, 0, 0, this.ClientSize.Width, this.ClientSize.Height, true);
            }
        }

        /// <summary>
        /// Old monolith's GLConfiguratorPane.RelaunchPane(). Kept as a Task-returning
        /// method (rather than void) so every existing call site (RibFSG_OnClick,
        /// SheetSelectionChange, AddinModule's ledger-change flow) keeps working
        /// unchanged with `_ = blpane.RelaunchPane();` - the actual cross-AppDomain call
        /// is synchronous (IGLSenseAddin.RelaunchConfiguratorPane returns void, dispatches
        /// fire-and-forget on the Addin.Core side), so this just wraps it in
        /// Task.CompletedTask.
        /// </summary>
        public Task RelaunchPane()
        {
            try
            {
                GlobalsEx.Context?.Logger?.LogDebug("GLConfiguratorPane.RelaunchPane: requesting relaunch from Addin.Core.");
                GlobalsEx.Addin?.RelaunchConfiguratorPane();
            }
            catch (Exception ex)
            {
                GlobalsEx.Context?.Logger?.LogException(ex, "GLConfiguratorPane.RelaunchPane");
            }
            return Task.CompletedTask;
        }

        /// <summary>Old monolith's GLConfiguratorPane.ResetPaneReference().</summary>
        public Task ResetPaneReference()
        {
            try
            {
                GlobalsEx.Context?.Logger?.LogDebug("GLConfiguratorPane.ResetPaneReference: requesting reset from Addin.Core.");
                GlobalsEx.Addin?.ResetConfiguratorPaneReference();
            }
            catch (Exception ex)
            {
                GlobalsEx.Context?.Logger?.LogException(ex, "GLConfiguratorPane.ResetPaneReference");
            }
            return Task.CompletedTask;
        }

        private void GLConfiguratorPane_ADXBeforeTaskPaneShow(object sender, ADXBeforeTaskPaneShowEventArgs e)
        {
            try
            {
                GlobalsEx.Context?.Logger?.LogDebug($"GLConfiguratorPane_ADXBeforeTaskPaneShow fired (DisplayConfigurator={DisplayConfigurator}).");
                if (sender is GLConfiguratorPane pane)
                {
                    pane.Visible = DisplayConfigurator;

                    if (pane.Visible)
                    {
                        pane.Width = pane.MinimumSize.Width;
                    }
                }
            }
            catch (Exception ex)
            {
                GlobalsEx.Context?.Logger?.LogException(ex, "GLConfiguratorPane_ADXBeforeTaskPaneShow");
            }
        }

        protected override void WndProc(ref Message m)
        {
            // NOTE: a WM_MOUSEWHEEL branch was tried here (forwarding into Addin.Core via
            // IGLSenseAddin.TryScrollOpenComboBoxPopup) as part of the ADXTaskPane
            // mouse-wheel fix saga (CLAUDE.md section 24.3). It was removed: disabling the
            // global low-level mouse hook (SuggestAppendComboBox's other fix attempt)
            // proved this WndProc override never actually received WM_MOUSEWHEEL for the
            // popup in the first place - the message simply never reaches this control's
            // window procedure via normal routing in this hosting context, only via a raw,
            // pre-routing low-level hook. See SuggestAppendComboBox.cs for the current fix
            // (a low-level hook on its own dedicated, non-blocking thread).
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
