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
        // Absolute floor/ceiling (DIP) - the hard safety rails GetWidthBoundsPx() clamps
        // the percentage-derived bounds into, and the fallback used outright when Excel's
        // own window width isn't available yet (e.g. early startup). _minWidthDip must
        // stay >= GLBalanceConfigurator.MinimumConfiguratorWidth, which the header content
        // actually needs to fit without overflowing (icon + "Balance Configurator" title +
        // close button, all Auto-sized Grid columns - WPF's Grid does not compress Auto
        // columns to force-fit). 595 is inherited from FinalWorkingCode's measured fix for
        // the identical header XAML shape (confirmed there via direct PointToScreen
        // measurement: the close button's own computed screen position sat a consistent
        // 36px past the container's right edge at a narrower floor, with 75 DIP of
        // headroom added as a safety margin) - this project's header uses the same icon,
        // title text and CustomWindowCloseButtonStyle, so the same floor is expected to
        // hold, but this has not been independently re-measured on this project's own
        // build/DPI matrix and should be verified the same way before shipping.
        private readonly int _minWidthDip = 595;
        private readonly int _maxWidthDip = 900;
        // The pane's width target as a fraction of Excel's own current window width - see
        // GetWidthBoundsPx(). Keeps the pane a consistent proportion of the available
        // space across any monitor/resolution/DPI, instead of a fixed DIP size that could
        // eat up a large fraction of the window at high display scaling on a modest
        // resolution screen. Ported from FinalWorkingCode's identical fix.
        private const float MinWidthPercent = 0.25f;
        private const float MaxWidthPercent = 0.35f;
        // Guarantees at least this much room between the computed min and max (DIP,
        // DPI-scaled in GetWidthBoundsPx()) even when the percentage-derived max would
        // otherwise fall below the absolute floor - without this, minPx and maxPx could
        // collapse to the exact same value on a non-maximized/narrower Excel window,
        // locking the splitter dead with zero room to drag.
        private readonly int _minRangeSlackDip = 100;
        // Width the pane opens at (GLConfiguratorPane_ADXBeforeTaskPaneShow) - distinct
        // from _minWidthDip so the pane can launch wider than its floor.
        private readonly int _defaultWidthDip = 610;
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
        [DllImport("user32.dll")] private static extern bool GetWindowRect(IntPtr hWnd, out RECT lpRect);

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

        // Control.Width can go stale relative to the pane's real on-screen size - during
        // a live DPI/resolution change, Add-in Express's own internal reflow can move the
        // real native window without going through whatever path keeps WinForms' cached
        // Control.Width in sync. Always read the real rect here instead of trusting
        // this.Width for clamp math. Ported from FinalWorkingCode's identical fix.
        private int GetActualWidthPx()
        {
            if (this.IsHandleCreated && GetWindowRect(this.Handle, out RECT rect))
            {
                return rect.Right - rect.Left;
            }

            return this.Width;
        }

        // Excel.Application.Width is documented as POINTS (72/inch), but doesn't reliably
        // reflect the window's actual current-DPI physical size (confirmed on
        // FinalWorkingCode via live testing - a points/dpi conversion came out ~1.6x too
        // large at 175% scale). GlobalsEx.Context.ExcelHandle + GetWindowRect instead
        // reads the window's real physical pixel rect directly via Win32 - the same unit
        // space this.Width/GetActualWidthPx already use. Ported from FinalWorkingCode's
        // identical fix (re-pointed from AppState.Instance.ExcelApp.Hwnd to
        // GlobalsEx.Context.ExcelHandle, the host-side equivalent already used elsewhere
        // in this project - see AddinModule.cs's SheetSelectionChange focus-reclaim fix).
        private int GetExcelWindowWidthPx()
        {
            try
            {
                IntPtr hwnd = GlobalsEx.Context?.ExcelHandle ?? IntPtr.Zero;
                if (hwnd == IntPtr.Zero || !GetWindowRect(hwnd, out RECT rect))
                {
                    return 0;
                }

                return rect.Right - rect.Left;
            }
            catch (Exception ex)
            {
                GlobalsEx.Context?.Logger?.LogException(ex, "GLConfiguratorPane.GetExcelWindowWidthPx");
                return 0;
            }
        }

        // Percentage-of-Excel-window bounds, clamped into the absolute [_minWidthDip,
        // _maxWidthDip] safety rails (DPI-scaled). Falls back to those absolute bounds
        // outright when Excel's own window width isn't available yet (e.g. early
        // startup, GetExcelWindowWidthPx() returning 0). Ported from FinalWorkingCode's
        // identical fix.
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
        // Unlike FinalWorkingCode's version, there is no _wpfControl/_host (ElementHost)
        // to invalidate directly here - the embedded content is a reparented native HWND
        // (_contentHwnd) in a different AppDomain, kept in sync via ResizeContent()'s own
        // MoveWindow call, which already runs on every resize (see GLConfiguratorPane_Resize).
        // FinalWorkingCode's own InvalidateMeasure/InvalidateArrange/InvalidateVisual calls
        // were confirmed (via direct measurement) to have zero effect on the underlying bug
        // they were added for, so no cross-AppDomain equivalent is needed here either.
        public void RecomputeWidthBounds()
        {
            var (minPx, maxPx) = GetWidthBoundsPx();
            int actualWidthPx = GetActualWidthPx();
            int clampedWidth = Math.Min(Math.Max(actualWidthPx, minPx), maxPx);
            if (clampedWidth != this.Width)
            {
                GlobalsEx.Context?.Logger?.LogDebug($"GLConfiguratorPane.RecomputeWidthBounds: this.Width={this.Width} (actual={actualWidthPx}) -> {clampedWidth} (minPx={minPx}, maxPx={maxPx})");
                this.Width = clampedWidth;
            }

            ResizeContent();
        }

        public GLConfiguratorPane()
        {
            InitializeComponent();

            this.AutoScaleMode = AutoScaleMode.Dpi;
            ApplyDpiAwareSizing(GetEffectiveDpi());
            this.DpiChanged += GLConfiguratorPane_DpiChanged;
            this.HandleCreated += GLConfiguratorPane_HandleCreated;
            this.HandleDestroyed += GLConfiguratorPane_HandleDestroyed;

            // Freezes the pane's width between MinimumSize.Width and a DPI-scaled
            // _maxWidthDip during a live splitter drag. This is Add-in Express's own
            // dedicated resize-constraint event - fired on every mouse-move while the
            // user drags the splitter, distinct from the WM_SIZING/WM_WINDOWPOSCHANGING
            // overrides below (which independently enforce only the minimum, and are
            // left unchanged). e.NewRegionSize is read-only - there is no way to
            // substitute a clamped size, only accept or reject the exact proposed one -
            // but since this fires continuously during the drag, rejecting anything
            // outside the [Min, Max] range stops the pane from tracking the mouse any
            // further once it hits either bound. Ported from FinalWorkingCode's
            // identical fix (738fbed).
            this.ADXSplitterMove += GLConfiguratorPane_ADXSplitterMove;
        }

        private void GLConfiguratorPane_ADXSplitterMove(object sender, ADXSplitterMoveEventArgs e)
        {
            var (minWidthPx, maxWidthPx) = GetWidthBoundsPx();
            GlobalsEx.Context?.Logger?.LogDebug($"GLConfiguratorPane_ADXSplitterMove: NewRegionSize.Width={e.NewRegionSize.Width}, minWidthPx={minWidthPx}, maxWidthPx={maxWidthPx}");
            if (e.NewRegionSize.Width < minWidthPx || e.NewRegionSize.Width > maxWidthPx)
            {
                e.Cancel = true;
            }
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
            // Re-clamping into [min, max] here, computed fresh every time, is what keeps
            // the pane correctly sized across a live DPI/resolution change (Excel's own
            // window resizing/moving is handled separately by RecomputeWidthBounds,
            // called from AddinModule's WindowResize handler). Ported from
            // FinalWorkingCode's identical fix.
            var (minWidthPx, maxWidthPx) = GetWidthBoundsPx();
            int actualWidthPx = GetActualWidthPx();
            GlobalsEx.Context?.Logger?.LogDebug($"GLConfiguratorPane_Resize: this.Width={this.Width} (actual={actualWidthPx}), minWidthPx={minWidthPx}, maxWidthPx={maxWidthPx}");

            int clampedWidth = Math.Min(Math.Max(actualWidthPx, minWidthPx), maxWidthPx);
            if (clampedWidth != this.Width)
            {
                this.Width = clampedWidth;
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
        public Task RelaunchPane(bool showBusyOverlay = true)
        {
            try
            {
                GlobalsEx.Context?.Logger?.LogDebug($"GLConfiguratorPane.RelaunchPane: requesting relaunch from Addin.Core (showBusyOverlay={showBusyOverlay}).");
                GlobalsEx.Addin?.RelaunchConfiguratorPane(showBusyOverlay);
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

        /// <summary>
        /// Old monolith's GLConfiguratorPane.HasSavedConfigurationSelected(). Synchronous
        /// (unlike RelaunchPane/ResetPaneReference above) - IGLSenseAddin.
        /// HasSavedConfigurationSelected() blocks on Addin.Core's WPF thread via
        /// ConfiguratorPaneHost's Dispatcher.Invoke, so the result is available immediately.
        /// Returns false (safe default) on any failure or if Addin.Core isn't reachable.
        /// </summary>
        public bool HasSavedConfigurationSelected()
        {
            try
            {
                return GlobalsEx.Addin?.HasSavedConfigurationSelected() ?? false;
            }
            catch (Exception ex)
            {
                GlobalsEx.Context?.Logger?.LogException(ex, "GLConfiguratorPane.HasSavedConfigurationSelected");
                return false;
            }
        }

        private void GLConfiguratorPane_ADXBeforeTaskPaneShow(object sender, ADXBeforeTaskPaneShowEventArgs e)
        {
            try
            {
                GlobalsEx.Context?.Logger?.LogDebug($"GLConfiguratorPane_ADXBeforeTaskPaneShow fired (DisplayConfigurator={DisplayConfigurator}).");
                if (sender is GLConfiguratorPane pane)
                {
                    pane.Visible = DisplayConfigurator;

                    // Launches at _defaultWidthDip, not the MinimumSize.Width floor, so
                    // the two can differ. Ported from FinalWorkingCode's identical fix.
                    if (pane.Visible)
                    {
                        int defaultWidthPx = (int)Math.Round(_defaultWidthDip * pane.GetEffectiveDpi() / (float)DefaultDpi);
                        pane.Width = Math.Max(defaultWidthPx, pane.MinimumSize.Width);
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
