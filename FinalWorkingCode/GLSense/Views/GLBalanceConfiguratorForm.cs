using GLSense.Utilities;
using System;
using System.Drawing;
using System.Runtime.InteropServices;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Windows.Forms.Integration;

namespace GLSense.Views
{
    /// <summary>
    /// Floating, non-modal home for the Balance Configurator - a plain WinForms Form
    /// hosting the existing, unmodified GLBalanceConfigurator WPF UserControl via
    /// ElementHost, instead of a pure WPF Window.
    /// </summary>
    /// <remarks>
    /// This replaces an earlier pure-WPF Window implementation (GLBalanceConfiguratorWindow,
    /// removed), which repeatedly lost real Win32 keyboard focus back to Excel's grid every
    /// few seconds while open. Root-caused via a controlled A/B test (confirmed twice, with
    /// opposite results both times): the actual cause was GLSense.Utilities.ComMessageFilter
    /// - a process-wide IOleMessageFilter that used to be registered for the life of the
    /// add-in. Merely having ANY IOleMessageFilter registered - regardless of what value its
    /// MessagePending callback returns - changes how COM pumps window-activation-related
    /// messages while an outgoing call to Excel is pending (which is constantly, since
    /// nearly every Excel property read is one), letting Excel's own window silently
    /// reclaim OS activation from any separate top-level window this add-in creates. This
    /// reproduced identically regardless of the window's own implementation (a WPF Window,
    /// this WinForms Form hosting WPF via ElementHost, and even a bare WinForms Form with
    /// zero WPF content all failed the same way with the filter registered, and all worked
    /// once it was removed) - it was never actually about WPF vs. WinForms, or about this
    /// Form's own code. The filter is now scoped narrowly instead: CommonMethods.cs's
    /// DisableExcelSettings/EnableExcelSettings register/revoke it only for the duration of
    /// the bulk operations that actually need its "retry Excel's busy rejections" benefit.
    ///
    /// This Form is kept as a plain WinForms Form (rather than reverting to a WPF Window)
    /// since it works correctly and there is no reason to reintroduce WPF's extra hosting
    /// complexity now that the real root cause is fixed. Deliberately does NOT replicate
    /// GLConfiguratorPane.cs's WM_SIZING/WM_WINDOWPOSCHANGING RECT-mutation overrides (the
    /// code implicated in this app's original task-pane DPI/resize ghosting bug) - a plain
    /// top-level Form's MinimumSize/MaximumSize are already natively enforced by Windows'
    /// own WM_GETMINMAXINFO handling during a live drag, with no custom interception needed,
    /// unlike the pane's ADXExcelTaskPane hosting.
    /// </remarks>
    public class GLBalanceConfiguratorForm : Form
    {
        private readonly GLBalanceConfigurator _configuratorControl;
        private readonly ElementHost _host;
        private readonly int _minWidthDip = 600;
        private readonly int _minHeightDip = 300;
        private readonly int _maxWidthDip = 1000;
        private readonly int _maxHeightDip = 900;
        private const int DefaultDpi = 96;

        public GLBalanceConfiguratorForm()
        {
            LogUtility.LogDebug("GLBalanceConfiguratorForm.ctor invoked");

            Text = "Balance Configurator";
            // Manual, not CenterParent: position is set explicitly below in
            // GLBalanceConfiguratorForm_HandleCreated, matching the VB.NET sibling's
            // FrmFSG_Load positioning exactly (flush against the right edge of the
            // primary screen's working area, 50px up from the bottom).
            StartPosition = FormStartPosition.Manual;
            // No native title bar/caption: GLBalanceConfigurator (the hosted WPF control)
            // already draws its own header (icon, title, close button) - a native caption
            // on top of that would duplicate it, the same bug the earlier WPF Window
            // implementation hit and fixed by removing its own duplicate header. The
            // CreateParams override below adds WS_THICKFRAME so the Form stays resizable
            // by dragging its edges despite having no visible border/caption. Since removing
            // the native caption also removes its native drag-to-move behavior, that is
            // restored below via OnHeaderDragRequested/StartCaptionDrag instead.
            FormBorderStyle = FormBorderStyle.None;
            AutoScaleMode = AutoScaleMode.Dpi;
            Size = new Size(680, 800);

            // See GLConfiguratorPane.cs's own "REVERT NOTE" comment for why this scope
            // matters: constructing the WPF object here does NOT create its native HWND
            // (WinForms/WPF interop realizes handles lazily, cascading down from this
            // Form's own handle) - only wrapping the managed construction in the
            // per-monitor-aware context would miss the real handle creation entirely.
            // HandleCreated below re-applies the same context to cover the normal case
            // where this Form's handle (and so the ElementHost's cascade-created handle)
            // isn't realized until after this constructor returns.
            using (DpiAwarenessHelper.SetPerMonitorAware())
            {
                _configuratorControl = new GLBalanceConfigurator();

                _host = new ElementHost
                {
                    Dock = DockStyle.Fill,
                    Child = _configuratorControl
                };

                Controls.Add(_host);
            }

            _configuratorControl.OnCloseRequested += () => Close();

            // FormBorderStyle=None removed the native caption, so there is no native
            // drag-to-move affordance left. A Form-level WM_NCHITTEST-to-HTCAPTION
            // override (the usual borderless-Form drag trick) cannot substitute for it
            // here: the header is WPF content hosted in an ElementHost that fills this
            // Form's entire client area, so mouse messages over the header go straight to
            // that child HWND and never reach this Form's own WndProc at all (confirmed
            // via diagnostic logging - see git history for GLBalanceConfiguratorForm.cs).
            // Instead, GLBalanceConfigurator raises OnHeaderDragRequested when its header
            // Border is pressed, and StartCaptionDrag below forwards that press to Windows
            // as a native caption drag on this Form's own window - the same mechanism
            // WM_NCLBUTTONDOWN/HTCAPTION drives for a real title bar.
            _configuratorControl.OnHeaderDragRequested += StartCaptionDrag;

            // Puts initial keyboard focus on the Ledger field once the Form is actually
            // shown, so the first keystroke after opening lands in a real data field
            // rather than depending on WinForms' own default "focus the first control"
            // behavior (which would land on the ElementHost itself, not a specific WPF
            // field inside it).
            Shown += (s, e) => _configuratorControl?.CmbLedgers?.Focus();

            ApplyDpiAwareSizing(DefaultDpi);
            HandleCreated += GLBalanceConfiguratorForm_HandleCreated;
            DpiChanged += (s, e) => ApplyDpiAwareSizing(e.DeviceDpiNew);
        }

        /// <summary>
        /// WS_EX_TOOLWINDOW matches the VB.NET sibling's FormFSG.vb CreateParams override
        /// (tested against the cell-edit-mode click issue - did NOT resolve it, that is
        /// still unexplained and needs real investigation rather than another guess; kept
        /// anyway since it's still a reasonable, harmless taskbar/Alt-Tab visibility
        /// choice matching FormFSG). WS_THICKFRAME restores native resize-by-dragging-edges
        /// now that FormBorderStyle=None has removed the visible border that normally
        /// provides it.
        /// </summary>
        protected override CreateParams CreateParams
        {
            get
            {
                const int WS_EX_TOOLWINDOW = 0x80;
                const int WS_THICKFRAME = 0x00040000;
                CreateParams cp = base.CreateParams;
                cp.ExStyle |= WS_EX_TOOLWINDOW;
                cp.Style |= WS_THICKFRAME;
                return cp;
            }
        }

        [DllImport("user32.dll")]
        private static extern bool ReleaseCapture();

        [DllImport("user32.dll")]
        private static extern int SendMessage(IntPtr hWnd, int msg, int wParam, int lParam);

        /// <summary>Forwards a WPF header press to Windows as a native caption drag on
        /// this Form's own window. ReleaseCapture drops WPF's mouse capture first (it
        /// would otherwise swallow the drag); the WM_NCLBUTTONDOWN/HTCAPTION combination
        /// is the standard trick for driving a native move-drag from content that isn't
        /// itself the window's non-client area - Windows runs the same drag loop, with
        /// the same edge snapping, it would for a real title bar.</summary>
        private void StartCaptionDrag()
        {
            const int WM_NCLBUTTONDOWN = 0x00A1;
            const int HTCAPTION = 2;
            ReleaseCapture();
            SendMessage(Handle, WM_NCLBUTTONDOWN, HTCAPTION, 0);
        }

        // TEMPORARY DIAGNOSTIC: capturing the raw Win32 message sequence this Form
        // actually receives (or doesn't) for a click while an Excel cell is in edit
        // mode, since source-level comparison against FormFSG.vb turned up no
        // explanation. Remove once the edit-mode-click issue is root-caused.
        [DllImport("user32.dll")]
        private static extern IntPtr GetForegroundWindow();

        [DllImport("user32.dll")]
        private static extern IntPtr GetFocus();

        private static void LogDiagnosticMsg(string msgName, IntPtr hwnd)
        {
            try
            {
                LogUtility.LogDebug(
                    $"[EditModeDiag] {msgName} thisHwnd=0x{hwnd.ToInt64():X} " +
                    $"foregroundHwnd=0x{GetForegroundWindow().ToInt64():X} " +
                    $"focusHwnd=0x{GetFocus().ToInt64():X}");
            }
            catch
            {
                // diagnostic-only, never let logging failure affect WndProc
            }
        }

        protected override void WndProc(ref Message m)
        {
            const int WM_MOUSEACTIVATE = 0x0021;
            const int WM_NCACTIVATE = 0x0086;
            const int WM_ACTIVATE = 0x0006;
            const int WM_SETFOCUS = 0x0007;
            const int WM_KILLFOCUS = 0x0008;
            const int WM_LBUTTONDOWN = 0x0201;
            const int WM_NCLBUTTONDOWN = 0x00A1;

            switch (m.Msg)
            {
                case WM_MOUSEACTIVATE: LogDiagnosticMsg("WM_MOUSEACTIVATE(before)", m.HWnd); break;
                case WM_NCACTIVATE: LogDiagnosticMsg("WM_NCACTIVATE(before)", m.HWnd); break;
                case WM_ACTIVATE: LogDiagnosticMsg("WM_ACTIVATE(before)", m.HWnd); break;
                case WM_SETFOCUS: LogDiagnosticMsg("WM_SETFOCUS(before)", m.HWnd); break;
                case WM_KILLFOCUS: LogDiagnosticMsg("WM_KILLFOCUS(before)", m.HWnd); break;
                case WM_LBUTTONDOWN: LogDiagnosticMsg("WM_LBUTTONDOWN(before)", m.HWnd); break;
                case WM_NCLBUTTONDOWN: LogDiagnosticMsg("WM_NCLBUTTONDOWN(before)", m.HWnd); break;
            }

            base.WndProc(ref m);

            switch (m.Msg)
            {
                case WM_MOUSEACTIVATE: LogDiagnosticMsg($"WM_MOUSEACTIVATE(after, result={m.Result.ToInt32()})", m.HWnd); break;
                case WM_NCACTIVATE: LogDiagnosticMsg($"WM_NCACTIVATE(after, result={m.Result.ToInt32()})", m.HWnd); break;
            }
        }

        private void GLBalanceConfiguratorForm_HandleCreated(object sender, EventArgs e)
        {
            using (DpiAwarenessHelper.SetPerMonitorAware())
            {
                // Forces the ElementHost's native window to be realized now, under the
                // per-monitor-aware context, if it hasn't been already - same fix
                // GLConfiguratorPane.cs's own HandleCreated handler applies.
                if (_host != null)
                {
                    _ = _host.Handle;
                }
            }

            ApplyDpiAwareSizing(GetEffectiveDpi());
            ApplyVbNetStylePosition();
        }

        /// <summary>Matches the VB.NET sibling's FrmFSG_Load exactly: "Me.Location = New
        /// Point(Screen.PrimaryScreen.WorkingArea.Width - Me.Width,
        /// Screen.PrimaryScreen.WorkingArea.Height - Me.Height - 50)" - flush against the
        /// right edge of the PRIMARY screen's working area (not necessarily the screen
        /// Excel itself is on - the VB.NET version has this same quirk), 50px up from the
        /// bottom. Called after ApplyDpiAwareSizing so Width/Height already reflect the
        /// real per-monitor-scaled size.</summary>
        private void ApplyVbNetStylePosition()
        {
            var workingArea = Screen.PrimaryScreen.WorkingArea;
            Location = new Point(
                workingArea.Width - Width,
                workingArea.Height - Height - 50);
        }

        private int GetEffectiveDpi()
        {
            try
            {
                if (IsHandleCreated)
                {
                    return DeviceDpi > 0 ? DeviceDpi : DefaultDpi;
                }
            }
            catch
            {
                // ignore and fall back
            }

            return DefaultDpi;
        }

        private void ApplyDpiAwareSizing(float dpi)
        {
            var scale = dpi / DefaultDpi;
            int minWidthPx = (int)Math.Round(_minWidthDip * scale);
            int minHeightPx = (int)Math.Round(_minHeightDip * scale);
            int maxWidthPx = (int)Math.Round(_maxWidthDip * scale);
            int maxHeightPx = (int)Math.Round(_maxHeightDip * scale);

            MinimumSize = new Size(minWidthPx, minHeightPx);
            MaximumSize = new Size(maxWidthPx, maxHeightPx);

            Size = new Size(
                Math.Min(Math.Max(Width, minWidthPx), maxWidthPx),
                Math.Min(Math.Max(Height, minHeightPx), maxHeightPx));
        }

        /// <summary>The Excel top-level window this Form is currently owned by. Mirrors
        /// the VB.NET sibling's FormFSG.ExcelHWND - RibFSGWindow_OnClick compares this
        /// against the currently active Excel window to decide whether the Form needs to
        /// be re-shown under a new owner (Book1/Book2 are separate top-level Excel
        /// windows; an owned window otherwise stays tied to its original owner even after
        /// that owner is no longer the active workbook).</summary>
        public IntPtr ExcelHwnd { get; private set; }

        /// <summary>
        /// Shows this Form non-modally, owned by Excel's main window. Safe to call again
        /// on an already-constructed Form to re-parent it to a different Excel window -
        /// same as the VB.NET sibling's FormFSG.Show(New WindowWrapper(excelHwnd)),
        /// called again whenever FSGForm.ExcelHWND changes.
        /// </summary>
        public void ShowFloating(IntPtr excelHwnd)
        {
            LogUtility.LogDebug($"GLBalanceConfiguratorForm.ShowFloating invoked. excelHwnd={excelHwnd}");
            ExcelHwnd = excelHwnd;
            Show(new Win32Window(excelHwnd));
        }

        /// <summary>Mirrors GLConfiguratorPane.RelaunchPane (GLConfiguratorPane.cs:216-234).
        /// showBusyOverlay=false is passed by AddinModule.ApplyBalanceWindowVisibility for
        /// the "Show Always, just clicked a different cell" case - see
        /// GLBalanceConfigurator.ReLoadConfigurator's own doc comment.</summary>
        public async Task RelaunchWindow(bool showBusyOverlay = true)
        {
            try
            {
                LogUtility.LogDebug($"GLBalanceConfiguratorForm.RelaunchWindow invoked. showBusyOverlay={showBusyOverlay}");
                // TEMPORARY DIAGNOSTIC: baseline checkpoint - the caller's thread (an
                // Excel COM event or ribbon click) before crossing into the WPF control.
                // See GLBalanceConfigurator.LogThreadDiag for the rest of the chain.
                LogUtility.LogDebug(
                    $"[ThreadDiag] RelaunchWindow: entry currentThread={System.Threading.Thread.CurrentThread.ManagedThreadId} " +
                    $"formInvokeRequired={InvokeRequired}");
                if (_configuratorControl != null)
                {
                    await _configuratorControl.ReLoadConfigurator(showBusyOverlay);
                }
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex, "GLBalanceConfiguratorForm.RelaunchWindow");
            }
        }

        /// <summary>Minimal IWin32Window wrapper for Show(owner) - same pattern as the
        /// VB.NET sibling's WindowWrapper.vb.</summary>
        private sealed class Win32Window : IWin32Window
        {
            public Win32Window(IntPtr handle) => Handle = handle;
            public IntPtr Handle { get; }
        }
    }
}
