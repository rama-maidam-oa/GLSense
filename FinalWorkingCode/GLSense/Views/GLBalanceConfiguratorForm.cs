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
    /// removed). That version repeatedly lost real Win32 keyboard focus back to Excel's
    /// grid every few seconds while open - confirmed via extensive diagnostic logging that
    /// GetForegroundWindow()/GetFocus() both correctly pointed at the WPF window, yet
    /// keystrokes still weren't reliably reaching it. A side-by-side comparison against
    /// this add-in's older VB.NET/WinForms sibling (C:\projects\Excel Add-ons\glsense,
    /// FormFSG.vb) - which shows an equivalent non-modal Balance Configurator form
    /// alongside Excel with ZERO special focus-handling code (no SetForegroundWindow/
    /// SetFocus/message filters/WndProc override while open) and has never had this
    /// problem - showed the leak is specific to WPF's own input pipeline in this exact
    /// in-process, same-thread Excel-hosting context (Excel owns the real Win32 message
    /// pump; a WPF Window only hooks into it via ComponentDispatcher rather than owning
    /// it, so a keystroke Excel's own message handling consumes internally can bypass
    /// WPF's HwndSource entirely even though raw Win32 focus state looks correct).
    /// Hosting the same WPF content inside a real WinForms Form sidesteps this: WinForms
    /// controls map directly onto native HWNDs with no such translation layer, which is
    /// exactly why FormFSG never needed a workaround.
    ///
    /// Deliberately does NOT replicate GLConfiguratorPane.cs's WM_SIZING/WM_WINDOWPOSCHANGING
    /// RECT-mutation overrides (the code implicated in this app's original task-pane DPI/
    /// resize ghosting bug) - a plain top-level Form's MinimumSize/MaximumSize are already
    /// natively enforced by Windows' own WM_GETMINMAXINFO handling during a live drag, with
    /// no custom interception needed, unlike the pane's ADXExcelTaskPane hosting.
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

        // TEMPORARY DIAGNOSTIC INSTRUMENTATION - remove once the typing-goes-to-Excel-cell
        // bug is root-caused for this WinForms Form. Same GetForegroundWindow/GetFocus
        // technique used to diagnose the earlier pure-WPF Window attempt.
        [DllImport("user32.dll")]
        private static extern IntPtr GetForegroundWindow();

        [DllImport("user32.dll")]
        private static extern IntPtr GetFocus();

        private void LogFocusState(string context)
        {
            try
            {
                var fg = GetForegroundWindow();
                var focus = GetFocus();
                LogUtility.LogDebug(
                    $"GLBalanceConfiguratorForm.LogFocusState[{context}]: " +
                    $"thisHwnd={Handle}, GetForegroundWindow={fg} (match={fg == Handle}), " +
                    $"GetFocus={focus} (match={focus == Handle}), ContainsFocus={ContainsFocus}, " +
                    $"ActiveControl={ActiveControl?.GetType().Name ?? "null"}");
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex, $"GLBalanceConfiguratorForm.LogFocusState[{context}]");
            }
        }

        public GLBalanceConfiguratorForm()
        {
            LogUtility.LogDebug("GLBalanceConfiguratorForm.ctor invoked");

            Text = "Balance Configurator";
            StartPosition = FormStartPosition.CenterParent;
            FormBorderStyle = FormBorderStyle.Sizable;
            MaximizeBox = false;
            AutoScaleMode = AutoScaleMode.Dpi;
            Size = new Size(650, 700);

            // TEMPORARY DIAGNOSTIC: log every keystroke the Form itself sees, and every
            // activation/deactivation/focus transition, so the next repro shows exactly
            // where this breaks rather than requiring another guess.
            KeyPreview = true;
            KeyDown += (s, e) => LogFocusState($"Form.KeyDown:{e.KeyCode}");
            Activated += (s, e) => LogFocusState("Form.Activated");
            Deactivate += (s, e) => LogFocusState("Form.Deactivate");
            GotFocus += (s, e) => LogFocusState("Form.GotFocus");
            LostFocus += (s, e) => LogFocusState("Form.LostFocus");

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

            // TEMPORARY DIAGNOSTIC: same as the Form-level hooks above, but on the
            // ElementHost itself - shows whether focus ever actually reaches the WPF
            // content's native host control, separate from the Form as a whole.
            _host.GotFocus += (s, e) => LogFocusState("ElementHost.GotFocus");
            _host.LostFocus += (s, e) => LogFocusState("ElementHost.LostFocus");

            _configuratorControl.OnCloseRequested += () => Close();

            ApplyDpiAwareSizing(DefaultDpi);
            HandleCreated += GLBalanceConfiguratorForm_HandleCreated;
            DpiChanged += (s, e) => ApplyDpiAwareSizing(e.DeviceDpiNew);
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

        /// <summary>
        /// Shows this Form non-modally, owned by Excel's main window.
        /// </summary>
        public void ShowFloating(IntPtr excelHwnd)
        {
            LogUtility.LogDebug($"GLBalanceConfiguratorForm.ShowFloating invoked. excelHwnd={excelHwnd}");
            Show(new Win32Window(excelHwnd));
            LogFocusState("ShowFloating:after-show");
        }

        /// <summary>Mirrors GLConfiguratorPane.RelaunchPane (GLConfiguratorPane.cs:216-234).</summary>
        public async Task RelaunchWindow()
        {
            try
            {
                LogUtility.LogDebug("GLBalanceConfiguratorForm.RelaunchWindow invoked.");
                if (_configuratorControl != null)
                {
                    await _configuratorControl.ReLoadConfigurator();
                }
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex, "GLBalanceConfiguratorForm.RelaunchWindow");
            }
        }

        /// <summary>Mirrors GLConfiguratorPane.ResetPaneReference (GLConfiguratorPane.cs:235-252).</summary>
        public Task ResetWindowReference()
        {
            try
            {
                LogUtility.LogDebug("GLBalanceConfiguratorForm.ResetWindowReference invoked.");
                GLBalanceConfigurator.ResetCellReference();
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex, "GLBalanceConfiguratorForm.ResetWindowReference");
            }

            return Task.CompletedTask;
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
