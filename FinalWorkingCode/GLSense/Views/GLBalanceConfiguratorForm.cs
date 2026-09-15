using GLSense.Utilities;
using System;
using System.Drawing;
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
            StartPosition = FormStartPosition.CenterParent;
            FormBorderStyle = FormBorderStyle.Sizable;
            MaximizeBox = false;
            AutoScaleMode = AutoScaleMode.Dpi;
            Size = new Size(650, 700);

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
