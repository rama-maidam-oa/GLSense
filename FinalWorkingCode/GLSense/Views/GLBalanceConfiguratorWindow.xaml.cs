using GLSense.Helpers;
using GLSense.Utilities;
using System;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Input;
using System.Windows.Interop;

namespace GLSense.Views
{
    public partial class GLBalanceConfiguratorWindow : DpiAwareWindow
    {
        // GLBalanceConfigurator's only constructor takes an optional GLConfiguratorPane
        // parameter (Views\GLBalanceConfigurator.xaml.cs:49), so it has no true
        // zero-argument constructor at the CLR metadata level and cannot be instantiated
        // directly as a XAML element carrying x:Name (WPF markup compiler error MC3054:
        // "cannot have a Name attribute... types without a default constructor"). It is
        // built here in code instead and hosted in the ConfiguratorHost placeholder
        // declared in the XAML, keeping GLBalanceConfigurator itself unmodified.
        private readonly GLBalanceConfigurator ConfiguratorControl;

        public GLBalanceConfiguratorWindow()
        {
            LogUtility.LogDebug("GLBalanceConfiguratorWindow.ctor invoked");

            // Deliberately leave DisableAutoSizing at its default (false) - it looks
            // tempting to set true here to stop DpiAwareWindow's auto-fit/recenter passes
            // from fighting a user's manual drag-resize, but that flag is an all-or-nothing
            // switch (Utilities\DpiAwareWindow.cs:54-58): it ALSO disables the WM_DPICHANGED
            // handler's ApplyScaleTransform call (:452,485), which is the actual mechanism
            // that rescales WPF content when this window moves to a different-DPI monitor.
            // Setting it true would silently reintroduce the exact DPI-glitch bug this
            // window exists to fix. None of the ~15 existing floating windows in this repo
            // set this flag (confirmed via grep) - the fixed MinWidth/MaxWidth/MinHeight/
            // MaxHeight in this window's XAML are plain WPF Window properties, already
            // enforced natively by ResizeMode="CanResize" independent of this flag, so
            // there is nothing to disable here.

            InitializeComponent();

            ConfiguratorControl = new GLBalanceConfigurator();
            ConfiguratorHost.Content = ConfiguratorControl;

            EnhancedDragDropHelper.EnableWindowDrag(this);

            ConfiguratorControl.OnCloseRequested += () => Close();
        }

        /// <summary>
        /// Shows this window non-modally owned by Excel's main window, then forces real
        /// OS foreground/keyboard-focus onto it - see the ForceSetForegroundWindow doc
        /// comment in ExcelWindowPositioning.cs for why a plain Activate() alone is not
        /// always enough here.
        /// </summary>
        public void ShowFloating(IntPtr excelHwnd)
        {
            LogUtility.LogDebug("GLBalanceConfiguratorWindow.ShowFloating invoked");
            ShowWithOwner(excelHwnd);
            ReclaimForegroundFocus();
        }

        private void ReclaimForegroundFocus()
        {
            try
            {
                Activate();
                var hwnd = new WindowInteropHelper(this).Handle;
                if (hwnd != IntPtr.Zero)
                {
                    ExcelWindowHelper.ForceSetForegroundWindow(hwnd);
                }
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex, "GLBalanceConfiguratorWindow.ReclaimForegroundFocus");
            }
        }

        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            LogUtility.LogDebug("GLBalanceConfiguratorWindow.Window_Loaded invoked");

            // A freshly-shown non-modal window can be the active window yet still have
            // no WPF element holding keyboard focus - explicitly move focus into the
            // content so the very first keystroke after opening lands in a field here,
            // not wherever WPF keyboard focus last was.
            ReclaimForegroundFocus();
            MoveFocus(new TraversalRequest(FocusNavigationDirection.First));
        }

        private void BtnClose_Click(object sender, RoutedEventArgs e)
        {
            LogUtility.LogDebug("GLBalanceConfiguratorWindow.BtnClose_Click invoked");
            Close();
        }

        public Task RelaunchWindow() => ConfiguratorControl.ReLoadConfigurator();

        public Task ResetWindowReference()
        {
            GLBalanceConfigurator.ResetCellReference();
            return Task.CompletedTask;
        }
    }
}
