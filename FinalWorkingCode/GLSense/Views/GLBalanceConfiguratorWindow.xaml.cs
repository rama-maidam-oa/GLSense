using GLSense.Helpers;
using GLSense.Utilities;
using System;
using System.Threading.Tasks;
using System.Windows;
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

            // DpiAwareWindow.FitToAvailableWorkArea overwrites MaxWidth/MaxHeight at
            // runtime from MaxWidthCap/MaxHeightCap (Utilities\DpiAwareWindow.cs:574-578,
            // 951-958), NOT from the MaxWidth/MaxHeight declared in this window's XAML.
            // Left at their defaults (1400d / null) the XAML's MaxWidth="1000"
            // MaxHeight="900" would silently be replaced by 1400 and by a work-area-derived
            // height. Mirror the XAML values here so the base class clamps to the intended
            // size instead of overriding it.
            MaxWidthCap = 1000d;
            MaxHeightCap = 900d;

            // The task pane cannot be Escape-closed at all, and Escape here is worse than
            // merely inconsistent: DpiAwareWindow.OnWindowPreviewKeyDown only suppresses
            // Escape while IsInteractionOverlayVisible() is true, and that method resolves
            // the overlay with FindName("AppOverlayControl") against the WINDOW's own XAML
            // namescope. Every other DpiAwareWindow declares <views:AppOverlay
            // x:Name="AppOverlayControl"/> in its own xaml; this one does not - its overlay
            // lives inside the hosted GLBalanceConfigurator control, a separate namescope -
            // so the lookup always returns null here and Escape would close the window even
            // mid-operation (e.g. while "Reloading Configurator" is showing). Reaching into
            // the hosted control's namescope would be fragile; disabling Escape entirely
            // matches the task pane and is the safe behaviour.
            EnableEscapeToClose = false;

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
            FocusFirstDataField();
        }

        /// <summary>
        /// Puts initial keyboard focus on the first real data field of the hosted
        /// configurator.
        /// </summary>
        /// <remarks>
        /// Deliberately NOT MoveFocus(FocusNavigationDirection.First). In visual-tree
        /// order the first focusable element under this window is the close button in
        /// GLBalanceConfigurator's own HeaderBar (Views\GLBalanceConfigurator.xaml:136),
        /// declared before any data control, and CustomWindowCloseButtonStyle does not set
        /// Focusable="False" - so a First traversal lands on the X and the very first
        /// Space/Enter keystroke after opening closes the window. Targeting the Ledger
        /// combo explicitly avoids that. If it cannot take focus (e.g. still disabled
        /// while the configurator loads), focus simply stays on the window itself, which
        /// is harmless - no button is focused, so no keystroke can activate one.
        /// </remarks>
        private void FocusFirstDataField()
        {
            try
            {
                bool focused = ConfiguratorControl?.CmbLedgers?.Focus() ?? false;
                if (!focused)
                {
                    LogUtility.LogDebug("GLBalanceConfiguratorWindow.FocusFirstDataField: first data field did not take focus; leaving focus on the window");
                }
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex, "GLBalanceConfiguratorWindow.FocusFirstDataField");
            }
        }

        /// <summary>
        /// Mirrors GLConfiguratorPane.RelaunchPane (GLConfiguratorPane.cs:216-234),
        /// including its try/catch + LogUtility.LogException wrapper: both are called
        /// fire-and-forget from AddinModule's SheetSelectionChange handler, so without the
        /// catch an exception would be swallowed by the discarded Task with no log line.
        /// </summary>
        public async Task RelaunchWindow()
        {
            try
            {
                LogUtility.LogDebug("GLBalanceConfiguratorWindow.RelaunchWindow invoked.");
                if (ConfiguratorControl != null && ConfiguratorControl.Dispatcher != null)
                {
                    // Dispatcher.InvokeAsync(async () => ...) doesn't wait for the inner Task -
                    // the DispatcherOperation completes as soon as the delegate hits its first
                    // await. Use a non-async delegate + Task.Unwrap() so the real completion is
                    // awaited (same pattern as GLConfiguratorPane.RelaunchPane).
                    await ConfiguratorControl.Dispatcher.InvokeAsync(() => ConfiguratorControl.ReLoadConfigurator()).Task.Unwrap();
                }
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex, "GLBalanceConfiguratorWindow.RelaunchWindow");
            }
        }

        /// <summary>
        /// Mirrors GLConfiguratorPane.ResetPaneReference (GLConfiguratorPane.cs:235-252).
        /// </summary>
        public async Task ResetWindowReference()
        {
            try
            {
                LogUtility.LogDebug("GLBalanceConfiguratorWindow.ResetWindowReference invoked.");
                if (ConfiguratorControl != null && ConfiguratorControl.Dispatcher != null)
                {
                    await ConfiguratorControl.Dispatcher.InvokeAsync(() =>
                    {
                        GLBalanceConfigurator.ResetCellReference();
                    });
                }
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex, "GLBalanceConfiguratorWindow.ResetWindowReference");
            }
        }
    }
}
