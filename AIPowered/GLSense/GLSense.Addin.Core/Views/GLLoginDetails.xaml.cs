// GLLoginDetails.xaml.cs in GLSense.Addin.Core
// Port of GLSense\Views\GLLoginDetails.xaml.cs (FinalWorkingCode) - a tiny read-only
// dialog (RibDBL1 ribbon button) showing the current AppState.Instance.LoginUserName/
// LoginUrl. No code-behind logic in the original beyond InitializeComponent + a close
// button - the two TextBoxes bind straight to AppState.Instance via {x:Static} in XAML
// (this project's AppState.cs, same pattern as the old monolith's), so nothing needs to
// be set up here either.
//   - DpiAwareWindow -> DpiAwareWindow (same as every other already-ported window).
//   - EnhancedDragDropHelper.EnableWindowDrag(this) -> TitleBar_MouseLeftButtonDown,
//     already wired in the XAML's title-bar Grid (see GLDailyRates.xaml.cs for the
//     identical pattern).
using GLSense.Addin.Core.Infrastructure;
using System;
using System.Windows.Input;

namespace GLSense.Addin.Core.Views
{
    /// <summary>
    /// Interaction logic for GLLoginDetails.xaml
    /// </summary>
    public partial class GLLoginDetails : DpiAwareWindow
    {
        public GLLoginDetails()
        {
            InitializeComponent();
            ServiceLocator.Logger?.LogDebug("GLLoginDetails constructor invoked");
        }

        private void TitleBar_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            try
            {
                if (e.LeftButton == MouseButtonState.Pressed)
                {
                    this.DragMove();
                }
            }
            catch (Exception ex)
            {
                ServiceLocator.Logger?.LogException(ex, "TitleBar_MouseLeftButtonDown error");
            }
        }

        private void BtnClose_Click(object sender, System.Windows.RoutedEventArgs e)
        {
            ServiceLocator.Logger?.LogDebug("GLLoginDetails.BtnClose_Click invoked - closing window");
            Close();
        }
    }
}
