using GLSense.Helpers;
using GLSense.Utilities;
using System.Windows;

namespace GLSense.Views
{
    /// <summary>
    /// Interaction logic for GLLoginDetails.xaml
    /// </summary>
    public partial class GLLoginDetails : DpiAwareWindow
    {
        public GLLoginDetails()
        {
            InitializeComponent();
            EnhancedDragDropHelper.EnableWindowDrag(this);
            LogUtility.LogDebug("GLLoginDetails constructor invoked");

        }

        private void BtnClose_Click(object sender, RoutedEventArgs e)
        {
            LogUtility.LogDebug("GLLoginDetails.BtnClose_Click invoked");
            Close();
        }
    }
}

