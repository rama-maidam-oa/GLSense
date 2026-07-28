using GLSense.Utilities;
using System;
using System.Windows;
using System.Windows.Controls;

namespace GLSense.Views
{
    /// <summary>
    /// Interaction logic for GLAccountsRef.xaml
    /// </summary>
    public partial class GLAccountsRef : UserControl
    {
        public GLAccountsRef()
        {
            InitializeComponent();
            Loaded += UserControl_Loaded;
            LogUtility.LogDebug("GLAccountsRef.GLAccountsRef: control constructed");
        }

        public static readonly DependencyProperty TextProperty =
            DependencyProperty.Register("Text", typeof(string), typeof(GLAccountsRef),
                new FrameworkPropertyMetadata(string.Empty, FrameworkPropertyMetadataOptions.BindsTwoWayByDefault, OnTextChanged));

        public string Text
        {
            get => (string)GetValue(TextProperty);
            set => SetValue(TextProperty, value);
        }

        public static readonly DependencyProperty TagNameProperty =
            DependencyProperty.Register("TagName", typeof(string), typeof(GLAccountsRef),
                new PropertyMetadata(string.Empty));

        public string TagName
        {
            get => (string)GetValue(TagNameProperty);
            set => SetValue(TagNameProperty, value);
        }

        // Event passed to hosting window when Text changes
        public event EventHandler<CellReferenceChangedEventArgs> CellReferenceChanged;

        private static void OnTextChanged(DependencyObject d, DependencyPropertyChangedEventArgs e)
        {
            var control = (GLAccountsRef)d;
            LogUtility.LogDebug($"GLAccountsRef.OnTextChanged: TagName={control.TagName}, NewValue={e.NewValue}");
            control.UpdateToolTipText();
            control.RaiseCellReferenceChanged(e.NewValue?.ToString());
        }

        private void RaiseCellReferenceChanged(string newRef)
        {
            CellReferenceChanged?.Invoke(this, new CellReferenceChangedEventArgs(TagName, newRef));
        }

        // Method to update tooltip text
        private void UpdateToolTipText()
        {
            if (toolTipTextBlock != null)
            {
                toolTipTextBlock.Text = string.IsNullOrEmpty(Text) ? "No accounts selected" : Text;
            }
        }

        private void UserControl_Loaded(object sender, RoutedEventArgs e)
        {
            LogUtility.LogDebug($"GLAccountsRef.UserControl_Loaded invoked - TagName={TagName}");
            // Initialize tooltip with current text
            UpdateToolTipText();
        }

        private void BtnEdit_Click(object sender, RoutedEventArgs e)
        {
            LogUtility.LogDebug($"GLAccountsRef.BtnEdit_Click invoked - TagName={TagName}");
            try
            {
                GLSegmentRef dlg = new(Text);
                dlg.EnableAutoLayoutRefresh = false;
                dlg.EnableExcelCentering = false;
                dlg.WindowStartupLocation = WindowStartupLocation.CenterOwner;

                if (dlg.ShowDialogWithOwner((IntPtr)AppState.Instance.ExcelApp.Hwnd) == true)
                {
                    Text = dlg.GLSegments_SelectedValue;
                    UpdateToolTipText();
                    LogUtility.LogDebug($"GLAccountsRef.BtnEdit_Click: user confirmed selection - GLSegments_SelectedValue={dlg.GLSegments_SelectedValue}");
                }
                else
                {
                    LogUtility.LogDebug("GLAccountsRef.BtnEdit_Click: dialog cancelled by user");
                }
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex, "Reference Button Click");
            }
        }

        private void BtnClear_Click(object sender, RoutedEventArgs e)
        {
            LogUtility.LogDebug($"GLAccountsRef.BtnClear_Click invoked - TagName={TagName}");
            this.Text = string.Empty;
            UpdateToolTipText();
        }
    }

    public class CellReferenceChangedEventArgs(string tag, string newRef) : EventArgs
    {
        public string TagName { get; set; } = tag;
        public string NewReference { get; set; } = newRef;
    }
}