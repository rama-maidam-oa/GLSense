// GLAccountsRef.xaml.cs in GLSense.Addin.Core
// Group H (Balance Configurator) - port of GLSense\Views\GLAccountsRef.xaml.cs
// (FinalWorkingCode). Read-only "selected GL accounts" display box (with tooltip,
// Edit and Clear buttons) used by GLBalanceConfigurator.xaml's Account Assignment(s) row.
//
// The Edit button - GLSegmentRef is now ported:
//   The old BtnEdit_Click opened GLSense.Views.GLSegmentRef (a full account/segment-value
//   picker dialog) directly and copied its GLSegments_SelectedValue back into Text. This
//   control still doesn't reference GLSegmentRef directly - it stays decoupled by raising
//   the public "EditRequested" event instead (mirroring the existing Action<string>/
//   Func<...> delegate-injection pattern this project already uses for AppOverlay hooks -
//   see GLJobsMonitor.xaml.cs's ShowWarningAction/ShowBusyAction, etc.). GLBalanceConfigurator
//   now subscribes to it (see GLBalanceConfigurator.xaml.cs's AcctsRef_EditRequested), which
//   opens Views\GLSegmentRef.xaml.cs and copies its GLSegments_SelectedValue back into
//   AcctsRef.Text on OK - same end result as the original inline call, just split across
//   the event boundary. The FindWarningHost fallback below now only fires in the
//   (currently never-hit) case where nothing has subscribed to EditRequested.
//   Text/TagName/Clear-button/tooltip behavior is otherwise an unchanged, verbatim port.
using GLSense.Addin.Core.Infrastructure;
using GLSense.Addin.Core.Interfaces;
using System;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;

namespace GLSense.Addin.Core.Views
{
    /// <summary>
    /// Interaction logic for GLAccountsRef.xaml
    /// </summary>
    public partial class GLAccountsRef : UserControl
    {
        public GLAccountsRef()
        {
            InitializeComponent();
            ServiceLocator.Logger?.LogDebug("GLAccountsRef constructor invoked");
            Loaded += UserControl_Loaded;
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

        // Event passed to hosting window when Text changes.
        // Reuses GLSense.Addin.Core.Views.CellReferenceChangedEventArgs, already defined
        // in Views\ExcelRefEditControl.xaml.cs for this same namespace/purpose.
        public event EventHandler<CellReferenceChangedEventArgs> CellReferenceChanged;

        // Raised when the user clicks the Edit (hand-pointer) button. See header comment -
        // this is a hook for a future GLSegmentRef-based picker, not yet subscribed to
        // anywhere in this project.
        public event EventHandler EditRequested;

        private static void OnTextChanged(DependencyObject d, DependencyPropertyChangedEventArgs e)
        {
            var control = (GLAccountsRef)d;
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
            // Initialize tooltip with current text
            UpdateToolTipText();
        }

        private void BtnEdit_Click(object sender, RoutedEventArgs e)
        {
            ServiceLocator.Logger?.LogDebug($"GLAccountsRef.BtnEdit_Click invoked (TagName={TagName})");
            if (EditRequested != null)
            {
                EditRequested.Invoke(this, EventArgs.Empty);
                return;
            }

            // No picker wired up yet (GLSegmentRef not ported - see header comment).
            // Surface that honestly instead of doing nothing.
            ServiceLocator.Logger?.LogWarn("GLAccountsRef.BtnEdit_Click: no EditRequested subscriber - picker not wired up yet");
            var host = FindWarningHost(this);
            host?.CellSelectionWarning("Account/segment picker is not available in this build yet. Enter a value via the cell reference field instead.");
        }

        private void BtnClear_Click(object sender, RoutedEventArgs e)
        {
            ServiceLocator.Logger?.LogDebug($"GLAccountsRef.BtnClear_Click invoked (TagName={TagName})");
            this.Text = string.Empty;
            UpdateToolTipText();
        }

        private static IWarningHost FindWarningHost(DependencyObject child)
        {
            DependencyObject parent = child;

            while (parent != null)
            {
                if (parent is IWarningHost warningHost)
                    return warningHost;

                parent = VisualTreeHelper.GetParent(parent);
            }

            return null;
        }
    }
}
