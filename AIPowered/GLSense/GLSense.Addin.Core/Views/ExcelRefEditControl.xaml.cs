// ExcelRefEditControl.xaml.cs in GLSense.Addin.Core
// Port of GLSense\Views\ExcelRefEditControl.xaml.cs (FinalWorkingCode) for Group C
// (Segment/Period pickers) - shared RefEdit-style dual text/cell-reference input custom
// control used by all 7 Group C windows. Ported FIRST since the other views' XAML
// declares it.
// Re-pointed vs. the original:
//   - GLSense.Interfaces.IWarningHost -> GLSense.Addin.Core.Interfaces.IWarningHost.
//   - GLSense.Utilities.AppState.Instance.ExcelApp -> GLSense.Addin.Core.Infrastructure.
//     ServiceLocator.ExcelApp (this project's AppState has no ExcelApp property).
//   - GLSense.Helpers.LogUtility.* -> ServiceLocator.Logger?.*.
//   - GLSense.Helpers.ExcelRefManager -> GLSense.Addin.Core.Helpers.ExcelRefManager
//     (ported alongside this file - see that file's header).
//   - CellReferenceChangedEventArgs: in the old project this lived in
//     Views\GLAccountsRef.xaml.cs (Group D/Balance Configurator territory, out of scope
//     for Group C). Defined directly in this file instead of dragging in that unrelated
//     class - it's a tiny 2-property EventArgs with no other dependencies.
// No other logic changes vs. the original.
using GLSense.Addin.Core.Helpers;
using GLSense.Addin.Core.Infrastructure;
using GLSense.Addin.Core.Interfaces;
using System;
using System.ComponentModel;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using Excel = Microsoft.Office.Interop.Excel;

namespace GLSense.Addin.Core.Views
{
    /// <summary>
    /// Interaction logic for ExcelRefEditControl.xaml
    /// </summary>
    public partial class ExcelRefEditControl : UserControl
    {
        private string _currentToolTipText = "No cell selected";
        public ExcelRefEditControl()
        {
            InitializeComponent(); // Don't forget this!
            // Not logged: a single window can host many of these controls (e.g. one
            // per ledger/activity/period field), so this fired once per instance with
            // no data - just noise duplicated across every field on every window.
            Loaded += ExcelRefEditControl_Loaded;
        }

        public static readonly DependencyProperty TextProperty = DependencyProperty.Register(
            "Text", typeof(string), typeof(ExcelRefEditControl),
            new FrameworkPropertyMetadata(string.Empty, OnTextChanged));

        public string Text
        {
            get { return (string)GetValue(TextProperty); }
            set { SetValue(TextProperty, value); }
        }

        public static readonly DependencyProperty TagNameProperty =
            DependencyProperty.Register("TagName", typeof(string), typeof(ExcelRefEditControl), new PropertyMetadata(string.Empty));

        public string TagName
        {
            get { return (string)GetValue(TagNameProperty); }
            set { SetValue(TagNameProperty, value); }
        }

        // Event passed to hosting window when Text changes
        public event EventHandler<CellReferenceChangedEventArgs> CellReferenceChanged;

        private static void OnTextChanged(DependencyObject d, DependencyPropertyChangedEventArgs e)
        {
            var control = (ExcelRefEditControl)d;

            // Update the tooltip text
            if (string.IsNullOrEmpty(e.NewValue?.ToString()))
            {
                control._currentToolTipText = "No cell selected";
            }
            else
            {
                control._currentToolTipText = e.NewValue.ToString();
            }

            // Update the tooltip directly
            if (control.txtBox != null && control.txtBox.ToolTip is ToolTip toolTip && toolTip.Content is StackPanel stackPanel)
            {
                foreach (var child in stackPanel.Children)
                {
                    if (child is TextBlock textBlock && textBlock.FontFamily?.Source == "Segoe UI")
                    {
                        textBlock.Text = control._currentToolTipText;
                        break;
                    }
                }
            }

            control.RaiseCellReferenceChanged(e.NewValue?.ToString());
        }

        public static readonly DependencyProperty ExcelAppProperty =
            DependencyProperty.Register("ExcelApp", typeof(Excel.Application), typeof(ExcelRefEditControl));

        public Excel.Application ExcelApp
        {
            get { return (Excel.Application)GetValue(ExcelAppProperty); }
            set { SetValue(ExcelAppProperty, value); }
        }

        public static readonly DependencyProperty RowAbsoluteProperty =
            DependencyProperty.Register(
            nameof(RowAbsolute),
            typeof(bool),
            typeof(ExcelRefEditControl),
            new PropertyMetadata(true)); // default: absolute row

        public bool RowAbsolute
        {
            get => (bool)GetValue(RowAbsoluteProperty);
            set => SetValue(RowAbsoluteProperty, value);
        }

        public static readonly DependencyProperty ColumnAbsoluteProperty =
            DependencyProperty.Register(
                nameof(ColumnAbsolute),
                typeof(bool),
                typeof(ExcelRefEditControl),
                new PropertyMetadata(true)); // default: absolute column

        public bool ColumnAbsolute
        {
            get => (bool)GetValue(ColumnAbsoluteProperty);
            set => SetValue(ColumnAbsoluteProperty, value);
        }

        private void RaiseCellReferenceChanged(string newRef)
        {
            CellReferenceChanged?.Invoke(this, new CellReferenceChangedEventArgs(TagName, newRef));
        }
        private void UpdateToolTipText(string text)
        {
            if (toolTipTextBlock != null)
            {
                toolTipTextBlock.Text = string.IsNullOrEmpty(text) ? "No cell selected" : text;
            }
        }
        private void BtnEdit_Click(object sender, RoutedEventArgs e)
        {
            ServiceLocator.Logger?.LogDebug($"ExcelRefEditControl.BtnEdit_Click invoked (TagName={TagName})");
            Window parentWindow = Window.GetWindow(this);
            FrameworkElement hostContainer = parentWindow == null ? FindHostContainer() : null;
            bool windowDisabled = false;
            bool hostDisabled = false;

            try
            {
                // Add null check for Excel app
                if (ServiceLocator.ExcelApp == null)
                {
                    ServiceLocator.Logger?.LogError("Excel application is not available.");
                    return;
                }

                if (parentWindow != null)
                {
                    parentWindow.IsEnabled = false;
                    windowDisabled = true;
                }
                else if (hostContainer != null)
                {
                    hostContainer.IsEnabled = false;
                    hostDisabled = true;
                }

                ServiceLocator.ExcelApp.EnableEvents = false;
                ServiceLocator.ExcelApp.DisplayAlerts = false;

                object result = ServiceLocator.ExcelApp.InputBox(
                    "Please select single cell..",
                    "Select excel cell",
                    Type: 8);

                // Handle cancellation (returns false)
                switch (result)
                {
                    case bool when !(bool)result:
                        return; // User cancelled
                    case Excel.Range rng:
                        {
                            // Check for single cell selection
                            if (rng.Cells.Count > 1)
                            {
                                // Optional: Show message to user
                                ServiceLocator.Logger?.LogError("Please select only a single cell.");
                                return;
                            }

                            string sheetName = ((Excel.Worksheet)rng.Parent).Name;
                            string cellAddress = rng.Address[RowAbsolute, ColumnAbsolute, Excel.XlReferenceStyle.xlA1, false];
                            string addr = $"'{sheetName}'!{cellAddress}";
                            this.Text = addr;
                            UpdateToolTipText(this.Text);
                            ServiceLocator.Logger?.LogDebug($"ExcelRefEditControl.BtnEdit_Click: cell reference selected: {addr}");
                            break;
                        }
                }
            }
            catch (Exception ex)
            {
                ServiceLocator.Logger?.LogException(ex, "ExcelRefEditControl.BtnEdit_Click");
            }
            finally
            {
                // Safely restore settings
                try
                {
                    if (ServiceLocator.ExcelApp != null)
                    {
                        ServiceLocator.ExcelApp.EnableEvents = true;
                        ServiceLocator.ExcelApp.DisplayAlerts = true;
                    }
                }
                catch (Exception ex)
                {
                    ServiceLocator.Logger?.LogError("Error restoring Excel settings: " + ex.Message);
                }
                finally
                {
                    if (windowDisabled && parentWindow != null)
                    {
                        parentWindow.IsEnabled = true;
                        parentWindow.Activate();
                    }
                    else if (hostDisabled && hostContainer != null)
                    {
                        hostContainer.IsEnabled = true;
                        hostContainer.Focus();
                    }
                }
            }
        }

        private void BtnClear_Click(object sender, RoutedEventArgs e)
        {
            ServiceLocator.Logger?.LogDebug($"ExcelRefEditControl.BtnClear_Click invoked (TagName={TagName})");
            this.Text = string.Empty;
            UpdateToolTipText(string.Empty);

            // Additionally clear associated combo selections when the refedit clear (eraser) is used.
            try
            {
                // Determine the binding path for our Text DP (e.g., "LedgerField.RefValue")
                var be = this.GetBindingExpression(TextProperty);
                if (be != null && this.DataContext != null)
                {
                    var path = be.ParentBinding.Path?.Path; // like "LedgerField.RefValue"
                    if (!string.IsNullOrEmpty(path))
                    {
                        var root = path.Split('.')[0]; // e.g., "LedgerField"
                        var vm = this.DataContext;
                        var vmType = vm.GetType();
                        var fieldProp = vmType.GetProperty(root);
                        if (fieldProp != null)
                        {
                            var field = fieldProp.GetValue(vm);
                            // Try to clear typical properties if this is a FieldBinding
                            var ftype = field?.GetType();
                            if (field != null && ftype != null)
                            {
                                var comboValueProp = ftype.GetProperty("ComboValue");
                                var comboTextProp = ftype.GetProperty("ComboText");
                                var refValueProp = ftype.GetProperty("RefValue");
                                try { comboValueProp?.SetValue(field, null); }
                                catch (Exception ex) { ServiceLocator.Logger?.LogWarn($"ExcelRefEditControl: failed to clear ComboValue (non-fatal): {ex.Message}"); }
                                try { comboTextProp?.SetValue(field, string.Empty); }
                                catch (Exception ex) { ServiceLocator.Logger?.LogWarn($"ExcelRefEditControl: failed to clear ComboText (non-fatal): {ex.Message}"); }
                                try { refValueProp?.SetValue(field, null); }
                                catch (Exception ex) { ServiceLocator.Logger?.LogWarn($"ExcelRefEditControl: failed to clear RefValue (non-fatal): {ex.Message}"); }

                                // If this is LedgerField, clear item selections on the VM
                                if (string.Equals(root, "LedgerField", StringComparison.OrdinalIgnoreCase))
                                {
                                    var ledgersProp = vmType.GetProperty("Ledgers");
                                    var ledgers = ledgersProp?.GetValue(vm) as System.Collections.IEnumerable;
                                    if (ledgers != null)
                                    {
                                        foreach (var it in ledgers)
                                        {
                                            var isp = it.GetType().GetProperty("IsSelected");
                                            if (isp != null)
                                            {
                                                try { isp.SetValue(it, false); }
                                                catch (Exception ex) { ServiceLocator.Logger?.LogWarn($"ExcelRefEditControl: failed to clear IsSelected on ledger (non-fatal): {ex.Message}"); }
                                            }
                                        }
                                    }
                                }

                                // Try to call a refresh on the VM (if available) to update enable states
                                var onFieldChanged = vmType.GetMethod("OnFieldChanged", new Type[] { typeof(string) });
                                if (onFieldChanged != null)
                                {
                                    try { onFieldChanged.Invoke(vm, new object[] { "" }); } catch (Exception ex) { ServiceLocator.Logger?.LogWarn($"ExcelRefEditControl: failed to invoke OnFieldChanged on VM (non-fatal): {ex.Message}"); }
                                }
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                ServiceLocator.Logger?.LogWarn($"ExcelRefEditControl: failed to refresh VM (non-fatal): {ex.Message}");
            }
        }

        private void ExcelRefEditControl_Loaded(object sender, RoutedEventArgs e)
        {
            UpdateToolTipText(this.Text);

            if (DesignerProperties.GetIsInDesignMode(this)) return;

            try
            {
                IWarningHost host = FindWarningHost(this);

                if (host != null)
                {
                    ExcelRefManager.SetupControl(this, TagName, host);
                }
            }
            catch (Exception ex)
            {
                // Never let a Loaded-event failure here crash the host window - the
                // control still works without duplicate-reference tracking.
                ServiceLocator.Logger?.LogException(ex, "ExcelRefEditControl.ExcelRefEditControl_Loaded: setup failed (non-fatal)");
            }
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

        private FrameworkElement FindHostContainer()
        {
            DependencyObject current = this;
            FrameworkElement lastFrameworkElement = null;

            while (current != null)
            {
                if (current is FrameworkElement fe)
                {
                    lastFrameworkElement = fe;
                }

                DependencyObject parent = VisualTreeHelper.GetParent(current);
                if (parent == null && current is FrameworkElement feCurrent)
                {
                    parent = feCurrent.Parent;
                }

                current = parent;
            }

            return lastFrameworkElement;
        }
    }

    public class CellReferenceChangedEventArgs : EventArgs
    {
        public CellReferenceChangedEventArgs(string tag, string newRef)
        {
            TagName = tag;
            NewReference = newRef;
        }

        public string TagName { get; set; }
        public string NewReference { get; set; }
    }
}
