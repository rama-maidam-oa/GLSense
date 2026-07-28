using GLSense.Helpers;
using GLSense.Interfaces;
using GLSense.Utilities;
using System;
using System.ComponentModel;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using Excel = Microsoft.Office.Interop.Excel;

namespace GLSense.Views
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
            Loaded += ExcelRefEditControl_Loaded;
            // Not logged: a single window can host many of these controls (e.g. one
            // per ledger/activity/period field), so this fired once per instance with
            // no data - just noise duplicated across every field on every window.
            // ExcelRefEditControl_Loaded's own log (with TagName) already identifies
            // which control did what.
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
            LogUtility.LogDebug($"ExcelRefEditControl.OnTextChanged: TagName={control.TagName}, NewValue={e.NewValue}");

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
            LogUtility.LogDebug($"ExcelRefEditControl.BtnEdit_Click invoked - TagName={TagName}");
            Window parentWindow = Window.GetWindow(this);
            FrameworkElement hostContainer = parentWindow == null ? FindHostContainer() : null;
            bool windowDisabled = false;
            bool hostDisabled = false;

            try
            {
                // Add null check for Excel app
                if (AppState.Instance.ExcelApp == null)
                {
                    LogUtility.LogError("ExcelRefEditControl.BtnEdit_Click: Excel application is not available.");
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

                AppState.Instance.ExcelApp.EnableEvents = false;
                AppState.Instance.ExcelApp.DisplayAlerts = false;

                object result = AppState.Instance.ExcelApp.InputBox(
                    "Please select single cell..",
                    "Select excel cell",
                    Type: 8);

                // Handle cancellation (returns false)
                switch (result)
                {
                    case bool when !(bool)result:
                        LogUtility.LogDebug("ExcelRefEditControl.BtnEdit_Click: user cancelled cell selection");
                        return; // User cancelled
                    case Excel.Range rng:
                        {
                            // Check for single cell selection
                            if (rng.Cells.Count > 1)
                            {
                                // Optional: Show message to user
                                LogUtility.LogError("ExcelRefEditControl.BtnEdit_Click: Please select only a single cell.");
                                return;
                            }

                            string sheetName = ((Excel.Worksheet)rng.Parent).Name;
                            string cellAddress = rng.Address[RowAbsolute, ColumnAbsolute, Excel.XlReferenceStyle.xlA1, false];
                            string addr = $"'{sheetName}'!{cellAddress}";
                            this.Text = addr;
                            UpdateToolTipText(this.Text);
                            LogUtility.LogDebug($"ExcelRefEditControl.BtnEdit_Click: cell selected - {addr}");
                            break;
                        }
                }
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex, "ExcelRefEditControl.BtnEdit_Click");
            }
            finally
            {
                // Safely restore settings
                try
                {
                    if (AppState.Instance.ExcelApp != null)
                    {
                        AppState.Instance.ExcelApp.EnableEvents = true;
                        AppState.Instance.ExcelApp.DisplayAlerts = true;
                    }
                }
                catch (Exception ex)
                {
                    LogUtility.LogError("ExcelRefEditControl.BtnEdit_Click: Error restoring Excel settings: " + ex.Message);
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
            LogUtility.LogDebug($"ExcelRefEditControl.BtnClear_Click invoked - TagName={TagName}");
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
                                catch (Exception ex) { LogUtility.LogWarn($"ExcelRefEditControl: failed to clear ComboValue (non-fatal): {ex.Message}"); }
                                try { comboTextProp?.SetValue(field, string.Empty); }
                                catch (Exception ex) { LogUtility.LogWarn($"ExcelRefEditControl: failed to clear ComboText (non-fatal): {ex.Message}"); }
                                try { refValueProp?.SetValue(field, null); }
                                catch (Exception ex) { LogUtility.LogWarn($"ExcelRefEditControl: failed to clear RefValue (non-fatal): {ex.Message}"); }

                                // If this is LedgerField or EncumbranceField, clear item selections on the VM
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
                                                catch (Exception ex) { LogUtility.LogWarn($"ExcelRefEditControl: failed to clear IsSelected on ledger (non-fatal): {ex.Message}"); }
                                            }
                                        }
                                    }
                                }
                                else if (string.Equals(root, "EncumbranceField", StringComparison.OrdinalIgnoreCase) || string.Equals(root, "Encumbrances", StringComparison.OrdinalIgnoreCase))
                                {
                                    var encProp = vmType.GetProperty("Encumbrances");
                                    var encs = encProp?.GetValue(vm) as System.Collections.IEnumerable;
                                    if (encs != null)
                                    {
                                        foreach (var it in encs)
                                        {
                                            var isp = it.GetType().GetProperty("IsSelected");
                                            if (isp != null)
                                            {
                                                try { isp.SetValue(it, false); }
                                                catch (Exception ex) { LogUtility.LogWarn($"ExcelRefEditControl: failed to clear IsSelected on encumbrance (non-fatal): {ex.Message}"); }
                                            }
                                        }
                                    }
                                }

                                // Try to call a refresh on the VM (if available) to update enable states
                                var onFieldChanged = vmType.GetMethod("OnFieldChanged", new Type[] { typeof(string) });
                                if (onFieldChanged != null)
                                {
                                    try { onFieldChanged.Invoke(vm, new object[] { "" }); } catch (Exception ex) { LogUtility.LogWarn($"ExcelRefEditControl: failed to invoke OnFieldChanged on VM (non-fatal): {ex.Message}"); }
                                }
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                LogUtility.LogWarn($"ExcelRefEditControl: failed to refresh VM (non-fatal): {ex.Message}");
            }
        }

        private void ExcelRefEditControl_Loaded(object sender, RoutedEventArgs e)
        {
            LogUtility.LogDebug($"ExcelRefEditControl.ExcelRefEditControl_Loaded invoked - TagName={TagName}");
            UpdateToolTipText(this.Text);

            if (DesignerProperties.GetIsInDesignMode(this))
            {
                LogUtility.LogDebug("ExcelRefEditControl.ExcelRefEditControl_Loaded: in designer mode, skipping warning host setup");
                return;
            }

            try
            {
                IWarningHost host = FindWarningHost(this);

                if (host != null)
                {
                    LogUtility.LogDebug("ExcelRefEditControl.ExcelRefEditControl_Loaded: warning host found, setting up ExcelRefManager control");
                    ExcelRefManager.SetupControl(this, TagName, host);
                }
                else
                {
                    LogUtility.LogDebug("ExcelRefEditControl.ExcelRefEditControl_Loaded: no warning host found");
                }
            }
            catch (Exception ex)
            {
                // Never let a Loaded-event failure here (e.g. a transient Excel COM/RCW
                // hiccup while wiring up the ref-edit tracker) crash the host window -
                // the control still works without duplicate-reference tracking.
                LogUtility.LogException(ex, "ExcelRefEditControl.ExcelRefEditControl_Loaded: setup failed (non-fatal)");
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
}
