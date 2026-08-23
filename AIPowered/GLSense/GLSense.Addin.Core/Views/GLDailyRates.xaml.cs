// GLDailyRates.xaml.cs in GLSense.Addin.Core
// Port of GLSense\Views\GLDailyRates.xaml.cs (FinalWorkingCode) for Group C
// (Segment/Period pickers) - GLSense_GetDailyRate formula picker (from/to currency +
// conversion type + date). Re-pointed the same way as GLGetPeriod.xaml.cs (see that
// file's header for the full mapping); additionally re-points GLSense.Extensions.
// DatePickerExtensions -> GLSense.Addin.Core.Extensions.DatePickerExtensions.
using GLSense.Addin.Core.Extensions;
using GLSense.Addin.Core.Infrastructure;
using GLSense.Addin.Core.Interfaces;
using GLSense.Addin.Core.Utilities;
using GLSense.Addin.Core.ViewModels;
using System;
using System.Collections.Generic;
using System.Windows;
using System.Windows.Input;
using Excel = Microsoft.Office.Interop.Excel;

namespace GLSense.Addin.Core.Views
{
    /// <summary>
    /// Interaction logic for GLDailyRates.xaml
    /// </summary>
    public partial class GLDailyRates : DpiAwareWindow, IWarningHost
    {
        private readonly GLDailyRatesViewModel vm;
        public GLDailyRates()
        {
            InitializeComponent();
            ServiceLocator.Logger?.LogDebug("GLDailyRates constructor invoked");

            vm = new GLDailyRatesViewModel(Dispatcher)
            {
                ExcelApp = ServiceLocator.ExcelApp,
                ShowWarningAction = (msg) => Dispatcher.Invoke(() => AppOverlayControl.ShowWarning(msg))
            };
            DataContext = vm;

            Loaded += Window_Loaded;
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

        private async void Window_Loaded(object sender, RoutedEventArgs e)
        {
            ServiceLocator.Logger?.LogDebug("GLDailyRates.Window_Loaded invoked");

            Excel.Range rng = ServiceLocator.ExcelApp.ActiveCell;
            string sheetName = ((Excel.Worksheet)rng.Parent).Name;
            string cellAddress = rng.Address[true, true, Excel.XlReferenceStyle.xlA1, false];
            string addr = $"'{sheetName}'!{cellAddress}";

            GlobalStateViewModel.Instance.ReferenceText = addr;
            List<string> FuncArgs = null;

            if (AppState.Instance.SelectedCube != null && AppState.Instance.SelectedLedger != null && (bool)rng.HasFormula && (rng.Formula.ToString().IndexOf("GLSense_GetDailyRate", StringComparison.OrdinalIgnoreCase) >= 0))
            {

                string FuncName = rng.Formula.ToString();

                FuncArgs = CommonFunctions.FormulaParameters(FuncName);
                ServiceLocator.Logger?.LogDebug($"GLDailyRates.Window_Loaded: existing GLSense_GetDailyRate formula found at {addr}, parsed {FuncArgs?.Count ?? 0} args");
            }

            dtpDate.SetupTooltip(
            title: "Currency Date",           // Appears in tooltip header
            dispatcher: this.Dispatcher,      // For UI thread safety
            dateFormat: "yyyy-MM-dd",         // Date format
            instructionText: "Click calendar icon to select/change date"  // Footer text
            );

            await vm.LoadDataAsync(FuncArgs);
        }
        public void CellSelectionWarning(string message)
        {
            try
            {
                AppOverlayControl.ShowWarning(message);
            }
            catch (Exception ex)
            {
                ServiceLocator.Logger?.LogException(ex, "GLDailyRates.CellSelectionWarning");
            }
        }
        private void BtnSubmit_Click(object sender, RoutedEventArgs e)
        {
            ServiceLocator.Logger?.LogDebug("GLDailyRates.BtnSubmit_Click invoked");
            if (string.IsNullOrWhiteSpace(CellReference.Text))
            {
                ServiceLocator.Logger?.LogDebug("GLDailyRates.BtnSubmit_Click: validation failed - cell reference is empty");
                AppOverlayControl.ShowWarning("Please select a cell reference for entering formula.");
            }
            else
            {
                try
                {
                    if (!string.IsNullOrWhiteSpace(CellReference.Text))
                    {
                        Microsoft.Office.Interop.Excel.Range rng = CommonFunctions.RemoveInDirect(CellReference.Text);
                        if (rng != null)
                        {
                            rng.NumberFormat = "0.00";

                            if (vm.WriteFormulaToCell(rng))
                            {
                                ServiceLocator.Logger?.LogDebug("GLDailyRates.BtnSubmit_Click: formula written successfully, closing window");
                                Close();
                            }
                            else
                            {
                                ServiceLocator.Logger?.LogWarn("GLDailyRates.BtnSubmit_Click: WriteFormulaToCell returned false");
                            }
                        }
                        else
                        {
                            ServiceLocator.Logger?.LogWarn($"GLDailyRates.BtnSubmit_Click: cell reference '{CellReference.Text}' did not resolve to a valid range");
                            AppOverlayControl.ShowWarning("The specified cell reference does not refer to a valid cell in the current workbook.");
                        }
                    }
                    else
                    {
                        AppOverlayControl.ShowWarning("Cell reference for get balance cannot be blank.Try providing a cell reference for generating balance formula.");
                    }
                }
                catch (Exception ex)
                {
                    ServiceLocator.Logger?.LogException(ex, "GLDailyRates.BtnSubmit_Click");
                }
            }
        }

        private void BtnClose_Click(object sender, RoutedEventArgs e)
        {
            ServiceLocator.Logger?.LogDebug("GLDailyRates.BtnClose_Click invoked - closing window");
            Close();
        }
    }
}
