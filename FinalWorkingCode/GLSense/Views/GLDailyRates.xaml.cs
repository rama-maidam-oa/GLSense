using GLSense.Extensions;
using GLSense.Helpers;
using GLSense.Interfaces;
using GLSense.Utilities;
using GLSense.ViewModels;
using System;
using System.Collections.Generic;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using Excel = Microsoft.Office.Interop.Excel;

namespace GLSense.Views
{
    /// <summary>
    /// Interaction logic for GLDailyRates.xaml
    /// </summary>
    public partial class GLDailyRates : DpiAwareWindow, IWarningHost
    {
        private readonly GLDailyRatesViewModel vm;
        public GLDailyRates()
        {
            LogUtility.LogDebug("GLDailyRates.ctor invoked");
            InitializeComponent();
            EnhancedDragDropHelper.EnableWindowDrag(this);

            vm = new GLDailyRatesViewModel(Dispatcher)
            {
                ExcelApp = AppState.Instance.ExcelApp.Application, // Pass the Excel application instance to the ViewModel
                ShowWarningAction = (msg) => Dispatcher.Invoke(() => AppOverlayControl.ShowWarning(msg))
            };
            DataContext = vm;

            Loaded += Window_Loaded;
        }
        private async void Window_Loaded(object sender, RoutedEventArgs e)
        {
            LogUtility.LogDebug("GLDailyRates.Window_Loaded invoked");

            Excel.Range rng = AppState.Instance.ExcelApp.ActiveCell;
            string sheetName = ((Excel.Worksheet)rng.Parent).Name;
            string cellAddress = rng.Address[true, true, Excel.XlReferenceStyle.xlA1, false];
            string addr = $"'{sheetName}'!{cellAddress}";

            GlobalStateViewModel.Instance.ReferenceText = addr;
            List<string> FuncArgs = null;

            if (AppState.Instance.SelectedCube != null && AppState.Instance.SelectedLedger != null && (bool)rng.HasFormula && (rng.Formula.ToString().IndexOf("GLSense_GetDailyRate", StringComparison.OrdinalIgnoreCase) >= 0))
            {
                LogUtility.LogDebug($"GLDailyRates.Window_Loaded: existing GLSense_GetDailyRate formula detected at {addr}, parsing parameters");
                string FuncName = rng.Formula.ToString();

                FuncArgs = CommonFunctions.FormulaParameters(FuncName);
            }
            else
            {
                LogUtility.LogDebug("GLDailyRates.Window_Loaded: no existing daily-rate formula detected, using defaults");
            }

            dtpDate.SetupTooltip(
            title: "Currency Date",           // Appears in tooltip header
            dispatcher: this.Dispatcher,      // For UI thread safety
            dateFormat: "yyyy-MM-dd",         // Date format
            instructionText: "Click calendar icon to select/change date"  // Footer text
            );

            LogUtility.LogDebug($"GLDailyRates.Window_Loaded: loading data - FuncArgs count={FuncArgs?.Count ?? 0}");
            await vm.LoadDataAsync(FuncArgs);
        }
        public void CellSelectionWarning(string message)
        {
            LogUtility.LogDebug($"GLDailyRates.CellSelectionWarning invoked - message={message}");
            try
            {
                AppOverlayControl.ShowWarning(message);
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex, "GLDailyRates.CellSelectionWarning");
            }
        }
        private void BtnSubmit_Click(object sender, RoutedEventArgs e)
        {
            LogUtility.LogDebug($"GLDailyRates.BtnSubmit_Click invoked - CellReference={CellReference.Text}");
            if (string.IsNullOrWhiteSpace(CellReference.Text))
            {
                LogUtility.LogDebug("GLDailyRates.BtnSubmit_Click: validation failed - cell reference is blank");
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
                                LogUtility.LogDebug("GLDailyRates.BtnSubmit_Click: formula written successfully, closing window");
                                Close();
                            }
                            else
                            {
                                LogUtility.LogDebug("GLDailyRates.BtnSubmit_Click: WriteFormulaToCell returned false, window not closed");
                            }
                        }
                        else
                        {
                            LogUtility.LogDebug($"GLDailyRates.BtnSubmit_Click: validation failed - cell reference '{CellReference.Text}' does not resolve to a valid range");
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
                    LogUtility.LogException(ex, "GLDailyRates.BtnSubmit_Click");
                }
            }
        }

        private void BtnClose_Click(object sender, RoutedEventArgs e)
        {
            LogUtility.LogDebug("GLDailyRates.BtnClose_Click invoked");
            Close();
        }
    }
}

