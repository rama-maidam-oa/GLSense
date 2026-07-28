using GLSense.Helpers;
using GLSense.Interfaces;
using GLSense.Models;
using GLSense.Repositories;
using GLSense.Utilities;
using GLSense.ViewModels;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using Excel = Microsoft.Office.Interop.Excel;

namespace GLSense.Views
{
    /// <summary>
    /// Interaction logic for GLGetPeriod.xaml
    /// </summary>
    public partial class GLGetPeriod : DpiAwareWindow, IWarningHost
    {
        private sealed class LedgerInfo
        {
            public List<string> FuncArgs { get; set; }
            public List<string> FuncValues { get; set; }
            public string LedgerName { get; set; }
            public dynamic LedgerRecord { get; set; }
        }
        private readonly GLGetPeriodModel vm;
        public GLGetPeriod()
        {
            LogUtility.LogDebug("GLGetPeriod.ctor invoked");
            InitializeComponent();
            EnhancedDragDropHelper.EnableWindowDrag(this);

            vm = new GLGetPeriodModel(Dispatcher)
            {
                ExcelApp = AppState.Instance.ExcelApp.Application, // Pass the Excel application instance to the ViewModel
                ShowWarningAction = (msg) => Dispatcher.Invoke(() => AppOverlayControl.ShowWarning(msg)),
                ShowWarningAsyncAction = async (msg) => await Dispatcher.InvokeAsync(async () => await AppOverlayControl.ShowWarningAsync(msg)),
                ShowBusyAction = async (txt, cancel) =>
                        await Dispatcher.InvokeAsync(async () =>
                            await AppOverlayControl.ShowBusyasynTask(txt, cancel)),
                HideBusyAsyncAction = async () => await Dispatcher.InvokeAsync(async () => await AppOverlayControl.HideBusyAsync())
            };
            DataContext = vm;

            Loaded += Window_Loaded;
        }
        public void CellSelectionWarning(string message)
        {
            LogUtility.LogDebug($"GLGetPeriod.CellSelectionWarning invoked - message={message}");
            try
            {
                AppOverlayControl.ShowWarning(message);
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex, "GLGetPeriod.CellSelectionWarning");
            }
        }
        private async void Window_Loaded(object sender, RoutedEventArgs e)
        {
            LogUtility.LogDebug("GLGetPeriod.Window_Loaded invoked");
            var (Address, Range) = GetActiveCellInfo();
            GlobalStateViewModel.Instance.ReferenceText = Address;

            if (!HasValidCubeAndLedger())
            {
                LogUtility.LogDebug("GLGetPeriod.Window_Loaded: no valid cube/ledger selected, aborting");
                return;
            }

            var ledgerInfo = await ProcessPeriodFormulaAsync(Range);
            string ledgerName = string.IsNullOrWhiteSpace(ledgerInfo.LedgerName)
                ? AppState.Instance.SelectedLedger.LedgerName
                : ledgerInfo.LedgerName;

            await vm.LoadDataAsync(ledgerInfo.FuncArgs, ledgerInfo.FuncValues);
            await Dispatcher.InvokeAsync(() => cmbLedgers.Text = ledgerName);
        }
        private static (string Address, Excel.Range Range) GetActiveCellInfo()
        {
            var rng = AppState.Instance.ExcelApp.ActiveCell;
            string sheetName = ((Excel.Worksheet)rng.Parent).Name;
            string cellAddress = rng.Address[true, true, Excel.XlReferenceStyle.xlA1, false];
            return ($"'{sheetName}'!{cellAddress}", rng);
        }

        private static bool HasValidCubeAndLedger()
        {
            return AppState.Instance.SelectedCube != null && AppState.Instance.SelectedLedger != null;
        }

        private async Task<LedgerInfo> ProcessPeriodFormulaAsync(Excel.Range rng)
        {
            if (!IsPeriodFormula(rng))
            {
                LogUtility.LogDebug("GLGetPeriod.ProcessPeriodFormulaAsync: no existing GLSense_GetPeriod formula detected, using defaults");
                return new LedgerInfo();
            }

            var ledgerInfo = ValidateLedgerFromFormula(rng);
            if (ledgerInfo.LedgerRecord == null)
            {
                LogUtility.LogDebug($"GLGetPeriod.ProcessPeriodFormulaAsync: ledger '{ledgerInfo.LedgerName}' from formula not found in selected cube");
                await ShowLedgerNotFoundWarningAsync(ledgerInfo.LedgerName);
                return new LedgerInfo();
            }

            await LoadLedgerDataIfNeededAsync(ledgerInfo.LedgerRecord);
            return ledgerInfo;
        }

        private static bool IsPeriodFormula(Excel.Range rng)
        {
            if (!(bool)rng.HasFormula) return false;
            return rng.Formula.ToString().IndexOf("GLSense_GetPeriod", StringComparison.OrdinalIgnoreCase) >= 0;
        }

        private LedgerInfo ValidateLedgerFromFormula(Excel.Range rng)
        {
            string formula = rng.Formula.ToString();
            var funcArgs = CommonFunctions.FormulaParameters(formula);
            var funcValues = CommonFunctions.FormulaValues(formula);

            string ledgerName = AppState.Instance.SelectedLedger.LedgerName;

            // If less than 3 arguments ? ledger is missing ? insert at index 2 or at the end (if funcArgs is shorter than funcValues)
            if (funcValues != null && funcValues.Count < 3)
            {
                funcArgs?.Add(ledgerName); // or appropriate argument name
                funcValues?.Add(ledgerName);
            }
            // If already 3 or more ? take ledger from formula
            else if (funcValues != null && funcValues.Count >= 3)
            {
                ledgerName = funcValues[2].Replace("\"", "");
            }

            var ledgerRecord = AppState.Instance.SelectedCube.Ledgers
                .FirstOrDefault(x => x.LedgerName == ledgerName);

            return new LedgerInfo
            {
                FuncArgs = funcArgs,
                FuncValues = funcValues,
                LedgerName = ledgerName,
                LedgerRecord = ledgerRecord
            };
        }

        private async Task LoadLedgerDataIfNeededAsync(dynamic ledgerRecord)
        {
            var segmentCount = DataRepository.GetTableItemsCount(
                AppState.Instance.SelectedCube.CubeId,
                ledgerRecord.LedgerId,
                "SEGMENTS");

            if (segmentCount == 0)
            {
                LogUtility.LogDebug("GLGetPeriod.LoadLedgerDataIfNeededAsync: segments not loaded for ledger, fetching responsibilities");
                await LoadLedgerSegmentsWithProgressAsync(ledgerRecord);
            }
        }

        private async Task LoadLedgerSegmentsWithProgressAsync(dynamic ledgerRecord)
        {
            LogUtility.LogDebug("GLGetPeriod.LoadLedgerSegmentsWithProgressAsync invoked");
            var ctsHelper = new CancellationHelper();
            var token = ctsHelper.GetToken();

            try
            {
                await Dispatcher.InvokeAsync(() =>
                {
                    AppOverlayControl.ShowBusyasyn(
                        "Please wait while we fetch the ledger data...  (click Cancel to stop)",
                        async () => {
                            if (!ctsHelper.IsCancellationRequested)
                            {
                                ctsHelper.Cancel();
                                LogUtility.LogWarn("Loading cancelled by user");
                            }
                            await Task.CompletedTask;
                        });
                    return Task.CompletedTask;
                });

                await CommonFunctions.FillResponsibilitiesAsync(
                    ledgerRecord.LedgerId,
                    AppState.Instance.SelectedCube.CubeId,
                    token);
            }
            catch (OperationCanceledException)
            {
                LogUtility.LogWarn("Loading ledger segments operation was cancelled.");
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex, "GLGetPeriod.LoadLedgerSegmentsWithProgressAsync");
            }
            finally
            {
                try
                {
                    if (ctsHelper != null && !ctsHelper.IsCancellationRequested)
                        ctsHelper.Cancel();

                    ctsHelper?.Dispose();
                }
                catch (Exception ex)
                {
                    // Swallow dispose exceptions (Excel COM weirdness) but still log for diagnostics.
                    LogUtility.LogWarn($"GLGetPeriod.LoadLedgerSegmentsWithProgressAsync: exception disposing CancellationHelper (ignored): {ex.Message}");
                }
                await Dispatcher.InvokeAsync(async () => await AppOverlayControl.HideBusyAsync());
                ctsHelper.Dispose();
            }
        }

        private async Task ShowLedgerNotFoundWarningAsync(string ledgerName)
        {
            LogUtility.LogWarn($"GLGetPeriod.ShowLedgerNotFoundWarningAsync: ledger '{ledgerName}' from formula not found in selected cube, setting default values");
            string message = $"Ledger \"{ledgerName}\" in the formula does not exist in the selected cube!" +
                            Environment.NewLine + "Setting default values.";
            await AppOverlayControl.ShowWarningAsync(message);
        }
        private void BtnSubmit_Click(object sender, RoutedEventArgs e)
        {
            LogUtility.LogDebug($"GLGetPeriod.BtnSubmit_Click invoked - CellReference={CellReference.Text}");
            if (string.IsNullOrWhiteSpace(CellReference.Text))
            {
                LogUtility.LogDebug("GLGetPeriod.BtnSubmit_Click: validation failed - cell reference is blank");
                AppOverlayControl.ShowWarning("Please select a cell reference for entering formula.");
            }
            else
            {
                try
                {
                    if (!string.IsNullOrWhiteSpace(CellReference.Text))
                    {
                        Excel.Range rng = CommonFunctions.RemoveInDirect(CellReference.Text);
                        if (rng != null)
                        {
                            rng.NumberFormat = AppConstants.General;

                            if (vm.WriteFormulaToCell(rng))
                            {
                                LogUtility.LogDebug("GLGetPeriod.BtnSubmit_Click: formula written successfully, closing window");
                                Close();
                            }
                            else
                            {
                                LogUtility.LogDebug("GLGetPeriod.BtnSubmit_Click: WriteFormulaToCell returned false, window not closed");
                            }

                        }
                        else
                        {
                            LogUtility.LogDebug($"GLGetPeriod.BtnSubmit_Click: validation failed - cell reference '{CellReference.Text}' does not resolve to a valid range");
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
                    LogUtility.LogException(ex, "GLGetPeriod.BtnSubmit_Click");
                }
            }
        }

        private void BtnClose_Click(object sender, RoutedEventArgs e)
        {
            LogUtility.LogDebug("GLGetPeriod.BtnClose_Click invoked");
            Close();
        }
    }
}

