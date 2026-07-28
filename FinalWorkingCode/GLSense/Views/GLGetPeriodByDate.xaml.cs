using GLSense.Extensions;
using GLSense.Helpers;
using GLSense.Interfaces;
using GLSense.Repositories;
using GLSense.Utilities;
using GLSense.ViewModels;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using Excel = Microsoft.Office.Interop.Excel;

namespace GLSense.Views
{
    /// <summary>
    /// Interaction logic for GLGetPeriodByDate.xaml
    /// </summary>
    public partial class GLGetPeriodByDate : DpiAwareWindow, IWarningHost
    {
        private sealed class LedgerInfo
        {
            public List<string> FuncArgs { get; set; }
            public List<string> FuncValues { get; set; }
            public string LedgerName { get; set; }
            public dynamic LedgerRecord { get; set; }
        }
        private readonly GLPeriodByDateModel vm;
        public GLGetPeriodByDate()
        {
            LogUtility.LogDebug("GLGetPeriodByDate.ctor invoked");
            InitializeComponent();
            EnhancedDragDropHelper.EnableWindowDrag(this);

            vm = new GLPeriodByDateModel(Dispatcher)
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
        private async void Window_Loaded(object sender, RoutedEventArgs e)
        {
            LogUtility.LogDebug("GLGetPeriodByDate.Window_Loaded invoked");
            var (Address, Range) = GetActiveCellInfo();
            GlobalStateViewModel.Instance.ReferenceText = Address;
            LogUtility.LogDebug($"GLGetPeriodByDate.Window_Loaded: active cell reference={Address}");

            if (!HasValidCubeAndLedger())
            {
                LogUtility.LogDebug("GLGetPeriodByDate.Window_Loaded: validation failed - no cube/ledger selected, aborting load");
                return;
            }

            var ledgerInfo = await ProcessYearFormulaAsync(Range);
            string ledgerName = string.IsNullOrWhiteSpace(ledgerInfo.LedgerName)
                ? AppState.Instance.SelectedLedger.LedgerName
                : ledgerInfo.LedgerName;

            dtpDate.SetupTooltip(
                title: "Period Date",           // Appears in tooltip header
                dispatcher: this.Dispatcher,      // For UI thread safety
                dateFormat: "yyyy-MM-dd",         // Date format
                instructionText: "Click calendar icon to select/change date"  // Footer text
                );

            await vm.LoadDataAsync(ledgerInfo.FuncArgs, ledgerInfo.FuncValues);
            await Dispatcher.InvokeAsync(() => cmbLedgers.Text = ledgerName);
            LogUtility.LogDebug($"GLGetPeriodByDate.Window_Loaded: completed - ledgerName={ledgerName}");
        }
        private static (string Address, Excel.Range Range) GetActiveCellInfo()
        {
            LogUtility.LogDebug("GLGetPeriodByDate.GetActiveCellInfo invoked");
            var rng = AppState.Instance.ExcelApp.ActiveCell;
            string sheetName = ((Excel.Worksheet)rng.Parent).Name;
            string cellAddress = rng.Address[true, true, Excel.XlReferenceStyle.xlA1, false];
            return ($"'{sheetName}'!{cellAddress}", rng);
        }

        private static bool HasValidCubeAndLedger()
        {
            return AppState.Instance.SelectedCube != null && AppState.Instance.SelectedLedger != null;
        }

        private async Task<LedgerInfo> ProcessYearFormulaAsync(Excel.Range rng)
        {
            LogUtility.LogDebug("GLGetPeriodByDate.ProcessYearFormulaAsync invoked");
            if (!IsYearFormula(rng))
            {
                LogUtility.LogDebug("GLGetPeriodByDate.ProcessYearFormulaAsync: cell formula is not a GetPeriodByDate formula, using defaults");
                return new LedgerInfo();
            }

            var ledgerInfo = ValidateLedgerFromFormula(rng);
            if (ledgerInfo.LedgerRecord == null)
            {
                LogUtility.LogWarn($"GLGetPeriodByDate.ProcessYearFormulaAsync: ledger '{ledgerInfo.LedgerName}' from formula not found in selected cube");
                await ShowLedgerNotFoundWarningAsync(ledgerInfo.LedgerName);
                return new LedgerInfo();
            }

            await LoadLedgerDataIfNeededAsync(ledgerInfo.LedgerRecord);
            LogUtility.LogDebug($"GLGetPeriodByDate.ProcessYearFormulaAsync: resolved ledger={ledgerInfo.LedgerName}");
            return ledgerInfo;
        }

        private static bool IsYearFormula(Excel.Range rng)
        {
            if (!(bool)rng.HasFormula) return false;
            return rng.Formula.ToString().IndexOf("GLSense_GetPeriodByDate", StringComparison.OrdinalIgnoreCase) >= 0;
        }

        private LedgerInfo ValidateLedgerFromFormula(Excel.Range rng)
        {
            LogUtility.LogDebug("GLGetPeriodByDate.ValidateLedgerFromFormula invoked");
            string formula = rng.Formula.ToString();
            var funcArgs = CommonFunctions.FormulaParameters(formula);
            var funcValues = CommonFunctions.FormulaValues(formula);

            string ledgerName = AppState.Instance.SelectedLedger.LedgerName;

            // If less than 3 arguments ? ledger is missing ? insert at index 0
            if (funcValues != null && funcValues.Count < 3)
            {
                funcArgs?.Insert(1, ledgerName); // or appropriate argument name
                funcValues?.Insert(1, ledgerName);
            }
            // If already 3 or more ? take ledger from formula
            else if (funcValues != null && funcValues.Count >= 3)
            {
                ledgerName = funcValues[1].Replace("\"", "");
            }

            var ledgerRecord = AppState.Instance.SelectedCube.Ledgers
                .FirstOrDefault(x => x.LedgerName == ledgerName);
            LogUtility.LogDebug($"GLGetPeriodByDate.ValidateLedgerFromFormula: ledgerName={ledgerName}, found={(ledgerRecord != null)}");

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
            LogUtility.LogDebug("GLGetPeriodByDate.LoadLedgerDataIfNeededAsync invoked");
            var segmentCount = DataRepository.GetTableItemsCount(
                AppState.Instance.SelectedCube.CubeId,
                ledgerRecord.LedgerId,
                "SEGMENTS");
            LogUtility.LogDebug($"GLGetPeriodByDate.LoadLedgerDataIfNeededAsync: segmentCount={segmentCount}");

            if (segmentCount == 0)
            {
                LogUtility.LogDebug("GLGetPeriodByDate.LoadLedgerDataIfNeededAsync: no segments cached, loading ledger segments");
                await LoadLedgerSegmentsWithProgressAsync(ledgerRecord);
            }
        }

        private async Task LoadLedgerSegmentsWithProgressAsync(dynamic ledgerRecord)
        {
            LogUtility.LogDebug("GLGetPeriodByDate.LoadLedgerSegmentsWithProgressAsync invoked");
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

                LogUtility.LogDebug("GLGetPeriodByDate.LoadLedgerSegmentsWithProgressAsync: calling CommonFunctions.FillResponsibilitiesAsync");
                await CommonFunctions.FillResponsibilitiesAsync(
                    ledgerRecord.LedgerId,
                    AppState.Instance.SelectedCube.CubeId,
                    token);
                LogUtility.LogDebug("GLGetPeriodByDate.LoadLedgerSegmentsWithProgressAsync: FillResponsibilitiesAsync completed successfully");
            }
            catch (OperationCanceledException)
            {
                LogUtility.LogWarn("Loading ledger segments operation was cancelled.");
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex, "GLGetPeriodByDate.LoadLedgerSegmentsWithProgressAsync");
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
                    LogUtility.LogWarn($"GLGetPeriodByDate.LoadLedgerSegmentsWithProgressAsync: exception disposing CancellationHelper (ignored): {ex.Message}");
                }
                await Dispatcher.InvokeAsync(async () => await AppOverlayControl.HideBusyAsync());
                ctsHelper.Dispose();
            }
        }

        private async Task ShowLedgerNotFoundWarningAsync(string ledgerName)
        {
            LogUtility.LogDebug($"GLGetPeriodByDate.ShowLedgerNotFoundWarningAsync invoked - ledgerName={ledgerName}");
            string message = $"Ledger \"{ledgerName}\" in the formula does not exist in the selected cube!" +
                            Environment.NewLine + "Setting default values.";
            await AppOverlayControl.ShowWarningAsync(message);
        }
        public void CellSelectionWarning(string message)
        {
            LogUtility.LogDebug($"GLGetPeriodByDate.CellSelectionWarning invoked - message={message}");
            try
            {
                AppOverlayControl.ShowWarning(message);
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex, "GLGetPeriodByDate.CellSelectionWarning");
            }
        }
        private void BtnClose_Click(object sender, RoutedEventArgs e)
        {
            LogUtility.LogDebug("GLGetPeriodByDate.BtnClose_Click invoked");
            this.Close();
        }
        private void BtnSubmit_Click(object sender, RoutedEventArgs e)
        {
            LogUtility.LogDebug("GLGetPeriodByDate.BtnSubmit_Click invoked");
            if (string.IsNullOrWhiteSpace(CellReference.Text))
            {
                LogUtility.LogDebug("GLGetPeriodByDate.BtnSubmit_Click: validation failed - cell reference is blank");
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
                            rng.NumberFormat = AppConstants.General;

                            if (vm.WriteFormulaToCell(rng))
                            {
                                LogUtility.LogDebug("GLGetPeriodByDate.BtnSubmit_Click: formula written successfully, closing window");
                                Close();
                            }
                            else
                            {
                                LogUtility.LogDebug("GLGetPeriodByDate.BtnSubmit_Click: WriteFormulaToCell returned false, formula not written");
                            }
                        }
                        else
                        {
                            LogUtility.LogDebug("GLGetPeriodByDate.BtnSubmit_Click: validation failed - cell reference does not resolve to a valid range");
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
                    LogUtility.LogException(ex, "GLGetPeriodByDate.BtnSubmit_Click");
                }
            }
        }
    }
}

