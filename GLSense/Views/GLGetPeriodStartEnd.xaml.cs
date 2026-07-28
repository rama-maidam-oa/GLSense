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
    /// Interaction logic for GLGetPeriodStartEnd.xaml
    /// </summary>
    public partial class GLGetPeriodStartEnd : DpiAwareWindow, IWarningHost
    {
        private sealed class LedgerInfo
        {
            public List<string> FuncArgs { get; set; }
            public List<string> FuncValues { get; set; }
            public string LedgerName { get; set; }
            public dynamic LedgerRecord { get; set; }
        }
        private readonly GLPeriodDetails vm;
        public GLGetPeriodStartEnd(string FormulaName)
        {
            LogUtility.LogDebug($"GLGetPeriodStartEnd.ctor invoked - FormulaName={FormulaName}");
            InitializeComponent();
            EnhancedDragDropHelper.EnableWindowDrag(this);

            switch (FormulaName)
            {
                case "START":
                    this.TxtGLFormulaType.Text = "Get Period Start";
                    break;
                case "END":
                    this.TxtGLFormulaType.Text = "Get Period End";
                    break;
            }

            vm = new GLPeriodDetails(Dispatcher, FormulaName)
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
            LogUtility.LogDebug("GLGetPeriodStartEnd.Window_Loaded invoked");

            var (Address, Range) = GetActiveCellInfo();
            GlobalStateViewModel.Instance.ReferenceText = Address;
            LogUtility.LogDebug($"GLGetPeriodStartEnd.Window_Loaded: active cell reference={Address}");

            if (!HasValidCubeAndLedger())
            {
                LogUtility.LogDebug("GLGetPeriodStartEnd.Window_Loaded: validation failed - no cube/ledger selected, aborting load");
                return;
            }

            var ledgerInfo = await ProcessCellFormulaAsync(Range);

            string ledgerName = GetLedgerName(ledgerInfo);
            await vm.LoadDataAsync(ledgerInfo.FuncArgs, ledgerInfo.FuncValues);

            await SetLedgerDropdownAsync(ledgerName);
            LogUtility.LogDebug($"GLGetPeriodStartEnd.Window_Loaded: completed - ledgerName={ledgerName}");
        }
        private static (string Address, Excel.Range Range) GetActiveCellInfo()
        {
            LogUtility.LogDebug("GLGetPeriodStartEnd.GetActiveCellInfo invoked");
            var rng = AppState.Instance.ExcelApp.ActiveCell;
            string sheetName = ((Excel.Worksheet)rng.Parent).Name;
            string cellAddress = rng.Address[true, true, Excel.XlReferenceStyle.xlA1, false];
            return ($"'{sheetName}'!{cellAddress}", rng);
        }

        private static bool HasValidCubeAndLedger()
        {
            return AppState.Instance.SelectedCube != null && AppState.Instance.SelectedLedger != null;
        }
        private async Task<LedgerInfo> ProcessCellFormulaAsync(Excel.Range rng)
        {
            LogUtility.LogDebug("GLGetPeriodStartEnd.ProcessCellFormulaAsync invoked");
            if (!IsPeriodFormula(rng))
            {
                LogUtility.LogDebug("GLGetPeriodStartEnd.ProcessCellFormulaAsync: cell formula is not a GetPeriodStart/End formula, using defaults");
                return new LedgerInfo();
            }

            var ledgerInfo = ValidateLedgerFromFormula(rng);
            if (ledgerInfo.LedgerRecord == null)
            {
                LogUtility.LogWarn($"GLGetPeriodStartEnd.ProcessCellFormulaAsync: ledger '{ledgerInfo.LedgerName}' from formula not found in selected cube");
                await ShowLedgerNotFoundWarningAsync(ledgerInfo.LedgerName);
            }

            await LoadLedgerDataIfNeededAsync(ledgerInfo.LedgerRecord);
            LogUtility.LogDebug($"GLGetPeriodStartEnd.ProcessCellFormulaAsync: resolved ledger={ledgerInfo.LedgerName}");
            return new LedgerInfo
            {
                FuncArgs = ledgerInfo.FuncArgs,
                FuncValues = ledgerInfo.FuncValues,
                LedgerName = ledgerInfo.LedgerName,
                LedgerRecord = ledgerInfo.LedgerRecord
            };
        }

        private static bool IsPeriodFormula(Excel.Range rng)
        {
            if (!(bool)rng.HasFormula) return false;
            string formula = rng.Formula.ToString();
            return formula.IndexOf("GLSense_GetPeriodStart", StringComparison.OrdinalIgnoreCase) >= 0 ||
                   formula.IndexOf("GLSense_GetPeriodEnd", StringComparison.OrdinalIgnoreCase) >= 0;
        }
        private LedgerInfo ValidateLedgerFromFormula(Excel.Range rng)
        {
            LogUtility.LogDebug("GLGetPeriodStartEnd.ValidateLedgerFromFormula invoked");
            string formula = rng.Formula.ToString();
            var funcArgs = CommonFunctions.FormulaParameters(formula);
            var funcValues = CommonFunctions.FormulaValues(formula);
            string ledgerName = AppState.Instance.SelectedLedger.LedgerName;

            // If less than 3 arguments ? ledger is missing ? insert at index 2 or at the end (if funcArgs is shorter than funcValues)
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
            LogUtility.LogDebug($"GLGetPeriodStartEnd.ValidateLedgerFromFormula: ledgerName={ledgerName}, found={(ledgerRecord != null)}");

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
            LogUtility.LogDebug("GLGetPeriodStartEnd.LoadLedgerDataIfNeededAsync invoked");
            if (ledgerRecord == null)
            {
                LogUtility.LogDebug("GLGetPeriodStartEnd.LoadLedgerDataIfNeededAsync: ledgerRecord is null, nothing to load");
                return;
            }

            var segmentCount = DataRepository.GetTableItemsCount(
                AppState.Instance.SelectedCube.CubeId,
                ledgerRecord.LedgerId,
                "SEGMENTS");
            LogUtility.LogDebug($"GLGetPeriodStartEnd.LoadLedgerDataIfNeededAsync: segmentCount={segmentCount}");

            if (segmentCount == 0)
            {
                LogUtility.LogDebug("GLGetPeriodStartEnd.LoadLedgerDataIfNeededAsync: no segments cached, loading ledger segments");
                await LoadLedgerSegmentsWithProgressAsync(ledgerRecord);
            }
        }

        private async Task LoadLedgerSegmentsWithProgressAsync(dynamic ledgerRecord)
        {
            LogUtility.LogDebug("GLGetPeriodStartEnd.LoadLedgerSegmentsWithProgressAsync invoked");
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

                LogUtility.LogDebug("GLGetPeriodStartEnd.LoadLedgerSegmentsWithProgressAsync: calling CommonFunctions.FillResponsibilitiesAsync");
                await CommonFunctions.FillResponsibilitiesAsync(
                    ledgerRecord.LedgerId,
                    AppState.Instance.SelectedCube.CubeId,
                    token);
                LogUtility.LogDebug("GLGetPeriodStartEnd.LoadLedgerSegmentsWithProgressAsync: FillResponsibilitiesAsync completed successfully");
            }
            catch (OperationCanceledException)
            {
                LogUtility.LogWarn("Loading ledger segments operation was cancelled.");
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex, "GLGetPeriodStartEnd.LoadLedgerSegmentsWithProgressAsync");
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
                    LogUtility.LogWarn($"GLGetPeriodStartEnd.LoadLedgerSegmentsWithProgressAsync: exception disposing CancellationHelper (ignored): {ex.Message}");
                }
                await Dispatcher.InvokeAsync(async () => await AppOverlayControl.HideBusyAsync());
                ctsHelper.Dispose();
            }
        }

        private async Task ShowLedgerNotFoundWarningAsync(string ledgerName)
        {
            LogUtility.LogDebug($"GLGetPeriodStartEnd.ShowLedgerNotFoundWarningAsync invoked - ledgerName={ledgerName}");
            string message = $"Ledger \"{ledgerName}\" in the formula does not exist in the selected cube!" +
                            Environment.NewLine + "Setting default values.";
            await AppOverlayControl.ShowWarningAsync(message);
        }

        private static string GetLedgerName(LedgerInfo ledgerInfo)
        {
            return string.IsNullOrWhiteSpace(ledgerInfo.LedgerName)
                ? AppState.Instance.SelectedLedger.LedgerName
                : ledgerInfo.LedgerName;
        }

        private async Task SetLedgerDropdownAsync(string ledgerName)
        {
            await Dispatcher.InvokeAsync(() => cmbLedgers.Text = ledgerName);
        }
        public void CellSelectionWarning(string message)
        {
            LogUtility.LogDebug($"GLGetPeriodStartEnd.CellSelectionWarning invoked - message={message}");
            try
            {
                AppOverlayControl.ShowWarning(message);
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex, "GLGetPeriodStartEnd.CellSelectionWarning");
            }
        }
        private void BtnSubmit_Click(object sender, RoutedEventArgs e)
        {
            LogUtility.LogDebug("GLGetPeriodStartEnd.BtnSubmit_Click invoked");
            if (string.IsNullOrWhiteSpace(CellReference.Text))
            {
                LogUtility.LogDebug("GLGetPeriodStartEnd.BtnSubmit_Click: validation failed - cell reference is blank");
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
                                LogUtility.LogDebug("GLGetPeriodStartEnd.BtnSubmit_Click: formula written successfully, closing window");
                                Close();
                            }
                            else
                            {
                                LogUtility.LogDebug("GLGetPeriodStartEnd.BtnSubmit_Click: WriteFormulaToCell returned false, formula not written");
                            }
                        }
                        else
                        {
                            LogUtility.LogDebug("GLGetPeriodStartEnd.BtnSubmit_Click: validation failed - cell reference does not resolve to a valid range");
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
                    LogUtility.LogException(ex, "GLGetPeriodStartEnd.BtnSubmit_Click");
                }
            }
        }

        private void BtnClose_Click(object sender, RoutedEventArgs e)
        {
            LogUtility.LogDebug("GLGetPeriodStartEnd.BtnClose_Click invoked");
            Close();
        }
    }
}

