// GLGetPeriodByDate.xaml.cs in GLSense.Addin.Core
// Port of GLSense\Views\GLGetPeriodByDate.xaml.cs (FinalWorkingCode) for Group C
// (Segment/Period pickers) - GLSense_GetPeriodByDate formula picker (date + ledger +
// numeric offset). Re-pointed the same way as GLGetPeriod.xaml.cs (see that file's header
// for the full mapping); additionally re-points GLSense.Extensions.DatePickerExtensions ->
// GLSense.Addin.Core.Extensions.DatePickerExtensions (dtpDate.SetupTooltip - already
// ported verbatim in this pass).
using GLSense.Addin.Core.Extensions;
using GLSense.Addin.Core.Helpers;
using GLSense.Addin.Core.Infrastructure;
using GLSense.Addin.Core.Interfaces;
using GLSense.Addin.Core.Repositories;
using GLSense.Addin.Core.Utilities;
using GLSense.Addin.Core.ViewModels;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Input;
using Excel = Microsoft.Office.Interop.Excel;

namespace GLSense.Addin.Core.Views
{
    /// <summary>
    /// Interaction logic for GLGetPeriodByDate.xaml
    /// </summary>
    public partial class GLGetPeriodByDate : BaseWindow, IWarningHost
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
            InitializeComponent();
            ServiceLocator.Logger?.LogDebug("GLGetPeriodByDate constructor invoked");

            vm = new GLPeriodByDateModel(Dispatcher)
            {
                ExcelApp = ServiceLocator.ExcelApp,
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
            ServiceLocator.Logger?.LogDebug("GLGetPeriodByDate.Window_Loaded invoked");
            var (Address, Range) = GetActiveCellInfo();
            GlobalStateViewModel.Instance.ReferenceText = Address;
            ServiceLocator.Logger?.LogDebug($"GLGetPeriodByDate.Window_Loaded: active cell={Address}");

            if (!HasValidCubeAndLedger())
            {
                ServiceLocator.Logger?.LogDebug("GLGetPeriodByDate.Window_Loaded: no cube/ledger selected, skipping formula pre-population");
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
        }
        private static (string Address, Excel.Range Range) GetActiveCellInfo()
        {
            var rng = ServiceLocator.ExcelApp.ActiveCell;
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
            if (!IsYearFormula(rng))
                return new LedgerInfo();

            var ledgerInfo = ValidateLedgerFromFormula(rng);
            if (ledgerInfo.LedgerRecord == null)
            {
                await ShowLedgerNotFoundWarningAsync(ledgerInfo.LedgerName);
                return new LedgerInfo();
            }

            await LoadLedgerDataIfNeededAsync(ledgerInfo.LedgerRecord);
            return ledgerInfo;
        }

        private static bool IsYearFormula(Excel.Range rng)
        {
            if (!(bool)rng.HasFormula) return false;
            return rng.Formula.ToString().IndexOf("GLSense_GetPeriodByDate", StringComparison.OrdinalIgnoreCase) >= 0;
        }

        private LedgerInfo ValidateLedgerFromFormula(Excel.Range rng)
        {
            string formula = rng.Formula.ToString();
            var funcArgs = CommonFunctions.FormulaParameters(formula);
            var funcValues = CommonFunctions.FormulaValues(formula);

            string ledgerName = AppState.Instance.SelectedLedger.LedgerName;

            // If less than 3 arguments -> ledger is missing -> insert at index 0
            if (funcValues != null && funcValues.Count < 3)
            {
                funcArgs?.Insert(1, ledgerName); // or appropriate argument name
                funcValues?.Insert(1, ledgerName);
            }
            // If already 3 or more -> take ledger from formula
            else if (funcValues != null && funcValues.Count >= 3)
            {
                ledgerName = funcValues[1].Replace("\"", "");
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
                await LoadLedgerSegmentsWithProgressAsync(ledgerRecord);
            }
        }

        private async Task LoadLedgerSegmentsWithProgressAsync(dynamic ledgerRecord)
        {
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
                                ServiceLocator.Logger?.LogWarn("Loading cancelled by user");
                            }
                            await Task.CompletedTask;
                        });
                    return Task.CompletedTask;
                });

                ServiceLocator.Logger?.LogDebug($"GLGetPeriodByDate: calling FillResponsibilitiesAsync for LedgerId={ledgerRecord.LedgerId}, CubeId={AppState.Instance.SelectedCube.CubeId}");
                await CommonFunctions.FillResponsibilitiesAsync(
                    ledgerRecord.LedgerId,
                    AppState.Instance.SelectedCube.CubeId,
                    token);
                ServiceLocator.Logger?.LogDebug("GLGetPeriodByDate: FillResponsibilitiesAsync completed");
            }
            catch (OperationCanceledException)
            {
                ServiceLocator.Logger?.LogWarn("Loading ledger segments operation was cancelled.");
            }
            catch (Exception ex)
            {
                ServiceLocator.Logger?.LogException(ex, "GLGetPeriodByDate.LoadLedgerSegmentsWithProgressAsync");
            }
            finally
            {
                try
                {
                    if (ctsHelper != null && !ctsHelper.IsCancellationRequested)
                        ctsHelper.Cancel();

                    ctsHelper?.Dispose();
                }
                catch
                {
                    // Swallow dispose exceptions (Excel COM weirdness)
                }
                await Dispatcher.InvokeAsync(async () => await AppOverlayControl.HideBusyAsync());
            }
        }

        private async Task ShowLedgerNotFoundWarningAsync(string ledgerName)
        {
            ServiceLocator.Logger?.LogWarn($"GLGetPeriodByDate: Ledger \"{ledgerName}\" from formula not found in selected cube; falling back to default values.");
            string message = $"Ledger \"{ledgerName}\" in the formula does not exist in the selected cube!" +
                            Environment.NewLine + "Setting default values.";
            await AppOverlayControl.ShowWarningAsync(message);
        }
        public void CellSelectionWarning(string message)
        {
            try
            {
                AppOverlayControl.ShowWarning(message);
            }
            catch (Exception ex)
            {
                ServiceLocator.Logger?.LogException(ex, "GLGetPeriodByDate.CellSelectionWarning");
            }
        }
        private void BtnClose_Click(object sender, RoutedEventArgs e)
        {
            ServiceLocator.Logger?.LogDebug("GLGetPeriodByDate.BtnClose_Click invoked - closing window");
            this.Close();
        }
        private void BtnSubmit_Click(object sender, RoutedEventArgs e)
        {
            ServiceLocator.Logger?.LogDebug("GLGetPeriodByDate.BtnSubmit_Click invoked");
            if (string.IsNullOrWhiteSpace(CellReference.Text))
            {
                ServiceLocator.Logger?.LogDebug("GLGetPeriodByDate.BtnSubmit_Click: validation failed - cell reference is empty");
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
                                ServiceLocator.Logger?.LogDebug("GLGetPeriodByDate.BtnSubmit_Click: formula written successfully, closing window");
                                Close();
                            }
                            else
                            {
                                ServiceLocator.Logger?.LogWarn("GLGetPeriodByDate.BtnSubmit_Click: WriteFormulaToCell returned false");
                            }
                        }
                        else
                        {
                            ServiceLocator.Logger?.LogWarn($"GLGetPeriodByDate.BtnSubmit_Click: cell reference '{CellReference.Text}' did not resolve to a valid range");
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
                    ServiceLocator.Logger?.LogException(ex, "GLGetPeriodByDate.BtnSubmit_Click");
                }
            }
        }
    }
}
