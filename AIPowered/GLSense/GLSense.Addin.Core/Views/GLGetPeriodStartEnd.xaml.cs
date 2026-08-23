// GLGetPeriodStartEnd.xaml.cs in GLSense.Addin.Core
// Port of GLSense\Views\GLGetPeriodStartEnd.xaml.cs (FinalWorkingCode) for Group C
// (Segment/Period pickers) - shared window for GLSense_GetPeriodStart/End, driven by the
// FormulaName ctor param ("START"/"END") into the shared GLPeriodDetails ViewModel (adds
// the "Include Adjacent Periods" checkbox vs. GLGetPeriodDetails). Re-pointed the same way
// as GLGetPeriod.xaml.cs (see that file's header for the full mapping).
using GLSense.Addin.Core.Helpers;
using GLSense.Addin.Core.Infrastructure;
using GLSense.Addin.Core.Interfaces;
using GLSense.Addin.Core.Models;
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
            InitializeComponent();
            ServiceLocator.Logger?.LogDebug($"GLGetPeriodStartEnd constructor invoked with FormulaName={FormulaName}");

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
            ServiceLocator.Logger?.LogDebug("GLGetPeriodStartEnd.Window_Loaded invoked");

            var (Address, Range) = GetActiveCellInfo();
            GlobalStateViewModel.Instance.ReferenceText = Address;
            ServiceLocator.Logger?.LogDebug($"GLGetPeriodStartEnd.Window_Loaded: active cell={Address}");

            if (!HasValidCubeAndLedger())
            {
                ServiceLocator.Logger?.LogDebug("GLGetPeriodStartEnd.Window_Loaded: no cube/ledger selected, skipping formula pre-population");
                return;
            }

            var ledgerInfo = await ProcessCellFormulaAsync(Range);

            string ledgerName = GetLedgerName(ledgerInfo);
            await vm.LoadDataAsync(ledgerInfo.FuncArgs, ledgerInfo.FuncValues);

            await SetLedgerDropdownAsync(ledgerName);
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
        private async Task<LedgerInfo> ProcessCellFormulaAsync(Excel.Range rng)
        {
            if (!IsPeriodFormula(rng))
                return new LedgerInfo();

            var ledgerInfo = ValidateLedgerFromFormula(rng);
            if (ledgerInfo.LedgerRecord == null)
                await ShowLedgerNotFoundWarningAsync(ledgerInfo.LedgerName);

            await LoadLedgerDataIfNeededAsync(ledgerInfo.LedgerRecord);
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
            string formula = rng.Formula.ToString();
            var funcArgs = CommonFunctions.FormulaParameters(formula);
            var funcValues = CommonFunctions.FormulaValues(formula);
            string ledgerName = AppState.Instance.SelectedLedger.LedgerName;

            // If less than 3 arguments -> ledger is missing -> insert at index 2 or at the end (if funcArgs is shorter than funcValues)
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
            if (ledgerRecord == null) return;

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

                ServiceLocator.Logger?.LogDebug($"GLGetPeriodStartEnd: calling FillResponsibilitiesAsync for LedgerId={ledgerRecord.LedgerId}, CubeId={AppState.Instance.SelectedCube.CubeId}");
                await CommonFunctions.FillResponsibilitiesAsync(
                    ledgerRecord.LedgerId,
                    AppState.Instance.SelectedCube.CubeId,
                    token);
                ServiceLocator.Logger?.LogDebug("GLGetPeriodStartEnd: FillResponsibilitiesAsync completed");
            }
            catch (OperationCanceledException)
            {
                ServiceLocator.Logger?.LogWarn("Loading ledger segments operation was cancelled.");
            }
            catch (Exception ex)
            {
                ServiceLocator.Logger?.LogException(ex, "GLGetPeriodStartEnd.LoadLedgerSegmentsWithProgressAsync");
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
            ServiceLocator.Logger?.LogWarn($"GLGetPeriodStartEnd: Ledger \"{ledgerName}\" from formula not found in selected cube; falling back to default values.");
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
            try
            {
                AppOverlayControl.ShowWarning(message);
            }
            catch (Exception ex)
            {
                ServiceLocator.Logger?.LogException(ex, "GLGetPeriodStartEnd.CellSelectionWarning");
            }
        }
        private void BtnSubmit_Click(object sender, RoutedEventArgs e)
        {
            ServiceLocator.Logger?.LogDebug("GLGetPeriodStartEnd.BtnSubmit_Click invoked");
            if (string.IsNullOrWhiteSpace(CellReference.Text))
            {
                ServiceLocator.Logger?.LogDebug("GLGetPeriodStartEnd.BtnSubmit_Click: validation failed - cell reference is empty");
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
                                ServiceLocator.Logger?.LogDebug("GLGetPeriodStartEnd.BtnSubmit_Click: formula written successfully, closing window");
                                Close();
                            }
                            else
                            {
                                ServiceLocator.Logger?.LogWarn("GLGetPeriodStartEnd.BtnSubmit_Click: WriteFormulaToCell returned false");
                            }
                        }
                        else
                        {
                            ServiceLocator.Logger?.LogWarn($"GLGetPeriodStartEnd.BtnSubmit_Click: cell reference '{CellReference.Text}' did not resolve to a valid range");
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
                    ServiceLocator.Logger?.LogException(ex, "GLGetPeriodStartEnd.BtnSubmit_Click");
                }
            }
        }

        private void BtnClose_Click(object sender, RoutedEventArgs e)
        {
            ServiceLocator.Logger?.LogDebug("GLGetPeriodStartEnd.BtnClose_Click invoked - closing window");
            Close();
        }
    }
}
