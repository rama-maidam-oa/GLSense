// PeriodsDiscoverer.cs in GLSense.Addin.Core
// Port of GLSense\Utilities\PeriodsDiscoverer.cs (FinalWorkingCode) for Group D
// (Segment/Period discoverers). Preserves the original's offset-based fill-direction
// logic verbatim - only the plumbing below was re-pointed (same set of changes as
// SegmentDiscoverer.cs in this same pass - see that file's header for the full mapping):
//   - GLSense.Helpers/.Models/.Views -> GLSense.Addin.Core.* equivalents.
//   - GLSense.Service.ServiceLocator.PeriodDataService -> Services.DataServiceLocator.
//     PeriodDataService (Group C's data-service layer).
//   - AppState.Instance.ExcelApp -> ServiceLocator.ExcelApp.
//   - LogUtility.* -> ServiceLocator.Logger.*.
//   - WinForms MessageBoxIcon/MessageBoxButtons -> WPF MessageBoxImage/MessageBoxButton.
//   - AddinModule.CurrentInstance.RibAsFormula.Pressed -> ServiceLocator.RibbonController.
//     GetControlPressed("RibAsFormula").
//   - CreateAndShowProgressWindow re-pointed the same way as SegmentDiscoverer's own
//     copy (WpfAppManager.InvokeOnWpfThread has no Func<T> overload here; GLWaitWindow
//     captured via closure; win.ShowWithOwner(hwnd) -> win.Show()).
//   - Dropped the unused "using ControlzEx.Standard;" - this project does not reference
//     ControlzEx and nothing in this file needs it.
using GLSense.Addin.Core.Helpers;
using GLSense.Addin.Core.Infrastructure;
using GLSense.Addin.Core.Models;
using GLSense.Addin.Core.Services;
using GLSense.Addin.Core.Views;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Threading;
using Excel = Microsoft.Office.Interop.Excel;

namespace GLSense.Addin.Core.Utilities
{
    public static class PeriodsDiscoverer
    {
        private static GLWaitWindow Win { get; set; }
        private static Excel.Application ExcelApp { get; set; }
        private static Excel.Workbook PrWorbook { get; set; }
        private static Excel.Worksheet PrWorksheet { get; set; }
        private static Excel.Range CellActive { get; set; }
        private static Excel.Range Selection { get; set; }
        private static List<PeriodModel> Periods { get; set; }
        private static PeriodModel BasePeriod { get; set; }
        private static int BasePeriodIndex { get; set; }
        private static string LedgerReference { get; set; }
        private static string LedgerString { get; set; }
        private static CancellationHelper _ctsHelper;
        private static CancellationToken Token => _ctsHelper?.GetToken() ?? default;
        public static async Task FillPeriods()
        {
            ServiceLocator.Logger?.LogDebug("PeriodsDiscoverer.FillPeriods: started.");
            try
            {
                _ctsHelper = new CancellationHelper();

                CommonMethods.DisableExcelSettings();

                ExcelApp = ServiceLocator.ExcelApp;
                CellActive = ExcelApp.ActiveCell;
                Selection = ExcelApp.Selection as Excel.Range;

                bool isValid = await ValidateAsync();

                Token.ThrowIfCancellationRequested();

                if (!isValid) return;

                string rngValue = CellActive.Value2.ToString();

                PrWorbook = ExcelApp.ActiveWorkbook;
                PrWorksheet = CellActive.Worksheet;

                Win = CreateAndShowProgressWindow(_ctsHelper);
                await InitializeProgressWindowAsync();

                BuildPeriodHelpers();

                if (string.IsNullOrWhiteSpace(LedgerString) && string.IsNullOrWhiteSpace(LedgerReference))
                {
                    await ShowWarnMessage("No ledger found for period discover.");
                    return;
                }

                Periods = PModel(LedgerString);

                Token.ThrowIfCancellationRequested();

                if (Periods == null || !Periods.Any())
                {
                    await ShowWarnMessage("Failed in fetching the period values.");
                    return;
                }

                BasePeriod = Periods.FirstOrDefault(p => p.PeriodName == rngValue);

                Token.ThrowIfCancellationRequested();

                if (BasePeriod == null)
                {
                    await ShowWarnMessage($"The selected item \"{rngValue}\" does not exists in the periods list.");
                    return;
                }

                BasePeriodIndex = Periods.FindIndex(p => p.PeriodName == rngValue);

                await RunPeriodDiscovery();

                await MessageProgressWindowAsync("Excel refreshing the formulas.");
            }
            catch (OperationCanceledException)
            {
                await ShowCancelledAsync();
            }
            catch (Exception ex)
            {
                await HandleUnexpectedErrorAsync(ex);
            }
            finally
            {
                try
                {
                    if (_ctsHelper != null && !_ctsHelper.IsCancellationRequested)
                        _ctsHelper.Cancel();

                    _ctsHelper?.Dispose();  // always safe - handles all cases
                }
                catch (Exception ex)
                {
                    ServiceLocator.Logger?.LogWarn($"FillPeriods: failed disposing CancellationHelper (non-fatal): {ex.Message}");
                }
                await SafelyCloseWindowAsync();
                CommonMethods.EnableExcelSettings();
            }
        }
        private static async Task RunPeriodDiscovery()
        {

            await MessageProgressWindowAsync("Extracting information.");
            await Task.Yield();

            Token.ThrowIfCancellationRequested();

            bool isMultiColumn = Selection.Columns.Count > 1;
            bool isMultiRow = Selection.Rows.Count > 1;

            if (!isMultiColumn && !isMultiRow)
                return; // Nothing to do

            bool isVertical = isMultiRow && !isMultiColumn;
            bool isReverse = isVertical
                ? Selection.Cells.Row < CellActive.Row
                : Selection.Cells.Column < CellActive.Column;

            string ledgerRef = GetLedgerReference();

            if (!string.IsNullOrWhiteSpace(ledgerRef))
            {
                await FillPeriodDiscoverValues(isReverse, Selection, CellActive, BasePeriodIndex, ledgerRef, Periods);
            }
        }

        private static string GetLedgerReference()
        {
            if (!string.IsNullOrWhiteSpace(LedgerReference))
                return LedgerReference;

            if (!string.IsNullOrWhiteSpace(LedgerString))
                return $"\"{LedgerString.Replace("\"", "")}\"";

            return string.Empty;
        }
        private static bool CanProcessSelection(Excel.Range loopRng)
        {
            if (loopRng?.Cells?.Count <= 1)
                return false;

            bool isHorizontal = IsHorizontal(loopRng);
            bool isVertical = IsVertical(loopRng);

            if (!isHorizontal && !isVertical)
            {
                ServiceLocator.Logger?.LogWarn("Selection must be a single row or a single column. Aborting to avoid unintended updates.");
                return false;
            }

            return true;
        }
        private static bool IsHorizontal(Excel.Range rng) => rng.Rows.Count == 1 && rng.Columns.Count >= 1;
        private static bool IsVertical(Excel.Range rng) => rng.Columns.Count == 1 && rng.Rows.Count >= 1;

        private static async Task FillPeriodDiscoverValues(
            bool isReverse,
            Excel.Range loopRng,
            Excel.Range formulaRange,
            int periodIndex,
            string ledgerRef,
            List<PeriodModel> periods)
        {
            await MessageProgressWindowAsync("Filling period details.");
            await Task.Yield();

            if (!CanProcessSelection(loopRng))
                return;

            // Base (active) cell: do not modify this cell.
            Excel.Range baseCell = formulaRange;

            bool writeAsFormula = ServiceLocator.RibbonController?.GetControlPressed("RibAsFormula") ?? false;
            string rangeRef = $"'{formulaRange.Worksheet.Name}'!{formulaRange.Address[true, true]}";

            try
            {
                int cellCount = loopRng.Cells.Count;

                // Generate offsets that ALWAYS skip the base cell (0)
                // Forward: 1..(cellCount-1)   Reverse: -(cellCount-1)..-1
                var offsets = GenerateOffsets(isReverse, cellCount);

                foreach (int offset in offsets)
                {
                    Token.ThrowIfCancellationRequested();

                    // Compute the target cell relative to the base cell
                    Excel.Range targetCell = IsHorizontal(loopRng)
                        ? baseCell.Offset[0, offset]     // columns: right(+), left(-)
                        : baseCell.Offset[offset, 0];    // rows: down(+), up(-)

                    if (targetCell == null)
                        continue;

                    // Optional: ensure we only write within the selected range
                    // If you want this guard, uncomment:

                    int targetPeriodIndex = periodIndex + offset;
                    if (targetPeriodIndex < 0 || targetPeriodIndex >= periods.Count)
                        continue;

                    try
                    {
                        if (!writeAsFormula)
                        {
                            targetCell.NumberFormat = "@";
                            targetCell.Value2 = periods[targetPeriodIndex].PeriodName;
                        }
                        else
                        {
                            targetCell.NumberFormat = AppConstants.General;
                            // Offset is the relative move used by the GLSense_GetPeriod function
                            targetCell.Value = $"=GLSense_GetPeriod({rangeRef}, {offset}, {ledgerRef})";
                        }
                    }
                    catch (Exception ex)
                    {
                        ServiceLocator.Logger?.LogException(ex, $"PeriodsDiscoverer.FillPeriodDiscoverValues: failed writing target cell at offset {offset} (targetPeriodIndex={targetPeriodIndex}) - clearing cell instead");
                        targetCell.Value = string.Empty;
                    }
                }

                // Keep async signature (optional)
                await Task.CompletedTask;
            }
            catch (Exception ex)
            {
                ServiceLocator.Logger?.LogException(ex, $"PeriodsDiscoverer.FillPeriodDiscoverValues: failed (isReverse={isReverse}, periodIndex={periodIndex}, ledgerRef='{ledgerRef}')");
            }
        }


        private static int[] GenerateOffsets(bool isReverse, int cellCount)
        {

            if (cellCount <= 0)
                return Array.Empty<int>();

            Token.ThrowIfCancellationRequested();
            // We will generate (cellCount - 1) offsets that exclude 0 (the base)
            // Forward: 1..(cellCount - 1)
            // Reverse: -(cellCount - 1)..-1

            int start = isReverse ? -(cellCount - 1) : 1;
            int end = isReverse ? -1 : (cellCount - 1);

            var offsets = new List<int>(cellCount - 1);
            for (int i = start; i <= end; i++)
                offsets.Add(i);

            return offsets.ToArray();


        }
        private static void BuildPeriodHelpers()
        {
            try
            {
                string defaultLedger = AppState.Instance.SelectedLedger.LedgerName;

                // Default values if no formula or no supported period function
                LedgerString = defaultLedger;
                LedgerReference = string.Empty;

                if (!(bool)CellActive.HasFormula)
                    return;

                string formula = CellActive.Formula.ToString();

                string[] periodFunctions =
                {
                    "GLSense_GetPeriod(",
                    "GLSense_GetPeriodByDate(",
                    "GLSense_GetPeriodStart(",
                    "GLSense_GetPeriodEnd(",
                    "GLSense_GetPeriodByYear("
                };

                if (!periodFunctions.Any(f => formula.Contains(f)))
                    return;

                // Extract parameters and actual values from the formula
                var parameters = CommonFunctions.FormulaParameters(formula);
                var values = CommonFunctions.FormulaValues(formula);

                bool isPeriodOnlyFunction = formula.Contains("GLSense_GetPeriod(")
                                         || formula.Contains("GLSense_GetPeriodByYear(");

                int ledgerParamIndex = isPeriodOnlyFunction ? 2 : 1;

                // Safely extract ledger reference (e.g., $A$1) and ledger string value
                ExtractLedgerInfo(parameters, values, ledgerParamIndex);
            }
            catch (Exception ex)
            {
                ServiceLocator.Logger?.LogException(ex, "PeriodsDiscoverer.BuildPeriodHelpers: failed");
                LedgerString = string.Empty;
                LedgerReference = string.Empty;
            }
        }

        private static void ExtractLedgerInfo(List<string> parameters, List<string> values, int ledgerIndex)
        {
            try
            {
                // Ledger reference: the raw parameter like "$B$10"
                if (ledgerIndex < parameters.Count)
                {
                    string param = parameters[ledgerIndex]?.ToString();
                    if (!string.IsNullOrWhiteSpace(param) && param.Contains("$"))
                    {
                        LedgerReference = param;
                    }
                    else
                    {
                        LedgerReference = string.Empty;
                    }
                }
                else
                {
                    LedgerReference = string.Empty;
                }

                // Ledger string value: the evaluated value, e.g., "Actual"
                if (ledgerIndex < values.Count)
                {
                    string value = values[ledgerIndex]?.ToString();
                    LedgerString = !string.IsNullOrWhiteSpace(value)
                        ? value.Replace("\"", "").Trim()
                        : string.Empty;
                }
                else
                {
                    LedgerString = string.Empty;
                }
            }
            catch (Exception ex)
            {
                ServiceLocator.Logger?.LogException(ex, $"PeriodsDiscoverer.ExtractLedgerInfo(ledgerIndex={ledgerIndex}): failed");
                LedgerString = string.Empty;
                LedgerReference = string.Empty;
            }
        }

        //Standard helpers

        private static async Task<bool> ValidateAsync()
        {
            if (ServiceLocator.ExcelApp == null)
            {
                await ShowErrorMessage("Unable to get excel instance.");
                return false;
            }

            if (!AppState.Instance.IsLoginCompleted)
            {
                await ShowErrorMessage("Please login to the instance.");
                return false;
            }

            if (Selection.Cells.Count == 1)
            {
                await ShowWarnMessage("Selection cannot be a single cell. It must be a range of multiple cells, either vertically or horizontally.");
                return false;
            }

            if (Selection.Rows.Count > 1 && Selection.Columns.Count > 1)
            {
                await ShowWarnMessage("Selection can be multiple rows with a single column or a single row with multiple columns.");
                return false;
            }

            if (Selection.Address != null && Selection.Address.ToString().Contains(","))
            {
                await ShowWarnMessage("Selection cannot be non-contagious.");
                return false;
            }

            if (CellActive.Value2 == null)
            {
                await ShowWarnMessage("The first cell of the selection cannot be empty.");
                return false;
            }
            return true;
        }

        private static async Task ShowErrorMessage(string message)
        {
            await SafelyCloseWindowAsync();
            CommonFunctions.GLSenseMessage(
                message,
                MessageBoxImage.Error,
                MessageBoxButton.OK);
        }
        private static async Task ShowWarnMessage(string message)
        {
            await SafelyCloseWindowAsync();
            CommonFunctions.GLSenseMessage(
                message,
                MessageBoxImage.Warning,
                MessageBoxButton.OK);
        }
        private static async Task ShowCancelledAsync()
        {
            await SafelyCloseWindowAsync();
            CommonFunctions.GLSenseMessage(
                "Operation cancelled!",
                MessageBoxImage.Error,
                MessageBoxButton.OK);
        }

        private static async Task HandleUnexpectedErrorAsync(Exception ex)
        {
            ServiceLocator.Logger?.LogException(ex, "PeriodsDiscoverer.FillPeriods: unexpected error");
            await SafelyCloseWindowAsync();
            CommonFunctions.GLSenseMessage(
                $"An unexpected error occurred.{Environment.NewLine}{ex.Message}",
                MessageBoxImage.Error,
                MessageBoxButton.OK);
        }
        private static GLWaitWindow CreateAndShowProgressWindow(CancellationHelper cts)
        {
            GLWaitWindow win = null;
            try
            {
                // WpfAppManager.InvokeOnWpfThread only takes an Action (no return value),
                // so capture the created window from inside the delegate - same pattern
                // AddinEntry.LedgerChanged / SegmentDiscoverer.CreateAndShowProgressWindow use.
                WpfAppManager.InvokeOnWpfThread(() =>
                {
                    try
                    {
                        // Use the passed-in cts, don't create a new one
                        win = new GLWaitWindow(cts);
                        win.Show();
                        win.StartMonitoring();
                    }
                    catch (Exception ex)
                    {
                        ServiceLocator.Logger?.LogException(ex, "PeriodsDiscoverer.CreateAndShowProgressWindow: failed on WPF thread");
                        win = null;
                    }
                });
            }
            catch (Exception ex)
            {
                ServiceLocator.Logger?.LogException(ex, "PeriodsDiscoverer.CreateAndShowProgressWindow: InvokeOnWpfThread failed");
            }
            return win;
        }

        private static Task InitializeProgressWindowAsync()
        {
            // Basic guards
            if (Win == null || Win.Dispatcher == null)
                return Task.CompletedTask;

            try
            {
                // Fire-and-forget: this is invoked from contexts that may run on a
                // thread with no captured SynchronizationContext, so awaiting the
                // dispatch would risk resuming subsequent Excel COM calls on an
                // arbitrary ThreadPool thread instead of the calling thread.
                _ = Win.Dispatcher.InvokeAsync(
                    () =>
                    {
                        Win.SetProcessTitle("Periods Discoverer");
                        Win.SetProcessMessage("Filling the periods from selected.");
                    },
                    DispatcherPriority.Normal);

                return Task.CompletedTask;
            }
            catch (TaskCanceledException)
            {
                // Dispatcher shutting down; nothing to do
                return Task.CompletedTask;
            }
            catch (Exception ex)
            {
                // Last-resort logging if something unexpected happens
                ServiceLocator.Logger?.LogException(ex, "PeriodsDiscoverer.InitializeProgressWindowAsync: failed");
                return Task.CompletedTask;
            }
        }
        private static Task MessageProgressWindowAsync(string message)
        {
            // Basic guards
            if (Win == null || Win.Dispatcher == null)
                return Task.CompletedTask;

            try
            {
                // Fire-and-forget: do not await the dispatcher operation itself.
                // Awaiting here would introduce a suspend point that can let the
                // caller resume on a different thread (e.g. a background worker
                // with no captured SynchronizationContext), which is unsafe when
                // the code right after the await touches Excel COM objects.
                _ = Win.Dispatcher.InvokeAsync(
                    () =>
                    {
                        Win.SetProcessMessage(message);
                    },
                    DispatcherPriority.Normal);

                return Task.CompletedTask;
            }
            catch (TaskCanceledException)
            {
                // Dispatcher shutting down; nothing to do
                return Task.CompletedTask;
            }
            catch (Exception ex)
            {
                // Last-resort logging if something unexpected happens
                ServiceLocator.Logger?.LogException(ex, $"PeriodsDiscoverer.MessageProgressWindowAsync(message='{message}'): failed");
                return Task.CompletedTask;
            }
        }
        private static async Task SafelyCloseWindowAsync()
        {
            if (Win == null)
                return;

            try
            {
                if (Win.Dispatcher.CheckAccess())  // Already on UI thread
                {
                    Win.RequestClose();
                }
                else
                {
                    await Win.Dispatcher.InvokeAsync(() => Win.RequestClose());
                }
            }
            catch (Exception ex)
            {
                ServiceLocator.Logger?.LogError($"Error closing wait window: {ex.Message}");
            }
            finally
            {
                ExcelWindowHelper.ActivateExcelMainWindow(ServiceLocator.ExcelApp);
                Win = null;
            }
        }
        private static List<PeriodModel> PModel(string lName)
        {
            try
            {
                List<PeriodModel> periods = LoadPeriodsForLedger(lName);
                if (periods == null)
                {
                    return new List<PeriodModel>();
                }
                else
                {
                    return periods;
                }
            }
            catch (Exception ex)
            {
                ServiceLocator.Logger?.LogException(ex, $"PeriodsDiscoverer.PModel(lName='{lName}'): failed");
                return new List<PeriodModel>();
            }
        }
        private static List<PeriodModel> LoadPeriodsForLedger(string ledgerName)
        {
            try
            {
                var dataService = DataServiceLocator.PeriodDataService;
                return new List<PeriodModel>(dataService.GetPeriodsForLedger(ledgerName));
            }
            catch (Exception ex)
            {
                ServiceLocator.Logger?.LogException(ex, $"Failed to load periods for ledger '{ledgerName}'");
                return new List<PeriodModel>();
            }
        }
    }
}
