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
        public static Task FillPeriods() => RunFillPeriods(isByDate: false);

        // OISR-19533 (ported from FinalWorkingCode): Discover > Periods By Date fills from a date in the active cell.
        public static Task FillPeriodsByDate() => RunFillPeriods(isByDate: true);

        private static async Task RunFillPeriods(bool isByDate)
        {
            string opName = isByDate ? "FillPeriodsByDate" : "FillPeriods";
            ServiceLocator.Logger?.LogDebug($"PeriodsDiscoverer.{opName}: started.");
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

                string dateArgument = string.Empty;
                int baseOffsetValue = 0;
                DateTime baseDate = default;
                string rngValue = string.Empty;
                Func<PeriodModel, bool> periodMatch;

                if (isByDate)
                {
                    if (!TryResolveBaseDate(out baseDate, out dateArgument, out baseOffsetValue))
                    {
                        ServiceLocator.Logger?.LogWarn($"PeriodsDiscoverer.{opName}: the selected cell does not contain a recognizable date.");
                        await ShowWarnMessage("The selected cell does not contain a valid date.");
                        return;
                    }

                    // Calendar-day containment: a period's EndDate is midnight of its last day.
                    periodMatch = p => p.StartDate.Date <= baseDate.Date && p.EndDate.Date >= baseDate.Date;
                }
                else
                {
                    rngValue = CellActive.Value2.ToString();
                    periodMatch = p => p.PeriodName == rngValue;
                }

                Periods = PModel(LedgerString, periodMatch);

                Token.ThrowIfCancellationRequested();

                if (Periods == null || !Periods.Any())
                {
                    await ShowWarnMessage("Failed in fetching the period values.");
                    return;
                }

                if (isByDate)
                {
                    int dateIndex = Periods.FindIndex(p => periodMatch(p));

                    Token.ThrowIfCancellationRequested();

                    if (dateIndex < 0)
                    {
                        ServiceLocator.Logger?.LogWarn($"PeriodsDiscoverer.{opName}: date '{baseDate:d}' does not fall within any period for ledger '{LedgerString}'.");
                        await ShowWarnMessage($"The selected date \"{baseDate:d}\" does not exists in the periods list.");
                        return;
                    }

                    // The active cell resolves to the date's period shifted by the offset already in its own
                    // GLSense_GetPeriodByDate formula (0 for a plain date); the fill continues from there.
                    BasePeriodIndex = dateIndex + baseOffsetValue;

                    if (BasePeriodIndex < 0 || BasePeriodIndex >= Periods.Count)
                    {
                        ServiceLocator.Logger?.LogWarn($"PeriodsDiscoverer.{opName}: date '{baseDate:d}' with offset {baseOffsetValue} resolves outside the periods list for ledger '{LedgerString}'.");
                        await ShowWarnMessage($"The selected date \"{baseDate:d}\" with offset {baseOffsetValue} does not exists in the periods list.");
                        return;
                    }

                    BasePeriod = Periods[BasePeriodIndex];
                    ServiceLocator.Logger?.LogDebug($"PeriodsDiscoverer.{opName}: base date '{baseDate:d}' (offset {baseOffsetValue}) resolved to period '{BasePeriod.PeriodName}' at index {BasePeriodIndex} of {Periods.Count}.");
                }
                else
                {
                    BasePeriod = Periods.FirstOrDefault(p => p.PeriodName == rngValue);

                    Token.ThrowIfCancellationRequested();

                    if (BasePeriod == null)
                    {
                        await ShowWarnMessage($"The selected item \"{rngValue}\" does not exists in the periods list.");
                        return;
                    }

                    BasePeriodIndex = Periods.FindIndex(p => p.PeriodName == rngValue);
                }

                await RunPeriodDiscovery(isByDate, dateArgument, baseOffsetValue);

                await MessageProgressWindowAsync("Excel refreshing the formulas.");
                ServiceLocator.Logger?.LogDebug($"PeriodsDiscoverer.{opName}: completed.");
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
                    ServiceLocator.Logger?.LogWarn($"{opName}: failed disposing CancellationHelper (non-fatal): {ex.Message}");
                }
                await SafelyCloseWindowAsync();
                CommonMethods.TryEnableExcelSettings($"PeriodsDiscoverer.{opName}");
            }
        }
        private static async Task RunPeriodDiscovery(bool isByDate, string dateArgument, int baseOffsetValue)
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
                await FillPeriodDiscoverValues(isReverse, Selection, CellActive, BasePeriodIndex, ledgerRef, Periods, isByDate, dateArgument, baseOffsetValue);
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
            List<PeriodModel> periods,
            bool isByDate,
            string dateArgument,
            int baseOffsetValue)
        {
            await MessageProgressWindowAsync("Filling period details.");
            await Task.Yield();

            if (!CanProcessSelection(loopRng))
                return;

            // Base (active) cell: do not modify this cell.
            Excel.Range baseCell = formulaRange;

            bool writeAsFormula = ServiceLocator.RibbonController?.GetControlPressed("RibAsFormula") ?? false;
            string rangeRef = $"'{formulaRange.Worksheet.Name.Replace("'", "''")}'!{formulaRange.Address[true, true]}"; // Excel doubles ' in quoted sheet names

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
                            // Offset is the relative move. GLSense_GetPeriodByDate's offset is relative to the date's
                            // own period (dateArgument is the same for every cell), so the base cell's offset is added.
                            targetCell.Value = isByDate
                                ? $"=GLSense_GetPeriodByDate({dateArgument}, {ledgerRef}, {baseOffsetValue + offset})"
                                : $"=GLSense_GetPeriod({rangeRef}, {offset}, {ledgerRef})";
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

        /// <summary>
        /// The base date for "Periods By Date" from the active cell: a date value or any formula that evaluates to one.
        /// If the cell already holds a GLSense_GetPeriodByDate(...) formula (which evaluates to a period name), the
        /// date argument and offset are read back from that formula instead.
        /// </summary>
        private static bool TryResolveBaseDate(out DateTime baseDate, out string dateArgument, out int baseOffsetValue)
        {
            baseDate = default;
            dateArgument = string.Empty;
            baseOffsetValue = 0;

            try
            {
                if ((bool)CellActive.HasFormula)
                {
                    string formula = CellActive.Formula.ToString();

                    if (formula.Contains("GLSense_GetPeriodByDate("))
                    {
                        var parameters = CommonFunctions.FormulaParameters(formula);
                        var values = CommonFunctions.FormulaValues(formula);

                        string param = parameters != null && parameters.Count > 0 ? parameters[0] : null;
                        string value = values != null && values.Count > 0 ? values[0]?.Trim() : null;
                        bool isCellRef = !string.IsNullOrWhiteSpace(param) && param.Contains("$");

                        if (!string.IsNullOrWhiteSpace(value) && TryParseDateArgument(value, out baseDate))
                        {
                            dateArgument = isCellRef ? param : value;

                            // The offset is the 3rd argument; 0 (the UDF's default) when missing or not a number.
                            string offsetText = values.Count > 2 ? values[2]?.Replace("\"", "").Trim() : null;
                            if (!string.IsNullOrWhiteSpace(offsetText) && int.TryParse(offsetText, out int parsedOffset))
                                baseOffsetValue = parsedOffset;

                            ServiceLocator.Logger?.LogDebug($"PeriodsDiscoverer.TryResolveBaseDate: from GLSense_GetPeriodByDate formula, dateArgument='{dateArgument}', baseDate={baseDate:d}, baseOffsetValue={baseOffsetValue}.");
                            return true;
                        }
                    }
                }

                object rawValue = CellActive.Value2;
                if (rawValue == null)
                    return false;

                baseDate = Udf.UdfDispatcher.ParsePeriodDate(rawValue);
                if (baseDate == default)
                    return false;

                dateArgument = $"'{CellActive.Worksheet.Name.Replace("'", "''")}'!{CellActive.Address[true, true]}"; // Excel doubles ' in quoted sheet names
                return true;
            }
            catch (Exception ex)
            {
                ServiceLocator.Logger?.LogException(ex, "PeriodsDiscoverer.TryResolveBaseDate");
                baseDate = default;
                dateArgument = string.Empty;
                baseOffsetValue = 0;
                return false;
            }
        }

        /// <summary>A formula argument's text as a date: a literal DATE(y,m,d), or anything ParsePeriodDate accepts.</summary>
        private static bool TryParseDateArgument(string text, out DateTime date)
        {
            date = default;
            if (string.IsNullOrWhiteSpace(text))
                return false;

            var match = System.Text.RegularExpressions.Regex.Match(text, @"DATE\(\s*(\d{4})\s*,\s*(\d{1,2})\s*,\s*(\d{1,2})\s*\)", System.Text.RegularExpressions.RegexOptions.IgnoreCase);
            if (match.Success)
            {
                try
                {
                    date = new DateTime(int.Parse(match.Groups[1].Value), int.Parse(match.Groups[2].Value), int.Parse(match.Groups[3].Value));
                    return true;
                }
                catch (Exception ex)
                {
                    ServiceLocator.Logger?.LogException(ex, $"PeriodsDiscoverer.TryParseDateArgument: invalid DATE(...) literal '{text}'");
                    return false;
                }
            }

            date = Udf.UdfDispatcher.ParsePeriodDate(text);
            return date != default;
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
        // FSG reports scheduled in the Orbit application can use several ledgers, and their
        // templates carry them as one ';'-joined Ledger value (e.g. "Vision Operations (USA);Vision Brazil SL (USD)").
        // Period data is stored per ledger, so try each ledger in the order given and use the first
        // one whose periods satisfy periodMatch; later ledgers are only checked when an earlier one
        // has no period data or lacks the requested period. A single ledger behaves exactly as PModel.
        // Mirrors UdfDispatcher.PModel(string, Func<PeriodModel, bool>).
        private static List<PeriodModel> PModel(string lName, Func<PeriodModel, bool> periodMatch)
        {
            if (string.IsNullOrEmpty(lName))
                return PModel(lName);

            string[] ledgerNames = lName
                .Split(new[] { ';' }, StringSplitOptions.RemoveEmptyEntries)
                .Select(l => l.Trim().Trim('"'))
                .Where(l => l.Length > 0)
                .ToArray();

            if (ledgerNames.Length <= 1)
                return PModel(lName);

            List<PeriodModel> firstLoaded = null;
            foreach (string ledgerName in ledgerNames)
            {
                var periods = PModel(ledgerName);
                if (!periods.Any())
                {
                    ServiceLocator.Logger?.LogDebug($"PeriodsDiscoverer.PModel: no period data for ledger \"{ledgerName}\", trying next ledger");
                    continue;
                }

                if (periods.Any(periodMatch))
                {
                    ServiceLocator.Logger?.LogDebug($"PeriodsDiscoverer.PModel: using ledger \"{ledgerName}\" from \"{lName}\"");
                    return periods;
                }

                ServiceLocator.Logger?.LogDebug($"PeriodsDiscoverer.PModel: requested period not found for ledger \"{ledgerName}\", trying next ledger");
                firstLoaded ??= periods;
            }

            // No ledger had the period: hand back the first ledger that had data so FillPeriods'
            // own "does not exist in the periods list" message shows, as for a single ledger.
            return firstLoaded ?? new List<PeriodModel>();
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
