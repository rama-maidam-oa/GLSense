using ControlzEx.Standard;
using GLSense.Helpers;
using GLSense.Models;
using GLSense.Service;
using GLSense.Views;
using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Windows.Interop;
using System.Windows.Threading;
using Excel = Microsoft.Office.Interop.Excel;

namespace GLSense.Utilities
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

        public static Task FillPeriodsByDate() => RunFillPeriods(isByDate: true);

        private static async Task RunFillPeriods(bool isByDate)
        {
            string opName = isByDate ? "FillPeriodsByDate" : "FillPeriods";
            try
            {
                LogUtility.LogDebug($"PeriodsDiscoverer.{opName} started.");
                _ctsHelper = new CancellationHelper();

                CommonMethods.DisableExcelSettings();

                ExcelApp = AppState.Instance.ExcelApp;
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
                    LogUtility.LogWarn($"PeriodsDiscoverer.{opName}: no ledger found for period discover (activeCellValue={CellActive.Value2}).");
                    await ShowWarnMessage("No ledger found for period discover.");
                    return;
                }

                Periods = PModel(LedgerString);

                Token.ThrowIfCancellationRequested();

                if (Periods == null || !Periods.Any())
                {
                    LogUtility.LogWarn($"PeriodsDiscoverer.{opName}: failed in fetching period values for ledger '{LedgerString}'.");
                    await ShowWarnMessage("Failed in fetching the period values.");
                    return;
                }

                string dateArgument = string.Empty;
                int baseOffsetValue = 0;

                if (isByDate)
                {
                    if (!TryResolveBaseDate(out DateTime baseDate, out dateArgument, out baseOffsetValue))
                    {
                        LogUtility.LogWarn($"PeriodsDiscoverer.{opName}: the selected cell does not contain a recognizable date.");
                        await ShowWarnMessage("The selected cell does not contain a valid date.");
                        return;
                    }

                    // Calendar-day containment (matches GLSenseExcelFunctions.FindPeriodName and
                    // GLPeriodByDateModel's period lookup): a period's stored EndDate is midnight of
                    // its last day, not the last instant of it, so comparing .Date keeps the last day
                    // included regardless of any time-of-day component on the selected date.
                    int dateIndex = Periods.FindIndex(p => p.StartDate.Date <= baseDate.Date && p.EndDate.Date >= baseDate.Date);

                    Token.ThrowIfCancellationRequested();

                    if (dateIndex < 0)
                    {
                        LogUtility.LogWarn($"PeriodsDiscoverer.{opName}: date '{baseDate:d}' does not fall within any period for ledger '{LedgerString}'.");
                        await ShowWarnMessage($"The selected date \"{baseDate:d}\" does not exists in the periods list.");
                        return;
                    }

                    // The active cell's own resolved/anchor period is the date's period
                    // shifted by whatever offset was already baked into its formula (0
                    // for a plain date). Filled cells' target periods - and any formulas
                    // written for them - are computed relative to THIS anchor, not the
                    // raw date's own period, so the fill continues the same sequence the
                    // active cell already represents.
                    BasePeriodIndex = dateIndex + baseOffsetValue;

                    if (BasePeriodIndex < 0 || BasePeriodIndex >= Periods.Count)
                    {
                        LogUtility.LogWarn($"PeriodsDiscoverer.{opName}: date '{baseDate:d}' with offset {baseOffsetValue} resolves outside the periods list for ledger '{LedgerString}'.");
                        await ShowWarnMessage($"The selected date \"{baseDate:d}\" with offset {baseOffsetValue} does not exists in the periods list.");
                        return;
                    }

                    BasePeriod = Periods[BasePeriodIndex];
                    LogUtility.LogDebug($"PeriodsDiscoverer.{opName}: base date '{baseDate:d}' (offset {baseOffsetValue}) resolved to period '{BasePeriod.PeriodName}' at index {BasePeriodIndex} of {Periods.Count} periods.");
                }
                else
                {
                    string rngValue = CellActive.Value2.ToString();

                    BasePeriod = Periods.FirstOrDefault(p => p.PeriodName == rngValue);

                    Token.ThrowIfCancellationRequested();

                    if (BasePeriod == null)
                    {
                        LogUtility.LogWarn($"PeriodsDiscoverer.{opName}: selected item \"{rngValue}\" does not exist in the periods list for ledger '{LedgerString}'.");
                        await ShowWarnMessage($"The selected item \"{rngValue}\" does not exists in the periods list.");
                        return;
                    }

                    BasePeriodIndex = Periods.FindIndex(p => p.PeriodName == rngValue);
                    LogUtility.LogDebug($"PeriodsDiscoverer.{opName}: base period '{rngValue}' resolved at index {BasePeriodIndex} of {Periods.Count} periods.");
                }

                await RunPeriodDiscovery(isByDate, dateArgument, baseOffsetValue);

                await MessageProgressWindowAsync("Excel refreshing the formulas.");
                LogUtility.LogDebug($"PeriodsDiscoverer.{opName} completed.");
            }
            catch (OperationCanceledException)
            {
                LogUtility.LogWarn($"PeriodsDiscoverer.{opName} cancelled by user.");
                await ShowCancelledAsync();
            }
            catch (Exception ex)
            {
                await HandleUnexpectedErrorAsync(ex, opName);
            }
            finally
            {
                try
                {
                    if (_ctsHelper != null && !_ctsHelper.IsCancellationRequested)
                        _ctsHelper.Cancel();

                    _ctsHelper?.Dispose();  // ? ALWAYS SAFE - handles ALL cases
                }
                catch (Exception ex)
                {
                    LogUtility.LogWarn($"{opName}: failed disposing CancellationHelper (non-fatal): {ex.Message}");
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
                LogUtility.LogDebug($"PeriodsDiscoverer.RunPeriodDiscovery: isVertical={isVertical}, isReverse={isReverse}, ledgerRef={ledgerRef}, isByDate={isByDate}, baseOffsetValue={baseOffsetValue}");
                await FillPeriodDiscoverValues(isReverse, Selection, CellActive, BasePeriodIndex, ledgerRef, Periods, isByDate, dateArgument, baseOffsetValue);
            }
            else
            {
                LogUtility.LogWarn("PeriodsDiscoverer.RunPeriodDiscovery: no ledger reference could be resolved; skipping fill.");
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
        private static bool CanProcessSelection(Range loopRng)
        {
            if (loopRng?.Cells?.Count <= 1)
                return false;

            bool isHorizontal = IsHorizontal(loopRng);
            bool isVertical = IsVertical(loopRng);

            if (!isHorizontal && !isVertical)
            {
                LogUtility.LogWarn("Selection must be a single row or a single column. Aborting to avoid unintended updates.");
                return false;
            }

            return true;
        }
        private static bool IsHorizontal(Range rng) => rng.Rows.Count == 1 && rng.Columns.Count >= 1;
        private static bool IsVertical(Range rng) => rng.Columns.Count == 1 && rng.Rows.Count >= 1;

        private static async Task FillPeriodDiscoverValues(
            bool isReverse,
            Range loopRng,
            Range formulaRange,
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
            Range baseCell = formulaRange;

            bool writeAsFormula = AddinModule.CurrentInstance.RibAsFormula.Pressed;
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
                    Range targetCell = IsHorizontal(loopRng)
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
                            // Offset is the relative move used by the GLSense_GetPeriod/GLSense_GetPeriodByDate
                            // functions. For "By Date", GLSense_GetPeriodByDate's offset argument is always
                            // relative to the raw date's own period (dateArgument never changes per target cell),
                            // so the base cell's own already-baked-in offset (baseOffsetValue) has to be added
                            // back in on top of the per-cell distance to keep the fill anchored on the period the
                            // base cell actually resolves to, not the date's period.
                            targetCell.Value = isByDate
                                ? $"=GLSense_GetPeriodByDate({dateArgument}, {ledgerRef}, {baseOffsetValue + offset})"
                                : $"=GLSense_GetPeriod({rangeRef}, {offset}, {ledgerRef})";
                        }
                    }
                    catch (Exception cellEx)
                    {
                        LogUtility.LogWarn($"PeriodsDiscoverer.FillPeriodDiscoverValues: failed writing period at offset {offset} (targetPeriodIndex={targetPeriodIndex}): {cellEx.Message}");
                        targetCell.Value = string.Empty;
                    }
                }

                // Keep async signature (optional)
                await Task.CompletedTask;
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex, "PeriodsDiscoverer.FillPeriodDiscoverValues");
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
                [
                    "GLSense_GetPeriod(",
                    "GLSense_GetPeriodByDate(",
                    "GLSense_GetPeriodStart(",
                    "GLSense_GetPeriodEnd(",
                    "GLSense_GetPeriodByYear("
                ];

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
                LogUtility.LogDebug($"PeriodsDiscoverer.BuildPeriodHelpers: LedgerString='{LedgerString}', LedgerReference='{LedgerReference}'");
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex, "PeriodsDiscoverer.BuildPeriodHelpers");
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
                LogUtility.LogException(ex, "PeriodsDiscoverer.ExtractLedgerInfo");
                LedgerString = string.Empty;
                LedgerReference = string.Empty;
            }
        }

        /// <summary>
        /// Resolves the base date for "Periods By Date" discovery from the active cell.
        /// Handles a constant date value and any formula that evaluates to a date (e.g.
        /// =DATE(...), a reference to another date cell) via the active cell's own
        /// evaluated Value2. If the active cell itself already holds a
        /// GLSense_GetPeriodByDate(...) formula, its evaluated Value2 would be a period
        /// name (not a date), so the date argument is reverse-parsed from the formula
        /// instead - mirroring how the ledger is already reverse-parsed in
        /// BuildPeriodHelpers/ExtractLedgerInfo.
        /// </summary>
        /// <param name="dateArgument">
        /// The text to embed as the date argument of a new GLSense_GetPeriodByDate
        /// formula: either the reverse-parsed cell reference / literal DATE(y,m,d) from
        /// an existing formula, or a reference to the active cell itself.
        /// </param>
        /// <param name="baseOffsetValue">
        /// The offset already baked into the active cell's own GLSense_GetPeriodByDate
        /// formula (0 for a plain date with no such formula), so the fill can anchor on
        /// the period the active cell actually resolves to, not the raw date's period.
        /// </param>
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

                        string param = parameters != null && parameters.Count > 0 ? parameters[0]?.ToString() : null;
                        string value = values != null && values.Count > 0 ? values[0]?.ToString()?.Trim() : null;

                        bool isCellRef = !string.IsNullOrWhiteSpace(param) && param.Contains("$");

                        if (!string.IsNullOrWhiteSpace(value) && TryParseDateArgument(value, out baseDate))
                        {
                            dateArgument = isCellRef ? param : value;

                            // offset is the 3rd GLSense_GetPeriodByDate argument (index 2); defaults to 0
                            // (the UDF's own DefaultOffset) when omitted or unparsable.
                            string offsetText = values != null && values.Count > 2 ? values[2]?.ToString()?.Replace("\"", "").Trim() : null;
                            if (!string.IsNullOrWhiteSpace(offsetText) && int.TryParse(offsetText, out int parsedOffset))
                                baseOffsetValue = parsedOffset;

                            LogUtility.LogDebug($"PeriodsDiscoverer.TryResolveBaseDate: resolved from existing GLSense_GetPeriodByDate formula, dateArgument='{dateArgument}', baseDate={baseDate:d}, baseOffsetValue={baseOffsetValue}.");
                            return true;
                        }
                    }
                }

                // Constant date value, or any other formula (e.g. =DATE(...), a reference to
                // another date cell) - Excel has already evaluated the cell to a numeric
                // OADate (or a formatted date string) by the time Value2 is read here,
                // regardless of whether the content is a literal or a formula. No prior
                // period offset applies here, so baseOffsetValue stays 0.
                object rawValue = CellActive.Value2;
                if (rawValue == null)
                    return false;

                baseDate = GLSenseExcelFunctions.XLLContainer.ParsePeriodDate(rawValue);
                if (baseDate == default)
                    return false;

                dateArgument = $"'{CellActive.Worksheet.Name}'!{CellActive.Address[true, true]}";
                return true;
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex, "PeriodsDiscoverer.TryResolveBaseDate");
                baseDate = default;
                dateArgument = string.Empty;
                baseOffsetValue = 0;
                return false;
            }
        }

        /// <summary>
        /// Parses a formula argument's evaluated text into a date. Handles a literal
        /// DATE(y,m,d) argument (ClsFormulaParser returns this as-is, unevaluated, when
        /// it appears nested inside another function's argument) in addition to
        /// everything XLLContainer.ParsePeriodDate already handles (OADate numerics,
        /// culture date strings, the explicit format list).
        /// </summary>
        private static bool TryParseDateArgument(string text, out DateTime date)
        {
            date = default;
            if (string.IsNullOrWhiteSpace(text))
                return false;

            Match match = Regex.Match(text, @"DATE\(\s*(\d{4})\s*,\s*(\d{1,2})\s*,\s*(\d{1,2})\s*\)", RegexOptions.IgnoreCase);
            if (match.Success)
            {
                try
                {
                    date = new DateTime(int.Parse(match.Groups[1].Value), int.Parse(match.Groups[2].Value), int.Parse(match.Groups[3].Value));
                    return true;
                }
                catch (Exception ex)
                {
                    LogUtility.LogException(ex, $"PeriodsDiscoverer.TryParseDateArgument: invalid DATE(...) literal '{text}'");
                    return false;
                }
            }

            date = GLSenseExcelFunctions.XLLContainer.ParsePeriodDate(text);
            return date != default;
        }

        //Standard helpers

        private static async Task<bool> ValidateAsync()
        {
            if (AppState.Instance.ExcelApp == null)
            {
                LogUtility.LogWarn("PeriodsDiscoverer.ValidateAsync: Excel instance unavailable.");
                await ShowErrorMessage("Unable to get excel instance.");
                return false;
            }

            if (!AppState.Instance.IsLoginCompleted)
            {
                LogUtility.LogWarn("PeriodsDiscoverer.ValidateAsync: login not completed.");
                await ShowErrorMessage("Please login to the instance.");
                return false;
            }

            if (Selection.Cells.Count == 1)
            {
                LogUtility.LogWarn("PeriodsDiscoverer.ValidateAsync: selection is a single cell.");
                await ShowWarnMessage("Selection cannot be a single cell. It must be a range of multiple cells, either vertically or horizontally.");
                return false;
            }

            if (Selection.Rows.Count > 1 && Selection.Columns.Count > 1)
            {
                LogUtility.LogWarn("PeriodsDiscoverer.ValidateAsync: selection spans multiple rows and columns.");
                await ShowWarnMessage("Selection can be multiple rows with a single column or a single row with multiple columns.");
                return false;
            }

            if (Selection.Address != null && Selection.Address.ToString().Contains(","))
            {
                LogUtility.LogWarn("PeriodsDiscoverer.ValidateAsync: selection is non-contiguous.");
                await ShowWarnMessage("Selection cannot be non-contagious.");
                return false;
            }

            if (CellActive.Value2 == null)
            {
                LogUtility.LogWarn("PeriodsDiscoverer.ValidateAsync: active cell value is empty.");
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
                MessageBoxIcon.Error,
                MessageBoxButtons.OK);
        }
        private static async Task ShowWarnMessage(string message)
        {
            await SafelyCloseWindowAsync();
            CommonFunctions.GLSenseMessage(
                message,
                MessageBoxIcon.Warning,
                MessageBoxButtons.OK);
        }
        private static async Task ShowCancelledAsync()
        {
            await SafelyCloseWindowAsync();
            CommonFunctions.GLSenseMessage(
                "Operation cancelled!",
                MessageBoxIcon.Error,
                MessageBoxButtons.OK);
        }

        private static async Task HandleUnexpectedErrorAsync(Exception ex, string opName = "FillPeriods")
        {
            LogUtility.LogException(ex, $"PeriodsDiscoverer.{opName} (unexpected)");
            await SafelyCloseWindowAsync();
            CommonFunctions.GLSenseMessage(
                $"An unexpected error occurred.{Environment.NewLine}{ex.Message}",
                MessageBoxIcon.Error,
                MessageBoxButtons.OK);
        }
        private static GLWaitWindow CreateAndShowProgressWindow(CancellationHelper cts)
        {
            try
            {
                // Use Invoke to get a return value
                return WpfAppManager.InvokeOnWpfThread(() =>
                {
                    try
                    {
                        // Use the passed-in cts, don't create a new one
                        var win = new GLWaitWindow(cts);
                        win.ShowWithOwner((IntPtr)AppState.Instance.ExcelApp.Hwnd);
                        win.StartMonitoring();
                        return win;
                    }
                    catch (Exception ex)
                    {
                        LogUtility.LogException(ex, "PeriodsDiscoverer.CreateAndShowProgressWindow (inner)");
                        return null;
                    }
                });
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex, "PeriodsDiscoverer.CreateAndShowProgressWindow");
            }
            return null;
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
                LogUtility.LogException(ex, "PeriodsDiscoverer.InitializeProgressWindowAsync");
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
                LogUtility.LogException(ex, "PeriodsDiscoverer.MessageProgressWindowAsync");
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
                LogUtility.LogException(ex, "PeriodsDiscoverer.SafelyCloseWindowAsync");
            }
            finally
            {
                ExcelWindowHelper.ActivateExcelMainWindow(GLSense.AppState.Instance.ExcelApp);
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
                LogUtility.LogException(ex, $"PeriodsDiscoverer.PModel: ledger={lName}");
                return new List<PeriodModel>();
            }
        }
        private static List<PeriodModel> LoadPeriodsForLedger(string ledgerName)
        {
            try
            {
                var dataService = ServiceLocator.PeriodDataService;
                return new List<PeriodModel>(dataService.GetPeriodsForLedger(ledgerName));
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex, $"Failed to load periods for ledger '{ledgerName}'");
                return new List<PeriodModel>();
            }
        }
    }
}
