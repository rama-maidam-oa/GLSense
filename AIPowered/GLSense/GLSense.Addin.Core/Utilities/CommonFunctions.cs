// CommonFunctions.cs in GLSense.Addin.Core
// Port of GLSense\Utilities\CommonFunctions.cs (FinalWorkingCode).
//   - BalanceFormulaExists / GetBalanceTotalRange / GetFormulaCellsWithinArea /
//     GLSenseMessage were ported first (for the RibCellHighlight template) and are
//     unchanged here.
//   - GLSenseMessage: simplified to a plain WPF MessageBox instead of the custom
//     GLMessageWindow view (that view hasn't been ported into this project yet - it's
//     a separate "port the Views" task). Swap this back to a GLMessageWindow-based
//     implementation once that view exists here.
//   - This pass adds: SheetExistsAsync, RemoveInDirect, FormulaParameters,
//     FormulaValues, MultiFormulaValues, GetBalancesCountInCells, SanitizeSheetName,
//     EnsureUniqueSheetName, SheetExists (two overloads), NameRangeExists, UnescapeXml,
//     EscapeXml, WorkbookBrokenLinks, NativeMethods.SetForegroundWindow.
//   - Re-pointed vs. the original: AppState.Instance.ExcelApp -> ServiceLocator.ExcelApp;
//     LogUtility.* -> ServiceLocator.Logger.*; ClsFormulaParser/ExcelComHelper resolve
//     via this project's Helpers namespace (both already ported).
//   - Group B pass adds: FillResponsibilitiesAsync + its private helpers
//     (ValidateTransportLevelResponse, ProcessLedgerApiResponseAsync, LogLedgerFailure,
//     TryGetRecordsNode), now that ApiHelper/ApiResponseHelper/JsonGlobals,
//     LedgerDataRepository.InsertLedgerDataAsync and AppState.Instance.LoginUrl all exist
//     in this project.
//
//   - Group F pass adds: NotValidBalancesDict + its private helpers
//     (GetExpectedArgumentCount, ValidateAllWorksheets, ValidateWorksheet, ValidateCell,
//     ValidateSingleBalanceFunction, ValidateMultipleBalanceFunctions, AddInvalid,
//     TryGetFormulaCells) - previously deferred pending DataRepository.GetTableItemsCount
//     and AppState.Instance.SelectedCube/SelectedLedger, both of which now exist. Used by
//     BalanceRefresh.ValidateAndUpdateFormulasIfNeededAsync.
//
// Deliberately NOT ported (needs a subsystem this project doesn't have at all):
//   - ConvertButtons/ConvertIcon (WinForms MessageBoxIcon/MessageBoxButtons ->
//     WPF MessageBoxImage/MessageBoxButton converters): not needed - this project's
//     GLSenseMessage already takes the WPF enums directly, no WinForms boundary exists.
using GLSense.Addin.Core.Helpers;
using GLSense.Addin.Core.Infrastructure;
using GLSense.Addin.Core.Repositories;
using GLSense.Addin.Core.Views;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text.Json;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using Excel = Microsoft.Office.Interop.Excel;

namespace GLSense.Addin.Core.Utilities
{
#nullable enable
    public static class CommonFunctions
    {
        /// <summary>
        /// Shows the custom-styled GLMessageWindow (Views\GLMessageWindow.xaml(.cs)),
        /// ported from the old monolith's GLSense\Views\GLMessageWindow.xaml(.cs) now that
        /// Views have been migrated. RESOLVED (was previously a plain WPF MessageBox.Show
        /// fallback pending this view's port - see git history/this method's old TODO).
        /// Signature/return type kept exactly as before so every existing call site across
        /// this project keeps compiling unchanged.
        /// </summary>
        public static MessageBoxResult GLSenseMessage(
                string msg,
                MessageBoxImage msgIcon,
                MessageBoxButton buttons = MessageBoxButton.OK)
        {
            // WpfAppManager.InvokeOnWpfThread only takes an Action (no return value),
            // so capture the result from inside the delegate, exactly like every other
            // InvokeOnWpfThread call site in this project that needs a return value.
            MessageBoxResult result = MessageBoxResult.None;
            WpfAppManager.InvokeOnWpfThread(() =>
            {
                var win = new GLMessageWindow(msg, msgIcon, buttons)
                {
                    CenterInExcel = true,
                    ModalToExcel = true,
                    ShowInTaskbar = false
                };

                win.ShowDialog();
                result = win.Result;
            });
            return result;
        }

        /// <summary>
        /// Fetches the ledger-setup-data payload for the given ledger/cube and persists
        /// it to SQLite via LedgerDataRepository. Called from GLCubeDetails
        /// (ProcessCubeSelectionNew/Reload) right after a cube+ledger is selected/changed.
        /// </summary>
        public static async Task FillResponsibilitiesAsync(
                 long ledgerId,
                 long cubeId,
                 CancellationToken cancellationToken)
        {
            try
            {
                string apiUrl =
                    $"{AppState.Instance.LoginUrl}/rest/secure/finance/ledger-setup-data" +
                    $"?ledgerId={ledgerId}&cubeId={cubeId}";

                cancellationToken.ThrowIfCancellationRequested();

                string responseData =
                    await ApiHelper.ServerAPI(apiUrl, "Form", "", "POST", cancellationToken);

                cancellationToken.ThrowIfCancellationRequested();

                ValidateTransportLevelResponse(responseData);

                cancellationToken.ThrowIfCancellationRequested();

                await ProcessLedgerApiResponseAsync(
                    responseData,
                    cubeId,
                    ledgerId,
                    apiUrl,
                    cancellationToken);
            }
            catch (OperationCanceledException)
            {
                throw;
            }
            catch (Exception ex)
            {
                ServiceLocator.Logger?.LogException(ex, "Error in FillResponsibilitiesAsync");
                throw; // Rethrow to allow caller to handle
            }
        }

        // ---------------------------------------------------------------------
        // Extracted validation methods (each does one thing)
        // ---------------------------------------------------------------------
        private static void ValidateTransportLevelResponse(string responseData)
        {
            if (string.IsNullOrWhiteSpace(responseData))
                throw new InvalidOperationException("Received empty response from server.");

            if (responseData.IndexOf("(401) Unauthorized", StringComparison.OrdinalIgnoreCase) >= 0)
                throw new UnauthorizedAccessException("Session expired! Please re-login.");

            if (responseData.StartsWith("Error", StringComparison.OrdinalIgnoreCase) ||
                responseData.IndexOf("<!DOCTYPE", StringComparison.OrdinalIgnoreCase) >= 0)
                throw new InvalidOperationException(responseData);
        }

        // ---------------------------------------------------------------------
        // Extracted JSON processing (the most complex part)
        // ---------------------------------------------------------------------
        private static async Task ProcessLedgerApiResponseAsync(
                    string responseData,
                    long cubeId,
                    long ledgerId,
                    string apiUrl,
                    CancellationToken cancellationToken)
        {
            try
            {
                var result =
                    ApiResponseHelper.Parse<JsonElement>(responseData, JsonGlobals.Options);

                if (!result.IsSuccess)
                {
                    LogLedgerFailure(apiUrl, responseData);

                    throw new InvalidOperationException(
                        result.ErrorMessage ?? "Ledger API returned failure status.");
                }

                JsonElement root = result.Value;

                if (!TryGetRecordsNode(root, out JsonElement recordsElem) ||
                    recordsElem.ValueKind != JsonValueKind.Object)
                {
                    throw new InvalidOperationException(
                        "Selected ledgers record count is 0");
                }

                cancellationToken.ThrowIfCancellationRequested();

                await LedgerDataRepository.InsertLedgerDataAsync(
                    cubeId,
                    ledgerId,
                    responseData);
            }
            catch (JsonException ex)
            {
                ServiceLocator.Logger?.LogWarn(
                    $"{nameof(FillResponsibilitiesAsync)}|Invalid JSON response: {responseData}");

                ServiceLocator.Logger?.LogException(ex,
                    "Failed to parse ledger API response");

                throw new InvalidOperationException(
                    "Invalid JSON response from API",
                    ex);
            }
        }

        private static void LogLedgerFailure(string apiUrl, string responseData)
        {
            string className =
                MethodBase.GetCurrentMethod()?.DeclaringType?.Name ?? "UnknownClass";

            ServiceLocator.Logger?.LogWarn(
                $"{className}.{nameof(ProcessLedgerApiResponseAsync)}|API: {apiUrl}");

            ServiceLocator.Logger?.LogWarn(
                $"{className}.{nameof(ProcessLedgerApiResponseAsync)}|Response: {responseData}");
        }

        private static bool TryGetRecordsNode(
            JsonElement root,
            out JsonElement recordsNode)
        {
            var recordProp = root.EnumerateObject()
                .FirstOrDefault(prop => string.Equals(prop.Name,
                                                     "records",
                                                     StringComparison.OrdinalIgnoreCase));

            if (recordProp.Value.ValueKind != JsonValueKind.Undefined)
            {
                recordsNode = recordProp.Value;
                return true;
            }

            recordsNode = default;
            return false;
        }

        public static bool BalanceFormulaExists(string sheetName)
        {
            if (ServiceLocator.ExcelApp == null || !AppState.Instance.IsLoginCompleted)
                return false;

            Excel.Workbook wb = ServiceLocator.ExcelApp.ActiveWorkbook;

            if (wb == null || string.IsNullOrWhiteSpace(sheetName))
                return false;

            Excel.Worksheet? ws;

            try
            {
                ws = wb.Worksheets[sheetName] as Excel.Worksheet;
            }
            catch (Exception ex)
            {
                ServiceLocator.Logger.LogException(ex, $"CommonFunctions.BalanceFormulaExists: failed accessing worksheet '{sheetName}'");
                return false;
            }

            if (ws == null)
                return false;

            Excel.Range? formulaCells = null;

            try
            {
                // Get all cells on sheet that contain formulas.
                // SpecialCells(xlCellTypeFormulas) throws if there are no such cells.
                formulaCells = ws.Cells.SpecialCells(
                    Excel.XlCellType.xlCellTypeFormulas,
                    Type.Missing);

                foreach (Excel.Range cell in formulaCells)
                {
                    bool hasFormula = (cell.HasFormula is bool hf) && hf;

                    string formula = cell.Formula as string ?? string.Empty;

                    if (hasFormula &&
                        formula.IndexOf(AppConstants.glBal, StringComparison.OrdinalIgnoreCase) >= 0)
                    {
                        return true;
                    }
                }
            }
            catch (COMException)
            {
                // Thrown when there are no formula cells on the sheet; treat as "not found".
                return false;
            }
            catch (Exception ex)
            {
                ServiceLocator.Logger.LogException(ex, $"CommonFunctions.BalanceFormulaExists: failed scanning formula cells on '{sheetName}'");
                return false;
            }
            finally
            {
                if (formulaCells != null)
                {
                    Marshal.ReleaseComObject(formulaCells);
                }
            }

            return false;
        }

        public static Excel.Range? GetBalanceTotalRange(string rngAddress)
        {
            if (string.IsNullOrWhiteSpace(rngAddress))
                return null;

            var res = ExcelExternalRef.ResolveRangeWithContext(rngAddress);
            if (res?.Range == null)
                return null;

            Excel.Range root = res.Range;
            Excel.Application app = root.Application;
            Excel.Range? totalRange = null;

            try
            {
                // Iterate each Area to handle multi-area input robustly.
                foreach (Excel.Range area in root.Areas)
                {
                    Excel.Range? formulasInArea = GetFormulaCellsWithinArea(area);

                    if (formulasInArea == null)
                        continue;

                    // Iterate the resulting cells and filter by the formula token
                    foreach (Excel.Range cell in formulasInArea.Cells)
                    {
                        if (cell.Formula is not string f)
                            continue;
                        if (!string.IsNullOrEmpty(f) &&
                            f?.IndexOf(AppConstants.glBal, StringComparison.OrdinalIgnoreCase) >= 0)
                        {
                            totalRange = totalRange == null ? cell : app.Union(totalRange, cell);
                        }
                    }
                }
            }
            catch (COMException ex)
            {
                ServiceLocator.Logger.LogException(ex, $"COM error in GetBalanceTotalRange for address: {rngAddress}");
                totalRange = null;
            }
            catch (Exception ex)
            {
                ServiceLocator.Logger.LogException(ex, $"Error in GetBalanceTotalRange for address: {rngAddress}");
                totalRange = null;
            }

            return totalRange;
        }

        private static Excel.Range? GetFormulaCellsWithinArea(Excel.Range area)
        {
            // Single-cell path: avoid SpecialCells quirk by using HasFormula
            if (area.Rows.Count == 1 && area.Columns.Count == 1)
            {
                try
                {
                    return (bool)area.HasFormula ? area : null;
                }
                catch (COMException ex)
                {
                    ServiceLocator.Logger?.LogWarn($"GetFormulaCellsWithinArea: HasFormula threw for a single-cell area, treating as not-a-formula-cell: {ex.Message}");
                    return null;
                }
            }

            // Multi-cell path: use SpecialCells but constrain using Intersect
            try
            {
                var cand = area.SpecialCells(Excel.XlCellType.xlCellTypeFormulas);
                // Belt-and-suspenders: ensure nothing leaked from outside the area
                var app = area.Application;
                var limited = app.Intersect(cand, area);
                return limited; // can be null if none matched inside the area
            }
            catch (COMException)
            {
                // Thrown if no formula cells exist in this area
                return null;
            }
        }

        /// <summary>
        /// Checks whether a worksheet named <paramref name="shtname"/> exists in the
        /// active workbook. Kept as Task-returning to match the original call sites,
        /// but the body runs synchronously on the calling thread - Excel COM objects
        /// are apartment-affinitized and must not be touched from a ThreadPool thread.
        /// </summary>
        public static async Task<bool> SheetExistsAsync(string shtname)
        {
            try
            {
                var wb = ServiceLocator.ExcelApp?.ActiveWorkbook;
                if (wb == null) return false;

                foreach (Excel.Worksheet worksheet in wb.Worksheets)
                {
                    try
                    {
                        if (worksheet.Name == shtname)
                        {
                            return true;
                        }
                    }
                    finally
                    {
                        ExcelComHelper.SafeRelease(worksheet, "Worksheet");
                    }
                }
                await Task.CompletedTask;
                return false;
            }
            catch (Exception ex)
            {
                ServiceLocator.Logger?.LogException(ex, $"CommonFunctions.SheetExistsAsync(shtname='{shtname}'): failed");
                return false;
            }
        }

        /// <summary>
        /// Resolves a range address that may be wrapped in INDIRECT("...") to the
        /// underlying Excel.Range, stripping the INDIRECT(...) wrapper and quotes.
        /// </summary>
        public static Excel.Range? RemoveInDirect(string strAddress)
        {
            try
            {
                Excel.Range? rng = null;

                if (ServiceLocator.ExcelApp == null || string.IsNullOrEmpty(strAddress))
                {
                    return rng;
                }

                string sRangeAddress = string.Empty;

                if (strAddress.IndexOf("INDIRECT", StringComparison.OrdinalIgnoreCase) == -1)
                {
                    rng = ServiceLocator.ExcelApp.Range[strAddress.Replace('"'.ToString(), "")];
                }
                else
                {
                    try
                    {
                        sRangeAddress = strAddress.Replace('"'.ToString(), "");
                        int openParenIndex = sRangeAddress.IndexOf("(") + 1;
                        sRangeAddress = sRangeAddress.Substring(openParenIndex);
                        sRangeAddress = sRangeAddress.Substring(0, sRangeAddress.Length - 1);
                        rng = ServiceLocator.ExcelApp.Range[sRangeAddress];
                    }
                    catch (Exception ex)
                    {
                        ServiceLocator.Logger?.LogException(ex, $"CommonFunctions.RemoveInDirect: failed resolving INDIRECT range from '{strAddress}'");
                    }
                }

                return rng;
            }
            catch (Exception ex)
            {
                ServiceLocator.Logger?.LogException(ex, $"CommonFunctions.RemoveInDirect(strAddress='{strAddress}'): failed");
                return null;
            }
        }

        /// <summary>Raw (unresolved) argument tokens of a GLSense_GetBalance(...) formula.</summary>
        public static List<string>? FormulaParameters(string formula)
        {
            try
            {
                var fncparser = new ClsFormulaParser(formula);
                return fncparser.FormulaArgs();
            }
            catch (Exception ex)
            {
                ServiceLocator.Logger?.LogException(ex, $"CommonFunctions.FormulaParameters(formula='{formula}'): failed");
                return null;
            }
        }

        /// <summary>Argument tokens with any cell-reference ($A$1-style) arguments resolved to their values.</summary>
        public static List<string>? FormulaValues(string formula)
        {
            try
            {
                var fncparser = new ClsFormulaParser(formula);
                return fncparser.FormulaArgs_Values();
            }
            catch (Exception ex)
            {
                ServiceLocator.Logger?.LogException(ex, $"CommonFunctions.FormulaValues(formula='{formula}'): failed");
                return null;
            }
        }

        /// <summary>
        /// Used when a single cell contains more than one GLSense_GetBalance(...) call:
        /// returnType selects which slice of a specific call to extract
        /// ("Functions" = every function call text; "Arguments" = raw args of one call;
        /// "Arguments_WithValues" = same, with cell references resolved).
        /// </summary>
        public static List<string>? MultiFormulaValues(string formula, string returnType)
        {
            try
            {
                if (string.IsNullOrWhiteSpace(formula) || string.IsNullOrWhiteSpace(returnType))
                    return null;

                var fncparser = new ClsFormulaParser(formula);

                return returnType switch
                {
                    "Functions" => fncparser.ExtractFunctions(),
                    "Arguments" => ClsFormulaParser.ExtractArguments(formula),
                    "Arguments_WithValues" => ClsFormulaParser.ExtractArguments_WithValues(formula),
                    _ => null,
                };
            }
            catch (Exception ex)
            {
                ServiceLocator.Logger?.LogException(ex, $"CommonFunctions.MultiFormulaValues(returnType='{returnType}'): failed");
                return null;
            }
        }

        public static int GetBalancesCountInCells(string input)
        {
            try
            {
                return Regex.Matches(input, Regex.Escape(AppConstants.glBal)).Count;
            }
            catch (Exception ex)
            {
                ServiceLocator.Logger?.LogException(ex, $"Failed in counting GLsense_GetBalance in given string {input}");
                return 0;
            }
        }

        /// <summary>
        /// Group F pass - ported from GLSense\Utilities\CommonFunctions.cs (FinalWorkingCode).
        /// Was deliberately deferred out of the earlier CommonFunctions.cs pass (see this
        /// file's header comment) pending DataRepository.GetTableItemsCount +
        /// AppState.Instance.SelectedCube/SelectedLedger, both of which now exist. Used by
        /// BalanceRefresh.ValidateAndUpdateFormulasIfNeededAsync to find balance formulas
        /// with fewer arguments than the current segment configuration expects (i.e. stale
        /// formulas written before a segment was added), so they can be offered an update.
        /// Re-pointed vs. the original: LogUtility.* -> ServiceLocator.Logger?.*;
        /// AppState.Instance.ExcelApp -> not needed here (ExcelExternalRef.
        /// ResolveRangeWithContext already resolves via ServiceLocator.ExcelApp).
        /// </summary>
        public static Dictionary<string, string> NotValidBalancesDict(string rngAddress)
        {
            ServiceLocator.Logger?.LogDebug("Balance validation started.");

            var result = new Dictionary<string, string>();

            var resolveResult = ExcelExternalRef.ResolveRangeWithContext(rngAddress);
            if (resolveResult.Workbook == null)
                return result;

            int expectedArgCount = GetExpectedArgumentCount();
            if (expectedArgCount == 0)
            {
                ServiceLocator.Logger?.LogWarn("Failed in fetching segments count for validation!");
                return result;
            }

            try
            {
                ValidateAllWorksheets(resolveResult.Workbook, expectedArgCount, result);
            }
            catch (Exception ex)
            {
                ServiceLocator.Logger?.LogException(ex, "Error while checking not valid balances in workbook.");
            }
            finally
            {
                ServiceLocator.Logger?.LogDebug("Balance validation completed.");
            }

            return result;
        }

        // Extracted: Compute expected argument count (9 + segments + 1)
        private static int GetExpectedArgumentCount()
        {
            int segmentsCount = DataRepository.GetTableItemsCount(
                AppState.Instance.SelectedCube.CubeId,
                AppState.Instance.SelectedLedger.LedgerId,
                "SEGMENTS");

            return segmentsCount > 0 ? 9 + segmentsCount + 1 : 0;
        }

        // Extracted: Main loop over worksheets
        private static void ValidateAllWorksheets(Excel.Workbook wb, int expectedArgCount, Dictionary<string, string> result)
        {
            foreach (Excel.Worksheet ws in wb.Worksheets)
            {
                ValidateWorksheet(ws, expectedArgCount, result);
            }
        }

        // Extracted: Validation logic for a single worksheet
        private static void ValidateWorksheet(Excel.Worksheet ws, int expectedArgCount, Dictionary<string, string> result)
        {
            Excel.Range? formulaCells = TryGetFormulaCells(ws);
            if (formulaCells == null)
                return;

            ServiceLocator.Logger?.LogDebug($"Validating balances for worksheet {ws.Name}.");

            foreach (Excel.Range cell in formulaCells)
            {
                ValidateCell(cell, expectedArgCount, result);
            }

            ServiceLocator.Logger?.LogDebug($"Validating balances for worksheet {ws.Name} completed.");
        }

        // Extracted: Validation logic for a single cell
        private static void ValidateCell(Excel.Range cell, int expectedArgCount, Dictionary<string, string> result)
        {
            string? cellFormula = cell.Formula?.ToString();
            if (string.IsNullOrWhiteSpace(cellFormula))
                return;

            if (cellFormula?.IndexOf(AppConstants.glBal, StringComparison.OrdinalIgnoreCase) < 0)
                return;

            if (cellFormula == null)
                return;

            int balanceCount = GetBalancesCountInCells(cellFormula);

            if (balanceCount == 1)
            {
                ValidateSingleBalanceFunction(cellFormula, expectedArgCount, cell, result);
            }
            else if (balanceCount > 1)
            {
                ValidateMultipleBalanceFunctions(cellFormula, expectedArgCount, cell, result);
            }
        }

        // Extracted: Single GLSENSE_GETBALANCE(...) call
        private static void ValidateSingleBalanceFunction(string formula, int expectedArgCount, Excel.Range cell, Dictionary<string, string> result)
        {
            var parser = new ClsFormulaParser(formula);
            int actualArgs = parser.Formula_ArgsCount();

            if (actualArgs <= expectedArgCount)
            {
                AddInvalid(result, cell, formula);
            }
        }

        // Extracted: Multiple GLSENSE_GETBALANCE(...) calls in one cell
        private static void ValidateMultipleBalanceFunctions(string cellFormula, int expectedArgCount, Excel.Range cell, Dictionary<string, string> result)
        {
            var functions = MultiFormulaValues(cellFormula, "Functions");
            if (functions == null || functions.Count == 0)
                return;

            foreach (var func in functions)
            {
                var args = MultiFormulaValues(func, "Arguments");
                if (args == null)
                    continue;

                if (args.Count <= expectedArgCount)
                {
                    AddInvalid(result, cell, cellFormula);
                    break; // No need to check further — cell is already invalid
                }
            }
        }

        // Optional: Keep AddInvalid as-is (it's simple)
        private static void AddInvalid(Dictionary<string, string> dict, Excel.Range cell, string formula)
        {
            string address = cell.Address[false, false];
            dict[address] = formula;
        }

        private static Excel.Range? TryGetFormulaCells(Excel.Worksheet ws)
        {
            if (ws == null) return null;

            try
            {
                Excel.Range rng = ws.Cells.SpecialCells(Excel.XlCellType.xlCellTypeFormulas);

                // Some versions of Excel return a single cell as Range with Count == 1
                // So we only return null if truly no formulas
                return (rng.Count > 0) ? rng : null;
            }
            catch (COMException ex) when (ex.HResult == -2146827284) // 0x800A03EC = No cells found
            {
                return null;
            }
            catch (Exception ex)
            {
                ServiceLocator.Logger?.LogException(ex, "Unexpected error in TryGetFormulaCells");
                return null;
            }
        }

        /// <summary>
        /// Sanitizes a worksheet name: removes illegal characters and truncates to Excel's 31 char limit.
        /// Illegal characters: : \ / ? * [ ]
        /// </summary>
        public static string SanitizeSheetName(string raw, Excel.Workbook? wb = null)
        {
            // Regression fix: this used to eagerly fetch ServiceLocator.ExcelApp.ActiveWorkbook
            // unconditionally, even though `wb` is only actually used in the (rare) empty-`raw`
            // branch immediately below. Every real caller (e.g. DrilldownBl.BuildDrilldownInfo,
            // called from the SheetBeforeDoubleClick drilldown chain) always passes a non-empty
            // `raw`, so that fetch ran for no reason on every single call - and specifically in
            // the double-click drilldown's async call chain, that ActiveWorkbook fetch crosses
            // the host<->Addin.Core AppDomain boundary via .NET Remoting and was throwing
            // "InvalidCastException: Unable to cast object of type 'System.__ComObject' to type
            // 'Microsoft.Office.Interop.Excel.WorkbookClass'" (see
            // GLSense_Logs_22-Jul-2026.log, Context: DrilldownBl.Balance_Drilldown), aborting
            // the drilldown. Moving the fetch inside the branch that actually needs `wb`
            // eliminates that unnecessary (and, in this call path, crash-prone) cross-domain
            // COM marshaling entirely for the common non-empty-name case.
            if (string.IsNullOrWhiteSpace(raw))
            {
                wb ??= ServiceLocator.ExcelApp?.ActiveWorkbook;
                return wb != null ? EnsureUniqueSheetName(wb, "Sheet") : "Sheet";
            }

            string name = raw.Trim();

            // Replace illegal characters
            char[] illegal = { ':', '\\', '/', '?', '*', '[', ']' };
            foreach (char ch in illegal)
                name = name.Replace(ch, '_');

            // Excel limits worksheet names to 29 chars reserve 2 chars (31 chars)
            if (name.Length > 29)
                name = name.Substring(0, 29);

            // Avoid empty after sanitization
            if (string.IsNullOrWhiteSpace(name))
                name = "Sheet";

            //removing this part as this is creating duplicate sheets for drilldown data
            //name = EnsureUniqueSheetName(wb, name);

            return name;
        }

        /// <summary>
        /// Ensures uniqueness by appending " (n)" while respecting the 31-character limit.
        /// </summary>
        private static string EnsureUniqueSheetName(Excel.Workbook wb, string baseName)
        {
            string name = baseName;
            int counter = 1;

            // Trim to allow room for suffix " (n)"
            string suffix = $" ({counter})";
            int maxBaseLen = Math.Max(0, 31 - suffix.Length);
            string trimmedBase = baseName.Length > maxBaseLen ? baseName.Substring(0, maxBaseLen) : baseName;

            while (SheetExists(wb, name))
            {
                suffix = $" ({counter})";
                int allowed = 31 - suffix.Length;
                string head = trimmedBase.Length > allowed ? trimmedBase.Substring(0, allowed) : trimmedBase;
                name = head + suffix;
                counter++;
            }

            return name;
        }

        private static bool SheetExists(Excel.Workbook wb, string sheetName)
        {
            if (wb == null || string.IsNullOrEmpty(sheetName)) return false;
            foreach (Excel.Worksheet ws in wb.Worksheets)
            {
                if (string.Equals(ws.Name, sheetName, StringComparison.OrdinalIgnoreCase))
                    return true;
            }
            return false;
        }

        /// <summary>Checks whether a worksheet named <paramref name="shtname"/> exists in the active workbook.</summary>
        public static bool SheetExists(string shtname)
        {
            try
            {
                if (string.IsNullOrWhiteSpace(shtname))
                    return false;

                var wb = ServiceLocator.ExcelApp?.ActiveWorkbook;
                if (wb == null)
                    return false;

                foreach (Excel.Worksheet ws in wb.Worksheets)
                {
                    if (string.Equals(ws.Name, shtname, StringComparison.OrdinalIgnoreCase))
                        return true;
                }

                return false;
            }
            catch (Exception ex)
            {
                ServiceLocator.Logger?.LogException(ex, $"CommonFunctions.SheetExists(shtname='{shtname}'): failed");
                return false;
            }
        }

        public static bool NameRangeExists(string rngvalue)
        {
            try
            {
                if (string.IsNullOrWhiteSpace(rngvalue))
                    return false;

                var wb = ServiceLocator.ExcelApp?.ActiveWorkbook;
                if (wb == null)
                    return false;

                Excel.Names names = wb.Names;
                if (names == null || names.Count == 0)
                    return false;

                foreach (Excel.Name nm in names)
                {
                    // Case-insensitive match, same semantics as VB's ToUpper comparison
                    if (string.Equals(nm.Name, rngvalue, StringComparison.OrdinalIgnoreCase))
                        return true;
                }

                return false;
            }
            catch (Exception ex)
            {
                ServiceLocator.Logger?.LogException(ex, $"CommonFunctions.NameRangeExists(rngvalue='{rngvalue}'): failed");
                return false;
            }
        }

        public static string UnescapeXml(string s)
        {
            var unxml = s;

            try
            {
                if (!string.IsNullOrEmpty(unxml))
                {
                    // Note: these are intentionally double-encoded entities (e.g., &amp;amp;)
                    unxml = unxml.Replace("&amp;amp;", "&amp;");
                    unxml = unxml.Replace("&amp;apos;", "'");
                    unxml = unxml.Replace("&amp;quot;", "\"");
                    unxml = unxml.Replace("&amp;gt;", "&gt;");
                    unxml = unxml.Replace("&amp;lt;", "&lt;");
                }

                return unxml;
            }
            catch (Exception ex)
            {
                ServiceLocator.Logger?.LogException(ex, "CommonFunctions.UnescapeXml: failed");
                return string.Empty;
            }
        }

        public static string EscapeXml(string s)
        {
            var toxml = s;

            try
            {
                if (!string.IsNullOrEmpty(toxml))
                {
                    // Intentionally producing double-encoded entities (e.g., & -> &amp;amp;)
                    toxml = toxml.Replace("&", "&amp;amp;");
                    toxml = toxml.Replace("'", "&amp;apos;");
                    toxml = toxml.Replace("\"", "&amp;quot;");
                    toxml = toxml.Replace("&gt;", "&amp;gt;");
                    toxml = toxml.Replace("&lt;", "&amp;lt;");
                }

                return toxml;
            }
            catch (Exception ex)
            {
                ServiceLocator.Logger?.LogException(ex, "CommonFunctions.EscapeXml: failed");
                return string.Empty;
            }
        }

        public static string WorkbookBrokenLinks()
        {
            try
            {
                var wb = ServiceLocator.ExcelApp?.ActiveWorkbook;
                if (wb == null) return string.Empty;

                // LinkSources returns object (usually object[] of strings), or null if none.
                object? linksObj;
                try
                {
                    linksObj = wb.LinkSources(Excel.XlLink.xlExcelLinks);
                }
                catch
                {
                    // If Excel throws (e.g., no links), treat as none.
                    linksObj = null;
                }

                if (linksObj == null) return string.Empty;

                var broken = new List<string>();

                // Safely iterate over the returned COM array of link strings.
                foreach (var link in (linksObj as Array)?.Cast<object>().Select(o => o as string) ?? Enumerable.Empty<string?>())
                {
                    if (string.IsNullOrWhiteSpace(link))
                    {
                        // Consider empty/invalid entries broken
                        broken.Add(link ?? string.Empty);
                        continue;
                    }

                    try
                    {
                        // Test if the link target exists on disk (for file-based links).
                        // Note: For URL links, File.Exists will return false; you may extend with URI checks if needed.
                        if (!File.Exists(link) && link != null)
                        {
                            broken.Add(link);
                        }
                    }
                    catch
                    {
                        // In case of an invalid path or other IO issues, mark as broken.
                        broken.Add(link ?? string.Empty);
                    }
                }

                if (broken.Count == 0) return string.Empty;

                // Format: "1.) <link>\r\n2.) <link>..."
                var formatted = string.Join(
                    Environment.NewLine,
                    broken.Select((link, index) => $"{index + 1}.) {link}")
                );
                if (!string.IsNullOrWhiteSpace(formatted))
                {
                    ServiceLocator.Logger?.LogWarn($"Broken Link Found: {formatted}");
                    return formatted;
                }
                else
                {
                    return string.Empty;
                }
            }
            catch (Exception ex)
            {
                ServiceLocator.Logger?.LogException(ex, "Exception in getting workbook broken links.");
                return string.Empty;
            }
        }

        internal static class NativeMethods
        {
            [DllImport("user32.dll")]
            internal static extern bool SetForegroundWindow(IntPtr hWnd);
        }
    }
#nullable disable
}
