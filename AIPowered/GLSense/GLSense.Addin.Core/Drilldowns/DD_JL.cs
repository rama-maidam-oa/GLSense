// DD_JL.cs in GLSense.Addin.Core
// Port of GLSense\Drilldowns\DD_JL.cs (FinalWorkingCode), class DrilldownJl.
// DrilldownJl does NOT share a base class with DrilldownBl/DrilldownSl in the old code
// (confirmed via grep across FinalWorkingCode\GLSense for ": Drilldown" inheritance -
// none found); each is its own independent class with parallel but separately
// implemented REST-call/progress-window/DDDatatoWorksheet logic. Ported that way here too.
//
// Re-pointed vs. the original (business logic/REST URL/payload shape/response handling
// unchanged):
//   - GLSense.Helpers/.Models/.Utilities/.Views -> GLSense.Addin.Core.* equivalents.
//   - LogUtility.* (static) -> ServiceLocator.Logger?.* (instance via context).
//   - AppState.Instance.ExcelApp -> ServiceLocator.ExcelApp (this project's AppState has
//     no ExcelApp field).
//   - System.Windows.Forms MessageBoxIcon/MessageBoxButtons -> System.Windows
//     MessageBoxImage/MessageBoxButton (this project has no WinForms reference).
//   - GLWaitWindow now derives from BaseWindow: win.ShowWithOwner(hwnd) -> win.Show()
//     (Excel owner set automatically via ServiceLocator.ExcelHandle). CreateAndShow-
//     ProgressWindow rewritten to the WpfAppManager.InvokeOnWpfThread(Action)-with-
//     captured-local pattern (InvokeOnWpfThread has no Func<T> overload here).
using GLSense.Addin.Core.Helpers;
using GLSense.Addin.Core.Infrastructure;
using GLSense.Addin.Core.Models;
using GLSense.Addin.Core.Utilities;
using GLSense.Addin.Core.Views;
using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Text.Json;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using Excel = Microsoft.Office.Interop.Excel;

namespace GLSense.Addin.Core.Drilldowns
{
#nullable enable
    public class DrilldownJl
    {
        private Excel.Application ExcelApp { get; }
        private Excel.Workbook? JlWorbook { get; set; }
        private Excel.Worksheet? JlWorksheet { get; set; }
        private Excel.Range? JlRange { get; set; }
        private CancellationHelper? _ctsHelper;
        private CancellationToken Token => _ctsHelper?.GetToken() ?? default;
        private ExternalResolveResult ExternalResolveResult { get; set; } = new ExternalResolveResult();
        private readonly string _jlAddress;

        // Restored during the end-to-end completeness pass: the old monolith's ddType
        // constructor parameter (used by RibJournalDD -> "JL", and by
        // RibBalancesDDToSubLedger/RibBalancesDDToUnified -> "BLDD_SL"/"BLDD_UF", which
        // route the SAME journal-drilldown-shaped payload to different REST endpoints -
        // see BuildApiUrl below) was dropped when this class was first ported, hardcoding
        // "JL" everywhere. The 2-arg constructor is kept as a convenience defaulting to
        // "JL" so the existing RunJournalDrilldown()/double-click-dispatch call sites are
        // unaffected.
        private readonly string _ddType;

        private static GLWaitWindow? Win { get; set; }

        private static readonly string[] JournalHeaders =
        {
            "Period_Name",        // 0
            "Actual_Flag",        // 1
            "Budget_Version_Id",  // 2  (not used in your object)
            "Encumbrance_Type_Id",// 3  (not used in your object)
            "Ccid",               // 4  -> codeCombinationId (long)
            "Ledger_Id",          // 5  -> ledgerId (long)
            "Balance_Type",       // 6
            "Currency_Code",      // 7
            "Translated_Flag",    // 8
            "Source_Name",        // 9
            "Category_Name",      // 10
            "Status",             // 11
            "StartDate",          // 12
            "EndDate"             // 13
        };

        public DrilldownJl(Excel.Application excelApp, string rngAddress)
            : this(excelApp, rngAddress, "JL")
        {
        }

        public DrilldownJl(Excel.Application excelApp, string rngAddress, string ddType)
        {
            ExcelApp = excelApp ?? throw new ArgumentNullException(nameof(excelApp));
            _jlAddress = rngAddress;
            _ddType = string.IsNullOrEmpty(ddType) ? "JL" : ddType;
        }

        public async Task ProcessJLDrilldown()
        {
            _ctsHelper = new CancellationHelper();

            ServiceLocator.Logger?.LogDebug($"DrilldownJl.ProcessJLDrilldown started. ddType={_ddType}, address={_jlAddress}");

            try
            {
                ExternalResolveResult = ExcelExternalRef.ResolveRangeWithContext(_jlAddress);
                JlRange = ExternalResolveResult.Range;

                if (JlRange == null)
                {
                    return;
                }

                JlWorksheet = ExternalResolveResult.Worksheet;
                JlWorbook = ExternalResolveResult.Workbook;

                CommonMethods.DisableExcelSettings();

                if (!IsValidSingleColumnSelection(JlRange))
                    return;

                if (!HasAnyValue(JlRange))
                    return;

                Win = CreateAndShowProgressWindow(_ctsHelper);

                try
                {
                    await InitializeProgressWindowAsync("Journals Drilldown", "Processing request...");

                    string headerText = TryGetColumnHeaderText(JlWorksheet, JlRange);
                    if (!await TryRunDrilldownAsync(headerText, JlWorksheet, JlRange))
                        return;
                }
                catch (Exception ex)
                {
                    await HandleUnexpectedErrorAsync(ex);
                }
                finally
                {
                    await SafelyCloseWindowAsync();
                }
            }
            catch (OperationCanceledException)
            {
                await ShowCancelledAsync();
                ServiceLocator.Logger?.LogWarn("JL Drilldown operation was cancelled by the user.");
            }
            catch (Exception ex)
            {
                ServiceLocator.Logger?.LogException(ex);
            }
            finally
            {
                CommonMethods.TryEnableExcelSettings("DrilldownJl.ProcessJLDrilldown");
            }
        }

        private static bool IsValidSingleColumnSelection(Excel.Range rng)
        {
            if (rng == null || rng.Columns.Count >= 2)
            {
                CommonFunctions.GLSenseMessage(
                    "Cannot fetch data for multiple column selections!"
                    + Environment.NewLine
                    + "Can be multiple rows with a single column.",
                    MessageBoxImage.Exclamation,
                    MessageBoxButton.OK);
                return false;
            }

            return true;
        }

        private bool HasAnyValue(Excel.Range rng)
        {
            try
            {
                int cnt = (int)ExcelApp.WorksheetFunction.CountA(rng);
                if (cnt == 0)
                {
                    CommonFunctions.GLSenseMessage(
                        "Current selection is empty!",
                        MessageBoxImage.Exclamation,
                        MessageBoxButton.OK);
                    return false;
                }

                return true;
            }
            catch (Exception ex)
            {
                ServiceLocator.Logger?.LogException(ex);
                return false;
            }
        }

        private static GLWaitWindow? CreateAndShowProgressWindow(CancellationHelper cts)
        {
            try
            {
                GLWaitWindow? win = null;

                // WpfAppManager.InvokeOnWpfThread only takes an Action (no return value),
                // so capture the created window from inside the delegate - same pattern
                // SegmentDiscoverer.cs/DrillCellHighlighter.cs use for GLWaitWindow.
                WpfAppManager.InvokeOnWpfThread(() =>
                {
                    try
                    {
                        win = new GLWaitWindow(cts);
                        win.Show();
                        win.StartMonitoring();
                    }
                    catch (Exception ex)
                    {
                        ServiceLocator.Logger?.LogException(ex);
                        win = null;
                    }
                });

                return win;
            }
            catch (Exception ex)
            {
                ServiceLocator.Logger?.LogException(ex);
            }
            return null;
        }

        private static Task InitializeProgressWindowAsync(string title, string message)
        {
            if (Win == null)
                return Task.CompletedTask;

            // Fire-and-forget: this only updates progress UI. Callers must not be
            // resumed on an arbitrary ThreadPool thread after awaiting this dispatch,
            // since immediately-following code reads Excel COM objects directly.
            _ = Win.Dispatcher.InvokeAsync(() =>
            {
                Win.SetProcessTitle(title);
                Win.SetProcessMessage(message);
            });
            return Task.CompletedTask;
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
                await Win.Dispatcher.InvokeAsync(() => Win.RequestClose());
            }
            catch (Exception ex)
            {
                ServiceLocator.Logger?.LogError($"Error closing wait window: {ex.Message}");
            }
        }

        private async Task<bool> TryRunDrilldownAsync(
            string headerText,
            Excel.Worksheet wrksheet,
            Excel.Range rng)
        {
            string? key = headerText?.Trim();

            if (string.IsNullOrEmpty(key))
            {
                ServiceLocator.Logger?.LogDebug("DrilldownJl.TryRunDrilldownAsync: header text empty, invalid selection.");
                return ShowInvalidSelection();
            }

            string ddType = GetDrilldownType(key);
            if (string.IsNullOrEmpty(ddType))
            {
                ServiceLocator.Logger?.LogDebug($"DrilldownJl.TryRunDrilldownAsync: header key '{key}' did not map to a known drilldown type, invalid selection.");
                return ShowInvalidSelection();
            }

            ServiceLocator.Logger?.LogDebug($"DrilldownJl.TryRunDrilldownAsync: resolved amount type '{ddType}' from header '{key}'.");
            await Journal_DrillDown(wrksheet, rng, ddType);
            return true;
        }

        private static string GetDrilldownType(string? key)
        {
            if (key == null)
                return string.Empty;

            if (key.Equals("ptd_net", StringComparison.OrdinalIgnoreCase))
                return "PTD";

            if (key.Equals("ytd_net", StringComparison.OrdinalIgnoreCase))
                return "YTD";

            if (key.Equals("quarter_to_date_net", StringComparison.OrdinalIgnoreCase) ||
                key.Equals("qtd_net", StringComparison.OrdinalIgnoreCase))
                return "QTD";

            var jlTypes = new[] { "ENTERED_DR", "ENTERED_CR", "ENTERED_NET", "ACCOUNTED_DR", "ACCOUNTED_CR", "ACCOUNTED_NET" };

            if (!string.IsNullOrEmpty(key) &&
                jlTypes.Any(valid => valid.Equals(key, StringComparison.OrdinalIgnoreCase)))
            {
                return "JED";
            }

            return string.Empty;
        }

        private static bool ShowInvalidSelection()
        {
            Win?.Dispatcher.InvokeAsync(() => Win.RequestClose());
            CommonFunctions.GLSenseMessage(
                $"Invalid selection. Please select values of either PTD_NET or YTD_NET or QUARTER_TO_DATE_NET" +
                $"ENTERED_DR or ENTERED_CR or ENTERED_NET or ACCOUNTED_DR or ACCOUNTED_CR or ACCOUNTED_NET for journals drilldown!",
                MessageBoxImage.Error,
                MessageBoxButton.OK);
            return false;
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
            ServiceLocator.Logger?.LogException(ex);
            await SafelyCloseWindowAsync();
            CommonFunctions.GLSenseMessage(
                $"An unexpected error occurred.{Environment.NewLine}{ex.Message}",
                MessageBoxImage.Error,
                MessageBoxButton.OK);
        }

        private async Task Journal_DrillDown(
            Excel.Worksheet journalsSheet,
            Excel.Range journalsRange,
            string amtType)
        {
            try
            {
                ServiceLocator.Logger?.LogDebug($"DrilldownJl.Journal_DrillDown started. amtType={amtType}, address={journalsRange.get_Address()}");

                int headerIndex = TryGetColumnHeaderIndex(journalsSheet, amtType);
                if (headerIndex == -1)
                {
                    ServiceLocator.Logger?.LogWarn($"DrilldownJl.Journal_DrillDown: could not find header column for amtType={amtType} in sheet '{journalsSheet.Name}'.");
                    await ShowHeaderNotFoundAsync(amtType);
                    return;
                }

                string journalsAddress = journalsRange.get_Address(
                    RowAbsolute: true,
                    ColumnAbsolute: true,
                    ReferenceStyle: XlReferenceStyle.xlA1,
                    External: true);

                List<JournalsQuerySubmit> drilldownList = BuildDrilldownList(journalsSheet, journalsRange, headerIndex, out int selectedCount, out string lastValue, out string[] lastJournalValues);

                ServiceLocator.Logger?.LogDebug($"DrilldownJl.Journal_DrillDown: built {drilldownList.Count} drilldown row(s), selectedCount={selectedCount}.");

                var payload = new JournalDD { journalDrilldowns = drilldownList.ToArray() };
                string httpPostText = JsonSerializer.Serialize(payload, JsonGlobals.Options);

                string worksheetName = BuildWorksheetName(journalsSheet, selectedCount);
                worksheetName = NormalizeWorksheetName(worksheetName);

                // More than one row selected -> "Multi Select.". Parsing lastValue/lastJournalValues
                // (only the LAST row's fields) here would be misleading, since each selected row can
                // carry different period/ledger/currency/etc. values - same reasoning DrilldownSl.
                // Subledger_DrillDown already applies ("Multi Select" when strBuilder.Count >= 2) and
                // DrilldownBl.GetDrilldownTitle applies ("Multi Select." when totalCount >= 2). Only
                // parse the individual field values when exactly one row is selected.
                string multiString = selectedCount >= 2
                    ? "Multi Select."
                    : BuildMultiStringSafe(lastValue, lastJournalValues);

                string consolidated = BuildConsolidatedObject(worksheetName, journalsAddress, _ddType, multiString);
                string apiUrl = BuildApiUrl(consolidated);

                await SendRequestAndHandleResponseAsync(apiUrl, httpPostText, consolidated);
            }
            catch (Exception ex)
            {
                ServiceLocator.Logger?.LogException(ex, "DrilldownJl.Journal_DrillDown");
            }
        }

        private static async Task ShowHeaderNotFoundAsync(string amtType)
        {
            if (Win == null)
                return;
            await Win.Dispatcher.InvokeAsync(() => Win.RequestClose());
            CommonFunctions.GLSenseMessage(
                $"Could not find header column for {amtType} in Journals sheet!",
                MessageBoxImage.Error,
                MessageBoxButton.OK);
        }

        private List<JournalsQuerySubmit> BuildDrilldownList(
            Excel.Worksheet sheet,
            Excel.Range journalsRange,
            int headerIndex,
            out int count,
            out string lastValue,
            out string[] lastJournalValues)
        {
            var list = new List<JournalsQuerySubmit>();
            count = 0;
            lastValue = string.Empty;
            lastJournalValues = Array.Empty<string>();

            foreach (Range loopCell in journalsRange.Cells)
            {
                object? cellVal = (sheet.Cells[loopCell.Row, headerIndex] as Range)?.Value;
                string value = cellVal != null ? Convert.ToString(cellVal) ?? string.Empty : string.Empty;
                if (string.IsNullOrEmpty(value))
                    continue;

                count++;
                var parts = value.Split(new[] { "~~" }, StringSplitOptions.None);
                lastValue = value;
                lastJournalValues = parts;

                list.Add(new JournalsQuerySubmit
                {
                    periodName = parts.ElementAtOrDefault(0),
                    actualFlag = parts.ElementAtOrDefault(1),
                    encumbranceTypeId = ToLongSafe(parts.ElementAtOrDefault(3)),
                    codeCombinationId = ToLongSafe(parts.ElementAtOrDefault(4)),
                    ledgerId = ToLongSafe(parts.ElementAtOrDefault(5)),
                    balanceType = parts.ElementAtOrDefault(6),
                    currencyCode = parts.ElementAtOrDefault(7),
                    translatedFlag = parts.ElementAtOrDefault(8),
                    jeSourceName = parts.ElementAtOrDefault(9),
                    jeCategoryName = parts.ElementAtOrDefault(10),
                    status = parts.ElementAtOrDefault(11),
                    startDate=parts.ElementAtOrDefault(12),
                    endDate = parts.ElementAtOrDefault(13)
                });
            }

            return list;
        }

        private string BuildWorksheetName(
            Excel.Worksheet sheet,
            int selectedCount)
        {
            string sheetName = sheet.Name;
            string cellStr = ExcelApp.ActiveCell.get_Address(
                RowAbsolute: false,
                ColumnAbsolute: false,
                ReferenceStyle: XlReferenceStyle.xlA1,
                External: false);

            string suffix = selectedCount >= 2
                ? $"_{_ddType}_" + cellStr + " +"
                : $"_{_ddType}_" + cellStr;

            return sheetName + suffix;
        }

        private static string BuildMultiStringSafe(string value, string[] journalValues)
        {
            if (string.IsNullOrEmpty(value) || !value.Contains("~"))
                return value;

            try
            {
                return BuildMultiString(JournalHeaders, journalValues);
            }
            catch (Exception ex)
            {
                ServiceLocator.Logger?.LogException(ex, "DrilldownJl.BuildMultiStringSafe");
                return value;
            }
        }

        private static string NormalizeWorksheetName(string worksheetName)
        {
            worksheetName = CommonFunctions.SanitizeSheetName(worksheetName);
            return worksheetName;
        }

        private static string BuildConsolidatedObject(
            string worksheetName,
            string journalsAddress,
            string ddType,
            string multiString)
        {
            return $"{worksheetName}*{journalsAddress}*{DateTime.Now:dd-MMM-yyyy hh:mm:ss tt}*{ddType}*{multiString}";
        }

        private string BuildApiUrl(string consolidatedObject)
        {
            string encodedUrl = WebUtility.UrlEncode(consolidatedObject);
            long cubeId = AppState.Instance.SelectedCube.CubeId;

            string endPoint;
            if (_ddType == "BLDD_SL")
            {
                endPoint = "balance-dd-subledger-dd";
            }
            else if (_ddType == "BLDD_UF")
            {
                endPoint = "balance-dd-unified-dd";
            }
            else
            {
                endPoint = "journal-drilldown"; // default endpoint for plain JL drilldown
            }

            return $"{AppState.Instance.LoginUrl}/rest/secure/finance/{endPoint}?cubeId={cubeId}&jobName={_ddType}&jobDescription={encodedUrl}";
        }

        private async Task SendRequestAndHandleResponseAsync(
            string apiUrl,
            string httpPostText,
            string consolidatedObject)
        {
            try
            {
                if (Win != null)
                    _ = Win.Dispatcher.InvokeAsync(() => Win.SetProcessMessage("Sending request to server..."));

                ServiceLocator.Logger?.LogDebug($"DrilldownJl.SendRequestAndHandleResponseAsync: POST {apiUrl}");
                string output = await ApiHelper.ServerAPI(apiUrl, "JSON", httpPostText, "POST", Token);
                ServiceLocator.Logger?.LogDebug("DrilldownJl.SendRequestAndHandleResponseAsync: response received from server.");

                if (Win != null)
                    _ = Win.Dispatcher.InvokeAsync(() => Win.SetProcessMessage("Response received..."));

                await HandleResponseAsync(output.Trim(), consolidatedObject);
            }
            catch (OperationCanceledException exOp)
            {
                ServiceLocator.Logger?.LogWarn($"DrilldownJl.SendRequestAndHandleResponseAsync: request to {apiUrl} was cancelled: {exOp.Message}");
            }
            catch (Exception ex)
            {
                ServiceLocator.Logger?.LogException(ex, "DrilldownJl.SendRequestAndHandleResponseAsync");
            }
        }

        private async Task HandleResponseAsync(string json, string consolidatedObject)
        {
            try
            {
                var result = ApiResponseHelper.Parse<JsonElement>(json, JsonGlobals.Options);

                if (!result.IsSuccess)
                {
                    await ShowErrorMessageAsync(result.ErrorMessage ?? "Unknown error");
                    return;
                }

                var root = result.Value;

                string msg = string.Empty;
                if (root.TryGetProperty("msg", out var msgEl))
                    msg = msgEl.GetString() ?? string.Empty;
                else if (root.TryGetProperty("message", out var messageEl))
                    msg = messageEl.GetString() ?? string.Empty;

                if (await HandleBackgroundProcessingAsync(msg))
                    return;

                if (!TryGetRecordsNode(root, out JsonElement recordsNode))
                {
                    await ShowErrorMessageAsync(string.IsNullOrWhiteSpace(msg) ? "No records returned." : msg);
                    return;
                }

                int recordCount = GetRecordCount(recordsNode);
                if (recordCount == 0)
                {
                    await ShowErrorMessageAsync(string.IsNullOrWhiteSpace(msg) ? "No records returned." : msg);
                    return;
                }

                var dataToSheet = new DDDatatoWorksheet(ExcelApp, json, "JL", consolidatedObject, Token, Win);
                await dataToSheet.DD_DatetoWorksheet();
            }
            catch (Exception ex)
            {
                ServiceLocator.Logger?.LogException(ex);
            }
        }

        private static int GetRecordCount(JsonElement recordsNode) =>
            recordsNode.ValueKind switch
            {
                JsonValueKind.Array => recordsNode.GetArrayLength(),
                JsonValueKind.Object => recordsNode.EnumerateObject().Count(),
                _ => 0
            };

        private static bool TryGetRecordsNode(JsonElement root, out JsonElement recordsNode)
        {
            var recordProp = root.EnumerateObject()
                .FirstOrDefault(prop => string.Equals(prop.Name, "records", StringComparison.OrdinalIgnoreCase));

            if (recordProp.Value.ValueKind != JsonValueKind.Undefined)
            {
                recordsNode = recordProp.Value;
                return true;
            }

            recordsNode = default;
            return false;
        }

        private static async Task ShowErrorMessageAsync(string msg)
        {
            if (Win != null)
                await Win.Dispatcher.InvokeAsync(() => Win.RequestClose());
            if (!string.IsNullOrWhiteSpace(msg))
            {
                CommonFunctions.GLSenseMessage(msg, MessageBoxImage.Error, MessageBoxButton.OK);
            }
        }

        private async Task<bool> HandleBackgroundProcessingAsync(string msg)
        {
            if (msg.IndexOf("Drilldown request is being processed in the background", StringComparison.OrdinalIgnoreCase) < 0)
                return false;

            Token.ThrowIfCancellationRequested();

            string displayMsg = msg;
            string jobId = ExtractJobId(msg);

            if (!string.IsNullOrWhiteSpace(jobId))
            {
                string rangeName = "GLSense_DD_" + jobId;
                if (!CommonFunctions.NameRangeExists(rangeName))
                {
                    // Regression-pattern fix: see DrilldownBl.HandleBackgroundProcessingAsync's
                    // identical fix in DD_BL.cs for the full writeup - fetching a fresh
                    // ExcelApp.ActiveWorkbook this deep in the async drilldown chain (after the
                    // HTTP round-trip) is the exact pattern that threw an InvalidCastException
                    // crossing the host<->Addin.Core AppDomain boundary in SanitizeSheetName.
                    // JlWorbook was already captured once, early, in the constructor - reuse it
                    // instead of re-fetching. Guarded and logged (non-fatal): this only adds a
                    // named range used by the "launch process window" follow-up.
                    try
                    {
                        JlWorbook?.Names.Add(Name: rangeName, RefersToR1C1: jobId);
                    }
                    catch (Exception ex)
                    {
                        ServiceLocator.Logger?.LogWarn($"DrilldownJl.HandleBackgroundProcessingAsync: failed to add named range '{rangeName}' (non-fatal): {ex.Message}");
                    }
                }
                displayMsg += Environment.NewLine + "Launch process window to check the status.";
            }

            await ShowMessageAsync(displayMsg, MessageBoxImage.Information);
            return true;
        }

        private static async Task ShowMessageAsync(string message, MessageBoxImage msgIcon)
        {
            if (string.IsNullOrWhiteSpace(message) || Win == null) return;
            await SafelyCloseWindowAsync();
            CommonFunctions.GLSenseMessage(message, msgIcon, MessageBoxButton.OK);
        }

        private static string ExtractJobId(string msg)
        {
            var parts = msg.Split(' ');
            if (parts.Length == 0) return string.Empty;
            // Replace usage of ^ operator (C# 8.0 Index) with classic array access
            return parts[parts.Length - 1].TrimEnd('.');
        }

        private static string BuildMultiString(string[] headers, string[] values)
        {
            if (values == null || values.Length == 0) return string.Empty;

            int len = Math.Min(headers.Length, values.Length);

            var pairs = Enumerable.Range(0, len)
                .Select(i => $"{headers[i]}:= {values[i]}");

            return string.Join(", ", pairs);
        }

        private static long ToLongSafe(object value)
        {
            if (value == null) return 0;

            if (value is double d) return (long)d;
            if (value is int i) return i;
            if (value is long l) return l;

            var s = Convert.ToString(value)?.Trim();
            if (string.IsNullOrEmpty(s)) return 0;

            s = s!.Replace(",", "").Replace(" ", "");

            return long.TryParse(s, out var result) ? result : 0;
        }

        private static string TryGetColumnHeaderText(Excel.Worksheet wrksheet, Excel.Range rng)
        {
            object? v4 = (wrksheet.Cells[4, rng.Column] as Range)?.Value;
            string HeaderText = v4 != null ? Convert.ToString(v4) ?? string.Empty : string.Empty;

            if (string.IsNullOrWhiteSpace(HeaderText))
            {
                object? v5 = (wrksheet.Cells[5, rng.Column] as Range)?.Value;
                HeaderText = v5 != null ? Convert.ToString(v5) ?? string.Empty : string.Empty;
            }

            return HeaderText;
        }

        private int TryGetColumnHeaderIndex(Excel.Worksheet JournalsSheet, string AmtType)
        {
            if (JournalsSheet == null) return -1;

            try
            {
                Excel.Range HeaderRange = JournalsSheet.Range["A5:ZZ5"];
                string[] searchTerms = GetSearchTerms(AmtType);

                foreach (string term in searchTerms)
                {
                    try
                    {
                        return (int)ExcelApp.WorksheetFunction.Match(term, HeaderRange, 0);
                    }
                    catch
                    {
                        // Expected/normal: WorksheetFunction.Match throws when the term isn't found in
                        // the header row - just try the next candidate term, nothing to log here.
                    }
                }

                ServiceLocator.Logger?.LogDebug($"DrilldownJl.TryGetColumnHeaderIndex: none of the search terms [{string.Join(", ", searchTerms)}] for amtType={AmtType} matched a header in sheet '{JournalsSheet.Name}'.");
            }
            catch (Exception ex)
            {
                ServiceLocator.Logger?.LogException(ex, "DrilldownJl.TryGetColumnHeaderIndex");
            }

            return -1;
        }

        private static string[] GetSearchTerms(string AmtType)
        {
            return AmtType.ToUpperInvariant() switch
            {
                "YTD" => new[] { "DRILL_DOWN1", "DRILLDOWN1" },
                "PTD"  => new[] { "DRILL_DOWN2", "DRILLDOWN2" },
                "JED" => new[] { "DRILL_DOWN2", "DRILLDOWN2" },
                "QTD" => new[] { "DRILL_DOWN3", "DRILLDOWN3" },
                _ => Array.Empty<string>()
            };
        }
    }

    public class JournalsQuerySubmit
    {
        public string? actualFlag { get; set; }
        public long ledgerId { get; set; }
        public long codeCombinationId { get; set; }
        public string? periodName { get; set; }
        public string? balanceType { get; set; }
        public string? currencyCode { get; set; }
        public string? translatedFlag { get; set; }
        public long encumbranceTypeId { get; set; }
        public string? jeSourceName { get; set; }
        public string? jeCategoryName { get; set; }
        public string? status { get; set; }
        public string? startDate { get; set; }
        public string? endDate { get; set; }
    }

    public class JournalDD
    {
        public JournalsQuerySubmit[]? journalDrilldowns { get; set; }
    }
#nullable disable
}
