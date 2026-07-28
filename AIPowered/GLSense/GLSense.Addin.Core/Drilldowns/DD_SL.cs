// DD_SL.cs in GLSense.Addin.Core
// Port of GLSense\Drilldowns\DD_SL.cs (FinalWorkingCode), class DrilldownSl.
// DrilldownSl does NOT share a base class with DrilldownBl/DrilldownJl in the old code
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
    public class DrilldownSl
    {
        private Excel.Application ExcelApp { get; }
        private Excel.Worksheet? SlWorksheet { get; set; }
        private Excel.Range? SlRange { get; set; }
        private CancellationHelper? _ctsHelper;
        private CancellationToken Token => _ctsHelper?.GetToken() ?? default;
        private ExternalResolveResult ExternalResolveResult { get; set; } = new ExternalResolveResult();
        private readonly string? _slAddress;
        private static GLWaitWindow? Win { get; set; }

        public DrilldownSl(Excel.Application excelApp, string rngAddress)
        {
            ExcelApp = excelApp ?? throw new ArgumentNullException(nameof(excelApp));
            _slAddress = rngAddress;
        }

        public async Task ProcessSLDrilldown()
        {
            ServiceLocator.Logger?.LogDebug($"DrilldownSl.ProcessSLDrilldown started. address={_slAddress}");

            ExternalResolveResult = ExcelExternalRef.ResolveRangeWithContext(_slAddress);
            SlRange = ExternalResolveResult.Range;

            if (SlRange == null)
            {
                return;
            }

            SlWorksheet = ExternalResolveResult.Worksheet;

            CommonMethods.DisableExcelSettings();

            if (!IsValidSingleColumnSelection(SlRange))
                return;

            if (!HasAnyValue(SlRange))
                return;

            if (!IsInRange(SlRange))
                return;

            int headerIndex = TryGetColumnHeaderIndex(SlWorksheet);
            if (headerIndex == -1)
            {
                await ShowHeaderNotFoundAsync();
                return;
            }

            _ctsHelper = new CancellationHelper();

            Win = CreateAndShowProgressWindow(_ctsHelper);

            try
            {
                await InitializeProgressWindowAsync("Subledgers Drilldown", "Processing request...");

                if (!await TryRunDrilldownAsync(SlWorksheet, SlRange, headerIndex))
                    return;
            }
            catch (OperationCanceledException)
            {
                await ShowCancelledAsync();
                ServiceLocator.Logger?.LogWarn("Subledger drilldown operation cancelled by user.");
            }
            catch (Exception ex)
            {
                await HandleUnexpectedErrorAsync(ex);
            }
            finally
            {
                try
                {
                    if (!_ctsHelper.IsCancellationRequested)
                        _ctsHelper.Cancel();

                    _ctsHelper.Dispose();  // safe
                }
                catch (Exception ex)
                {
                    // Swallow dispose exceptions (Excel COM weirdness) but still log for diagnostics.
                    ServiceLocator.Logger?.LogWarn($"DrilldownSl.ProcessSLDrilldown: exception disposing CancellationHelper (ignored): {ex.Message}");
                }
                await SafelyCloseWindowAsync();
                CommonMethods.EnableExcelSettings();
            }
        }

        private async Task<bool> TryRunDrilldownAsync(
                    Excel.Worksheet wrksheet,
                    Excel.Range rng,
                    int HeaderIndex)
        {
            await Subledger_DrillDown(wrksheet, rng, HeaderIndex);
            return true;
        }

        private async Task Subledger_DrillDown(
                Excel.Worksheet SubledgersSheet,
                Excel.Range SubledgersRange,
                int HeaderIndex)
        {
            string? SubledgersAddress = SubledgersRange.get_Address(
                RowAbsolute: true,
                ColumnAbsolute: true,
                ReferenceStyle: XlReferenceStyle.xlA1,
                External: true);

            string? CellStr = SubledgersRange.get_Address(
                RowAbsolute: false,
                ColumnAbsolute: false,
                External: false);

            var strView = new List<string>();
            var strBuilder = new List<string>();

            ExtractViewData(SubledgersRange, HeaderIndex, SubledgersSheet, ref strView, ref strBuilder);

            ServiceLocator.Logger?.LogDebug($"DrilldownSl.Subledger_DrillDown: extracted {strView.Count} view(s), {strBuilder.Count} row(s) from selection.");

            if (strView.Count == 0)
            {
                ServiceLocator.Logger?.LogDebug("DrilldownSl.Subledger_DrillDown: no views found in selection, aborting.");
                await HandleEmptyViewsAsync();
                return;
            }

            Token.ThrowIfCancellationRequested();

            string? worksheetName = BuildWorksheetName(strBuilder, SubledgersSheet.Name, CellStr);
            string? multiString = strBuilder.Count >= 2 ? "Multi Select" : strBuilder.ElementAtOrDefault(0);

            string? consolidatedObject = BuildConsolidatedObject(worksheetName, SubledgersAddress, multiString);
            Token.ThrowIfCancellationRequested();

            var payload = BuildSubledgerPayload(strView, strBuilder);
            Token.ThrowIfCancellationRequested();

            string? httpPostText = JsonSerializer.Serialize(payload, JsonGlobals.Options);

            string? apiUrl = BuildApiUrl(consolidatedObject);
            Token.ThrowIfCancellationRequested();

            await SendRequestAndHandleResponseAsync(apiUrl, httpPostText, consolidatedObject);
            Token.ThrowIfCancellationRequested();
        }

        private async Task SendRequestAndHandleResponseAsync(
            string apiUrl,
            string httpPostText,
            string consolidatedObject)
        {
            try
            {
                if (Win != null)
                {
                    _ = Win.Dispatcher.InvokeAsync(() => Win.SetProcessMessage("Sending request to server..."));
                }
                ServiceLocator.Logger?.LogDebug($"DrilldownSl.SendRequestAndHandleResponseAsync: POST {apiUrl}");
                var output = await ApiHelper.ServerAPI(apiUrl, "JSON", httpPostText, "POST", Token);
                ServiceLocator.Logger?.LogDebug("DrilldownSl.SendRequestAndHandleResponseAsync: response received from server.");
                if (Win != null)
                {
                    _ = Win.Dispatcher.InvokeAsync(() => Win.SetProcessMessage("Response received..."));
                }

                if (!IsValidOutput(output))
                {
                    ServiceLocator.Logger?.LogWarn($"DrilldownSl.SendRequestAndHandleResponseAsync: invalid/unexpected response from {apiUrl}");
                    return;
                }

                Token.ThrowIfCancellationRequested();

                await HandleResponseAsync(output.ToString().Trim(), consolidatedObject);
            }
            catch (OperationCanceledException exOp)
            {
                ServiceLocator.Logger?.LogWarn($"DrilldownSl.SendRequestAndHandleResponseAsync: request to {apiUrl} was cancelled: {exOp.Message}");
            }
            catch (Exception ex)
            {
                ServiceLocator.Logger?.LogException(ex, "DrilldownSl.SendRequestAndHandleResponseAsync");
            }
        }

        private static bool IsValidOutput(object output)
        {
            if (output == null)
                return false;

            string? s = output.ToString();
            return !string.IsNullOrWhiteSpace(s) && s.Length > 3 && s.IndexOf("records", StringComparison.Ordinal) >= 0;
        }

        private async Task HandleResponseAsync(string json, string consolidatedObject)
        {
            try
            {
                if (string.IsNullOrWhiteSpace(json))
                    return;

                var result = ApiResponseHelper.Parse<JsonElement>(json, JsonGlobals.Options);

                if (!result.IsSuccess)
                {
                    await ShowErrorMessageAsync(
                        result.ErrorMessage ?? "Unknown error");
                    return;
                }

                var root = result.Value;

                // Extract message safely
                string msg = string.Empty;
                if (root.TryGetProperty("msg", out var msgEl))
                {
                    msg = msgEl.GetString() ?? string.Empty;
                }
                else if (root.TryGetProperty("message", out var messageEl))
                {
                    msg = messageEl.GetString() ?? string.Empty;
                }

                // Background processing
                if (await HandleBackgroundProcessingAsync(msg))
                    return;

                // Records detection
                if (!TryGetRecordsNode(root, out JsonElement recordsNode))
                {
                    await ShowErrorMessageAsync(string.IsNullOrWhiteSpace(msg) ? "No records returned." : msg);
                    return;
                }

                int recordCount = GetRecordCount(recordsNode);

                if (recordCount == 0)
                {
                    await ShowErrorMessageAsync("No records returned.");
                    return;
                }

                var dataToSheet = new DDDatatoWorksheet(
                    ExcelApp,
                    json,
                    "SL",
                    consolidatedObject,
                    Token,
                    Win);

                await dataToSheet.DD_DatetoWorksheet();
            }
            catch (Exception ex)
            {
                ServiceLocator.Logger?.LogException(ex, "DrilldownSl.HandleResponseAsync");
            }
        }

        private static int GetRecordCount(JsonElement recordsNode)
        {
            return recordsNode.ValueKind switch
            {
                JsonValueKind.Array => recordsNode.GetArrayLength(),
                JsonValueKind.Object => recordsNode.EnumerateObject().Count(),
                _ => 0
            };
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

        private async Task<bool> HandleBackgroundProcessingAsync(string msg)
        {
            if (string.IsNullOrWhiteSpace(msg))
                return false;

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
                    // DrilldownSl has no dedicated *Worbook field like DrilldownBl/DrilldownJl,
                    // but ExternalResolveResult.Workbook was already captured once, early, in
                    // ProcessSLDrilldown (via ExcelExternalRef.ResolveRangeWithContext, before any
                    // async work began) - reuse that instead of re-fetching. Guarded and logged
                    // (non-fatal): this only adds a named range used by the "launch process
                    // window" follow-up.
                    try
                    {
                        ExternalResolveResult?.Workbook?.Names.Add(Name: rangeName, RefersToR1C1: jobId);
                    }
                    catch (Exception ex)
                    {
                        ServiceLocator.Logger?.LogWarn($"DrilldownSl.HandleBackgroundProcessingAsync: failed to add named range '{rangeName}' (non-fatal): {ex.Message}");
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

            string lastPart = parts[parts.Length - 1].TrimEnd('.');
            return lastPart;
        }

        private static async Task ShowErrorMessageAsync(string? msg)
        {
            if (Win != null)
                await Win.Dispatcher.InvokeAsync(() => Win.RequestClose());
            if (msg != null && !string.IsNullOrWhiteSpace(msg))
            {
                CommonFunctions.GLSenseMessage(
                    msg,
                    MessageBoxImage.Error,
                    MessageBoxButton.OK);
            }
        }

        private static string BuildApiUrl(string consolidatedObject)
        {
            string? encodedUrl = WebUtility.UrlEncode(consolidatedObject);
            long? cubeId = AppState.Instance.SelectedCube.CubeId;
            return $"{AppState.Instance.LoginUrl}/rest/secure/finance/subledger-drilldown?cubeId={cubeId}&jobName=SL&jobDescription={encodedUrl}";
        }

        private SubLedgerDD BuildSubledgerPayload(
                List<string> strView,
                List<string> strBuilder)
        {
            var uniqueViews = strView.Distinct().ToList();
            var jsonList = BuildJsonList(uniqueViews, strBuilder.ToArray());

            return new SubLedgerDD
            {
                subledgerDrilldowns = jsonList.ToArray()
            };
        }

        private List<SubledgerQuerySubmit> BuildJsonList(
                    List<string> uniqueView,
                    string[] arraylist)
        {
            var jsonList = new List<SubledgerQuerySubmit>();

            foreach (string element in uniqueView)
            {
                Token.ThrowIfCancellationRequested();

                var elementSplit = element.Split(new[] { "~~" }, StringSplitOptions.None);
                var journalLineNumbers = GetMatchingJournalLines(elementSplit, arraylist);

                jsonList.Add(new SubledgerQuerySubmit
                {
                    subledgerViewName = elementSplit.ElementAtOrDefault(0),
                    journalHeaderId = elementSplit.ElementAtOrDefault(1),
                    jounalLineNumberList = journalLineNumbers.ToArray()
                });
            }

            return jsonList;
        }

        private List<object> GetMatchingJournalLines(string[] elementSplit, string[] arraylist)
        {
            var journalLineNumbers = new List<object>();

            for (int i = 0; i < arraylist.Length; i++)
            {
                Token.ThrowIfCancellationRequested();
                var strValue = arraylist[i].Split(new[] { "~~" }, StringSplitOptions.None);
                if (strValue.ElementAtOrDefault(0) == elementSplit.ElementAtOrDefault(0) &&
                    strValue.ElementAtOrDefault(1) == elementSplit.ElementAtOrDefault(1))
                {
                    journalLineNumbers.Add(strValue.ElementAtOrDefault(2));
                }
            }

            return journalLineNumbers;
        }

        private static string BuildConsolidatedObject(
                string worksheetName,
                string slAddress,
                string multiString)
        {
            return $"{worksheetName}*{slAddress}*{DateTime.Now:dd-MMM-yyyy hh:mm:ss tt}*SL*{multiString}";
        }

        private static string BuildWorksheetName(List<string> arraylist, string sheetName, string cellStr)
        {
            string baseName = $"{sheetName}_SL_{cellStr}";

            if (arraylist.Count >= 2)
                baseName += " +";

            return NormalizeWorksheetName(baseName);
        }

        private static string NormalizeWorksheetName(string worksheetName)
        {
            worksheetName = CommonFunctions.SanitizeSheetName(worksheetName);
            return worksheetName;
        }

        private static async Task HandleEmptyViewsAsync()
        {
            await SafelyCloseWindowAsync();
            CommonFunctions.GLSenseMessage(
                "Views in the selection are empty for subledger drilldown!",
                MessageBoxImage.Error,
                MessageBoxButton.OK);
        }

        private void ExtractViewData(
                Excel.Range rng,
                int lcCol,
                Excel.Worksheet actSheet,
                ref List<string> strView,
                ref List<string> strBuilder)
        {
            foreach (Excel.Range rng1 in rng.Rows)
            {
                Token.ThrowIfCancellationRequested();

                Excel.Range rowRange = (Excel.Range)actSheet.Cells[rng1.Row, lcCol];
                if (rng1.Row < 6 || rowRange.Value2 == null)
                    continue;

                string? rngValue = rowRange.Value2?.ToString() ?? string.Empty;
                strBuilder.Add(rngValue);

                if (!string.IsNullOrWhiteSpace(rngValue))
                {
                    var parts = rngValue.Split(new[] { "~~" }, StringSplitOptions.None);
                    strView.Add($"{parts.ElementAtOrDefault(0)}~~{parts.ElementAtOrDefault(1)}");
                }
            }
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

            try
            {
                // Fire-and-forget: progress UI update only. Do not introduce a
                // suspend point here - the caller reads Excel COM objects
                // (SlWorksheet/SlRange) immediately after awaiting this method.
                _ = Win.Dispatcher.InvokeAsync(() =>
                {
                    Win.SetProcessTitle(title);
                    Win.SetProcessMessage(message);
                }, System.Windows.Threading.DispatcherPriority.Normal);

                return Task.CompletedTask;
            }
            catch (Exception ex)
            {
                ServiceLocator.Logger?.LogException(ex, "DrilldownSl.InitializeProgressWindowAsync");
                return Task.CompletedTask;
            }
        }

        private static async Task SafelyCloseWindowAsync()
        {
            if (Win == null)
                return;

            try
            {
                if (Win.Dispatcher.CheckAccess())
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

        private static async Task ShowHeaderNotFoundAsync()
        {
            if (Win != null)
                await Win.Dispatcher.InvokeAsync(() => Win.RequestClose());
            CommonFunctions.GLSenseMessage(
                $"Could not find header column in Subledgers sheet!",
                MessageBoxImage.Error,
                MessageBoxButton.OK);
        }

        private int TryGetColumnHeaderIndex(Excel.Worksheet SlSheet)
        {
            if (SlSheet == null) return -1;

            try
            {
                Excel.Range HeaderRange = SlSheet.Range["A5:ZZ5"];
                string[] searchTerms = new[] { "DRILL_DOWN1", "DRILLDOWN1" };

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

                ServiceLocator.Logger?.LogDebug($"DrilldownSl.TryGetColumnHeaderIndex: none of the search terms [{string.Join(", ", searchTerms)}] matched a header in sheet '{SlSheet.Name}'.");
            }
            catch (Exception ex)
            {
                ServiceLocator.Logger?.LogException(ex, "DrilldownSl.TryGetColumnHeaderIndex");
            }

            return -1;
        }

        private static bool IsInRange(Excel.Range rng)
        {
            if (rng == null || rng.Row <= 5)
            {
                CommonFunctions.GLSenseMessage(
                    "Invalid selection for sub-ledgers drilldown!",
                    MessageBoxImage.Exclamation,
                    MessageBoxButton.OK);
                return false;
            }

            return true;
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
    }

    public class SubledgerQuerySubmit
    {
        public string? subledgerViewName { get; set; }
        public string? journalHeaderId { get; set; }
        public object[]? jounalLineNumberList { get; set; }
    }

    public class SubLedgerDD
    {
        public SubledgerQuerySubmit[]? subledgerDrilldowns { get; set; }
    }
#nullable disable
}
