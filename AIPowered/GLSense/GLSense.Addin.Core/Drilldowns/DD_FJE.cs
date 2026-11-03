// DD_FJE.cs in GLSense.Addin.Core
// OISR-22823 (parent OISR-19286): "View Full Journal Entry" drilldown from a journal sheet.
//
// A journal drilldown sheet carries two columns from the server: VIEW_FULL_JOURNAL_ENTRY
// (shown, turned into a hyperlink by DDDatatoWorksheet) and FJE_DD_PARAMETERS (hidden,
// "ledgerId~~headerId"). Clicking the hyperlink makes the host's SheetFollowHyperlink handler
// send the clicked cell's external address here (AddinEntry "RunFullJournalEntry"). This class
// reads that row's FJE_DD_PARAMETERS, posts
//   POST {LoginUrl}/rest/secure/finance/journal-entry-drilldown?cubeId={cubeId}
//   { "journalEntries": [ { "ledgerId": n, "headerId": n } ], "metadata": [ ... ] }
// and writes the result like a subledger drilldown from journals (ddType "SL", sheet name
// "{sheet}_SL_{cell}"). "metadata" is the JOURNAL metadata saved locally for the cube when
// "Overwrite drilldown metadata with locally saved" is on, otherwise empty. The title
// "{link text} ({FJE_DD_PARAMETERS})" goes on the progress window and, through the 5th part
// of the consolidated object, on A3 of the new sheet.
//
// Progress window, cancellation, background-job and error handling follow DD_SL.cs.
using GLSense.Addin.Core.Common;
using GLSense.Addin.Core.Helpers;
using GLSense.Addin.Core.Infrastructure;
using GLSense.Addin.Core.Models;
using GLSense.Addin.Core.Utilities;
using GLSense.Addin.Core.Views;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text.Json;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using Excel = Microsoft.Office.Interop.Excel;

namespace GLSense.Addin.Core.Drilldowns
{
#nullable enable
    public class DrilldownFje
    {
        private Excel.Application ExcelApp { get; }
        private readonly string? _cellAddress;
        private CancellationHelper? _ctsHelper;
        private CancellationToken Token => _ctsHelper?.GetToken() ?? default;
        private ExternalResolveResult ExternalResolveResult { get; set; } = new ExternalResolveResult();
        private static GLWaitWindow? Win { get; set; }

        public DrilldownFje(Excel.Application excelApp, string cellAddress)
        {
            ExcelApp = excelApp ?? throw new ArgumentNullException(nameof(excelApp));
            _cellAddress = cellAddress;
        }

        public async Task ProcessFullJournalEntryDrilldown()
        {
            ServiceLocator.Logger?.LogDebug($"DrilldownFje.ProcessFullJournalEntryDrilldown started. address={_cellAddress}");

            ExternalResolveResult = ExcelExternalRef.ResolveRangeWithContext(_cellAddress);
            Excel.Range? cell = ExternalResolveResult.Range;
            Excel.Worksheet? sheet = ExternalResolveResult.Worksheet;

            if (cell == null || sheet == null)
                return;

            if (cell.Row <= 5)
            {
                ShowMessage("Invalid selection for full journal entry drilldown!", MessageBoxImage.Exclamation);
                return;
            }

            int paramsColumn = FindParamsColumn(sheet);
            if (paramsColumn == -1)
            {
                ShowMessage($"Could not find the {DDDatatoWorksheet.FullJournalEntryParamsColumn} column in the journals sheet!", MessageBoxImage.Error);
                return;
            }

            string linkText = ((cell.Value2 ?? cell.Value)?.ToString() ?? string.Empty).Trim();
            var paramsCell = (Excel.Range)sheet.Cells[cell.Row, paramsColumn];
            string fjeParams = ((paramsCell.Value2 ?? paramsCell.Value)?.ToString() ?? string.Empty).Trim();

            if (!TryParseParams(fjeParams, out long ledgerId, out long headerId))
            {
                ServiceLocator.Logger?.LogWarn($"DrilldownFje: invalid {DDDatatoWorksheet.FullJournalEntryParamsColumn} value '{fjeParams}' on row {cell.Row} of '{sheet.Name}'.");
                ShowMessage($"Invalid full journal entry parameters '{fjeParams}' for the selected row!", MessageBoxImage.Error);
                return;
            }

            string title = $"{linkText} ({fjeParams})";

            if (!CommonMethods.TryDisableExcelSettings("DrilldownFje.ProcessFullJournalEntryDrilldown"))
                return;

            _ctsHelper = new CancellationHelper();
            Win = CreateAndShowProgressWindow(_ctsHelper);

            try
            {
                InitializeProgressWindow(title, "Processing request...");

                string cellStr = cell.get_Address(RowAbsolute: false, ColumnAbsolute: false, External: false);
                string external = cell.get_Address(
                    RowAbsolute: true,
                    ColumnAbsolute: true,
                    ReferenceStyle: Excel.XlReferenceStyle.xlA1,
                    External: true);

                string worksheetName = CommonFunctions.SanitizeSheetName($"{sheet.Name}_SL_{cellStr}");
                string consolidatedObject = $"{worksheetName}*{external}*{DateTime.Now:dd-MMM-yyyy hh:mm:ss tt}*SL*{title}";

                var payload = new FullJournalEntryRequest
                {
                    journalEntries = new[] { new FullJournalEntryKey { ledgerId = ledgerId, headerId = headerId } },
                    metadata = GetLocalJournalMetadata()
                };
                Token.ThrowIfCancellationRequested();

                string httpPostText = JsonSerializer.Serialize(payload, JsonGlobals.Options);
                long? cubeId = AppState.Instance.SelectedCube?.CubeId;
                string apiUrl = $"{AppState.Instance.LoginUrl}/rest/secure/finance/journal-entry-drilldown?cubeId={cubeId}";

                await SendRequestAndHandleResponseAsync(apiUrl, httpPostText, consolidatedObject);
            }
            catch (OperationCanceledException)
            {
                await SafelyCloseWindowAsync();
                ShowMessage("Operation cancelled!", MessageBoxImage.Error);
                ServiceLocator.Logger?.LogWarn("Full journal entry drilldown cancelled by user.");
            }
            catch (Exception ex)
            {
                ServiceLocator.Logger?.LogException(ex, "DrilldownFje.ProcessFullJournalEntryDrilldown");
                await SafelyCloseWindowAsync();
                ShowMessage($"An unexpected error occurred.{Environment.NewLine}{ex.Message}", MessageBoxImage.Error);
            }
            finally
            {
                try
                {
                    if (!_ctsHelper.IsCancellationRequested)
                        _ctsHelper.Cancel();
                    _ctsHelper.Dispose();
                }
                catch (Exception ex)
                {
                    ServiceLocator.Logger?.LogWarn($"DrilldownFje: exception disposing CancellationHelper (ignored): {ex.Message}");
                }
                await SafelyCloseWindowAsync();
                CommonMethods.TryEnableExcelSettings("DrilldownFje.ProcessFullJournalEntryDrilldown");
            }
        }

        // "ledgerId~~headerId", both whole numbers.
        internal static bool TryParseParams(string value, out long ledgerId, out long headerId)
        {
            ledgerId = 0;
            headerId = 0;

            if (string.IsNullOrWhiteSpace(value))
                return false;

            var parts = value.Split(new[] { "~~" }, StringSplitOptions.None);
            if (parts.Length != 2)
                return false;

            return long.TryParse(parts[0].Trim(), NumberStyles.Integer, CultureInfo.InvariantCulture, out ledgerId) &&
                   long.TryParse(parts[1].Trim(), NumberStyles.Integer, CultureInfo.InvariantCulture, out headerId);
        }

        // Row 5 holds the display name, row 4 the actual column name (DDDatatoWorksheet).
        private int FindParamsColumn(Excel.Worksheet sheet)
        {
            foreach (string headerRow in new[] { "A5:ZZ5", "A4:ZZ4" })
            {
                try
                {
                    return (int)ExcelApp.WorksheetFunction.Match(
                        DDDatatoWorksheet.FullJournalEntryParamsColumn, sheet.Range[headerRow], 0);
                }
                catch
                {
                    // Match throws when not found - try the next header row.
                }
            }

            ServiceLocator.Logger?.LogDebug($"DrilldownFje.FindParamsColumn: {DDDatatoWorksheet.FullJournalEntryParamsColumn} not found in sheet '{sheet.Name}'.");
            return -1;
        }

        private Dictionary<string, object>[] GetLocalJournalMetadata()
        {
            try
            {
                if (!UserConfig.OverwriteDrilldownMetadata)
                    return new Dictionary<string, object>[0];

                var cube = AppState.Instance.SelectedCube;
                var wb = ExternalResolveResult?.Workbook;
                if (cube == null || wb == null)
                    return new Dictionary<string, object>[0];

                if (!DrilldownMetadataXmlStore.TryRead(wb, cube.CubeId, "JOURNAL", out string rawJson))
                    return new Dictionary<string, object>[0];

                var metadata = DrilldownMetadataXmlStore.ExtractDrilldownTypeMetadata(rawJson, "JOURNAL");
                ServiceLocator.Logger?.LogDebug($"DrilldownFje.GetLocalJournalMetadata: sending {metadata?.Length ?? 0} locally saved JOURNAL column(s).");
                return metadata ?? new Dictionary<string, object>[0];
            }
            catch (Exception ex)
            {
                ServiceLocator.Logger?.LogException(ex, "DrilldownFje.GetLocalJournalMetadata");
                return new Dictionary<string, object>[0];
            }
        }

        private async Task SendRequestAndHandleResponseAsync(string apiUrl, string httpPostText, string consolidatedObject)
        {
            try
            {
                SetProcessMessage("Sending request to server...");
                ServiceLocator.Logger?.LogDebug($"DrilldownFje.SendRequestAndHandleResponseAsync: POST {apiUrl}");
                var output = await ApiHelper.ServerAPI(apiUrl, "JSON", httpPostText, "POST", Token);
                SetProcessMessage("Response received...");

                string? json = output?.ToString()?.Trim();
                if (string.IsNullOrWhiteSpace(json) || json!.Length <= 3 || json.IndexOf("records", StringComparison.Ordinal) < 0)
                {
                    ServiceLocator.Logger?.LogWarn($"DrilldownFje.SendRequestAndHandleResponseAsync: invalid/unexpected response from {apiUrl}");
                    return;
                }

                Token.ThrowIfCancellationRequested();
                await HandleResponseAsync(json, consolidatedObject);
            }
            catch (OperationCanceledException exOp)
            {
                ServiceLocator.Logger?.LogWarn($"DrilldownFje.SendRequestAndHandleResponseAsync: request to {apiUrl} was cancelled: {exOp.Message}");
            }
            catch (Exception ex)
            {
                ServiceLocator.Logger?.LogException(ex, "DrilldownFje.SendRequestAndHandleResponseAsync");
            }
        }

        private async Task HandleResponseAsync(string json, string consolidatedObject)
        {
            var result = ApiResponseHelper.Parse<JsonElement>(json, JsonGlobals.Options);
            if (!result.IsSuccess)
            {
                await ShowErrorAsync(result.ErrorMessage ?? "Unknown error");
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

            var recordProp = root.EnumerateObject()
                .FirstOrDefault(p => string.Equals(p.Name, "records", StringComparison.OrdinalIgnoreCase));

            if (recordProp.Value.ValueKind == JsonValueKind.Undefined)
            {
                await ShowErrorAsync(string.IsNullOrWhiteSpace(msg) ? "No records returned." : msg);
                return;
            }

            int recordCount = recordProp.Value.ValueKind switch
            {
                JsonValueKind.Array => recordProp.Value.GetArrayLength(),
                JsonValueKind.Object => recordProp.Value.EnumerateObject().Count(),
                _ => 0
            };

            if (recordCount == 0)
            {
                await ShowErrorAsync("No records returned.");
                return;
            }

            var dataToSheet = new DDDatatoWorksheet(ExcelApp, json, "SL", consolidatedObject, Token, Win);
            await dataToSheet.DD_DatetoWorksheet();
        }

        // Same as DrilldownSl.HandleBackgroundProcessingAsync.
        private async Task<bool> HandleBackgroundProcessingAsync(string msg)
        {
            if (string.IsNullOrWhiteSpace(msg) ||
                msg.IndexOf("Drilldown request is being processed in the background", StringComparison.OrdinalIgnoreCase) < 0)
                return false;

            Token.ThrowIfCancellationRequested();

            string displayMsg = msg;
            var parts = msg.Split(' ');
            string jobId = parts.Length == 0 ? string.Empty : parts[parts.Length - 1].TrimEnd('.');

            if (!string.IsNullOrWhiteSpace(jobId))
            {
                string rangeName = "GLSense_DD_" + jobId;
                if (!CommonFunctions.NameRangeExists(rangeName))
                {
                    try
                    {
                        ExternalResolveResult?.Workbook?.Names.Add(Name: rangeName, RefersToR1C1: jobId);
                    }
                    catch (Exception ex)
                    {
                        ServiceLocator.Logger?.LogWarn($"DrilldownFje.HandleBackgroundProcessingAsync: failed to add named range '{rangeName}' (non-fatal): {ex.Message}");
                    }
                }

                displayMsg += Environment.NewLine + "Launch process window to check the status.";
            }

            await SafelyCloseWindowAsync();
            ShowMessage(displayMsg, MessageBoxImage.Information);
            return true;
        }

        private static async Task ShowErrorAsync(string msg)
        {
            await SafelyCloseWindowAsync();
            ShowMessage(msg, MessageBoxImage.Error);
        }

        private static void ShowMessage(string message, MessageBoxImage icon)
        {
            if (string.IsNullOrWhiteSpace(message))
                return;

            CommonFunctions.GLSenseMessage(message, icon, MessageBoxButton.OK);
        }

        private static void SetProcessMessage(string message)
        {
            if (Win != null)
                _ = Win.Dispatcher.InvokeAsync(() => Win.SetProcessMessage(message));
        }

        private static GLWaitWindow? CreateAndShowProgressWindow(CancellationHelper cts)
        {
            GLWaitWindow? win = null;
            try
            {
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
            }
            catch (Exception ex)
            {
                ServiceLocator.Logger?.LogException(ex);
            }
            return win;
        }

        // Fire-and-forget, as in DrilldownSl: no suspend point before the Excel COM reads that follow.
        private static void InitializeProgressWindow(string title, string message)
        {
            if (Win == null)
                return;

            try
            {
                _ = Win.Dispatcher.InvokeAsync(() =>
                {
                    Win.SetProcessTitle(title);
                    Win.SetProcessMessage(message);
                }, System.Windows.Threading.DispatcherPriority.Normal);
            }
            catch (Exception ex)
            {
                ServiceLocator.Logger?.LogException(ex, "DrilldownFje.InitializeProgressWindow");
            }
        }

        private static async Task SafelyCloseWindowAsync()
        {
            if (Win == null)
                return;

            try
            {
                await Win.Dispatcher.InvokeAsync(() => Win.RequestClose());
            }
            catch (Exception ex)
            {
                ServiceLocator.Logger?.LogError($"DrilldownFje: error closing wait window: {ex.Message}");
            }
        }
    }

    public class FullJournalEntryKey
    {
        public long ledgerId { get; set; }
        public long headerId { get; set; }
    }

    public class FullJournalEntryRequest
    {
        public FullJournalEntryKey[]? journalEntries { get; set; }
        public Dictionary<string, object>[]? metadata { get; set; }
    }
#nullable disable
}
