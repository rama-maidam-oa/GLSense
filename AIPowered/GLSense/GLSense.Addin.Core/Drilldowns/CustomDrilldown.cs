// CustomDrilldown.cs in GLSense.Addin.Core
// Ported from GLSense\AddinModule.cs (FinalWorkingCode) - the custom-drilldown hyperlink
// flow: GLSenseCustomDrillDown/ProcessCustomDrilldownAsync/FindDrilldownXmlForSheet/
// ExtractDrilldownSheet_XPath/ExecuteCustomDrilldownAsync/UpdateJsonParameters/
// UpdateColumnParameter/FindColumnNumber/SendCustomDrilldownRequestAsync/
// WriteCustomDrilldownToWorksheet/MergeJsonAndGenerateString. Deliberately deferred out
// of the earlier Group E "Drilldowns" pass along with the journal-attachment flow (see
// JournalAttachments.cs) - both are triggered by SheetFollowHyperlink, whose host-side
// classification/dispatch is a separate, later pass (see AddinEntry.cs OnExcelEvent's
// "SheetFollowHyperlink" case).
//
// PUBLIC ENTRY POINT (host-side wiring depends on this exact signature):
//   public static async Task RunCustomDrilldown(string tableSheetName, string cellExternalAddress, string headerLabel)
//
// Preconditions the later pass's AddinEntry.OnExcelEvent("SheetFollowHyperlink", ...)
// handler is expected to have already checked, exactly like the old monolith's
// adxExcelAppEvents1_SheetFollowHyperlink did before calling GLSenseCustomDrillDown
// (see AddinModule.cs's IsValidDrilldownSheet/IsCustomDrilldownHyperlink):
//   - IsValidDrilldownSheet(sht): sht.ListObjects.Count > 0 &&
//     sht.ListObjects[1].Name.StartsWith("ORB_DD_").
//   - IsCustomDrilldownHyperlink(hyperlink): !string.IsNullOrEmpty(hyperlink.ScreenTip) &&
//     hyperlink.ScreenTip.IndexOf("CUSTOM DRILLDOWN", StringComparison.OrdinalIgnoreCase) >= 0
//     - as opposed to a plain journal-attachment hyperlink (no such ScreenTip), which is
//     routed instead to JournalAttachments.RunJournalAttachmentFlow.
// The host is also expected to compute headerLabel exactly like the old
// HandleCustomDrilldownHyperlink did before crossing the AppDomain boundary:
//   Excel.Range rng = sht.Range[hyperlink.SubAddress];
//   string headerLabel = ((Excel.Range)sht.Cells[5, rng.Column]).Value2?.ToString();
// and cellExternalAddress via that same rng's fully-qualified external address (e.g.
// "[Book1.xlsx]Sheet1!$C$7", via Excel.Range.Address[External:=true] /
// ExcelExternalRef.BuildExternalAddress) - only strings/primitives cross the AppDomain
// boundary (see IGLSenseAddin.OnRibbonAction's existing convention, which this mirrors).
//
// Re-pointed vs. the original (business logic/URLs/JSON shapes unchanged):
//   - namespace GLSense -> GLSense.Addin.Core; GLSense.Helpers/.Utilities/.Views ->
//     GLSense.Addin.Core.* equivalents.
//   - AppState.Instance.ExcelApp -> ServiceLocator.ExcelApp.
//   - LogUtility.* (static) -> ServiceLocator.Logger?.* (instance via context).
//   - Excel.ListObject tableObj / Excel.Range rng (previously received live from the
//     SheetFollowHyperlink event) are now re-derived here instead of passed in: tableObj
//     via ServiceLocator.ExcelApp.Worksheets[tableSheetName].ListObjects[1] (matching the
//     old sht.ListObjects[1] convention - see IsValidDrilldownSheet above), rng via
//     ExcelExternalRef.ResolveRangeWithContext(cellExternalAddress).Range. Excel.Worksheet
//     ws (previously AppState.Instance.ExcelApp.ActiveSheet) is likewise now
//     ServiceLocator.ExcelApp.Worksheets[tableSheetName], which is more robust than
//     ActiveSheet since it can't drift if focus changes between the event firing and this
//     method running.
//   - GLWaitWindow/CreateAndShowWaitWindow/InitializeWaitWindowAsync/
//     SafelyCloseWaitWindowAsync/ShowErrorMessageAsync/MessageWaitWindowAsync/
//     GuardLoginAndExcel: duplicated here as private statics - this project's established
//     per-file convention (see BalanceHighlighter.cs/RowVisibilityProcessor.cs/
//     RangeRefresher.cs/DrillCellHighlighter.cs for the same idiom) rather than a shared
//     helper class. GLWaitWindow is captured from a closure-local inside
//     WpfAppManager.InvokeOnWpfThread(Action) exactly like those files, since this
//     project's InvokeOnWpfThread has no Func<T> overload.
//   - CommonMethods.DisableExcelSettings()/EnableExcelSettings() and the CancellationHelper
//     scope, previously owned by the outer adxExcelAppEvents1_SheetFollowHyperlink handler
//     in the old monolith (which wrapped both the custom-drilldown AND journal-attachment
//     branches together), are now owned directly by RunCustomDrilldown, since that outer
//     handler's responsibilities are now split between the host (event classification) and
//     this entry point (everything else) - each of RunCustomDrilldown/
//     JournalAttachments.RunJournalAttachmentFlow owns its own Disable/Enable + cts scope.
//   - ApiHelper.ServerAPI, ApiResponseHelper.Parse<T>, JsonGlobals.Options,
//     AppConstants.RestSecure/AppConstants.value, DDDatatoWorksheet: all already ported -
//     used as-is, no changes needed.
using GLSense.Addin.Core.Helpers;
using GLSense.Addin.Core.Infrastructure;
using GLSense.Addin.Core.Utilities;
using GLSense.Addin.Core.Views;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.Json;
using System.Text.Json.Nodes;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Xml;
using System.Xml.Linq;
using Excel = Microsoft.Office.Interop.Excel;

namespace GLSense.Addin.Core.Drilldowns
{
    public static class CustomDrilldown
    {
        /// <summary>
        /// Entry point called (by a later pass) from AddinEntry.OnExcelEvent's
        /// SheetFollowHyperlink handler once the host has verified IsValidDrilldownSheet/
        /// IsCustomDrilldownHyperlink and resolved headerLabel. See file header for the
        /// exact preconditions this method assumes have already been checked, and for the
        /// exact meaning of each parameter.
        /// </summary>
        public static async Task RunCustomDrilldown(string tableSheetName, string cellExternalAddress, string headerLabel)
        {
            ServiceLocator.Logger?.LogDebug($"CustomDrilldown.RunCustomDrilldown started. tableSheetName='{tableSheetName}', cellExternalAddress='{cellExternalAddress}', headerLabel='{headerLabel}'.");

            if (!GuardLoginAndExcel())
                return;

            if (string.IsNullOrWhiteSpace(headerLabel))
            {
                ServiceLocator.Logger?.LogWarn("Exiting the hyperlink sub since the label header is null or empty");
                return;
            }

            GLWaitWindow win = null;
            using var ctsHelper = new CancellationHelper();
            CancellationToken token = ctsHelper.GetToken();

            CommonMethods.DisableExcelSettings();

            try
            {
                win = CreateAndShowWaitWindow(ctsHelper);
                await InitializeWaitWindowAsync(win, "Custom Drilldown", "GLSense custom drilldown...");
                ServiceLocator.Logger?.LogDebug("Custom Drilldown Process Started.");

                token.ThrowIfCancellationRequested();

                if (ServiceLocator.ExcelApp.Worksheets[tableSheetName] is not Excel.Worksheet ws)
                {
                    await ShowErrorMessageAsync(win, $"Worksheet '{tableSheetName}' not found.");
                    return;
                }

                if (ws.ListObjects == null || ws.ListObjects.Count == 0)
                {
                    await ShowErrorMessageAsync(win, $"No drilldown table found on worksheet '{tableSheetName}'.");
                    return;
                }

                var resolveResult = ExcelExternalRef.ResolveRangeWithContext(cellExternalAddress);
                Excel.Range rng = resolveResult?.Range;
                if (rng == null)
                {
                    await ShowErrorMessageAsync(win, $"Unable to resolve cell address '{cellExternalAddress}'.");
                    return;
                }

                Excel.ListObject tableObj = ws.ListObjects[1];

                await ProcessCustomDrilldownAsync(tableObj, rng, headerLabel, ws, win, token);
            }
            catch (OperationCanceledException)
            {
                ServiceLocator.Logger?.LogWarn("Custom Drilldown operation was canceled by the user.");
            }
            catch (Exception ex)
            {
                ServiceLocator.Logger?.LogException(ex, "CustomDrilldown.RunCustomDrilldown");
            }
            finally
            {
                await SafelyCloseWaitWindowAsync(win);
                CommonMethods.EnableExcelSettings();
            }
        }

        private static async Task ProcessCustomDrilldownAsync(Excel.ListObject tableObj, Excel.Range rng, string headerLabel, Excel.Worksheet ws, GLWaitWindow win, CancellationToken token)
        {
            ServiceLocator.Logger?.LogDebug($"CustomDrilldown.ProcessCustomDrilldownAsync started. sheet='{ws?.Name}', headerLabel='{headerLabel}'.");

            // Regression-pattern fix: this used to fetch a fresh ServiceLocator.ExcelApp.
            // ActiveWorkbook here - the same "COM object fetched fresh deep in an async
            // hyperlink/drilldown chain" pattern that threw an InvalidCastException crossing
            // the host<->Addin.Core AppDomain boundary in SanitizeSheetName (see that fix's
            // comment). Using ws.Parent instead avoids the risky re-fetch AND is more
            // correct: this method needs the workbook that owns `ws` (the sheet the
            // drilldown metadata actually lives on), not whatever workbook happens to be
            // "active" by the time this async continuation runs - those can differ if focus
            // moved to a different window during the preceding awaits.
            Excel.Workbook wb = ws?.Parent as Excel.Workbook;
            if (wb?.CustomXMLParts?.Count <= 0)
            {
                ServiceLocator.Logger?.LogDebug("CustomDrilldown.ProcessCustomDrilldownAsync: workbook has no CustomXMLParts, aborting.");
                return;
            }

            string xmlString = FindDrilldownXmlForSheet(wb, ws.Name);

            token.ThrowIfCancellationRequested();

            if (string.IsNullOrEmpty(xmlString))
            {
                string errorMsg = "Unable to find the drilldown meta " +
                    "object check if sheet renamed! If renamed then regenerate drilldown again to get meta object.";
                await ShowErrorMessageAsync(win, errorMsg);
                return;
            }

            await ExecuteCustomDrilldownAsync(tableObj, rng, headerLabel, xmlString, ws, win, token);
        }

        private static string FindDrilldownXmlForSheet(Excel.Workbook wb, string sheetName)
        {
            var cxps = wb.CustomXMLParts;
            string sheetNameEncoded = System.Net.WebUtility.HtmlEncode(sheetName);

            for (int i = cxps.Count; i >= 1; i--)
            {
                string xml = cxps[i].XML;
                if (string.IsNullOrEmpty(xml) || !xml.Contains("DRILLDOWNSHEET"))
                    continue;

                string shtNameInXml = System.Net.WebUtility.HtmlEncode(ExtractDrilldownSheet_XPath(xml));
                if (string.IsNullOrEmpty(shtNameInXml) || !shtNameInXml.Equals(sheetNameEncoded, StringComparison.OrdinalIgnoreCase))
                    continue;

                ServiceLocator.Logger?.LogDebug($"Found XML in internal memory location: {xml}");
                return System.Net.WebUtility.HtmlDecode(xml);
            }

            return null;
        }

        private static string ExtractDrilldownSheet_XPath(string xml)
        {
            if (string.IsNullOrWhiteSpace(xml)) return null;

            var doc = new XmlDocument { XmlResolver = null };

            try
            {
                doc.LoadXml(xml);
            }
            catch (XmlException ex)
            {
                ServiceLocator.Logger?.LogWarn($"CustomDrilldown.ExtractDrilldownSheet_XPath: failed to parse drilldown XML: {ex.Message}");
                return null;
            }

            try
            {
                var nav = doc.CreateNavigator();
                var node = nav.SelectSingleNode("//DRILLDOWNSHEET/text()[1]");
                var value = node?.Value?.Trim();
                return string.IsNullOrEmpty(value) ? null : value;
            }
            catch (Exception ex)
            {
                ServiceLocator.Logger?.LogException(ex);
                return null;
            }
        }

        private static async Task ExecuteCustomDrilldownAsync(Excel.ListObject tableObj, Excel.Range rng, string headerLabel, string xmlString, Excel.Worksheet ws, GLWaitWindow win, CancellationToken token)
        {
            try
            {
                XDocument doc = XDocument.Parse(xmlString);
                XElement columnNameElement = doc.Descendants("COLUMNNAME")
                    .FirstOrDefault(x => x.Attribute("Name")?.Value == headerLabel);

                if (columnNameElement == null)
                    return;

                token.ThrowIfCancellationRequested();

                string jsonData = columnNameElement.Value.Trim();
                ServiceLocator.Logger?.LogDebug($"JSON to Parse: {jsonData}");

                JsonNode jsonObject = JsonNode.Parse(jsonData);
                UpdateJsonParameters(jsonObject, tableObj, rng, ws, token);

                string updatedJsonString = JsonSerializer.Serialize(jsonObject, JsonGlobals.Options);
                ServiceLocator.Logger?.LogDebug($"Request Sent to Server with JSON Body: {updatedJsonString}");

                string responseString = await SendCustomDrilldownRequestAsync(updatedJsonString, token);
                token.ThrowIfCancellationRequested();

                ServiceLocator.Logger?.LogDebug($"Response Received: {responseString}");

                var result = ApiResponseHelper.Parse<JsonElement>(responseString, JsonGlobals.Options);

                if (!result.IsSuccess)
                {
                    string errorMsg = responseString?.Contains("A task was canceled.") == true
                        ? "Task has been canceled."
                        : result.ErrorMessage;

                    await ShowErrorMessageAsync(win, errorMsg);
                    return;
                }

                await WriteCustomDrilldownToWorksheet(ws, rng, jsonData, updatedJsonString, responseString, win, token);
            }
            catch (Exception ex)
            {
                ServiceLocator.Logger?.LogException(ex, "CustomDrilldown.ExecuteCustomDrilldownAsync");
            }
        }

        private static void UpdateJsonParameters(JsonNode jsonObject, Excel.ListObject tableObj, Excel.Range rng, Excel.Worksheet ws, CancellationToken token)
        {
            JsonArray parameters = jsonObject["parameters"].AsArray();

            foreach (JsonObject param in parameters.Cast<JsonObject>())
            {
                token.ThrowIfCancellationRequested();

                string paramType = param["type"]?.GetValue<string>();

                if (paramType == "COLUMN")
                {
                    UpdateColumnParameter(param, tableObj, rng, ws, token);
                }

                param.Remove("id");
            }
        }

        private static void UpdateColumnParameter(JsonObject param, Excel.ListObject tableObj, Excel.Range rng, Excel.Worksheet ws, CancellationToken token)
        {
            string colName = param[AppConstants.value]?.GetValue<string>();

            try
            {
                int colNumber = FindColumnNumber(tableObj, colName, token);

                if (colNumber > 0)
                {
                    Excel.Range rng1 = ws.Cells[rng.Row, colNumber] as Excel.Range;
                    object cellValue = rng1.Value2;
                    string val = cellValue?.ToString() ?? "";
                    param[AppConstants.value] = val;
                }
            }
            catch (Exception ex)
            {
                ServiceLocator.Logger?.LogException(ex, $"CustomDrilldown.UpdateColumnParameter: column '{colName}'");
            }
        }

        private static int FindColumnNumber(Excel.ListObject tableObj, string colName, CancellationToken token)
        {
            foreach (Excel.Range hdrRange in tableObj.HeaderRowRange)
            {
                token.ThrowIfCancellationRequested();
                object offsetValue = hdrRange.Offset[-1, 0].Value;
                if (offsetValue != null && offsetValue.ToString().Trim() == colName?.Trim())
                {
                    return hdrRange.Column;
                }
            }
            return 0;
        }

        private static async Task<string> SendCustomDrilldownRequestAsync(string jsonPayload, CancellationToken token)
        {
            string url = $"{AppState.Instance.LoginUrl}{AppConstants.RestSecure}custom-drilldown?cubeId={AppState.Instance.SelectedCube.CubeId}";
            ServiceLocator.Logger?.LogDebug($"CustomDrilldown.SendCustomDrilldownRequestAsync: POST {url}");
            string response = await ApiHelper.ServerAPI(url, "JSON", jsonPayload, "POST", token);
            ServiceLocator.Logger?.LogDebug("CustomDrilldown.SendCustomDrilldownRequestAsync: response received from server.");
            return response;
        }

        private static async Task WriteCustomDrilldownToWorksheet(Excel.Worksheet ws, Excel.Range rng, string jsonData, string updatedJsonString, string responseString, GLWaitWindow win, CancellationToken token)
        {
            string obj1 = $"{ws.Name}_CM_{rng.Address[false, false, Excel.XlReferenceStyle.xlA1]}";
            string obj2 = rng.Address[true, true, Excel.XlReferenceStyle.xlA1, true];
            string obj3 = DateTime.Now.ToString("dd-MMM-yyyy hh:mm:ss tt");
            string obj4 = "CM";
            string obj5 = MergeJsonAndGenerateString(jsonData, updatedJsonString);
            string finalData = $"{obj1}*{obj2}*{obj3}*{obj4}*{obj5}";

            var dataToSheet = new DDDatatoWorksheet(ServiceLocator.ExcelApp, responseString, "CM", finalData, token, win);
            await dataToSheet.DD_DatetoWorksheet();

            token.ThrowIfCancellationRequested();
        }

        private static string MergeJsonAndGenerateString(string jsonNames, string jsonValues)
        {
            JsonNode namesNode = JsonNode.Parse(jsonNames);
            JsonNode valuesNode = JsonNode.Parse(jsonValues);

            string identifierName = namesNode["identifierName"]?.GetValue<string>();

            var paramMap = new Dictionary<string, string>();
            foreach (JsonObject param in namesNode["parameters"].AsArray().Cast<JsonObject>())
            {
                string name = param["name"]?.GetValue<string>();
                string value = param[AppConstants.value]?.GetValue<string>();
                if (name != null)
                    paramMap[name] = value;
            }

            var result = new StringBuilder();
            result.Append($"Report Name:= {identifierName}");

            foreach (JsonObject param in valuesNode["parameters"].AsArray().Cast<JsonObject>())
            {
                string paramName = param["name"]?.GetValue<string>();
                string paramValue = param[AppConstants.value]?.GetValue<string>();
                string paramType = param["type"]?.GetValue<string>();

                if (paramMap.TryGetValue(paramName, out string mappedName))
                {
                    if (paramType == "CONSTANT")
                        result.Append($", Constant:= {paramValue}");
                    else
                        result.Append($", {mappedName}:= {paramValue}");
                }
            }

            return result.ToString();
        }

        // ---------------------------------------------------------------------
        // Local helper duplicates (this project's per-file convention - see
        // BalanceHighlighter.cs/RowVisibilityProcessor.cs/RangeRefresher.cs/
        // DrillCellHighlighter.cs for the same idiom).
        // ---------------------------------------------------------------------
        private static bool GuardLoginAndExcel() => AppState.Instance.IsLoginCompleted && ServiceLocator.ExcelApp != null;

        private static GLWaitWindow CreateAndShowWaitWindow(CancellationHelper cts)
        {
            try
            {
                GLWaitWindow win = null;
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
                return null;
            }
        }

        private static Task InitializeWaitWindowAsync(GLWaitWindow win, string title, string message)
        {
            if (win == null || win.Dispatcher == null) return Task.CompletedTask;
            try
            {
                return win.Dispatcher.InvokeAsync(() =>
                {
                    win.SetProcessTitle(title);
                    win.SetProcessMessage(message);
                }).Task;
            }
            catch (TaskCanceledException)
            {
                ServiceLocator.Logger?.LogDebug("CustomDrilldown.InitializeWaitWindowAsync: dispatcher invoke was cancelled (window likely closing).");
                return Task.CompletedTask;
            }
            catch (Exception ex)
            {
                ServiceLocator.Logger?.LogException(ex, "CustomDrilldown.InitializeWaitWindowAsync");
                return Task.CompletedTask;
            }
        }

        private static async Task SafelyCloseWaitWindowAsync(GLWaitWindow win)
        {
            if (win == null) return;
            try
            {
                if (win.Dispatcher.CheckAccess())
                    win.RequestClose();
                await win.Dispatcher.InvokeAsync(() => win.RequestClose());
            }
            catch (Exception ex)
            {
                ServiceLocator.Logger?.LogError($"Error closing wait window: {ex.Message}");
            }
        }

        private static Task MessageWaitWindowAsync(GLWaitWindow win, string message)
        {
            if (win == null || win.Dispatcher == null) return Task.CompletedTask;
            try
            {
                return win.Dispatcher.InvokeAsync(() => win.SetProcessMessage(message)).Task;
            }
            catch (TaskCanceledException)
            {
                ServiceLocator.Logger?.LogDebug("CustomDrilldown.MessageWaitWindowAsync: dispatcher invoke was cancelled (window likely closing).");
                return Task.CompletedTask;
            }
            catch (Exception ex)
            {
                ServiceLocator.Logger?.LogException(ex, "CustomDrilldown.MessageWaitWindowAsync");
                return Task.CompletedTask;
            }
        }

        private static async Task ShowErrorMessageAsync(GLWaitWindow win, string message)
        {
            await SafelyCloseWaitWindowAsync(win);
            CommonFunctions.GLSenseMessage(message, MessageBoxImage.Error, MessageBoxButton.OK);
        }
    }
}
