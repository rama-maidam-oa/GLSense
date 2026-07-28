using GLSense.Helpers;
using GLSense.Models;
using GLSense.Repositories;
using GLSense.Service;
using GLSense.Utilities;
using GLSense.Views;
using MahApps.Metro.Controls;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Net;
using System.Text.Json;
using System.Text.Json.Nodes;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Windows.Interop;
using Excel = Microsoft.Office.Interop.Excel;

namespace GLSense.Drilldowns
{
#nullable enable
    public class DrilldownBl
    {
        //After coming from vacation refactor the complete code
        private Excel.Application ExcelApp { get; }
        private Excel.Workbook? BlWorbook { get; set; }
        private Excel.Worksheet? BlWorksheet { get; set; }
        private Excel.Range? BlRange { get; set; }
        private CancellationHelper? _ctsHelper;
        private CancellationToken Token => _ctsHelper?.GetToken() ?? default;
        private ExternalResolveResult? ExternalResolveResult { get; set; }
        private readonly string _blAddress;
        private readonly string _DDType;
        private static GLWaitWindow? Win { get; set; }
        public DrilldownBl(Excel.Application xlapp, string rngAddress, string dDType)
        {
            ExcelApp = xlapp;
            _blAddress = rngAddress;
            _DDType = dDType;

            ExternalResolveResult = ExcelExternalRef.ResolveRangeWithContext(_blAddress);
            BlRange = ExternalResolveResult.Range;

            if (BlRange == null)
            {
                BlWorksheet = null;
                BlWorbook = null;
            }
            else
            {
                BlWorksheet = ExternalResolveResult.Worksheet;
                BlWorbook = ExternalResolveResult.Workbook;
            }
        }
        private void GuardValidInputs()
        {
            if (BlRange == null || BlWorksheet == null || BlWorbook == null)
                throw new InvalidOperationException("Required ranges/worksheets missing.");

            if (!BLFormulaExists(BlRange))
                throw new InvalidOperationException("No GLSense balance formulas found.");
        }
        private static GLWaitWindow? CreateAndShowProgressWindow(CancellationHelper cts)
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
                        LogUtility.LogException(ex);
                        return null;
                    }
                });
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex);
            }
            return null;
        }
        private async Task ValidateAndFixFormulasAsync(GLWaitWindow progress)
        {
            if (!AppState.Instance.VersionCheck)
            {
                LogUtility.LogDebug("DrilldownBl.ValidateAndFixFormulasAsync: VersionCheck disabled, skipping formula validation.");
                return;
            }

            await SetMessageAsync("Validating Balance formulas...");

            var invalidFormulas = CommonFunctions.NotValidBalancesDict(_blAddress);
            if (invalidFormulas?.Count > 0)
            {
                bool shouldFix = await EnsureCompatibilityAsync(progress);
                if (shouldFix)
                {
                    await SetMessageAsync("Correcting invalid balance formulas...");
                    CommonMethods.BalanceFormulas_Updation(invalidFormulas);
                }
                else
                {
                    await SetMessageAsync("Skipping formulas updation...");
                }
            }
        }
        private string CreateBalancePayload()
        {
            var balanceDto = BalanceDto.CreateBalanceDto(_blAddress);
            return JsonSerializer.Serialize(balanceDto, JsonGlobals.Options);
        }
        private static Task SetMessageAsync(string message)
        {
            if (string.IsNullOrWhiteSpace(message) || Win == null) return Task.CompletedTask;

            // Fire-and-forget: progress label update only, do not introduce a
            // suspend point here (callers may run on threads with no captured
            // SynchronizationContext, e.g. background worker threads).
            _ = Win.Dispatcher.InvokeAsync(() => Win.SetProcessMessage(message));
            return Task.CompletedTask;
        }
        public async Task ProcessBLDrilldown()
        {
            LogUtility.LogDebug($"DrilldownBl.ProcessBLDrilldown started. Address={_blAddress}, DDType={_DDType}");
            try
            {
                _ctsHelper = new CancellationHelper();

                GuardValidInputs();
                CommonMethods.DisableExcelSettings();

                Win = CreateAndShowProgressWindow(_ctsHelper ?? new());

                if (Win == null)
                {
                    LogUtility.LogWarn("Unable to set progress window");
                    return;
                }

                await ValidateAndFixFormulasAsync(Win);

                await SetMessageAsync("Creating balance dto...");
                string jsonPayload = CreateBalancePayload();

                var totalRange = CommonFunctions.GetBalanceTotalRange(_blAddress);
                if (totalRange == null)
                {
                    await ShowMessageAsync("Failed to get balances range. Check the excel logs.", MessageBoxIcon.Error);
                    return;
                }

                await Balance_Drilldown(totalRange, jsonPayload, _DDType);
            }
            catch (OperationCanceledException)
            {
                await ShowMessageAsync("Operation Cancelled!", MessageBoxIcon.Warning);
                LogUtility.LogWarn("BL Drilldown operation was cancelled by the user.");
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex);
                await ShowMessageAsync($"An unexpected error occurred.{Environment.NewLine}{ex.Message}", MessageBoxIcon.Error);
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
                    // Swallow dispose exceptions (Excel COM weirdness) but still log for diagnostics.
                    LogUtility.LogWarn($"DrilldownBl.ProcessBLDrilldown: exception disposing CancellationHelper (ignored): {ex.Message}");
                }
                CommonMethods.EnableExcelSettings();
                await SafelyCloseWindowAsync();
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
                await Win.Dispatcher.InvokeAsync(() => Win.RequestClose());
            }
            catch (Exception ex)
            {
                LogUtility.LogError($"Error closing wait window: {ex.Message}");
            }
        }
        private static async Task<bool> EnsureCompatibilityAsync(GLWaitWindow win)
        {

            string quest =
                "This GLSense workbook is not compatible with the GLSense version.\n" +
                "Do you wish to make it compatible?\n\n" +
                "Note: This can take some time depending on the number of GLSense formulas in the workbook.";


            bool? action = await win.ShowConfirmToastAsync(quest);

            if (action == true)
            {
                // YES
                return true;
            }
            else if (action == false)
            {
                // NO
                return false;
            }
            else // action == null
            {
                // CANCEL (or dialog dismissed)
                return false;
            }
        }
        private sealed class DrilldownInfo
        {
            public string? Title { get; set; }
            public string? Address { get; set; }
            public string? WorksheetName { get; set; }
            public string? Timestamp { get; set; }
            public string? DDType { get; set; }

            public override string ToString()
            {
                return $"{WorksheetName}*{Address}*{Timestamp}*{DDType}*{Title}";
            }
        }
        private DrilldownInfo BuildDrilldownInfo(Excel.Range balanceRange, string ddType)
        {
            Token.ThrowIfCancellationRequested();

            int balancesCount = BalanceFormulaCountInCells(balanceRange);
            int totalCount = balanceRange.Cells.Count;

            Token.ThrowIfCancellationRequested();

            string title = GetDrilldownTitle(balanceRange, balancesCount, totalCount);
            string address = balanceRange.Address[true, true, Excel.XlReferenceStyle.xlA1, true];
            string currentCell = ExcelApp.ActiveCell.Address[false, false, Excel.XlReferenceStyle.xlA1, false];

            Token.ThrowIfCancellationRequested();

            string sheetName = balanceRange.Worksheet.Name;
            string worksheetName = totalCount > 1
                ? $"{sheetName}_{ddType}_{currentCell} +"
                : $"{sheetName}_{ddType}_{currentCell}";

            worksheetName = CommonFunctions.SanitizeSheetName(worksheetName);

            return new DrilldownInfo
            {
                Title = title,
                Address = address,
                WorksheetName = worksheetName,
                Timestamp = DateTime.Now.ToString("dd-MMM-yyyy hh:mm:ss tt"),
                DDType = ddType
            };
        }
        private string GetDrilldownTitle(Excel.Range balanceRange, int balancesCount, int totalCount)
        {
            if (balancesCount > 1) return "Multiple Balance Formulas in the cell.";
            if (totalCount >= 2) return "Multi Select.";

            Token.ThrowIfCancellationRequested();

            string formula = balanceRange.Formula.ToString();
            var funcValues = CommonFunctions.FormulaValues(formula);

            if (funcValues == null || funcValues.Count <= 11) return string.Empty;

            string ledgerName = funcValues[1];
            string ldgerNameNormalized = NormalizeStrings(ledgerName);
            var ledgerRecord = AppState.Instance.SelectedCube.Ledgers;

            long ledgerId = 0;

            if (ledgerRecord == null || ledgerRecord.Count == 0)
            {
                return string.Empty;
            }
            else
            {
                var ledgerNames = ldgerNameNormalized.ToString().Split(';').Select(name => name.Trim());
                var matchingLedgers = ledgerRecord.Where(l => ledgerNames.Contains(l.LedgerName));
                ledgerId = matchingLedgers.FirstOrDefault()?.LedgerId ?? 0; // Default to 0 if no match
            }

            if (ledgerId == 0)
            {
                return string.Empty;
            }

            var repo = new DataRepository();
            var segs = repo.GetSegments(AppState.Instance.SelectedCube.CubeId, ledgerId);
            int accountIndex = segs.Select((seg, idx) => new { seg, idx })
                                  .FirstOrDefault(x => string.Equals(x.seg.SegmentName, "Account", StringComparison.OrdinalIgnoreCase))
                                  ?.idx ?? -1;

            Token.ThrowIfCancellationRequested();

            string actualFlag = funcValues[7]?.Replace("\"", "") ?? string.Empty;
            string actualFlagValue = funcValues[8]?.Replace("\"", "") ?? string.Empty;

            string aF = string.Empty;
            if (!string.IsNullOrWhiteSpace(actualFlag))
            {
                switch (actualFlag.ToLowerInvariant())
                {
                    case "budget":
                    case "b":
                        aF = "Budget:= " + actualFlagValue;
                        break;
                    case "encumbrance":
                    case "e":
                    case "actual+encumbrance":
                    case "a+e":
                        aF = "Encumbrance:= " + actualFlagValue;
                        break;
                    case "actual":
                    case "a":
                    case null:
                    case "":
                        break;
                    default:
                        break;
                }
            }

            var parts = new List<string>
            {
                $"Period:= {funcValues[3]?.Replace("\"", "")}",
                $"Balance Type:= {funcValues[4]?.Replace("\"", "")}",
                $"Currency Code:= {funcValues[5]?.Replace("\"", "")}",
                $"Actual Flag:= {funcValues[7]?.Replace("\"", "")}"
            };

            if (!string.IsNullOrWhiteSpace(actualFlagValue))
            {
                parts.Add(aF);
            }

            parts.Add($"Journal Source:= {funcValues[9]?.Replace("\"", "")}");
            parts.Add($"Journal Category:= {funcValues[10]?.Replace("\"", "")}");

            // Add account description
            if (accountIndex >= 0)
            {
                string segments = funcValues[11];
                string accountValue = ExtractAccountValue(segments, funcValues, accountIndex);
                parts.Add(Account_Desc(accountValue));
            }

            return string.Join(", ", parts);
        }
        private static string ExtractAccountValue(string segments, List<string> funcValues, int accountIndex)
        {
            if (segments.Contains(";"))
            {
                var splits = segments.Split(';');
                return splits[accountIndex].Trim().Replace("\"", "");
            }
            return funcValues[11 + accountIndex];
        }
        private static string BuildApiUrl(DrilldownInfo info, string ddType)
        {
            string encoded = WebUtility.UrlEncode($"{info.WorksheetName}*{info.Address}*{info.Timestamp}*{info.DDType}*{info.Title}");
            long cubeId = AppState.Instance.SelectedCube.CubeId;

            return ddType switch
            {
                "BL" => $"{AppState.Instance.LoginUrl}{AppConstants.RestSecure}balance-drilldown?cubeId={cubeId}&jobName={ddType}&jobDescription={encoded}",
                "BL_JL" => $"{AppState.Instance.LoginUrl}{AppConstants.RestSecure}balance-journal-drilldown?cubeId={cubeId}&jobName={ddType}&jobDescription={encoded}",
                "BL_SL" => $"{AppState.Instance.LoginUrl}{AppConstants.RestSecure}balance-subledger-drilldown?cubeId={cubeId}&jobName={ddType}&jobDescription={encoded}",
                _ => $"{AppState.Instance.LoginUrl}{AppConstants.RestSecure}total-drilldown?cubeId={cubeId}&jobName={ddType}&jobDescription={encoded}"
            };
        }
        private async Task<string> FetchDrilldownDataAsync(string apiUrl, string httpPostBody)
        {
            if (Win == null) return string.Empty;

            Token.ThrowIfCancellationRequested();

            _ = Win.Dispatcher.InvokeAsync(() => Win.SetProcessMessage("Sending request to server..."));
            LogUtility.LogDebug($"DrilldownBl.FetchDrilldownDataAsync: sending request to {apiUrl}");
            var response = await ApiHelper.ServerAPI(apiUrl, "JSON", httpPostBody, "POST", Token);

            Token.ThrowIfCancellationRequested();

            _ = Win.Dispatcher.InvokeAsync(() => Win.SetProcessMessage("Response received..."));
            LogUtility.LogDebug($"DrilldownBl.FetchDrilldownDataAsync: response received from {apiUrl}. Length={response?.ToString()?.Length ?? 0}");

            return response?.ToString()?.Trim() ?? string.Empty;
        }
        private async Task ProcessDrilldownResponseAsync(
                string json,
                DrilldownInfo info,
                string ddType)
        {

            try
            {
                var result = ApiResponseHelper.Parse<JsonElement>(json, JsonGlobals.Options);

                if (!result.IsSuccess)
                {
                    await ShowMessageAsync(
                        result.ErrorMessage ?? "Unknown error",
                        MessageBoxIcon.Error);
                    return;
                }

                var root = result.Value;

                // ---------------------------------------------
                // Extract message safely
                // ---------------------------------------------

                string msg;
                if (root.TryGetProperty("msg", out var msgEl))
                {
                    msg = msgEl.GetString() ?? string.Empty;
                }
                else if (root.TryGetProperty("message", out var messageEl))
                {
                    msg = messageEl.GetString() ?? string.Empty;
                }
                else
                {
                    msg = string.Empty;
                }

                // ---------------------------------------------
                // Background Processing
                // ---------------------------------------------

                if (await HandleBackgroundProcessingAsync(msg))
                    return;

                // ---------------------------------------------
                // Records detection (auto-safe)
                // ---------------------------------------------

                string responseMsg = msg; // Store the original message

                if (!TryGetRecordsNode(root, out JsonElement recordsNode))
                {
                    await ShowMessageAsync(
                        string.IsNullOrWhiteSpace(responseMsg)
                            ? "No records returned. Unable to read response message."
                            : responseMsg,
                        MessageBoxIcon.Error);
                    return;
                }

                int recordCount = GetRecordCount(recordsNode);

                if (recordCount == 0)
                {
                    await ShowMessageAsync(
                        string.IsNullOrWhiteSpace(responseMsg)
                            ? "No records returned.Unable to read response message."
                            : responseMsg,
                        MessageBoxIcon.Error);
                    return;
                }

                // ---------------------------------------------
                // SUCCESS - Write to worksheet
                // ---------------------------------------------

                var dataToSheet = new DDDatatoWorksheet(
                    ExcelApp,
                    json,
                    ddType,
                    info.ToString(),
                    Token,
                    Win);

                await dataToSheet.DD_DatetoWorksheet();
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex);
            }


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
        private static int GetRecordCount(JsonElement recordsNode)
        {
            return recordsNode.ValueKind switch
            {
                JsonValueKind.Array => recordsNode.GetArrayLength(),
                JsonValueKind.Object => recordsNode.EnumerateObject().Count(),
                _ => 0
            };
        }

        private async Task<bool> HandleBackgroundProcessingAsync(string msg)
        {
            if (msg.IndexOf(
                    "Drilldown request is being processed in the background",
                    StringComparison.OrdinalIgnoreCase) < 0)
                return false;

            Token.ThrowIfCancellationRequested();

            string displayMsg = msg;
            string jobId = ExtractJobId(msg);

            if (!string.IsNullOrWhiteSpace(jobId))
            {
                LogUtility.LogDebug($"DrilldownBl.HandleBackgroundProcessingAsync: drilldown queued in background with jobId={jobId}");
                string rangeName = "GLSense_DD_" + jobId;

                if (!CommonFunctions.NameRangeExists(rangeName))
                {
                    ExcelApp.ActiveWorkbook.Names.Add(
                        Name: rangeName,
                        RefersToR1C1: jobId);
                }

                displayMsg += Environment.NewLine +
                              "Launch process window to check the status.";
            }

            await ShowMessageAsync(displayMsg, MessageBoxIcon.Information);

            return true;
        }


        private static string ExtractJobId(string msg)
        {
            var parts = msg.Split(' ');
            if (parts.Length == 0) return string.Empty;

            string lastPart = parts[parts.Length - 1].TrimEnd('.');  // ✅ Works everywhere
            return lastPart;
        }
        private async Task Balance_Drilldown(Excel.Range BalanceRange, string HTTPPostBody, string DDType)
        {
            try
            {
                Token.ThrowIfCancellationRequested();

                var drilldownInfo = BuildDrilldownInfo(BalanceRange, DDType);
                string apiUrl = BuildApiUrl(drilldownInfo, DDType);

                Token.ThrowIfCancellationRequested();

                if (Win == null) return;

                string jsonResponse = await FetchDrilldownDataAsync(apiUrl, HTTPPostBody);
                if (string.IsNullOrWhiteSpace(jsonResponse)) return;

                await ProcessDrilldownResponseAsync(jsonResponse, drilldownInfo, DDType);
            }
            catch (OperationCanceledException)
            {
                LogUtility.LogError("Balance drilldown operation was cancelled.");
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex);
            }
        }
        private static async Task ShowMessageAsync(string message, MessageBoxIcon msgIcon)
        {
            if (string.IsNullOrWhiteSpace(message) || Win == null) return;

            await SafelyCloseWindowAsync();
            CommonFunctions.GLSenseMessage(message, msgIcon, MessageBoxButtons.OK);
        }

        private string Account_Desc(string sValue)
        {
            try
            {
                string newvalue = NormalizeStrings(sValue);
                if (sValue.Contains(",") || sValue.Contains('%'))
                {
                    return newvalue;
                }
                else if (sValue.Contains("--"))
                {
                    newvalue = NormalizeStrings(sValue.Replace("--", ""));
                }

                Token.ThrowIfCancellationRequested();

                var segmentValues = LoadSegmentValues(AppState.Instance.SelectedLedger.LedgerName);
                string finalStrin = string.Empty;

                if (newvalue.Contains("|"))
                {
                    var newValueSplit = newvalue.Split('|');

                    var value1 = segmentValues.FirstOrDefault(sv =>
                                        sv.SegmentName.Equals("ACCOUNT", StringComparison.OrdinalIgnoreCase) &&
                                        sv.SegmentValue.Equals(newValueSplit[0], StringComparison.OrdinalIgnoreCase));
                    var value2 = segmentValues.FirstOrDefault(sv =>
                                sv.SegmentName.Equals("ACCOUNT", StringComparison.OrdinalIgnoreCase) &&
                                sv.SegmentValue.Equals(newValueSplit[1], StringComparison.OrdinalIgnoreCase));

                    finalStrin = $"Account: {newValueSplit[0]}~{value1?.Description ?? string.Empty}||{newValueSplit[1]}~{value2?.Description ?? string.Empty}";
                }
                else
                {
                    var value = segmentValues.FirstOrDefault(sv =>
                                sv.SegmentName.Equals("ACCOUNT", StringComparison.OrdinalIgnoreCase) &&
                                sv.SegmentValue.Equals(newvalue, StringComparison.OrdinalIgnoreCase));

                    finalStrin = $"Account: {newvalue}~{value?.Description ?? string.Empty}";
                }

                return finalStrin;
            }
            catch (Exception ex1)
            {
                LogUtility.LogException(ex1);
            }
            return string.Empty;
        }
        static string NormalizeStrings(string s)
        {
            if (string.IsNullOrWhiteSpace(s)) return string.Empty;

            // Convert \" to ", \\ to \ (handles common escaped inputs)
            // Regex.Unescape handles most escaped sequences
            var unescaped = Regex.Unescape(s);

            // Trim whitespace and any leading/trailing quotes
            var trimmed = unescaped.Trim().Trim('"', '“', '”', '\'');

            return trimmed;
        }
        private static ObservableCollection<SegmentValueModel> LoadSegmentValues(string ledgerName)
        {
            try
            {
                var task = Task.Run(() =>
                {
                    var dataService = ServiceLocator.SegmentDataService;
                    return dataService.GetSegmentValues(ledgerName);
                });

                if (task.Wait(TimeSpan.FromSeconds(180)))
                {
                    return task.Result;
                }
                else
                {
                    throw new TimeoutException("Timeout loading segment values from service");
                }
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex, "Failed to load segment values");
                return new ObservableCollection<SegmentValueModel>();
            }
        }
        private int BalanceFormulaCountInCells(Excel.Range inputRange)
        {
            int FncCount = 0;
            try
            {
                foreach (Excel.Range cell in inputRange.Cells)
                {
                    Token.ThrowIfCancellationRequested();

                    bool hasFormula = (bool)cell.HasFormula;
                    if (hasFormula && cell.Formula?.ToString().Contains(AppConstants.glBal) == true)
                    {
                        FncCount = CommonFunctions.GetBalancesCountInCells(cell.Formula.ToString());

                        if (FncCount > 1)
                            return FncCount;

                    }
                }
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex);
            }
            return FncCount;
        }
        private bool BLFormulaExists(Excel.Range rng)
        {
            try
            {
                foreach (Excel.Range cell in rng.Cells)
                {
                    Token.ThrowIfCancellationRequested();

                    bool hasFormula = (cell.HasFormula is bool hf) && hf;
                    string formula;
                    if (cell.Formula is string s2)
                    {
                        formula = s2;
                    }
                    else if (cell.Formula is string s)
                    {
                        formula = s;
                    }
                    else
                    {
                        formula = string.Empty;
                    }
                    if (hasFormula && formula.IndexOf(AppConstants.glBal, StringComparison.OrdinalIgnoreCase) >= 0)
                    {
                        return true;
                    }
                }
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex);
            }
            return false;
        }
    }
#nullable disable
}
