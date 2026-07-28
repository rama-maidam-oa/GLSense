// BalanceRefresh.cs in GLSense.Addin.Core
// Ported from GLSense\Drilldowns\BalanceRefresh.cs (FinalWorkingCode). Public API:
// RefreshingBalancesAsync(refreshType, refreshMode) - wired now (RibRefreshAll/
// RibRefreshBook, see AddinEntry.cs) - and SubmitSnapAsync(refreshMode) - ported here
// but left UNWIRED; Group G's task explicitly says "reuse ported BalanceRefresh.cs" for
// RibSnapSubmit/RibSnapShot/RibSnapWorksheet/RibSnapWorkbook, so no ribbon dispatch case
// is added for it in this pass.
//
// Group F also required porting 3 further transitive dependencies discovered via grep
// while reading this file (see their own header comments for re-pointing details):
//   - Drilldowns\BulkRefreshProcess.cs (the actual bulk refresh/snapshot engine this
//     class delegates to via BulkRefreshProcess.StartRefreshing).
//   - Utilities\CommonFunctions.NotValidBalancesDict (previously deferred; now ported -
//     see CommonFunctions.cs's updated header comment) - used by
//     ValidateAndUpdateFormulasIfNeededAsync below.
//   - Helpers\SafeDeleteHelper/QuarantineDeleteHelper, Helpers\SnapshotDialogHelper -
//     used by ClearFiles/TryDeleteFileWithFallbackAsync and PromptForSnapshotPathAsync
//     respectively.
//
// Re-pointed vs. the original (business logic/validation order/API URLs/response
// handling unchanged):
//   - GLSense.Helpers/.Utilities/.Views -> GLSense.Addin.Core.Helpers/.Utilities/.Views.
//   - LogUtility.* (static) -> ServiceLocator.Logger?.*.
//   - AppPaths.TempFilesPath -> ServiceLocator.Paths.Temp.
//   - System.Windows.Forms MessageBoxIcon/MessageBoxButtons -> System.Windows
//     MessageBoxImage/MessageBoxButton (this project has no WinForms reference, same as
//     every other Group E/F file).
//   - GLWaitWindow now derives from BaseWindow: win.ShowWithOwner(hwnd) -> win.Show()
//     (Excel owner set automatically via ServiceLocator.ExcelHandle). CreateAndShow-
//     ProgressWindow rewritten to the WpfAppManager.InvokeOnWpfThread(Action)-with-
//     captured-local pattern (InvokeOnWpfThread has no Func<T> overload in this project -
//     same adaptation DD_JL.cs/DrillCellHighlighter.cs already made).
//   - _state.ExcelApp (AppState.Instance.ExcelApp in the original) -> ServiceLocator.
//     ExcelApp (this project's AppState has no ExcelApp field).
//   - BalanceRefresh's own ExcelStaDispatcher private helper (marshals calls back to
//     Excel's STA thread via a captured SynchronizationContext) is ported as-is - this is
//     an existing, deliberate thread-safety mechanism, not something to replace.
//
// 2026-07-15 fix: ValidateAndUpdateFormulasIfNeededAsync was passing just
// ServiceLocator.ExcelApp.ActiveCell.Address to CommonFunctions.NotValidBalancesDict -
// re-synced against the FinalWorkingCode fix, which instead builds a proper external,
// range-qualified address from the full Selection via ExcelExternalRef.BuildExternalAddress
// (same helper DD_JL.cs/DD_BL.cs already use), so invalid formulas across a multi-cell
// selection are all caught, not just the one active cell.
using GLSense.Addin.Core.Helpers;
using GLSense.Addin.Core.Infrastructure;
using GLSense.Addin.Core.Utilities;
using GLSense.Addin.Core.Views;
using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Net;
using System.Text.Json;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Threading;
using Excel = Microsoft.Office.Interop.Excel;

namespace GLSense.Addin.Core.Drilldowns
{
#nullable enable
    public class BalanceRefresh
    {
        private readonly AppState _state;
        private readonly SynchronizationContext? _excelContext;

        private GLWaitWindow? Win { get; set; }
        private Excel.Application? ExcelApp { get; set; }
        private Excel.Workbook? BrWorbook { get; set; }
        private Excel.Worksheet? BrWorksheet { get; set; }

        private CancellationHelper? _ctsHelper;
        private CancellationToken Token => _ctsHelper?.GetToken() ?? default;

        private string _Title = string.Empty;

        private static readonly Dictionary<string, string> TitleLookup
            = new(StringComparer.OrdinalIgnoreCase)
            {
                ["Refresh-Sheet"] = "Worksheet Refresh",
                ["Refresh-Book"] = "Workbook Refresh",
                ["Snapshot-Sheet"] = "Worksheet Snapshot",
                ["Snapshot-Book"] = "Workbook Snapshot"
            };

        public BalanceRefresh(AppState? state = null, SynchronizationContext? excelContext = null)
        {
            _state = state ?? AppState.Instance;
            _excelContext = excelContext ?? SynchronizationContext.Current;
        }

        public static Task SubmitSnapAsync(string refreshMode) => new BalanceRefresh().SubmitSnapshotInternalAsync(refreshMode);

        public static Task RefreshingBalancesAsync(string refreshType, string refreshMode) => new BalanceRefresh().RefreshBalancesInternalAsync(refreshType, refreshMode);

        private static string GetTitle(string refreshType, string refreshMode)
        {
            string key = $"{refreshType}-{refreshMode}";
            return TitleLookup.TryGetValue(key, out string title) ? title : "Processing...";
        }

        private async Task SubmitSnapshotInternalAsync(string RefreshMode)
        {
            ServiceLocator.Logger?.LogDebug($"BalanceRefresh.SubmitSnapshotInternalAsync started. RefreshMode={RefreshMode}");
            await InitializeAsync("Snapshot", RefreshMode);

            try
            {
                Win = CreateAndShowProgressWindow(_ctsHelper);
                if (Win == null)
                {
                    ServiceLocator.Logger?.LogWarn("Unable to set progress window");
                    return;
                }

                await InitializeProgressWindowAsync("Processing request...");

                Token.ThrowIfCancellationRequested();

                if (!await ValidateBalanceFormulasExistAsync())
                    return;

                if (!await ValidateNoBrokenLinksAsync())
                    return;

                Token.ThrowIfCancellationRequested();

                if (_state.VersionCheck)
                {
                    await ValidateAndUpdateFormulasIfNeededAsync();
                }

                await MessageProgressWindowAsync("Saving and zipping the file copy...");

                if (!await SavingAndZippingAsync())
                {
                    await ShowErrorMessage("Exception encountered while saving the file copy.");
                    return;
                }

                string postURL = await BuildApiUrlAsync(RefreshMode);
                await SendRequestAndHandleResponseAsync(postURL);
                Token.ThrowIfCancellationRequested();
            }
            catch (OperationCanceledException)
            {
                ServiceLocator.Logger?.LogWarn("Snapshot submission cancelled by user.");
            }
            catch (Exception ex)
            {
                await HandleUnexpectedErrorAsync(ex);
            }
            finally
            {
                await CleanupAsync();
            }
        }

        private async Task InitializeAsync(string refreshType, string refreshMode)
        {
            _ctsHelper = new CancellationHelper();
            _Title = GetTitle(refreshType, refreshMode);

            await RunExcelAsync(() =>
            {
                CommonMethods.DisableExcelSettings();
                ExcelApp = ServiceLocator.ExcelApp;
                BrWorbook = ExcelApp?.ActiveWorkbook;
                BrWorksheet = ExcelApp?.ActiveSheet as Excel.Worksheet;
            });

            LogDebugs();
        }

        private async Task CleanupAsync(Func<Task>? postCleanup = null)
        {
            try
            {
                if (_ctsHelper != null && !_ctsHelper.IsCancellationRequested)
                    _ctsHelper.Cancel();

                _ctsHelper?.Dispose();
            }
            catch (Exception ex)
            {
                // Swallow dispose exceptions (Excel COM weirdness) but still log for diagnostics.
                ServiceLocator.Logger?.LogWarn($"BalanceRefresh.CleanupAsync: exception disposing CancellationHelper (ignored): {ex.Message}");
            }

            await ClearFiles();
            await SafelyCloseWindowAsync();
            await RunExcelAsync(() => CommonMethods.EnableExcelSettings());

            if (postCleanup != null)
            {
                await postCleanup();
            }
        }

        private async Task<string> BuildApiUrlAsync(string RefreshMode)
        {
            await MessageProgressWindowAsync("Creating api url");

            return await RunExcelAsync(() =>
            {
                string bookName = WebUtility.UrlEncode(BrWorbook?.Name ?? string.Empty);
                string sheetName = WebUtility.UrlEncode(BrWorksheet?.Name ?? string.Empty);
                string LoginURL = _state.LoginUrl;
                long RespID = _state.SelectedCube.CubeId;
                long ChartofAccountID = _state.SelectedLedger.CoaId;

                Token.ThrowIfCancellationRequested();

                if (RefreshMode == "Book")
                {
                    ServiceLocator.Logger?.LogDebug($"Submitting snapshot for the workbook \"{BrWorbook?.Name}\"");
                    return $"{LoginURL}/rest/secure/finance/sheet-snapshot?cubeId={RespID}&coaid={ChartofAccountID}&bookName={bookName}";
                }

                ServiceLocator.Logger?.LogDebug($"Submitting snapshot for the worksheet \"{BrWorksheet?.Name}\"");
                return $"{LoginURL}/rest/secure/finance/sheet-snapshot?cubeId={RespID}&coaid={ChartofAccountID}&sheetName={sheetName}&bookName={bookName}";
            });
        }

        private async Task SendRequestAndHandleResponseAsync(string apiUrl)
        {
            Token.ThrowIfCancellationRequested();

            await MessageProgressWindowAsync("Uploading the file to server...");
            ServiceLocator.Logger?.LogDebug($"BalanceRefresh.SendRequestAndHandleResponseAsync: uploading file to {apiUrl}");
            try
            {
                string response = await ApiHelper.HttpUploadFileAsync(apiUrl, Token);

                Token.ThrowIfCancellationRequested();
                await MessageProgressWindowAsync("Response received from server...");

                if (!IsValidOutput(response))
                {
                    ServiceLocator.Logger?.LogWarn($"BalanceRefresh.SendRequestAndHandleResponseAsync: invalid/unexpected response from {apiUrl}");
                    await ShowErrorMessage(response);
                }
                else
                {
                    ServiceLocator.Logger?.LogDebug("BalanceRefresh.SendRequestAndHandleResponseAsync: upload succeeded, response received.");
                    await HandleResponseAsync(response.ToString().Trim());
                }
            }
            catch (Exception ex)
            {
                ServiceLocator.Logger?.LogException(ex, "BalanceRefresh.SendRequestAndHandleResponseAsync");
            }
        }

        private async Task HandleResponseAsync(string json)
        {
            await MessageProgressWindowAsync("Parsing the response.");

            JsonDocument document;
            try
            {
                document = JsonDocument.Parse(json);
            }
            catch (JsonException ex)
            {
                ServiceLocator.Logger?.LogException(ex);
                ServiceLocator.Logger?.LogRawJson("BalanceRefresh.HandleResponseAsync", json);
                return;
            }
            catch (Exception ex)
            {
                ServiceLocator.Logger?.LogException(ex);
                ServiceLocator.Logger?.LogRawJson("BalanceRefresh.HandleResponseAsync", json);
                return;
            }

            var root = document.RootElement;

            string status = root.TryGetProperty("status", out var statusElement)
                ? statusElement.GetString() ?? string.Empty
                : string.Empty;

            string msg = root.TryGetProperty("msg", out var msgElement)
                ? msgElement.GetString() ?? string.Empty
                : string.Empty;

            if (!IsSuccessStatus(status))
            {
                await ShowErrorMessage(msg);
                return;
            }

            await HandleBackgroundMessageAsync(msg);

            document.Dispose();
        }

        private async Task HandleBackgroundMessageAsync(string msg)
        {
            string msgStr = msg ?? string.Empty;
            string id = ExtractLastToken(msgStr);
            if (!string.IsNullOrEmpty(id))
            {
                await RunExcelAsync(() =>
                {
                    string rangeName = "GLSense_DD_" + id;
                    if (!CommonFunctions.NameRangeExists(rangeName))
                    {
                        try
                        {
                            BrWorbook?.Names.Add(Name: rangeName, RefersToR1C1: id);
                        }
                        catch (Exception ex)
                        {
                            ServiceLocator.Logger?.LogException(ex, "Exception adding the named range");
                        }
                    }
                });
            }

            if (!string.IsNullOrWhiteSpace(msgStr))
            {
                msgStr += Environment.NewLine + "Launch process window to check the status.";
            }

            await ShowInfoMessage(msgStr);
        }

        private static string ExtractLastToken(string msg)
        {
            if (string.IsNullOrWhiteSpace(msg))
                return string.Empty;

            var parts = msg.Split(' ');
            return parts.Length == 0 ? string.Empty : parts[parts.Length - 1];
        }

        private static bool IsSuccessStatus(string status) =>
              string.Equals(status, AppConstants.Success, StringComparison.OrdinalIgnoreCase);

        private static bool IsValidOutput(object output)
        {
            if (output == null)
                return false;

            string s = output.ToString();
            return s.Length > 3 && s.IndexOf("status", StringComparison.Ordinal) >= 0;
        }

        private async Task<bool> SavingAndZippingAsync()
        {
            var path = ServiceLocator.Paths.Temp ?? string.Empty;
            var filePath1 = Path.Combine(path, AppConstants.RefreshFileName);
            var zipPath = Path.Combine(path, AppConstants.RefreshZipFileName);

            try
            {
                await RunExcelAsync(() => BrWorbook?.SaveCopyAs(filePath1));
            }
            catch (Exception ex)
            {
                ServiceLocator.Logger?.LogException(ex);
                return false;
            }

            try
            {
                if (File.Exists(zipPath))
                {
                    File.Delete(zipPath);
                }

                using var newfile = ZipFile.Open(zipPath, ZipArchiveMode.Create);
                newfile.CreateEntryFromFile(filePath1, AppConstants.RefreshFileName);
                return true;
            }
            catch (Exception ex)
            {
                ServiceLocator.Logger?.LogException(ex);
            }
            return false;
        }

        public async Task RefreshBalancesInternalAsync(string RefreshType, string RefreshMode)
        {
            bool showMessage = false;
            string completionMessage = string.Empty;

            ServiceLocator.Logger?.LogDebug($"BalanceRefresh.RefreshBalancesInternalAsync started. RefreshType={RefreshType}, RefreshMode={RefreshMode}");

            try
            {
                await InitializeAsync(RefreshType, RefreshMode);

                // Regression fix: "workbook is saved" and "balance formulas exist" are both
                // preconditions that should stop the operation before the progress window
                // ever appears - previously both ran AFTER CreateAndShowProgressWindow, so
                // the user would see a progress window flash up only to immediately get an
                // error message. ValidateWorkbookIsSavedAsync/ValidateBalanceFormulasExistAsync
                // both call MessageProgressWindowAsync internally, which is a safe no-op
                // when Win is still null (see its own null-check), so it's safe to run these
                // before the window exists. Broken-links and the snapshot-path prompt are
                // NOT part of this reordering - they intentionally still run after the
                // window is up, unchanged from before.
                if (!await ValidateWorkbookIsSavedAsync()) return;
                if (!await ValidateBalanceFormulasExistAsync()) return;

                Win = CreateAndShowProgressWindow(_ctsHelper);

                if (Win == null)
                {
                    ServiceLocator.Logger?.LogWarn("Unable to set progress window");
                    return;
                }

                await InitializeProgressWindowAsync("Processing request...");

                if (!await ValidateNoBrokenLinksAsync()) return;

                Token.ThrowIfCancellationRequested();

                bool isSnapshot = IsSnapshotMode();
                string snapshotFilePath = string.Empty;

                if (isSnapshot)
                {
                    snapshotFilePath = await PromptForSnapshotPathAsync(Token);
                    Token.ThrowIfCancellationRequested();
                    if (string.IsNullOrEmpty(snapshotFilePath)) return;
                }

                Token.ThrowIfCancellationRequested();

                await CollectBalanceFormulasAsync(RefreshMode);

                if (_state.VersionCheck)
                {
                    Token.ThrowIfCancellationRequested();
                    await ValidateAndUpdateFormulasIfNeededAsync();
                }

                await MessageProgressWindowAsync("Saving copy and preparing for refresh...");

                Exception? refreshException = await BulkRefreshProcess.StartRefreshing(
                    Win, RefreshMode, isSnapshot, snapshotFilePath, Token, _state, _excelContext);

                if (refreshException != null)
                {
                    await HandleRefreshErrorAsync(refreshException);
                }
                else
                {
                    showMessage = isSnapshot;
                    completionMessage = await DisplayMessageAsync(RefreshType);
                    await LogSuccessAsync(RefreshType);
                }
            }
            catch (OperationCanceledException)
            {
                await ShowCancelledAsync();
                ServiceLocator.Logger?.LogWarn($"{_Title} cancelled by user.");
            }
            catch (Exception ex)
            {
                await HandleUnexpectedErrorAsync(ex);
            }
            finally
            {
                await CleanupAsync(() =>
                {
                    if (showMessage && !string.IsNullOrEmpty(completionMessage))
                    {
                        CommonFunctions.GLSenseMessage(completionMessage, MessageBoxImage.Information, MessageBoxButton.OK);
                    }

                    return Task.CompletedTask;
                });
            }
        }

        private bool IsSnapshotMode()
        {
            return _Title.IndexOf("snapshot", StringComparison.CurrentCultureIgnoreCase) >= 0;
        }

        private async Task<bool> ValidateWorkbookIsSavedAsync()
        {
            await MessageProgressWindowAsync("Checking if workbook is saved or not.");

            bool isSaved = await RunExcelAsync(() => !string.IsNullOrWhiteSpace(BrWorbook?.Path));

            if (isSaved)
                return true;

            await ShowErrorMessage("Save this file before executing this function.");
            return false;
        }

        private async Task<bool> ValidateNoBrokenLinksAsync()
        {
            await MessageProgressWindowAsync("Checking for broken links in the workbook.");
            string brokenLinks = await RunExcelAsync(() => CommonFunctions.WorkbookBrokenLinks());

            if (string.IsNullOrWhiteSpace(brokenLinks))
                return true;

            string linksText = string.Join("\", \"", brokenLinks);
            await ShowErrorMessage(
                $"The workbook has broken links:{Environment.NewLine}\"{linksText}\"{Environment.NewLine}Please fix them before proceeding.");

            return false;
        }

        private async Task<string> PromptForSnapshotPathAsync(CancellationToken token)
        {
            await MessageProgressWindowAsync("Select save location for snapshot.");
            return await SnapshotDialogHelper.PromptSnapshotAsync(Win, token);
        }

        private async Task<bool> ValidateBalanceFormulasExistAsync()
        {
            await MessageProgressWindowAsync("Checking for balance formulas...");
            bool exists = await ExistsBalanceFormulasAsync();

            if (exists)
                return true;

            await ShowErrorMessage($"No balance formulas found in the worksheet for {_Title}");
            return false;
        }

        private async Task CollectBalanceFormulasAsync(string refreshMode)
        {
            await RunExcelAsync(() =>
            {
                if (refreshMode.Equals("Sheet", StringComparison.OrdinalIgnoreCase))
                {
                    if (BrWorksheet != null)
                        CommonMethods.Get_GLSense_MultiFormulas(BrWorksheet.Name);
                }
                else
                {
                    if (BrWorbook == null) return;

                    foreach (Excel.Worksheet ws in BrWorbook.Worksheets)
                    {
                        CommonMethods.Get_GLSense_MultiFormulas(ws.Name);
                    }
                }
            });
        }

        private async Task ValidateAndUpdateFormulasIfNeededAsync()
        {
            await MessageProgressWindowAsync("Validating balance formulas...");

            // 2026-07-15 fix (ported from FinalWorkingCode): was passing just
            // ServiceLocator.ExcelApp.ActiveCell.Address - the single active cell only,
            // with no workbook/sheet qualification. That misses invalid formulas anywhere
            // else in a multi-cell Selection and can misresolve when the active sheet
            // isn't the one the formulas actually live on. Build the same external,
            // range-qualified address DD_JL.cs/DD_BL.cs already use for this purpose.
            var invalidFormulas = await RunExcelAsync(() =>
            {
                Excel.Range rng = (Excel.Range)ServiceLocator.ExcelApp.Selection;
                string external = ExcelExternalRef.BuildExternalAddress(rng);
                return CommonFunctions.NotValidBalancesDict(external);
            });

            if (invalidFormulas == null || invalidFormulas.Count == 0 || Win == null)
                return;

            bool shouldUpdate = await EnsureCompatibilityAsync(Win);
            if (shouldUpdate)
            {
                await MessageProgressWindowAsync("Updating invalid balance formulas...");
                await RunExcelAsync(() => CommonMethods.BalanceFormulas_Updation(invalidFormulas));
            }
            else
            {
                await MessageProgressWindowAsync("Skipping formula updates as requested...");
            }
        }

        private async Task HandleRefreshErrorAsync(Exception ex)
        {
            string errorMessage = ParseRefreshErrorMessage(ex.Message);
            await ShowErrorMessage(errorMessage);
        }

        private async Task<string> DisplayMessageAsync(string refreshType)
        {
            return await RunExcelAsync(() =>
            {
                if (BrWorksheet == null || BrWorbook == null)
                    return string.Empty;

                string scope = refreshType.Equals("Sheet", StringComparison.OrdinalIgnoreCase)
                    ? $"worksheet \"{BrWorksheet.Name}\""
                    : $"workbook \"{BrWorbook.Name}\"";

                return $"{_Title} completed successfully for {scope}";
            });
        }

        private Task LogSuccessAsync(string refreshType)
        {
            return RunExcelAsync(() =>
            {
                if (BrWorksheet == null || BrWorbook == null)
                    return;

                string scope = refreshType.Equals("Sheet", StringComparison.OrdinalIgnoreCase)
                    ? $"worksheet \"{BrWorksheet.Name}\""
                    : $"workbook \"{BrWorbook.Name}\"";

                ServiceLocator.Logger?.LogDebug($"{_Title} completed successfully for {scope}");
            });
        }

        private async Task SafelyCloseWindowAsync()
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

            await RunExcelAsync(() =>
            {
                if (ServiceLocator.ExcelApp != null)
                {
                    ServiceLocator.ExcelApp.WindowState = Excel.XlWindowState.xlMaximized;
                    ServiceLocator.ExcelApp.ActiveWorkbook?.Activate();
                }
            });
        }

        private static class ExcelStaDispatcher
        {
            public static Task Run(SynchronizationContext? context, Action action)
            {
                if (context == null || context == SynchronizationContext.Current)
                {
                    action();
                    return Task.CompletedTask;
                }

                var tcs = new TaskCompletionSource<object?>();
                context.Post(_ =>
                {
                    try
                    {
                        action();
                        tcs.SetResult(null);
                    }
                    catch (Exception ex)
                    {
                        tcs.SetException(ex);
                    }
                }, null);
                return tcs.Task;
            }

            public static Task<T> Run<T>(SynchronizationContext? context, Func<T> func)
            {
                if (context == null || context == SynchronizationContext.Current)
                {
                    return Task.FromResult(func());
                }

                var tcs = new TaskCompletionSource<T>();
                context.Post(_ =>
                {
                    try
                    {
                        T result = func();
                        tcs.SetResult(result);
                    }
                    catch (Exception ex)
                    {
                        tcs.SetException(ex);
                    }
                }, null);
                return tcs.Task;
            }
        }

        private Task RunExcelAsync(Action action) => ExcelStaDispatcher.Run(_excelContext, action);

        private Task<T> RunExcelAsync<T>(Func<T> func) => ExcelStaDispatcher.Run(_excelContext, func);

        private static string ParseRefreshErrorMessage(string rawMessage)
        {
            if (string.IsNullOrEmpty(rawMessage))
                return "An unknown error occurred during refresh.";

            if (rawMessage.IndexOf("<!DOCTYPE", StringComparison.OrdinalIgnoreCase) >= 0)
            {
                ServiceLocator.Logger?.LogError("Server returned HTML error page: " + rawMessage);
                return "A server error occurred. Please check logs for details.";
            }

            try
            {
                using JsonDocument doc = JsonDocument.Parse(rawMessage);
                var root = doc.RootElement;

                string msg = root.TryGetProperty("msg", out var msgProp) ? msgProp.GetString() ?? "" : "";
                string msg1 = root.TryGetProperty("msg1", out var msg1Prop) ? msg1Prop.GetString() ?? "" : "";

                if (!string.IsNullOrWhiteSpace(msg) && !string.IsNullOrWhiteSpace(msg1))
                {
                    // Check if msg1 is contained within msg (case-insensitive)
                    if (msg.ToLowerInvariant().Contains(msg1.ToLowerInvariant()))
                    {
                        return msg;
                    }
                    else
                    {
                        return $"{msg}{Environment.NewLine}{msg1}";
                    }
                }

                return !string.IsNullOrEmpty(msg) ? msg : msg1;
            }
            catch (JsonException ex)
            {
                ServiceLocator.Logger?.LogException(ex, "Failed to parse error message as JSON");
                ServiceLocator.Logger?.LogRawJson("BalanceRefresh.ParseRefreshErrorMessage", rawMessage);
                return rawMessage;
            }
            catch (Exception ex)
            {
                ServiceLocator.Logger?.LogException(ex, "Unexpected error while parsing refresh error message");
                ServiceLocator.Logger?.LogRawJson("BalanceRefresh.ParseRefreshErrorMessage", rawMessage);
            }

            return rawMessage;
        }
        private static async Task ClearFiles()
        {
            try
            {
                string workingFile = Path.Combine(ServiceLocator.Paths.Temp, AppConstants.RefreshFileName);
                string zipFile = Path.Combine(ServiceLocator.Paths.Temp, AppConstants.RefreshZipFileName);

                await TryDeleteFileWithFallbackAsync(workingFile);
                await TryDeleteFileWithFallbackAsync(zipFile);
            }
            catch (Exception ex)
            {
                ServiceLocator.Logger?.LogException(ex, "Exception while deleting the bulk refresh files");
            }
        }

        private static async Task TryDeleteFileWithFallbackAsync(string filePath)
        {
            if (string.IsNullOrWhiteSpace(filePath) || !File.Exists(filePath))
                return;

            var timeout = TimeSpan.FromSeconds(300);

            try
            {
                bool deleted = await SafeDeleteHelper
                    .TryDeleteFileAsync(filePath, timeout)
                    .ConfigureAwait(false);

                if (deleted)
                    return;

                string? quarantinedPath = QuarantineDeleteHelper.TryQuarantine(filePath);
                if (quarantinedPath != null)
                {
                    await SafeDeleteHelper.TryDeleteFileAsync(quarantinedPath, TimeSpan.FromSeconds(300));
                }

                ServiceLocator.Logger?.LogWarn($"Could not delete '{filePath}' within timeout.");
            }
            catch (OperationCanceledException)
            {
                ServiceLocator.Logger?.LogWarn($"Timed out while trying to delete file: {filePath}");
            }
            catch (Exception ex)
            {
                ServiceLocator.Logger?.LogException(ex, $"Error deleting file: {filePath}");
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
                return true;
            }
            else if (action == false)
            {
                return false;
            }
            else
            {
                return false;
            }
        }

        private void LogDebugs()
        {
            ServiceLocator.Logger?.LogDebug($"User : {_state.LoginUserName}");
            ServiceLocator.Logger?.LogDebug($"Instance : {_state.LoginUrl}");
            ServiceLocator.Logger?.LogDebug($"Cube Selected : {_state.SelectedCube.CubeName}");
            ServiceLocator.Logger?.LogDebug($"Ledger Selected : {_state.SelectedLedger.LedgerName}");
        }

        private async Task<bool> ExistsBalanceFormulasAsync()
        {
            try
            {
                return await RunExcelAsync(() =>
                {
                    if (BrWorbook == null || BrWorksheet == null)
                    {
                        return false;
                    }

                    if (_Title == "Worksheet Refresh" || _Title == "Worksheet Snapshot")
                    {
                        return CommonFunctions.BalanceFormulaExists(BrWorksheet.Name);
                    }

                    bool formulasExists = false;
                    foreach (Excel.Worksheet wsheet in BrWorbook.Worksheets)
                    {
                        if (CommonFunctions.BalanceFormulaExists(wsheet.Name))
                        {
                            formulasExists = true;
                            break;
                        }
                    }

                    return formulasExists;
                });
            }
            catch (Exception ex)
            {
                await HandleUnexpectedErrorAsync(ex);
                return false;
            }
        }

        private async Task ShowErrorMessage(string message)
        {
            await SafelyCloseWindowAsync();
            CommonFunctions.GLSenseMessage(
                message,
                MessageBoxImage.Error,
                MessageBoxButton.OK);
        }

        private async Task ShowInfoMessage(string message)
        {
            await SafelyCloseWindowAsync();
            CommonFunctions.GLSenseMessage(
                message,
                MessageBoxImage.Information,
                MessageBoxButton.OK);
        }

        private async Task ShowCancelledAsync()
        {
            await SafelyCloseWindowAsync();
            CommonFunctions.GLSenseMessage(
                "Operation cancelled!",
                MessageBoxImage.Error,
                MessageBoxButton.OK);
        }

        private async Task HandleUnexpectedErrorAsync(Exception ex)
        {
            ServiceLocator.Logger?.LogException(ex);
            await SafelyCloseWindowAsync();
            CommonFunctions.GLSenseMessage(
                $"An unexpected error occurred.{Environment.NewLine}{ex.Message}",
                MessageBoxImage.Error,
                MessageBoxButton.OK);
        }

        private GLWaitWindow? CreateAndShowProgressWindow(CancellationHelper? cts)
        {
            try
            {
                GLWaitWindow? win = null;

                // WpfAppManager.InvokeOnWpfThread only takes an Action (no return value),
                // so capture the created window from inside the delegate - same pattern
                // DrillCellHighlighter.cs/DD_JL.cs use for GLWaitWindow. win.Show() replaces
                // the original's win.ShowWithOwner(hwnd) - GLWaitWindow (BaseWindow-derived)
                // sets its Excel owner automatically via ServiceLocator.ExcelHandle.
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

        private Task InitializeProgressWindowAsync(string message)
        {
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
                        Win.SetProcessTitle(_Title);
                        Win.SetProcessMessage(message);
                    },
                    DispatcherPriority.Normal);

                return Task.CompletedTask;
            }
            catch (TaskCanceledException)
            {
                ServiceLocator.Logger?.LogDebug("BalanceRefresh.InitializeProgressWindowAsync: dispatcher invoke was cancelled (window likely closing).");
                return Task.CompletedTask;
            }
            catch (Exception ex)
            {
                ServiceLocator.Logger?.LogException(ex, "BalanceRefresh.InitializeProgressWindowAsync");
                return Task.CompletedTask;
            }
        }

        private Task MessageProgressWindowAsync(string message)
        {
            if (Win == null || Win.Dispatcher == null)
                return Task.CompletedTask;

            try
            {
                // Fire-and-forget: this is a best-effort progress label update only.
                // Callers on non-UI threads (e.g. raw background threads with no
                // captured SynchronizationContext) must not be resumed on an
                // arbitrary ThreadPool thread after awaiting this dispatch, since
                // downstream code may touch thread-affinitized Excel COM objects.
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
                ServiceLocator.Logger?.LogDebug("BalanceRefresh.MessageProgressWindowAsync: dispatcher invoke was cancelled (window likely closing).");
                return Task.CompletedTask;
            }
            catch (Exception ex)
            {
                ServiceLocator.Logger?.LogException(ex, "BalanceRefresh.MessageProgressWindowAsync");
                return Task.CompletedTask;
            }
        }
    }
#nullable restore
}
