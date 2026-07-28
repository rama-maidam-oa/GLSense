// BalanceHighlighter.cs in GLSense.Addin.Core
// Group F - extracted from GLSense\AddinModule.cs (FinalWorkingCode), RibHighlight_OnClick
// (old lines ~1288-1642) and its private helpers (ValidateHighlightPreconditions,
// FindAdaptiveMemoryCellsFast, BuildRangeEfficiently, BuildRangeCellByCell,
// ParseExcelCell, SelectAdaptiveBalanceRange), plus the shared wait-window helpers the
// old host also used (CreateAndShowWaitWindow/InitializeWaitWindowAsync/
// SafelyCloseWaitWindowAsync/MessageWaitWindowAsync/ShowErrorMessageAsync/
// GuardLoginAndExcel) - ported here as private copies, per the established
// per-file-duplication convention (DD_BL.cs/DD_JL.cs/DD_SL.cs/DDDatatoWorksheet.cs all
// carry their own copies rather than sharing one utility class).
//
// Named/shaped after DrillCellHighlighter.cs (the closest sibling pattern for a
// "ribbon click -> find+select cells on the active sheet" feature): a static class in
// this namespace with one public async entry point named after the ribbon action
// (RibHighlight_OnClick), matching that file's shape exactly.
//
// Finds every cell on the active sheet whose cached balance value (from
// AppState.Instance.CalculatedBalances, written by BulkRefreshProcess/DataTableBuilder -
// also ported this pass) is present ("adaptive memory" - i.e. already refreshed), builds
// an Excel range from them efficiently (batched by address-string length to avoid COM
// range-string limits), and selects it.
//
// Re-pointed vs. the original:
//   - AppState.Instance.ExcelApp -> ServiceLocator.ExcelApp (this project's AppState has
//     no ExcelApp field).
//   - LogUtility.* (static) -> ServiceLocator.Logger?.*.
//   - System.Windows.Forms MessageBoxIcon/MessageBoxButtons -> System.Windows
//     MessageBoxImage/MessageBoxButton (this project has no WinForms reference).
//   - GLWaitWindow now derives from BaseWindow: win.ShowWithOwner(hwnd) -> win.Show()
//     (Excel owner set automatically via ServiceLocator.ExcelHandle). CreateAndShow-
//     ProgressWindow rewritten to the WpfAppManager.InvokeOnWpfThread(Action)-with-
//     captured-local pattern (InvokeOnWpfThread has no Func<T> overload here).
using GLSense.Addin.Core.Helpers;
using GLSense.Addin.Core.Infrastructure;
using GLSense.Addin.Core.Utilities;
using GLSense.Addin.Core.Views;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Threading;
using Excel = Microsoft.Office.Interop.Excel;

namespace GLSense.Addin.Core.Drilldowns
{
    public static class BalanceHighlighter
    {
        /// <summary>
        /// Ribbon click (RibHighlight): selects balance cells on the active sheet whose
        /// values are already present in adaptive memory (AppState.Instance.CalculatedBalances).
        /// </summary>
        public static async Task RibHighlight_OnClick()
        {
            if (!GuardLoginAndExcel())
                return;

            Excel.Range adaptiveBalanceRange = null;
            CommonMethods.DisableExcelSettings();
            GLWaitWindow win = null;
            using var ctsHelper = new CancellationHelper();
            CancellationToken token = ctsHelper.GetToken();

            try
            {
                Excel.Worksheet wrkSheet = ServiceLocator.ExcelApp.ActiveSheet as Excel.Worksheet;

                ServiceLocator.Logger?.LogDebug("Selecting balance cells whose values are from adaptive memory.");
                ServiceLocator.Logger?.LogDebug($"Worksheet Name : {wrkSheet?.Name}");

                if (!ValidateHighlightPreconditions(wrkSheet))
                    return;

                win = CreateAndShowWaitWindow(ctsHelper);
                await InitializeWaitWindowAsync(win, "Highlighting Adaptive Memory", "Searching adaptive memory cells...");

                adaptiveBalanceRange = await FindAdaptiveMemoryCellsFast(wrkSheet, win, token);
            }
            catch (OperationCanceledException)
            {
                ServiceLocator.Logger?.LogWarn("Highlight Adaptive Memory operation was canceled by the user.");
            }
            catch (Exception ex)
            {
                ServiceLocator.Logger?.LogException(ex, "BalanceHighlighter.RibHighlight_OnClick");
            }
            finally
            {
                await SafelyCloseWaitWindowAsync(win);
                SelectAdaptiveBalanceRange(adaptiveBalanceRange);
                CommonMethods.EnableExcelSettings();
            }
        }

        private static bool ValidateHighlightPreconditions(Excel.Worksheet wrkSheet)
        {
            bool balancesExists = CommonFunctions.BalanceFormulaExists(wrkSheet?.Name);

            if (!balancesExists)
            {
                CommonFunctions.GLSenseMessage("No balance formulas exists in the current worksheet.", MessageBoxImage.Exclamation, MessageBoxButton.OK);
                return false;
            }

            DataTable dt = AppState.Instance.CalculatedBalances;

            if (dt == null || !dt.Columns.Contains("cache") || dt.Rows.Count == 0)
            {
                ServiceLocator.Logger?.LogDebug("No calculated values or 'cache' column found in balance refresh memory, or no rows present.");
                CommonFunctions.GLSenseMessage("Worksheet has to be refreshed first.", MessageBoxImage.Exclamation, MessageBoxButton.OK);
                return false;
            }

            return true;
        }

        private static async Task<Excel.Range> FindAdaptiveMemoryCellsFast(Excel.Worksheet wrkSheet, GLWaitWindow win, CancellationToken token)
        {
            DataTable dt = AppState.Instance.CalculatedBalances;
            string sheetNameEscaped = wrkSheet.Name.Replace("'", "''");
            string dataTableFilter = $"[excelSheet]='{sheetNameEscaped}' AND [cache] = True";

            ServiceLocator.Logger?.LogDebug($"BalanceHighlighter.FindAdaptiveMemoryCellsFast started for sheet '{wrkSheet.Name}'. Total cached rows in memory: {dt?.Rows.Count ?? 0}.");

            token.ThrowIfCancellationRequested();

            if (win != null)
            {
                await MessageWaitWindowAsync(win, "Filtering data...");
                await Task.Delay(1, token);
            }

            DataRow[] sheetRows = dt.Select(dataTableFilter);

            if (sheetRows == null || sheetRows.Length == 0)
            {
                ServiceLocator.Logger?.LogDebug($"BalanceHighlighter.FindAdaptiveMemoryCellsFast: no rows matched filter \"{dataTableFilter}\", aborting.");
                await ShowErrorMessageAsync(win, "No balances from adaptive memory");
                return null;
            }

            if (win != null)
            {
                await MessageWaitWindowAsync(win, $"Parsing {sheetRows.Length} cell addresses...");
                await Task.Delay(1, token);
            }

            // STEP 1: Parse and deduplicate in one pass
            var cellMap = new Dictionary<string, (int col, int row)>(StringComparer.OrdinalIgnoreCase);
            var addressOrder = new List<string>(sheetRows.Length);

            for (int i = 0; i < sheetRows.Length; i++)
            {
                token.ThrowIfCancellationRequested();

                string cellAddress = sheetRows[i]["excelCell"]?.ToString();
                if (string.IsNullOrWhiteSpace(cellAddress)) continue;

                if (!cellMap.ContainsKey(cellAddress))
                {
                    var (col, row) = ParseExcelCell(cellAddress);
                    cellMap[cellAddress] = (col, row);
                    addressOrder.Add(cellAddress);
                }
            }

            if (cellMap.Count == 0)
            {
                await ShowErrorMessageAsync(win, "No valid cell addresses found");
                return null;
            }

            if (win != null)
            {
                await MessageWaitWindowAsync(win, $"Sorting {cellMap.Count} unique cells...");
                await Task.Delay(1, token);
            }

            // STEP 2: Sort unique addresses by column then row

            var sortedAddresses = addressOrder
                .OrderBy(addr => cellMap[addr].col)
                .ThenBy(addr => cellMap[addr].row)
                .ToList();

            if (win != null)
            {
                await MessageWaitWindowAsync(win, $"Creating range from {sortedAddresses.Count} cells...");
                await Task.Delay(1, token);
            }

            // STEP 3: Build range efficiently
            ServiceLocator.Logger?.LogDebug($"BalanceHighlighter.FindAdaptiveMemoryCellsFast: {cellMap.Count} unique cell(s) found, building range.");
            Excel.Range result = await BuildRangeEfficiently(wrkSheet, sortedAddresses, win, token);

            if (win != null)
            {
                await MessageWaitWindowAsync(win, "Ready");
                await Task.Delay(1, token);
            }

            return result;
        }

        private static async Task<Excel.Range> BuildRangeEfficiently(Excel.Worksheet wrkSheet, List<string> addresses, GLWaitWindow win, CancellationToken token)
        {
            try
            {
                Excel.Range finalRange = null;
                int currentIndex = 0;
                int totalBatches = 0;

                while (currentIndex < addresses.Count)
                {
                    token.ThrowIfCancellationRequested();

                    // Build batch based on character count (max 200 chars for safety)
                    var batch = new List<string>();
                    int currentLength = 0;

                    while (currentIndex < addresses.Count)
                    {
                        string address = addresses[currentIndex];
                        int addedLength = address.Length;

                        // Add 1 for comma separator (except first item)
                        if (batch.Count > 0)
                        {
                            addedLength += 1;
                        }

                        if (currentLength + addedLength > 200)
                        {
                            break;
                        }

                        batch.Add(address);
                        currentLength += addedLength;
                        currentIndex++;
                    }

                    // Ensure we process at least one address
                    if (batch.Count == 0 && currentIndex < addresses.Count)
                    {
                        batch.Add(addresses[currentIndex]);
                        currentIndex++;
                    }

                    totalBatches++;

                    if (win != null)
                    {
                        await MessageWaitWindowAsync(win, $"Building range: batch {totalBatches} ({batch.Count} cells)");
                        await Task.Delay(1, token);
                    }

                    string batchAddress = string.Join(",", batch);

                    try
                    {
                        Excel.Range batchRange = wrkSheet.Range[batchAddress];

                        if (finalRange == null)
                        {
                            finalRange = batchRange;
                        }
                        else
                        {
                            finalRange = wrkSheet.Application.Union(finalRange, batchRange);
                        }
                    }
                    catch (Exception ex)
                    {
                        ServiceLocator.Logger?.LogException(ex, $"BalanceHighlighter.BuildRangeEfficiently: batch {totalBatches} ({batch.Count} cells) failed, falling back to per-cell processing");

                        // Fallback: Process each cell individually for this batch
                        foreach (string cell in batch)
                        {
                            try
                            {
                                Excel.Range cellRange = wrkSheet.Range[cell];
                                if (finalRange == null)
                                {
                                    finalRange = cellRange;
                                }
                                else
                                {
                                    finalRange = wrkSheet.Application.Union(finalRange, cellRange);
                                }
                            }
                            catch (Exception cellEx)
                            {
                                ServiceLocator.Logger?.LogException(cellEx, $"BalanceHighlighter.BuildRangeEfficiently: fallback failed for cell '{cell}'");
                            }
                        }
                    }
                }
                ServiceLocator.Logger?.LogDebug($"BalanceHighlighter.BuildRangeEfficiently: completed, {totalBatches} batch(es) processed for {addresses.Count} address(es).");
                return finalRange;
            }
            catch (Exception ex)
            {
                ServiceLocator.Logger?.LogException(ex, "BalanceHighlighter.BuildRangeEfficiently: batching failed entirely, falling back to cell-by-cell");
                return await BuildRangeCellByCell(wrkSheet, addresses, win, token);
            }
        }

        private static async Task<Excel.Range> BuildRangeCellByCell(Excel.Worksheet wrkSheet, List<string> addresses, GLWaitWindow win, CancellationToken token)
        {
            Excel.Range finalRange = null;
            int total = addresses.Count;

            for (int i = 0; i < total; i++)
            {
                token.ThrowIfCancellationRequested();

                if (win != null && i % 100 == 0)
                {
                    await MessageWaitWindowAsync(win, $"Processing cell {i}/{total}");
                    await Task.Delay(1, token);
                }

                try
                {
                    Excel.Range cellRange = wrkSheet.Range[addresses[i]];

                    if (finalRange == null)
                    {
                        finalRange = cellRange;
                    }
                    else
                    {
                        finalRange = wrkSheet.Application.Union(finalRange, cellRange);
                    }
                }
                catch (Exception ex)
                {
                    ServiceLocator.Logger?.LogException(ex, $"BalanceHighlighter.BuildRangeCellByCell: failed for cell '{addresses[i]}'");
                }
            }

            return finalRange;
        }

        private static (int col, int row) ParseExcelCell(string cellAddress)
        {
            int col = 0;
            int row = 0;
            int idx = 0;
            int len = cellAddress.Length;

            // Skip leading $
            if (len > 0 && cellAddress[idx] == '$') idx++;

            // Parse column letters (A-Z)
            while (idx < len)
            {
                char c = cellAddress[idx];
                if (c >= 'A' && c <= 'Z')
                {
                    col = col * 26 + (c - 'A' + 1);
                    idx++;
                }
                else if (c >= 'a' && c <= 'z')
                {
                    col = col * 26 + (c - 'a' + 1);
                    idx++;
                }
                else
                {
                    break;
                }
            }

            // Skip $ before row
            if (idx < len && cellAddress[idx] == '$') idx++;

            // Parse row digits
            while (idx < len)
            {
                char c = cellAddress[idx];
                if (c >= '0' && c <= '9')
                {
                    row = row * 10 + (c - '0');
                    idx++;
                }
                else
                {
                    break;
                }
            }

            return (col, row);
        }

        private static void SelectAdaptiveBalanceRange(Excel.Range adaptiveBalanceRange)
        {
            if (adaptiveBalanceRange == null)
                return;

            ServiceLocator.Logger?.LogDebug($"Balance cells found with adaptive memory : {adaptiveBalanceRange?.Address}");
            try
            {
                // Regression-pattern fix: this used to fetch a fresh ServiceLocator.ExcelApp.
                // ActiveWorkbook/ActiveSheet here - reached after two awaits from the original
                // ribbon click (FindAdaptiveMemoryCellsFast, then SafelyCloseWaitWindowAsync in
                // the finally block), the same "COM object fetched fresh deep in an async chain"
                // pattern that threw an InvalidCastException crossing the host<->Addin.Core
                // AppDomain boundary in SanitizeSheetName (see that fix's comment). Using
                // adaptiveBalanceRange's own Worksheet/Workbook instead avoids the risky
                // re-fetch AND is more correct: this needs to activate the sheet/workbook that
                // actually owns the found range, not whatever happens to be "active" by the
                // time this continuation runs (which is a no-op if focus never moved, but wrong
                // if it did).
                Excel.Worksheet ws = adaptiveBalanceRange.Worksheet;
                (ws?.Parent as Excel.Workbook)?.Activate();
                ws?.Activate();

                adaptiveBalanceRange.Select();
            }
            catch (Exception ex)
            {
                ServiceLocator.Logger?.LogException(ex, "BalanceHighlighter.SelectAdaptiveBalanceRange");
            }
        }

        private static bool GuardLoginAndExcel() => AppState.Instance.IsLoginCompleted && ServiceLocator.ExcelApp != null;

        private static GLWaitWindow CreateAndShowWaitWindow(CancellationHelper cts)
        {
            try
            {
                GLWaitWindow win = null;

                // WpfAppManager.InvokeOnWpfThread only takes an Action (no return value),
                // so capture the created window from inside the delegate - same pattern
                // DrillCellHighlighter.cs/DD_JL.cs use for GLWaitWindow.
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

        private static Task InitializeWaitWindowAsync(GLWaitWindow win, string title, string message)
        {
            if (win == null || win.Dispatcher == null) return Task.CompletedTask;
            try
            {
                return win.Dispatcher.InvokeAsync(() => { win.SetProcessTitle(title); win.SetProcessMessage(message); }, DispatcherPriority.Normal).Task;
            }
            catch (TaskCanceledException) { ServiceLocator.Logger?.LogDebug("BalanceHighlighter.InitializeWaitWindowAsync: dispatcher invoke was cancelled (window likely closing)."); return Task.CompletedTask; }
            catch (Exception ex) { ServiceLocator.Logger?.LogException(ex, "BalanceHighlighter.InitializeWaitWindowAsync"); return Task.CompletedTask; }
        }

        private static async Task SafelyCloseWaitWindowAsync(GLWaitWindow win)
        {
            if (win == null) return;
            try
            {
                if (win.Dispatcher.CheckAccess()) win.RequestClose();
                await win.Dispatcher.InvokeAsync(() => win.RequestClose());
            }
            catch (Exception ex) { ServiceLocator.Logger?.LogError($"Error closing wait window: {ex.Message}"); }
        }

        private static Task MessageWaitWindowAsync(GLWaitWindow win, string message)
        {
            if (win == null || win.Dispatcher == null) return Task.CompletedTask;
            try
            {
                return win.Dispatcher.InvokeAsync(() => win.SetProcessMessage(message), DispatcherPriority.Normal).Task;
            }
            catch (TaskCanceledException) { ServiceLocator.Logger?.LogDebug("BalanceHighlighter.MessageWaitWindowAsync: dispatcher invoke was cancelled (window likely closing)."); return Task.CompletedTask; }
            catch (Exception ex) { ServiceLocator.Logger?.LogException(ex, "BalanceHighlighter.MessageWaitWindowAsync"); return Task.CompletedTask; }
        }

        private static async Task ShowErrorMessageAsync(GLWaitWindow win, string message)
        {
            await SafelyCloseWaitWindowAsync(win);
            CommonFunctions.GLSenseMessage(message, MessageBoxImage.Error, MessageBoxButton.OK);
        }
    }
}
