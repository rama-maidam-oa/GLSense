// RowVisibilityProcessor.cs in GLSense.Addin.Core
// Group F - extracted from GLSense\AddinModule.cs (FinalWorkingCode), RibHideRows_OnClick/
// RibUnHideRows_OnClick (old lines ~2829-3006), including the RowProcessor/
// HideRowProcessor/UnhideRowProcessor class hierarchy and helpers (FindHideRows,
// ShouldHideRow, IsGetBalanceFormula, IsZero, GetSelection, GetFormulaAndValueArraysAsync,
// CoerceTo2D, BalancesRangeAsync, ProcessHideRowsAsync, ProcessUnhideRowsByBatchesAsync),
// plus the shared wait-window helpers these two handlers also used
// (CreateAndShowWaitWindow/InitializeWaitWindowAsync/SafelyCloseWaitWindowAsync/
// MessageWaitWindowAsync/GuardLoginAndExcel) - ported here as private copies, per the
// established per-file-duplication convention (DD_BL.cs/DD_JL.cs/DD_SL.cs/
// DDDatatoWorksheet.cs/BalanceHighlighter.cs all carry their own copies rather than
// sharing one utility class). ShowErrorMessageAsync was not needed - neither
// RibHideRows_OnClick nor RibUnHideRows_OnClick's code path in the original ever called
// it (only RibHighlight/ResetBalances did, both ported separately).
//
// Hides rows in the current selection where every balance-formula cell in that row
// evaluates to zero (RibHideRows), or restores rows previously hidden this way
// (RibUnHideRows). Entry points named after their ribbon actions (RibHideRows_OnClick/
// RibUnHideRows_OnClick), matching DrillCellHighlighter.cs's naming convention.
//
// Re-pointed vs. the original:
//   - AppState.Instance.ExcelApp -> ServiceLocator.ExcelApp (this project's AppState has
//     no ExcelApp field).
//   - LogUtility.* (static) -> ServiceLocator.Logger?.*.
//   - System.Windows.Forms MessageBoxIcon/MessageBoxButtons -> System.Windows
//     MessageBoxImage/MessageBoxButton (this project has no WinForms reference).
//   - GLWaitWindow now derives from BaseWindow: win.ShowWithOwner(hwnd) -> win.Show()
//     (Excel owner set automatically via ServiceLocator.ExcelHandle). CreateAndShow-
//     WaitWindow rewritten to the WpfAppManager.InvokeOnWpfThread(Action)-with-captured-
//     local pattern (InvokeOnWpfThread has no Func<T> overload here).
using GLSense.Addin.Core.Helpers;
using GLSense.Addin.Core.Infrastructure;
using GLSense.Addin.Core.Utilities;
using GLSense.Addin.Core.Views;
using System;
using System.Collections.Generic;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Threading;
using Excel = Microsoft.Office.Interop.Excel;

namespace GLSense.Addin.Core.Drilldowns
{
    public static class RowVisibilityProcessor
    {
        /// <summary>Ribbon click (RibHideRows).</summary>
        public static async Task RibHideRows_OnClick()
        {
            ServiceLocator.Logger?.LogDebug("RowVisibilityProcessor.RibHideRows_OnClick started.");
            var processor = new HideRowProcessor();
            await processor.ExecuteAsync("Hiding Rows");
        }

        /// <summary>Ribbon click (RibUnHideRows).</summary>
        public static async Task RibUnHideRows_OnClick()
        {
            ServiceLocator.Logger?.LogDebug("RowVisibilityProcessor.RibUnHideRows_OnClick started.");
            var processor = new UnhideRowProcessor();
            await processor.ExecuteAsync("Unhiding Rows");
        }

        public abstract class RowProcessor
        {
            public async Task ExecuteAsync(string operationName)
            {
                ServiceLocator.Logger?.LogDebug($"RowVisibilityProcessor.ExecuteAsync started. operationName='{operationName}'.");

                GLWaitWindow win = null;
                using var ctsHelper = new CancellationHelper();
                CancellationToken token = ctsHelper.GetToken();
                if (!CommonMethods.TryDisableExcelSettings($"RowVisibilityProcessor.ExecuteAsync ({operationName})"))
                    return;
                try
                {
                    if (!GuardLoginAndExcel()) return;
                    win = CreateAndShowWaitWindow(ctsHelper);
                    await InitializeWaitWindowAsync(win, operationName, operationName == "Hiding Rows" ? "Hiding rows for 0 balances…" : "Unhiding rows please wait...");
                    await Task.Yield();
                    Excel.Range selection = GetSelection();
                    Excel.Worksheet sheet = selection.Worksheet;
                    ServiceLocator.Logger?.LogDebug($"RowVisibilityProcessor.ExecuteAsync: selection address={selection.Address}, sheet='{sheet?.Name}'.");
                    Excel.Range balanceRange = await BalancesRangeAsync(selection);
                    if (balanceRange == null)
                    {
                        ServiceLocator.Logger?.LogDebug("RowVisibilityProcessor.ExecuteAsync: no balance formulas found in the selection, aborting.");
                        await SafelyCloseWaitWindowAsync(win);
                        CommonFunctions.GLSenseMessage("No balance formula's in the selection!", MessageBoxImage.Exclamation, MessageBoxButton.OK);
                        return;
                    }
                    await ProcessRowsCoreAsync(sheet, balanceRange, win, token);
                }
                catch (OperationCanceledException) { ServiceLocator.Logger?.LogWarn($"RowVisibilityProcessor.ExecuteAsync ({operationName}): operation cancelled by user."); }
                catch (Exception ex) { ServiceLocator.Logger?.LogException(ex, $"RowVisibilityProcessor.ExecuteAsync ({operationName})"); }
                finally
                {
                    await SafelyCloseWaitWindowAsync(win);
                    CommonMethods.TryEnableExcelSettings($"RowVisibilityProcessor.ExecuteAsync ({operationName})");
                    ServiceLocator.Logger?.LogDebug($"RowVisibilityProcessor.ExecuteAsync completed. operationName='{operationName}'.");
                }
            }
            protected abstract Task ProcessRowsCoreAsync(Excel.Worksheet sheet, Excel.Range selection, GLWaitWindow win, CancellationToken token);
        }

        public sealed class HideRowProcessor : RowProcessor
        {
            protected override async Task ProcessRowsCoreAsync(Excel.Worksheet sheet, Excel.Range selection, GLWaitWindow win, CancellationToken token)
            {
                token.ThrowIfCancellationRequested();
                var (formulas, values) = await GetFormulaAndValueArraysAsync(selection);
                if (formulas == null || values == null) return;
                token.ThrowIfCancellationRequested();
                var hideRows = FindHideRows(formulas, values, selection.Row);
                ServiceLocator.Logger?.LogDebug($"HideRowProcessor.ProcessRowsCoreAsync: found {hideRows.Count} row(s) to hide.");
                await ProcessHideRowsAsync(sheet, hideRows, win, token);
            }
        }

        public sealed class UnhideRowProcessor : RowProcessor
        {
            protected override async Task ProcessRowsCoreAsync(Excel.Worksheet sheet, Excel.Range selection, GLWaitWindow win, CancellationToken token)
            {
                token.ThrowIfCancellationRequested();
                var (formulas, values) = await GetFormulaAndValueArraysAsync(selection);
                if (formulas == null || values == null) return;
                token.ThrowIfCancellationRequested();
                var rowsToUnhide = FindHideRows(formulas, values, selection.Row);
                ServiceLocator.Logger?.LogDebug($"UnhideRowProcessor.ProcessRowsCoreAsync: found {rowsToUnhide.Count} row(s) to unhide.");
                if (rowsToUnhide.Count == 0) { await MessageWaitWindowAsync(win, "Nothing to unhide in the current selection."); return; }
                double standardHeight = sheet.StandardHeight;
                await ProcessUnhideRowsByBatchesAsync(sheet, rowsToUnhide, standardHeight, win, token);
            }
        }

        private static List<int> FindHideRows(object[,] formulas, object[,] values, int startRow)
        {
            var hideRows = new List<int>();
            int rLo = formulas.GetLowerBound(0), rHi = formulas.GetUpperBound(0), cLo = formulas.GetLowerBound(1), cHi = formulas.GetUpperBound(1);
            if (values.GetLength(0) != formulas.GetLength(0) || values.GetLength(1) != formulas.GetLength(1))
                throw new InvalidOperationException("Formulas and values arrays have different shapes.");
            for (int r = rLo; r <= rHi; r++)
            {
                if (ShouldHideRow(formulas, values, r, cLo, cHi)) hideRows.Add(startRow + (r - rLo));
            }
            return hideRows;
        }

        private static bool ShouldHideRow(object[,] formulas, object[,] values, int r, int cLo, int cHi)
        {
            for (int c = cLo; c <= cHi; c++)
            {
                if (IsGetBalanceFormula(formulas[r, c]) && IsZero(values[r, c])) return true;
            }
            return false;
        }

        private static bool IsGetBalanceFormula(object f)
        {
            if (f == null) return false;
            var s = f as string ?? f.ToString();
            if (string.IsNullOrWhiteSpace(s)) return false;
            return s.TrimStart('=', '@').IndexOf(AppConstants.glBal, StringComparison.OrdinalIgnoreCase) >= 0;
        }

        private static bool IsZero(object value) => value switch
        {
            null => true,
            double d => Math.Abs(d) < 1e-9,
            int i => i == 0,
            decimal m => m == 0m,
            string s => double.TryParse(s, out double parsed) && Math.Abs(parsed) < 1e-9,
            _ => false,
        };

        private static bool GuardLoginAndExcel() => AppState.Instance.IsLoginCompleted && ServiceLocator.ExcelApp != null;

        private static Excel.Range GetSelection() => ServiceLocator.ExcelApp.Selection as Excel.Range ?? throw new InvalidOperationException("No selection available");

        private static async Task<(object[,] formulas, object[,] values)> GetFormulaAndValueArraysAsync(Excel.Range selection)
        {
            await Task.Yield();
            return (CoerceTo2D(selection.Formula), CoerceTo2D(selection.Value2));
        }

        private static object[,] CoerceTo2D(object value) => value switch
        {
            object[,] array2d => array2d,
            null => new object[1, 1] { { null } },
            _ => new object[1, 1] { { value } }
        };

        private static async Task<Excel.Range> BalancesRangeAsync(Excel.Range selection)
        {
            string rngAddress = ExcelExternalRef.BuildExternalAddress(selection);
            Excel.Range totalRange = CommonFunctions.GetBalanceTotalRange(rngAddress);
            if (totalRange != null) return totalRange;
            await Task.Yield();
            return null;
        }

        private static async Task ProcessHideRowsAsync(Excel.Worksheet sheet, List<int> hideRows, GLWaitWindow win, CancellationToken token)
        {
            if (hideRows == null || hideRows.Count == 0) return;
            hideRows.Sort();
            int i = 0;
            while (i < hideRows.Count)
            {
                token.ThrowIfCancellationRequested();
                int start = hideRows[i], end = start;
                i++;
                while (i < hideRows.Count && hideRows[i] == end + 1) { end = hideRows[i]; i++; }
                sheet.Range[$"{start}:{end}"].RowHeight = 0.1;
                await MessageWaitWindowAsync(win, $"Hid rows {start}:{end}…");
                await Task.Yield();
            }
        }

        private static async Task ProcessUnhideRowsByBatchesAsync(Excel.Worksheet sheet, List<int> unhideRows, double standardHeight, GLWaitWindow win, CancellationToken token)
        {
            if (unhideRows == null || unhideRows.Count == 0) return;
            unhideRows.Sort();
            int i = 0;
            while (i < unhideRows.Count)
            {
                token.ThrowIfCancellationRequested();
                int start = unhideRows[i], end = start;
                i++;
                while (i < unhideRows.Count && unhideRows[i] == end + 1) { end = unhideRows[i]; i++; }
                var rng = sheet.Rows[$"{start}:{end}"] as Excel.Range;
                if (rng != null) rng.RowHeight = standardHeight;
                await MessageWaitWindowAsync(win, $"Unhid rows {start}:{end}…");
                await Task.Yield();
            }
            await MessageWaitWindowAsync(win, "Completed successfully.");
        }

        private static GLWaitWindow CreateAndShowWaitWindow(CancellationHelper cts)
        {
            try
            {
                GLWaitWindow win = null;

                // WpfAppManager.InvokeOnWpfThread only takes an Action (no return value),
                // so capture the created window from inside the delegate - same pattern
                // DrillCellHighlighter.cs/BalanceHighlighter.cs use for GLWaitWindow.
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
            catch (Exception ex) { ServiceLocator.Logger?.LogException(ex); return null; }
        }

        private static Task InitializeWaitWindowAsync(GLWaitWindow win, string title, string message)
        {
            if (win == null || win.Dispatcher == null) return Task.CompletedTask;
            try
            {
                return win.Dispatcher.InvokeAsync(() => { win.SetProcessTitle(title); win.SetProcessMessage(message); }, DispatcherPriority.Normal).Task;
            }
            catch (TaskCanceledException) { ServiceLocator.Logger?.LogDebug("RowVisibilityProcessor.InitializeWaitWindowAsync: dispatcher invoke was cancelled (window likely closing)."); return Task.CompletedTask; }
            catch (Exception ex) { ServiceLocator.Logger?.LogException(ex, "RowVisibilityProcessor.InitializeWaitWindowAsync"); return Task.CompletedTask; }
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
            catch (TaskCanceledException) { ServiceLocator.Logger?.LogDebug("RowVisibilityProcessor.MessageWaitWindowAsync: dispatcher invoke was cancelled (window likely closing)."); return Task.CompletedTask; }
            catch (Exception ex) { ServiceLocator.Logger?.LogException(ex, "RowVisibilityProcessor.MessageWaitWindowAsync"); return Task.CompletedTask; }
        }
    }
}
