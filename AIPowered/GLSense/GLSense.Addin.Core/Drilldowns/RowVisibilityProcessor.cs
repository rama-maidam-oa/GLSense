// RowVisibilityProcessor.cs in GLSense.Addin.Core
// Group F - extracted from GLSense\AddinModule.cs (FinalWorkingCode), RibHideRows_OnClick/
// RibUnHideRows_OnClick (old lines ~2829-3006), including the RowProcessor/
// HideRowProcessor/UnhideRowProcessor class hierarchy and helpers (ScanBalanceRowsAsync,
// GetFormulaCellsWithinArea, IsGetBalanceFormula, IsZero, GetSelection, ProcessHideRowsAsync,
// ProcessUnhideRowsByBatchesAsync),
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
using System.Runtime.InteropServices;
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
                    BalanceScanResult scan = await ScanBalanceRowsAsync(selection, win, token);
                    if (!scan.AnyBalanceFormulaFound)
                    {
                        ServiceLocator.Logger?.LogDebug("RowVisibilityProcessor.ExecuteAsync: no balance formulas found in the selection, aborting.");
                        await SafelyCloseWaitWindowAsync(win);
                        CommonFunctions.GLSenseMessage("No balance formula's in the selection!", MessageBoxImage.Exclamation, MessageBoxButton.OK);
                        return;
                    }
                    await ProcessRowsCoreAsync(sheet, scan, win, token);
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
            protected abstract Task ProcessRowsCoreAsync(Excel.Worksheet sheet, BalanceScanResult scan, GLWaitWindow win, CancellationToken token);
        }

        public sealed class HideRowProcessor : RowProcessor
        {
            protected override async Task ProcessRowsCoreAsync(Excel.Worksheet sheet, BalanceScanResult scan, GLWaitWindow win, CancellationToken token)
            {
                token.ThrowIfCancellationRequested();
                // Hide criterion: EVERY GLSense_GetBalance cell found in that row (within
                // the selected columns) must be zero - a row is left alone if even one
                // balance formula cell in it is non-zero.
                var hideRows = new List<int>();
                foreach (var kvp in scan.RowHasNonZero)
                {
                    if (!kvp.Value) hideRows.Add(kvp.Key);
                }
                hideRows.Sort();
                ServiceLocator.Logger?.LogDebug($"HideRowProcessor.ProcessRowsCoreAsync: found {hideRows.Count} row(s) to hide.");
                await ProcessHideRowsAsync(sheet, hideRows, win, token);
            }
        }

        public sealed class UnhideRowProcessor : RowProcessor
        {
            protected override async Task ProcessRowsCoreAsync(Excel.Worksheet sheet, BalanceScanResult scan, GLWaitWindow win, CancellationToken token)
            {
                token.ThrowIfCancellationRequested();
                // Unhide criterion is deliberately NOT "every balance cell in the row is
                // still zero" (that's Hide's job, already done when the row was
                // collapsed). Re-checking zero-ness here means the result depends on
                // exactly which columns happen to be selected THIS time - e.g. hiding via
                // J10:K10 then trying to unhide via a wider I9:J11 would drag in column
                // I's own (possibly non-zero) balance cell for row 10 and wrongly refuse
                // to unhide it, even though row 10 was never hidden based on column I at
                // all. Instead: any row that has a GLSense_GetBalance formula anywhere in
                // the selection AND is currently collapsed (RowHeight ~0.1, the height
                // ProcessHideRowsAsync sets) gets restored - independent of which columns
                // are selected or what today's live balance value is.
                var candidateRows = new List<int>(scan.RowHasNonZero.Keys);
                candidateRows.Sort();
                var rowsToUnhide = new List<int>();
                foreach (int row in candidateRows)
                {
                    token.ThrowIfCancellationRequested();
                    var rowRange = sheet.Rows[row] as Excel.Range;
                    if (rowRange != null && (double)rowRange.RowHeight < 1.0) rowsToUnhide.Add(row);
                }
                ServiceLocator.Logger?.LogDebug($"UnhideRowProcessor.ProcessRowsCoreAsync: found {rowsToUnhide.Count} row(s) to unhide.");
                if (rowsToUnhide.Count == 0) { await MessageWaitWindowAsync(win, "Nothing to unhide in the current selection."); return; }
                double standardHeight = sheet.StandardHeight;
                await ProcessUnhideRowsByBatchesAsync(sheet, rowsToUnhide, standardHeight, win, token);
            }
        }

        // ComVisible(false): this assembly is [assembly: ComVisible(true)] overall, but
        // this struct is an internal RowProcessor implementation detail, never touched via
        // COM - it only had to become `public` to satisfy C#'s "parameter type must be at
        // least as accessible as the method" rule on ProcessRowsCoreAsync. Its
        // Dictionary<int,bool> property can't be represented in a COM type library anyway
        // (generic types aren't COM-exportable) - matches the identical attribute/reasoning
        // added to FinalWorkingCode's copy of this same struct, where the type library
        // exporter (RegisterForComInterop=true there) actually warns about this at build
        // time.
        [ComVisible(false)]
        public readonly struct BalanceScanResult
        {
            public BalanceScanResult(bool anyBalanceFormulaFound, Dictionary<int, bool> rowHasNonZero)
            {
                AnyBalanceFormulaFound = anyBalanceFormulaFound;
                RowHasNonZero = rowHasNonZero;
            }
            public bool AnyBalanceFormulaFound { get; }
            /// <summary>Every distinct row (within the selection) that has at least one
            /// GLSense_GetBalance formula cell, mapped to whether ANY such cell in that row
            /// is non-zero (true) or every one of them is currently zero (false).</summary>
            public Dictionary<int, bool> RowHasNonZero { get; }
        }

        /// <summary>
        /// Scans <paramref name="selection"/>'s own Areas directly (via SpecialCells, same
        /// as CommonFunctions.GetFormulaCellsWithinArea) for GLSense_GetBalance formula
        /// cells and groups them by row - shared by both Hide (needs "is every cell in
        /// this row zero") and Unhide (needs "which rows have a balance formula at all",
        /// then separately checks each candidate row's current height - see
        /// UnhideRowProcessor's own comment for why it doesn't reuse the zero flag).
        ///
        /// Deliberately does NOT go through CommonFunctions.GetBalanceTotalRange: that
        /// helper reconstructs its result by calling Application.Union() once per matching
        /// cell, which for a large selection (e.g. E7:L393, ~3,000+ matching cells) means
        /// thousands of sequential, un-yielded COM calls before this method would even get
        /// to look at a single value - this is what caused the reported "Excel freezes for
        /// a few seconds before the wait window starts animating" symptom (all of that
        /// Union-building ran synchronously before ProcessHideRowsAsync's own per-batch
        /// yields ever got a chance to run). Scanning the selection's Areas/SpecialCells
        /// directly needs exactly one COM round trip per matching cell (Formula + Value2 +
        /// Row), not two-plus-a-Union, and periodically yields + updates the wait window
        /// with real progress so the UI never looks stuck. Still stays cheap and bounded
        /// even if the user selects a whole column (e.g. E:L): SpecialCells only returns
        /// cells that actually contain a formula, never the millions of blank rows between
        /// them.
        /// </summary>
        private static async Task<BalanceScanResult> ScanBalanceRowsAsync(Excel.Range selection, GLWaitWindow win, CancellationToken token)
        {
            bool anyBalanceFormulaFound = false;
            var rowHasNonZero = new Dictionary<int, bool>();
            Excel.Application app = selection.Application;
            int scanned = 0;

            foreach (Excel.Range area in selection.Areas)
            {
                Excel.Range formulaCells = GetFormulaCellsWithinArea(area, app);
                if (formulaCells == null) continue;

                foreach (Excel.Range cell in formulaCells.Cells)
                {
                    token.ThrowIfCancellationRequested();

                    if (IsGetBalanceFormula(cell.Formula))
                    {
                        anyBalanceFormulaFound = true;
                        int row = cell.Row;
                        bool alreadyNonZero = rowHasNonZero.TryGetValue(row, out bool prior) && prior;
                        rowHasNonZero[row] = alreadyNonZero || !IsZero(cell.Value2);
                    }

                    scanned++;
                    if (scanned % 200 == 0)
                    {
                        await MessageWaitWindowAsync(win, $"Scanning balance formulas… ({scanned} checked)");
                        await Task.Yield();
                    }
                }
            }

            return new BalanceScanResult(anyBalanceFormulaFound, rowHasNonZero);
        }

        /// <summary>Local copy of CommonFunctions.GetFormulaCellsWithinArea (private there) -
        /// per this file's established per-file-duplication convention.</summary>
        private static Excel.Range GetFormulaCellsWithinArea(Excel.Range area, Excel.Application app)
        {
            if (area.Rows.Count == 1 && area.Columns.Count == 1)
            {
                try { return (bool)area.HasFormula ? area : null; }
                catch (COMException ex)
                {
                    ServiceLocator.Logger?.LogWarn($"GetFormulaCellsWithinArea: HasFormula threw for a single-cell area, treating as not-a-formula-cell: {ex.Message}");
                    return null;
                }
            }

            try
            {
                var cand = area.SpecialCells(Excel.XlCellType.xlCellTypeFormulas);
                return app.Intersect(cand, area);
            }
            catch (COMException)
            {
                // Thrown if no formula cells exist in this area.
                return null;
            }
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
