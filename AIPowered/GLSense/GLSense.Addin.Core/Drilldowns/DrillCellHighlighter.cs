// DrillCellHighlighter.cs in GLSense.Addin.Core
// Ported from GLSense\Drilldowns\DrillCellHighlighter.cs (FinalWorkingCode), including
// the two fixes already made there (deterministic Range.DirectPrecedents-based cell
// scan instead of the flaky Range.NavigateArrow tracer-arrow walk, and iterating
// SpecialCells results via Areas instead of the unreliable multi-area Cells indexer).
//
// Adjustments made when porting into this project's architecture:
//   - AppState.Instance.ExcelApp / IsLoginCompleted -> ServiceLocator.ExcelApp / this
//     project's own AppState.Instance.IsLoginCompleted.
//   - LogUtility.* (static) -> ServiceLocator.Logger.* (instance via context).
//   - GLWaitWindow now derives from BaseWindow (WPF-UI FluentWindow) rather than the
//     old DpiAwareWindow, and sets its Excel owner automatically via
//     ServiceLocator.ExcelHandle/ModalToExcel - no explicit ShowWithOwner() call needed.
//   - System.Windows.Forms.MessageBoxIcon/MessageBoxButtons -> System.Windows
//     MessageBoxImage/MessageBoxButton (this project has no WinForms reference; the
//     enum member names used here - Exclamation/Error/OK - exist under both).
using GLSense.Addin.Core.Helpers;
using GLSense.Addin.Core.Infrastructure;
using GLSense.Addin.Core.Utilities;
using GLSense.Addin.Core.Views;
using System;
using System.Collections.Generic;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using Excel = Microsoft.Office.Interop.Excel;

namespace GLSense.Addin.Core.Drilldowns
{
    public static class DrillCellHighlighter
    {
        private static GLWaitWindow Win { get; set; }
        private static CancellationHelper _ctsHelper;
        private static CancellationToken Token => _ctsHelper?.GetToken() ?? default;

        /// <summary>
        /// Ribbon click: Selects cells that reference balance formulas (GLSense_GetBalance) on the active sheet.
        /// </summary>
        public static async Task RibCellHighlight_OnClick()
        {
            using (ServiceLocator.Logger.BeginLogScope("DrillCellHighlighter.RibCellHighlight_OnClick"))
            {
            _ctsHelper = new CancellationHelper();
            var excelApp = ServiceLocator.ExcelApp;
            bool snapshotTaken = false;
            bool originalScreenUpdating = true;
            bool originalDisplayAlerts = true;
            bool originalEnableEvents = true;
            bool originalDisplayStatusBar = true;
            bool originalInteractive = true;
            Excel.XlCalculation originalCalculation = Excel.XlCalculation.xlCalculationAutomatic;

            try
            {
                if (!GuardLoginAndExcel()) return;

                excelApp = ServiceLocator.ExcelApp;
                if (excelApp != null)
                {
                    snapshotTaken = true;
                    originalScreenUpdating = excelApp.ScreenUpdating;
                    originalDisplayAlerts = excelApp.DisplayAlerts;
                    originalEnableEvents = excelApp.EnableEvents;
                    originalDisplayStatusBar = excelApp.DisplayStatusBar;
                    originalInteractive = excelApp.Interactive;
                    originalCalculation = excelApp.Calculation;

                    excelApp.ScreenUpdating = false;
                    excelApp.DisplayAlerts = false;
                    excelApp.EnableEvents = false;
                    excelApp.DisplayStatusBar = false;
                    excelApp.Interactive = false;
                    excelApp.Calculation = Excel.XlCalculation.xlCalculationManual;
                }

                CommonMethods.DisableExcelSettings();

                var actCell = ServiceLocator.ExcelApp.ActiveCell;
                if (ServiceLocator.ExcelApp.ActiveSheet is not Excel.Worksheet wrksht || actCell == null) return;

                ServiceLocator.Logger.LogDebug("Selecting cells which have reference to balance formula(s).");
                ServiceLocator.Logger.LogDebug($"Worksheet Name : {wrksht.Name}");

                if (!CommonFunctions.BalanceFormulaExists(wrksht.Name))
                {
                    await ShowMessageAsync("No getbalance formulas in the current sheet.", MessageBoxImage.Exclamation);
                    return;
                }

                Win = CreateAndShowProgressWindow(_ctsHelper);

                Excel.Range formulaCells = await fRange(wrksht);
                if (formulaCells == null) return;
                Token.ThrowIfCancellationRequested();

                await InitializeProgressWindowAsync("Drill Cell Highlighter", "Processing request...");

                string external = ExcelExternalRef.BuildExternalAddress(formulaCells);

                Token.ThrowIfCancellationRequested();

                Excel.Range totalRange = null;

                totalRange = CommonFunctions.GetBalanceTotalRange(external);

                if (totalRange == null)
                {
                    await ShowMessageAsync("No cells found which have reference to balance formula(s).", MessageBoxImage.Exclamation);
                    actCell.Select();
                    return;
                }

                ServiceLocator.Logger.LogDebug($"Balance cells found : {totalRange.Cells.Address}");

                var uniqueDependentAddresses = CollectDependentsAddresses(totalRange, wrksht);

                if (uniqueDependentAddresses.Count > 0)
                {
                    if (TryBuildUnionRangeFromAddresses(uniqueDependentAddresses, out var selection))
                    {
                        selection?.Select();
                    }
                }
                else
                {
                    ServiceLocator.Logger.LogDebug("No dependent cells found.");
                    await ShowMessageAsync("No cells found which have reference to balance formula(s).", MessageBoxImage.Exclamation);
                    actCell.Select();
                }
            }
            catch (OperationCanceledException)
            {
                ServiceLocator.Logger.LogWarn("Drill cell highlight operation cancelled by user.");
            }
            catch (Exception ex)
            {
                ServiceLocator.Logger.LogException(ex, "Exception in finding balance formulas dependencies.");
            }
            finally
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
                    ServiceLocator.Logger.LogWarn($"DrillCellHighlighter.RibCellHighlight_OnClick: exception disposing CancellationHelper (ignored): {ex.Message}");
                }
                CommonMethods.TryEnableExcelSettings("DrillCellHighlighter.RibCellHighlight_OnClick");

                if (snapshotTaken && excelApp != null)
                {
                    try
                    {
                        excelApp.ScreenUpdating = originalScreenUpdating;
                        excelApp.DisplayAlerts = originalDisplayAlerts;
                        excelApp.EnableEvents = originalEnableEvents;
                        excelApp.DisplayStatusBar = originalDisplayStatusBar;
                        excelApp.Interactive = originalInteractive;
                        excelApp.Calculation = originalCalculation;
                    }
                    catch (Exception restoreEx)
                    {
                        ServiceLocator.Logger.LogException(restoreEx, "Error restoring Excel settings after drill cell highlighting.");
                    }
                }

                ServiceLocator.Logger.LogDebug("Selecting cells which have reference to balance formula(s) completed.");
                await SafelyCloseWindowAsync();
            }
            }
        }

        private static async Task<Excel.Range> fRange(Excel.Worksheet ws)
        {
            try
            {
                Excel.Range rng = null;
                // Get all cells on sheet that contain formulas.
                // SpecialCells(xlCellTypeFormulas) throws if there are no such cells.
                rng = ws.Cells.SpecialCells(
                    Excel.XlCellType.xlCellTypeFormulas,
                    Type.Missing);

                return rng;
            }
            catch (System.Runtime.InteropServices.COMException ex)
            {
                // Thrown when there are no formula cells on the sheet; treat as "not found".
                ServiceLocator.Logger.LogException(ex);
                await ShowMessageAsync($"An unexpected error occurred.{Environment.NewLine}{ex.Message}", MessageBoxImage.Error);
            }
            catch (Exception ex)
            {
                ServiceLocator.Logger.LogException(ex);
                await ShowMessageAsync($"An unexpected error occurred.{Environment.NewLine}{ex.Message}", MessageBoxImage.Error);
            }
            return null;
        }

        private static bool GuardLoginAndExcel()
        {
            return AppState.Instance.IsLoginCompleted && ServiceLocator.ExcelApp != null;
        }

        private static async Task ShowMessageAsync(string message, MessageBoxImage icon)
        {
            await SafelyCloseWindowAsync();
            CommonFunctions.GLSenseMessage(
                message,
                icon,
                MessageBoxButton.OK);
        }

        private static GLWaitWindow CreateAndShowProgressWindow(CancellationHelper cts)
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
                        ServiceLocator.Logger.LogException(ex);
                        win = null;
                    }
                });

                return win;
            }
            catch (Exception ex)
            {
                ServiceLocator.Logger.LogException(ex);
            }
            return null;
        }

        private static Task InitializeProgressWindowAsync(string title, string message)
        {
            // Fire-and-forget: progress UI update only. Do not introduce a
            // suspend point here - the caller reads Excel COM objects
            // immediately after awaiting this method.
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
                ServiceLocator.Logger.LogError($"Error closing wait window: {ex.Message}");
            }
        }

        /// <summary>
        /// Returns unique addresses (external A1 form) of every formula cell on <paramref name="wrksht"/>
        /// whose formula references (directly or indirectly) one or more cells in <paramref name="balancesRng"/>.
        /// </summary>
        /// <remarks>
        /// Uses Range.DirectPrecedents (deterministic) rather than tracer arrows - see the
        /// FinalWorkingCode fix history for why the tracer-arrow approach was unreliable.
        /// </remarks>
        private static HashSet<string> CollectDependentsAddresses(Excel.Range balancesRng, Excel.Worksheet wrksht)
        {
            Token.ThrowIfCancellationRequested();

            var result = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            if (balancesRng == null || wrksht == null)
                return result;

            Excel.Range allFormulaCells;
            try
            {
                allFormulaCells = wrksht.Cells.SpecialCells(Excel.XlCellType.xlCellTypeFormulas, Type.Missing);
            }
            catch (Exception ex)
            {
                ServiceLocator.Logger.LogException(ex);
                return result;
            }

            var app = wrksht.Application;

            // SpecialCells almost always returns a multi-area range (formula cells are rarely
            // one contiguous block). Indexing a multi-area range's Cells collection directly
            // (Cells[i]) is unreliable in Excel COM automation - it only reaches the first area.
            // Iterate Areas explicitly instead.
            foreach (Excel.Range area in allFormulaCells.Areas)
            {
                foreach (Excel.Range candidate in area.Cells)
                {
                    Token.ThrowIfCancellationRequested();

                    if (candidate == null)
                        continue;

                    // Skip the balance formula cells themselves - only cells that *reference* them qualify.
                    if (app.Intersect(candidate, balancesRng) != null)
                        continue;

                    var visited = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                    if (!CellReferencesBalanceRange(candidate, balancesRng, visited))
                        continue;

                    string address = candidate.Address[RowAbsolute: false,
                                                  ColumnAbsolute: false,
                                                  ReferenceStyle: Excel.XlReferenceStyle.xlA1,
                                                  External: true]; // external A1-style address

                    ServiceLocator.Logger.LogDebug($"Found dependencies at : {address}");
                    result.Add(address);
                }
            }

            return result;
        }

        /// <summary>
        /// True if <paramref name="cell"/> directly or indirectly (through one or more other formulas)
        /// references any cell in <paramref name="balancesRng"/>.
        /// </summary>
        private static bool CellReferencesBalanceRange(Excel.Range cell, Excel.Range balancesRng, HashSet<string> visited)
        {
            if (cell == null) return false;

            string key = cell.Address[true, true, Excel.XlReferenceStyle.xlA1, true];
            if (!visited.Add(key))
                return false; // already inspected on this path - guards against circular references

            Excel.Range precedents;
            try
            {
                if (cell.HasFormula is not bool hasFormula || !hasFormula)
                    return false;

                precedents = cell.DirectPrecedents;
            }
            catch
            {
                // No precedents (e.g. a formula made up only of literal/constant operands).
                return false;
            }

            if (precedents == null) return false;

            var app = cell.Application;

            // Fast path: does this cell reference a balance cell directly?
            if (app.Intersect(precedents, balancesRng) != null)
                return true;

            // Otherwise recurse into each precedent cell to catch indirect references
            // (e.g. a helper cell that itself sums a balance-formula cell).
            foreach (Excel.Range area in precedents.Areas)
            {
                Token.ThrowIfCancellationRequested();

                foreach (Excel.Range precCell in area.Cells)
                {
                    if (precCell.HasFormula is bool hf && hf &&
                        CellReferencesBalanceRange(precCell, balancesRng, visited))
                    {
                        return true;
                    }
                }
            }

            return false;
        }

        /// <summary>
        /// Builds a union range from a set of external A1 addresses.
        /// </summary>
        private static bool TryBuildUnionRangeFromAddresses(IEnumerable<string> addresses, out Excel.Range unionRange)
        {
            unionRange = null;

            var app = ServiceLocator.ExcelApp;
            if (app == null || addresses == null)
                return false;

            Excel.Range built = null;

            foreach (var addr in addresses)
            {
                // Skip null/empty addresses
                if (string.IsNullOrWhiteSpace(addr)) continue;

                // Resolve address to a Range (external A1 addresses supported)
                var rng = app.Range[addr];
                if (rng == null) continue;

                built = (built == null) ? rng : app.Union(built, rng);
            }

            unionRange = built;
            return unionRange != null;
        }
    }
}
