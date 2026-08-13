// RangeRefresher.cs in GLSense.Addin.Core
// Group F - extracted from GLSense\AddinModule.cs (FinalWorkingCode), RibRefreshRange_OnClick
// (old lines ~1644-1745) and its private helpers (CountBalanceFormulas,
// ValidateRefreshRange, RefreshFormulaCells), plus the shared wait-window helpers the old
// host also used (CreateAndShowWaitWindow/InitializeWaitWindowAsync/
// SafelyCloseWaitWindowAsync/MessageWaitWindowAsync) - ported here as private copies, per
// the established per-file-duplication convention (DD_BL.cs/DD_JL.cs/DD_SL.cs/
// DDDatatoWorksheet.cs/BalanceHighlighter.cs/RowVisibilityProcessor.cs all carry their own
// copies rather than sharing one utility class).
//
// Given its own class (rather than folded into AddinEntry.cs) to keep it consistent with
// BalanceHighlighter.cs/RowVisibilityProcessor.cs above - all three Group F
// extractions need the same wait-window scaffolding, and AddinEntry.cs already has one
// inline wait-window user (LedgerChanged, Group B) that intentionally stays lean; adding
// three more full progress-window flows directly there would make it unwieldy.
//
// Refreshes (recalculates) only the balance formulas in the current selection, respecting
// UserConfig.RefreshCells as a max-cells guard, and sets AppState.Instance.SingleRefresh
// true for the duration of the refresh loop.
//
// Re-pointed vs. the original:
//   - AppState.Instance.ExcelApp -> ServiceLocator.ExcelApp (this project's AppState has
//     no ExcelApp field).
//   - LogUtility.* (static) -> ServiceLocator.Logger?.*.
//   - System.Windows.Forms MessageBoxIcon/MessageBoxButtons -> System.Windows
//     MessageBoxImage/MessageBoxButton (this project has no WinForms reference).
//   - GLWaitWindow now derives from BaseWindow: win.ShowWithOwner(hwnd) -> win.Show().
using GLSense.Addin.Core.Helpers;
using GLSense.Addin.Core.Infrastructure;
using GLSense.Addin.Core.Utilities;
using GLSense.Addin.Core.Views;
using System;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Threading;
using Excel = Microsoft.Office.Interop.Excel;

namespace GLSense.Addin.Core.Drilldowns
{
    public static class RangeRefresher
    {
        /// <summary>Ribbon click (RibRefreshRange).</summary>
        public static async Task RibRefreshRange_OnClick()
        {
            ServiceLocator.Logger?.LogDebug("RangeRefresher.RibRefreshRange_OnClick started.");

            if (!GuardLoginAndExcel())
            {
                CommonFunctions.GLSenseMessage("Please log in to the instance.", MessageBoxImage.Exclamation, MessageBoxButton.OK);
                return;
            }

            GLWaitWindow win = null;
            using var ctsHelper = new CancellationHelper();
            CancellationToken token = ctsHelper.GetToken();

            if (!CommonMethods.TryDisableExcelSettings("RangeRefresher.RibRefreshRange_OnClick"))
                return;

            try
            {
                Excel.Range formulaCells = ServiceLocator.ExcelApp.Selection as Excel.Range;
                int glBalCount = CountBalanceFormulas(formulaCells, token);

                if (!ValidateRefreshRange(glBalCount))
                    return;

                token.ThrowIfCancellationRequested();

                win = CreateAndShowWaitWindow(ctsHelper);

                await InitializeWaitWindowAsync(win, "Range Refresh", "Refreshing selected range...");
                await Task.Yield();

                await RefreshFormulaCells(formulaCells, win, token);
            }
            catch (OperationCanceledException)
            {
                ServiceLocator.Logger?.LogWarn("Range refresh operation was canceled by the user.");
            }
            catch (Exception ex)
            {
                ServiceLocator.Logger?.LogException(ex);
            }
            finally
            {
                CommonMethods.TryEnableExcelSettings("RangeRefresher.RibRefreshRange_OnClick");
                await SafelyCloseWaitWindowAsync(win);
            }
        }

        private static int CountBalanceFormulas(Excel.Range formulaCells, CancellationToken token)
        {
            int count = 0;
            foreach (Excel.Range cell in formulaCells)
            {
                token.ThrowIfCancellationRequested();

                if (cell.HasFormula is true &&
                    cell.Formula is string formula &&
                    formula.IndexOf(AppConstants.glBal, StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    count++;
                }
            }
            return count;
        }

        private static bool ValidateRefreshRange(int glBalCount)
        {
            if (glBalCount == 0)
            {
                CommonFunctions.GLSenseMessage("No balance formulas exists in the selected range.", MessageBoxImage.Exclamation, MessageBoxButton.OK);
                return false;
            }

            if (glBalCount > UserConfig.RefreshCells)
            {
                string msg = $"The selected range contains {glBalCount} balance formulas.\nThe configured refresh range is {UserConfig.RefreshCells}.\nChange the configuration and try again! Max refresh range limit is 100.";
                CommonFunctions.GLSenseMessage(msg, MessageBoxImage.Warning, MessageBoxButton.OK);
                return false;
            }

            return true;
        }

        private static async Task RefreshFormulaCells(Excel.Range formulaCells, GLWaitWindow win, CancellationToken token)
        {
            try
            {
                // NOTE: matches the original exactly - SingleRefresh is intentionally only
                // reset to false in the catch block below (not in a finally), same as
                // GLSense\AddinModule.cs's RefreshFormulaCells. Not "fixed" here per the
                // porting rule to preserve logic exactly and only re-point plumbing.
                AppState.Instance.SingleRefresh = true;
                foreach (Excel.Range cell in formulaCells)
                {
                    token.ThrowIfCancellationRequested();
                    await MessageWaitWindowAsync(win, $"Refreshing range {cell.Address}.");
                    await Task.Yield();
                    cell.Dirty();
                    cell.Calculate();
                }
                await MessageWaitWindowAsync(win, "Completed refreshing the range.");
            }
            catch (Exception ex)
            {
                AppState.Instance.SingleRefresh = false;
                ServiceLocator.Logger?.LogException(ex);
            }
        }

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
            catch (Exception ex) { ServiceLocator.Logger?.LogException(ex); return null; }
        }

        private static Task InitializeWaitWindowAsync(GLWaitWindow win, string title, string message)
        {
            if (win == null || win.Dispatcher == null) return Task.CompletedTask;
            try
            {
                return win.Dispatcher.InvokeAsync(() => { win.SetProcessTitle(title); win.SetProcessMessage(message); }, DispatcherPriority.Normal).Task;
            }
            catch (TaskCanceledException) { ServiceLocator.Logger?.LogDebug("RangeRefresher.InitializeWaitWindowAsync: dispatcher invoke was cancelled (window likely closing)."); return Task.CompletedTask; }
            catch (Exception ex) { ServiceLocator.Logger?.LogException(ex, "RangeRefresher.InitializeWaitWindowAsync"); return Task.CompletedTask; }
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
            catch (TaskCanceledException) { ServiceLocator.Logger?.LogDebug("RangeRefresher.MessageWaitWindowAsync: dispatcher invoke was cancelled (window likely closing)."); return Task.CompletedTask; }
            catch (Exception ex) { ServiceLocator.Logger?.LogException(ex, "RangeRefresher.MessageWaitWindowAsync"); return Task.CompletedTask; }
        }
    }
}
