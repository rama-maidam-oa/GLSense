// CommonMethods.cs in GLSense.Addin.Core
// Port of GLSense\Utilities\CommonMethods.cs (FinalWorkingCode).
// Ported: DisableExcelSettings/EnableExcelSettings (already here), Clear_Sheet,
// BalanceFormulas_Updation, Get_GLSense_MultiFormulas (+ its nested MultiFormulaDetector).
// Intentionally NOT ported: EnsureExcelApp - it re-fetched AppState.Instance.ExcelApp
// from AddinModule.CurrentInstance/ADXAddinModule.CurrentInstance, which was the old
// project's way of recovering from a stale Excel reference. In this project
// ServiceLocator.ExcelApp is supplied by the host via IGLSenseContext and is always
// current (the host re-Initializes the AppDomain on every load), so there is no
// equivalent "reinitialize" operation needed here - callers should just null-check
// ServiceLocator.ExcelApp instead.
// Changes vs. the original: AppState.Instance.ExcelApp -> ServiceLocator.ExcelApp;
// LogUtility.* -> ServiceLocator.Logger.*; ClsFormulaParser/ExcelComHelper resolve via
// this project's Helpers namespace (both already ported, see GLSense.Addin.Core\Helpers).
using GLSense.Addin.Core.Helpers;
using GLSense.Addin.Core.Infrastructure;
using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using Excel = Microsoft.Office.Interop.Excel;

namespace GLSense.Addin.Core.Utilities
{
    public static class CommonMethods
    {
        /// <summary>
        /// Clears all cells on the active worksheet.
        /// </summary>
        public static void Clear_Sheet()
        {
            try
            {
                ServiceLocator.Logger?.LogDebug("Clearing active worksheet");
                Excel.Worksheet wrksht = ServiceLocator.ExcelApp?.ActiveSheet as Excel.Worksheet;
                if (wrksht != null)
                {
                    ServiceLocator.Logger?.LogDebug($"Clearing sheet: {wrksht.Name}");
                    wrksht.Cells.Clear();
                    ServiceLocator.Logger?.LogDebug("Sheet cleared successfully");
                }
                else
                {
                    ServiceLocator.Logger?.LogWarn("Active worksheet is null");
                }
            }
            catch (Exception ex)
            {
                ServiceLocator.Logger?.LogException(ex, "Failed to clear sheet");
                throw;
            }
        }

        /// <summary>
        /// Corrects GLSense_GetBalance formulas (adds missing trailing args) for every
        /// cell address in <paramref name="BalancesDict"/> (address -> current formula text).
        /// </summary>
        public static void BalanceFormulas_Updation(Dictionary<string, string> BalancesDict)
        {
            try
            {
                if (BalancesDict == null || BalancesDict.Keys.Count == 0)
                {
                    ServiceLocator.Logger?.LogDebug("No balance formulas to update");
                    return;
                }

                ServiceLocator.Logger?.LogDebug($"Updating {BalancesDict.Keys.Count} balance formulas");

                var app = ServiceLocator.ExcelApp ?? throw new InvalidOperationException("Excel unavailable");

                foreach (var key in BalancesDict.Keys)
                {
                    try
                    {
                        ServiceLocator.Logger?.LogDebug($"Processing formula at {key}");

                        var parser = new ClsFormulaParser(BalancesDict[key]);
                        int ArgsCount = parser.Formula_ArgsCount();
                        int expectedArgs = ArgsCount + 2;

                        string correctedFormula = parser.Formula_Correction(expectedArgs, 8);
                        Excel.Range rng = app.Range[key];

                        if (rng != null)
                        {
                            ServiceLocator.Logger?.LogDebug($"Old Formula: {BalancesDict[key]}");
                            ServiceLocator.Logger?.LogDebug($"New Formula: {correctedFormula}");
                            rng.Value2 = correctedFormula;
                            ServiceLocator.Logger?.LogDebug($"Formula updated successfully at {key}");
                        }
                        else
                        {
                            ServiceLocator.Logger?.LogWarn($"Range is null at {key}");
                        }
                    }
                    catch (Exception ex)
                    {
                        ServiceLocator.Logger?.LogException(ex, $"Failed to update formula at {key}");
                    }
                }
            }
            catch (Exception ex)
            {
                ServiceLocator.Logger?.LogException(ex, "Balance formulas update process failed");
                throw;
            }
        }

        private static bool CanAccessWorksheet(string sheetName, out Excel.Worksheet ws)
        {
            ws = null;

            if (ServiceLocator.ExcelApp == null || !AppState.Instance.IsLoginCompleted)
                return false;

            Excel.Workbook wb = ServiceLocator.ExcelApp.ActiveWorkbook;
            if (wb?.Worksheets[sheetName] is not Excel.Worksheet worksheet)
            {
                ServiceLocator.Logger?.LogWarn("Worksheet is null. Aborting Get_GLSense_MultiFormulas.");
                if (wb != null)
                    ExcelComHelper.SafeRelease(wb, "Workbook");
                return false;
            }

            ws = worksheet;
            return true;
        }

        private static Excel.Range GetFormulaCells(Excel.Worksheet ws)
        {
            try
            {
                if (ws == null) return null;

                var rngFind = ws.Cells.SpecialCells(Excel.XlCellType.xlCellTypeFormulas);
                if (rngFind == null)
                {
                    ServiceLocator.Logger?.LogWarn("No formula cells found or error accessing them.");
                    return null;
                }
                return rngFind;
            }
            catch (Exception)
            {
                ServiceLocator.Logger?.LogWarn($"No formula cells found or error accessing them in worksheet {ws?.Name}.");
                return null;
            }
        }

        private static class MultiFormulaDetector
        {
            private static readonly Regex BalanceFormulaRegex =
                new(@"[=@]?\bGLSense_GetBalance\b\s*\(", RegexOptions.IgnoreCase | RegexOptions.Compiled);

            public static void ProcessFormulaAreas(Excel.Range rngFind)
            {
                try
                {
                    foreach (Excel.Range area in rngFind.Areas)
                    {
                        ProcessArea(area);
                    }
                }
                catch (Exception ex)
                {
                    ServiceLocator.Logger?.LogException(ex, "CommonMethods.MultiFormulaDetector.ProcessFormulaAreas: failed");
                }
                finally
                {
                    ExcelComHelper.SafeRelease(rngFind, "formulaCells");
                }
            }

            private static void ProcessArea(Excel.Range area)
            {
                if (area == null) return;

                if (area?.Cells?.Count == 1)
                {
                    ProcessSingleCell(area);
                }
                else
                {
                    if (area?.Cells?.Count > 1)
                        ProcessMultiCellArea(area);
                }
            }

            private static void ProcessSingleCell(Excel.Range area)
            {
                if (area.Formula is string formula && HasMultipleMatches(formula))
                {
                    RecordMultiFormulaCell(area);
                }
            }

            private static void ProcessMultiCellArea(Excel.Range area)
            {
                if (area?.Formula is not object[,] formulaData)
                    return;

                var (RowLower, RowUpper, ColLower, ColUpper) = GetArrayBounds(formulaData);
                for (int i = RowLower; i <= RowUpper; i++)
                {
                    for (int j = ColLower; j <= ColUpper; j++)
                    {
                        ProcessCellFormula(formulaData, i, j, area);
                    }
                }
            }

            private static void ProcessCellFormula(object[,] formulaData, int i, int j, Excel.Range area)
            {
                if (formulaData[i, j] is string formula && HasMultipleMatches(formula))
                {
                    SafeProcessCell(area, i, j);
                }
            }

            private static bool HasMultipleMatches(string formula) =>
                BalanceFormulaRegex.Matches(formula).Count >= 2;

            private static (int RowLower, int RowUpper, int ColLower, int ColUpper) GetArrayBounds(object[,] array)
            {
                return (
                    array.GetLowerBound(0),
                    array.GetUpperBound(0),
                    array.GetLowerBound(1),
                    array.GetUpperBound(1)
                );
            }

            private static void SafeProcessCell(Excel.Range area, int row, int col)
            {
                Excel.Range cell = null;
                try
                {
                    cell = area.Cells[row, col] as Excel.Range;
                    RecordMultiFormulaCell(cell);
                }
                finally
                {
                    if (cell != null)
                        ExcelComHelper.SafeRelease(cell, "cell");
                }
            }

            private static void RecordMultiFormulaCell(Excel.Range cell)
            {
                if (cell == null) return;

                try
                {
                    var addr = cell.get_Address(RowAbsolute: true, ColumnAbsolute: true);
                    ServiceLocator.Logger?.LogDebug($"Range {cell.Worksheet.Name}!{addr} has multiple balance formulas.");
                }
                catch (Exception ex)
                {
                    ServiceLocator.Logger?.LogException(ex, "CommonMethods.MultiFormulaDetector.RecordMultiFormulaCell: failed");
                }
            }
        }

        /// <summary>
        /// Logs a warning for every cell on <paramref name="sheetName"/> whose formula
        /// contains more than one GLSense_GetBalance(...) call (those cells can't be
        /// reliably tracer-arrow'd / drilled into as a single balance reference).
        /// </summary>
        public static void Get_GLSense_MultiFormulas(string sheetName)
        {
            ServiceLocator.Logger?.LogDebug($"Checking for multiple balance formulas in sheet: {sheetName}");

            if (!CanAccessWorksheet(sheetName, out Excel.Worksheet ws))
            {
                return;
            }

            var formulaCells = GetFormulaCells(ws);
            if (formulaCells == null)
            {
                return;
            }

            ServiceLocator.Logger?.LogDebug($"Found formula cells, processing areas");
            MultiFormulaDetector.ProcessFormulaAreas(formulaCells);

            ServiceLocator.Logger?.LogDebug("Checking range for multiple balance formulas completed.");
        }

        public static void DisableExcelSettings()
        {
            try
            {
                ServiceLocator.Logger.LogDebug("Disabling Excel settings for batch operation");
                var app = ServiceLocator.ExcelApp ?? throw new InvalidOperationException("Excel unavailable");

                app.ScreenUpdating = false;
                app.DisplayAlerts = false;
                app.EnableEvents = false;

                ServiceLocator.Logger.LogDebug("Excel settings disabled successfully");
            }
            catch (Exception ex)
            {
                ServiceLocator.Logger.LogException(ex, "Failed to disable Excel settings");
                throw;
            }
        }

        public static void EnableExcelSettings()
        {
            try
            {
                ServiceLocator.Logger.LogDebug("Re-enabling Excel settings");
                var app = ServiceLocator.ExcelApp ?? throw new InvalidOperationException("Excel unavailable");

                app.ScreenUpdating = true;
                app.DisplayAlerts = true;
                app.EnableEvents = true;

                ServiceLocator.Logger.LogDebug("Excel settings enabled successfully");
            }
            catch (Exception ex)
            {
                ServiceLocator.Logger.LogException(ex, "Failed to enable Excel settings");
                throw;
            }
        }

        // Non-throwing counterparts to Disable/EnableExcelSettings, for call sites that
        // have no caller able to observe the exception (fire-and-forget drilldown/ribbon
        // entry points - `_ = SomeAsyncMethod()` - or a finally block), where an escaping
        // exception either becomes an unobserved Task exception or, worse, an unhandled
        // exception in this AppDomain - exactly what FinalWorkingCode's identical
        // unguarded call sites did in practice (COMException 0x800A03EC restoring
        // DisplayAlerts after a drilldown hyperlink click crashed the whole add-in there;
        // see FinalWorkingCode's CLAUDE.md). Logs and swallows instead of throwing.
        public static bool TryDisableExcelSettings(string context = null)
        {
            try
            {
                DisableExcelSettings();
                return true;
            }
            catch (Exception ex)
            {
                ServiceLocator.Logger?.LogException(ex, string.IsNullOrEmpty(context)
                    ? "Failed to disable Excel settings"
                    : $"Failed to disable Excel settings before {context}");
                return false;
            }
        }

        public static void TryEnableExcelSettings(string context = null)
        {
            try
            {
                EnableExcelSettings();
            }
            catch (Exception ex)
            {
                ServiceLocator.Logger?.LogException(ex, string.IsNullOrEmpty(context)
                    ? "Failed to restore Excel settings"
                    : $"Failed to restore Excel settings after {context}");
            }
        }
    }
}
