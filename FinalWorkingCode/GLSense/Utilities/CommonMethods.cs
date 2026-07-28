using AddinExpress.MSO;
using GLSense.Helpers;
using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using System.Threading;
using Excel = Microsoft.Office.Interop.Excel;

namespace GLSense.Utilities
{
#nullable enable
    public static class CommonMethods
    {
        public static void Clear_Sheet()
        {
            using (new LogUtility.LogScope("Clear_Sheet"))
            {
                try
                {
                    LogUtility.LogDebug("Clearing active worksheet");
                    Excel.Worksheet? wrksht = (Excel.Worksheet)AddinModule.CurrentInstance.ExcelApp.ActiveSheet;
                    if (wrksht != null)
                    {
                        LogUtility.LogDebug($"Clearing sheet: {wrksht.Name}");
                        wrksht.Cells.Clear();
                        LogUtility.LogDebug("Sheet cleared successfully");
                    }
                    else
                    {
                        LogUtility.LogWarn("Active worksheet is null");
                    }
                }
                catch (Exception ex)
                {
                    LogUtility.LogException(ex, "Failed to clear sheet");
                    throw;
                }
            }
        }

        public static void EnsureExcelApp()
        {
            using (new LogUtility.LogScope("EnsureExcelApp"))
            {
                try
                {
                    LogUtility.LogDebug("Checking Excel application instance");
                    if (AppState.Instance.ExcelApp == null || !IsExcelAppAlive(AppState.Instance.ExcelApp))
                    {
                        LogUtility.LogDebug("Excel app is null or not alive, reinitializing");
                        AppState.Instance.ExcelApp = (Excel.Application)((AddinModule)ADXAddinModule.CurrentInstance).HostApplication;
                        LogUtility.LogDebug("Excel app reinitialized successfully");
                    }
                    else
                    {
                        LogUtility.LogDebug("Excel app is alive and ready");
                    }
                }
                catch (Exception ex)
                {
                    // Do NOT rethrow: this is a best-effort "make sure the Excel reference
                    // is fresh" check called from places like a UserControl's Loaded event
                    // (e.g. ExcelRefEditControl -> ExcelRefManager.SetupControl). Historically
                    // this swallowed reinit failures so a transient COM/RCW hiccup here
                    // wouldn't crash the calling window; callers already null-check
                    // AppState.Instance.ExcelApp afterward. Rethrowing here caused
                    // "Object reference not set to an instance of an object" to surface on
                    // every window hosting ExcelRefEditControl whenever this reinit failed.
                    LogUtility.LogException(ex, "Failed to ensure Excel application (non-fatal, continuing without rethrow)");
                }
            }
        }
        private static bool IsExcelAppAlive(Excel._Application app)
        {
            try
            {
                _ = app.Hwnd; // Accessing property forces COM check
                return true;
            }
            catch (Exception ex)
            {
                LogUtility.LogWarn($"IsExcelAppAlive: COM access check failed (treating as not alive): {ex.Message}");
                return false;
            }
        }

        public static void DisableExcelSettings()
        {
            using (new LogUtility.LogScope("DisableExcelSettings"))
            {
                try
                {
                    LogUtility.LogDebug("Disabling Excel settings for batch operation");
                    var app = AppState.Instance.ExcelApp ?? throw new InvalidOperationException("Excel unavailable");
                    
                    LogUtility.LogDebug("Setting ScreenUpdating = false");
                    app.ScreenUpdating = false;
                    
                    LogUtility.LogDebug("Setting DisplayAlerts = false");
                    app.DisplayAlerts = false;
                    
                    LogUtility.LogDebug("Setting EnableEvents = false");
                    app.EnableEvents = false;
                                        
                    LogUtility.LogDebug("Excel settings disabled successfully");
                }
                catch (Exception ex)
                {
                    LogUtility.LogException(ex, "Failed to disable Excel settings");
                    throw;
                }
            }
        }

        public static void EnableExcelSettings()
        {
            using (new LogUtility.LogScope("EnableExcelSettings"))
            {
                try
                {
                    LogUtility.LogDebug("Re-enabling Excel settings");
                    var app = AppState.Instance.ExcelApp ?? throw new InvalidOperationException("Excel unavailable");
                    
                    LogUtility.LogDebug("Setting ScreenUpdating = true");
                    app.ScreenUpdating = true;
                    
                    LogUtility.LogDebug("Setting DisplayAlerts = true");
                    app.DisplayAlerts = true;
                    
                    LogUtility.LogDebug("Setting EnableEvents = true");
                    app.EnableEvents = true;
                    
                    
                    LogUtility.LogDebug("Excel settings enabled successfully");
                }
                catch (Exception ex)
                {
                    LogUtility.LogException(ex, "Failed to enable Excel settings");
                    throw;
                }
            }
        }
        
        public static void BalanceFormulas_Updation(Dictionary<string, string> BalancesDict)
        {
            using (new LogUtility.LogScope("BalanceFormulas_Updation"))
            {
                try
                {
                    if (BalancesDict == null || BalancesDict.Keys.Count == 0)
                    {
                        LogUtility.LogDebug("No balance formulas to update");
                        return;
                    }

                    LogUtility.LogDebug($"Updating {BalancesDict.Keys.Count} balance formulas");

                    int successCount = 0;
                    int failureCount = 0;

                    foreach (var key in BalancesDict.Keys)
                    {
                        try
                        {
                            LogUtility.LogDebug($"Processing formula at {key}");
                            
                            var parser = new ClsFormulaParser(BalancesDict[key]);
                            int ArgsCount = parser.Formula_ArgsCount();
                            int expectedArgs = ArgsCount + 2;

                            string correctedFormula = parser.Formula_Correction(expectedArgs, 8);
                            Excel.Range rng = AppState.Instance.ExcelApp.Range[key];

                            if (rng != null)
                            {
                                LogUtility.LogDebug($"Old Formula: {BalancesDict[key]}");
                                LogUtility.LogDebug($"New Formula: {correctedFormula}");
                                rng.Value2 = correctedFormula;
                                successCount++;
                                LogUtility.LogDebug($"Formula updated successfully at {key}");
                            }
                            else
                            {
                                LogUtility.LogWarn($"Range is null at {key}");
                                failureCount++;
                            }
                        }
                        catch (Exception ex)
                        {
                            failureCount++;
                            LogUtility.LogException(ex, $"Failed to update formula at {key}");
                        }
                    }

                }
                catch (Exception ex)
                {
                    LogUtility.LogException(ex, "Balance formulas update process failed");
                    throw;
                }
            }
        }
        private static bool CanAccessWorksheet(string sheetName, out Excel.Worksheet? ws)
        {
            ws = null;

            if (AppState.Instance.ExcelApp == null || !AppState.Instance.IsLoginCompleted)
                return false;

            Excel.Workbook? wb = AppState.Instance.ExcelApp.ActiveWorkbook;
            if (wb?.Worksheets[sheetName] is not Excel.Worksheet worksheet)
            {
                LogUtility.LogWarn("Worksheet is null. Aborting Get_GLSense_MultiFormulas.");
                if (wb != null)
                    SafeFinalReleaseCom(wb);
                return false;
            }

            ws = worksheet;
            return true;
        }
        private static Excel.Range? GetFormulaCells(Excel.Worksheet? ws)
        {
            try
            {
                if (ws == null) return null;

                var rngFind = ws.Cells.SpecialCells(Excel.XlCellType.xlCellTypeFormulas);
                if (rngFind == null)
                {
                    LogUtility.LogWarn("No formula cells found or error accessing them.");
                    return null;
                }
                return rngFind;
            }
            catch (Exception)
            {
                LogUtility.LogWarn($"No formula cells found or error accessing them in worksheet {ws?.Name}.");
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
                    LogUtility.LogException(ex, "CommonMethods.MultiFormulaDetector.ProcessFormulaAreas");
                }
                finally
                {
                    SafeFinalReleaseCom(rngFind);
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
                Excel.Range? cell = null;
                try
                {
                    cell = area.Cells[row, col] as Excel.Range;
                    RecordMultiFormulaCell(cell);
                }
                finally
                {
                    if (cell != null)
                        SafeFinalReleaseCom(cell);
                }
            }

            private static void RecordMultiFormulaCell(Excel.Range? cell)
            {
                if (cell == null) return;

                try
                {
                    var addr = cell.get_Address(RowAbsolute: true, ColumnAbsolute: true);
                    LogUtility.LogDebug($"Range {cell.Worksheet.Name}!{addr} has multiple balance formulas.");
                }
                catch (Exception ex)
                {
                    LogUtility.LogException(ex, "CommonMethods.MultiFormulaDetector.RecordMultiFormulaCell");
                }
            }
        }
        public static void Get_GLSense_MultiFormulas(string sheetName)
        {
            using (new LogUtility.LogScope("Get_GLSense_MultiFormulas"))
            {
                LogUtility.LogDebug($"Checking for multiple balance formulas in sheet: {sheetName}");
                
                if (!CanAccessWorksheet(sheetName, out Excel.Worksheet? ws))
                {
                    return;
                }

                var formulaCells = GetFormulaCells(ws);
                if (formulaCells == null)
                {
                    return;
                }

                LogUtility.LogDebug($"Found formula cells, processing areas");
                MultiFormulaDetector.ProcessFormulaAreas(formulaCells);

                LogUtility.LogDebug("Checking range for multiple balance formulas completed.");
            }
        }

        // Optional helper if not already in your project:
        private static void SafeFinalReleaseCom(object comObj)
        {
            if (comObj == null) return;
            try
            {
                System.Runtime.InteropServices.Marshal.FinalReleaseComObject(comObj);
            }
            catch (Exception ex)
            {
                LogUtility.LogWarn($"SafeFinalReleaseCom: failed to release COM object (non-fatal): {ex.Message}");
            }
        }
    }
#nullable disable
}
