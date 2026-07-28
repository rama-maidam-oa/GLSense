using GLSense.Utilities;
using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;

namespace GLSense.Helpers
{
    /// <summary>
    /// Helper for common Excel operations to reduce code duplication
    /// </summary>
    public static class ExcelOperationsHelper
    {
        /// <summary>
        /// Safely gets a worksheet by name
        /// </summary>
        public static Worksheet GetWorksheet(Workbook workbook, string sheetName)
        {
            using (new LogUtility.LogScope($"GetWorksheet({sheetName})"))
            {
                try
                {
                    if (workbook == null)
                    {
                        LogUtility.LogWarn("Workbook is null");
                        return null;
                    }

                    LogUtility.LogDebug($"Retrieving worksheet: {sheetName}");

                    var worksheet = workbook.Worksheets[sheetName] as Worksheet;
                    
                    if (worksheet == null)
                    {
                        LogUtility.LogWarn($"Worksheet not found: {sheetName}");
                    }
                    else
                    {
                        LogUtility.LogDebug($"Worksheet retrieved: {sheetName}");
                    }

                    return worksheet;
                }
                catch (Exception ex)
                {
                    ExceptionHelper.LogDetailedException(ex, $"GetWorksheet: {sheetName}");
                    return null;
                }
            }
        }

        /// <summary>
        /// Checks if a worksheet exists
        /// </summary>
        public static bool WorksheetExists(Workbook workbook, string sheetName)
        {
            using (new LogUtility.LogScope($"WorksheetExists({sheetName})"))
            {
                try
                {
                    if (workbook == null)
                    {
                        LogUtility.LogWarn("Workbook is null");
                        return false;
                    }

                    foreach (Worksheet ws in workbook.Worksheets)
                    {
                        if (string.Equals(ws.Name, sheetName, StringComparison.OrdinalIgnoreCase))
                        {
                            LogUtility.LogDebug($"Worksheet exists: {sheetName}");
                            ExcelComHelper.SafeRelease(ws, $"Worksheet {sheetName}");
                            return true;
                        }
                        ExcelComHelper.SafeRelease(ws, "Worksheet");
                    }

                    LogUtility.LogDebug($"Worksheet does not exist: {sheetName}");
                    return false;
                }
                catch (Exception ex)
                {
                    ExceptionHelper.LogDetailedException(ex, $"WorksheetExists: {sheetName}");
                    return false;
                }
            }
        }

        /// <summary>
        /// Creates or gets existing worksheet
        /// </summary>
        public static Worksheet CreateOrGetWorksheet(Application excelApp, string sheetName)
        {
            using (new LogUtility.LogScope($"CreateOrGetWorksheet({sheetName})"))
            {
                try
                {
                    if (excelApp?.ActiveWorkbook == null)
                    {
                        LogUtility.LogError("No active workbook");
                        return null;
                    }

                    var workbook = excelApp.ActiveWorkbook;

                    if (WorksheetExists(workbook, sheetName))
                    {
                        return GetWorksheet(workbook, sheetName);
                    }

                    LogUtility.LogDebug($"Creating new worksheet: {sheetName}");
                    
                    var newSheet = workbook.Worksheets.Add() as Worksheet;
                    if (newSheet != null)
                    {
                        newSheet.Name = sheetName;
                    }

                    return newSheet;
                }
                catch (Exception ex)
                {
                    ExceptionHelper.LogDetailedException(ex, $"CreateOrGetWorksheet: {sheetName}");
                    return null;
                }
            }
        }

        /// <summary>
        /// Safely deletes a worksheet
        /// </summary>
        public static bool DeleteWorksheet(Workbook workbook, string sheetName)
        {
            using (new LogUtility.LogScope($"DeleteWorksheet({sheetName})"))
            {
                try
                {
                    if (workbook == null)
                    {
                        LogUtility.LogWarn("Workbook is null");
                        return false;
                    }

                    var worksheet = GetWorksheet(workbook, sheetName);
                    if (worksheet == null)
                    {
                        LogUtility.LogDebug($"Worksheet not found: {sheetName}");
                        return false;
                    }

                    LogUtility.LogDebug($"Deleting worksheet: {sheetName}");
                    worksheet.Delete();
                    
                    ExcelComHelper.SafeRelease(worksheet, $"Worksheet {sheetName}");
                   
                    return true;
                }
                catch (Exception ex)
                {
                    ExceptionHelper.LogDetailedException(ex, $"DeleteWorksheet: {sheetName}");
                    return false;
                }
            }
        }

        /// <summary>
        /// Writes data to a range efficiently
        /// </summary>
        public static void WriteArrayToRange(Range startCell, object[,] data, string description = "")
        {
            using (new LogUtility.LogScope($"WriteArrayToRange: {description}"))
            {
                try
                {
                    if (startCell == null || data == null)
                    {
                        LogUtility.LogWarn("StartCell or data is null");
                        return;
                    }

                    int rows = data.GetLength(0);
                    int cols = data.GetLength(1);

                    LogUtility.LogDebug($"Writing array - Rows: {rows}, Columns: {cols}");

                    var targetRange = startCell.Resize[rows, cols];
                    targetRange.Value2 = data;

                    LogUtility.LogDebug($"Array written successfully{(string.IsNullOrEmpty(description) ? "" : $": {description}")}");
                }
                catch (Exception ex)
                {
                    ExceptionHelper.LogDetailedException(ex, $"WriteArrayToRange: {description}");
                }
            }
        }

        /// <summary>
        /// Reads data from a range efficiently
        /// </summary>
        public static object[,] ReadRangeValues(Range range, string description = "")
        {
            using (new LogUtility.LogScope($"ReadRangeValues: {description}"))
            {
                try
                {
                    if (range == null)
                    {
                        LogUtility.LogWarn("Range is null");
                        return null;
                    }

                    LogUtility.LogDebug($"Reading range values - Rows: {range.Rows.Count}, Columns: {range.Columns.Count}");

                    var values = range.Value2 as object[,];

                    if (values != null)
                    {
                        LogUtility.LogDebug($"Range values read successfully{(string.IsNullOrEmpty(description) ? "" : $": {description}")}");
                    }

                    return values;
                }
                catch (Exception ex)
                {
                    ExceptionHelper.LogDetailedException(ex, $"ReadRangeValues: {description}");
                    return null;
                }
            }
        }

        /// <summary>
        /// Gets used range efficiently
        /// </summary>
        public static Range GetUsedRange(Worksheet worksheet)
        {
            using (new LogUtility.LogScope("GetUsedRange"))
            {
                try
                {
                    if (worksheet == null)
                    {
                        LogUtility.LogWarn("Worksheet is null");
                        return null;
                    }

                    LogUtility.LogDebug($"Getting used range for worksheet: {worksheet.Name}");

                    var usedRange = worksheet.UsedRange;

                    if (usedRange != null)
                    {
                        LogUtility.LogDebug($"Used range: {usedRange.Rows.Count} rows x {usedRange.Columns.Count} columns");
                    }

                    return usedRange;
                }
                catch (Exception ex)
                {
                    ExceptionHelper.LogDetailedException(ex, "GetUsedRange");
                    return null;
                }
            }
        }

        /// <summary>
        /// Formats a range with number format
        /// </summary>
        public static void ApplyNumberFormat(Range range, string format, string description = "")
        {
            using (new LogUtility.LogScope($"ApplyNumberFormat: {description}"))
            {
                try
                {
                    if (range == null)
                    {
                        LogUtility.LogWarn("Range is null");
                        return;
                    }

                    LogUtility.LogDebug($"Applying format '{format}' to range{(string.IsNullOrEmpty(description) ? "" : $": {description}")}");

                    range.NumberFormat = format;

                    LogUtility.LogDebug("Number format applied successfully");
                }
                catch (Exception ex)
                {
                    ExceptionHelper.LogDetailedException(ex, $"ApplyNumberFormat: {description}");
                }
            }
        }

        /// <summary>
        /// Copies range values (not formulas) efficiently
        /// </summary>
        public static void CopyRangeValues(Range source, Range destination)
        {
            using (new LogUtility.LogScope("CopyRangeValues"))
            {
                try
                {
                    if (source == null || destination == null)
                    {
                        LogUtility.LogWarn("Source or destination range is null");
                        return;
                    }

                    LogUtility.LogDebug($"Copying values from {ExcelComHelper.GetRangeAddress(source, false, false)} to {ExcelComHelper.GetRangeAddress(destination, false, false)}");

                    source.Copy();
                    destination.PasteSpecial(XlPasteType.xlPasteValues);

                    LogUtility.LogDebug("Range values copied successfully");
                }
                catch (Exception ex)
                {
                    ExceptionHelper.LogDetailedException(ex, "CopyRangeValues");
                }
            }
        }

        /// <summary>
        /// Finds cells containing specific text
        /// </summary>
        public static List<Range> FindCellsWithText(Worksheet worksheet, string searchText, bool caseSensitive = false)
        {
            using (new LogUtility.LogScope($"FindCellsWithText(text={searchText}, caseSensitive={caseSensitive})"))
            {
                var results = new List<Range>();

                try
                {
                    if (worksheet == null || string.IsNullOrEmpty(searchText))
                    {
                        LogUtility.LogWarn("Worksheet is null or search text is empty");
                        return results;
                    }

                    LogUtility.LogDebug($"Searching for '{searchText}' in worksheet: {worksheet.Name}");

                    var usedRange = worksheet.UsedRange;
                    if (usedRange == null)
                    {
                        LogUtility.LogDebug("No used range in worksheet");
                        return results;
                    }

                    var foundCell = usedRange.Find(
                        What: searchText,
                        LookIn: XlFindLookIn.xlValues,
                        LookAt: XlLookAt.xlPart,
                        SearchOrder: XlSearchOrder.xlByRows,
                        MatchCase: caseSensitive);

                    if (foundCell == null)
                    {
                        LogUtility.LogDebug($"No cells found containing: {searchText}");
                        return results;
                    }

                    var firstAddress = foundCell.Address;
                    results.Add(foundCell);

                    while (true)
                    {
                        var nextCell = usedRange.FindNext(foundCell);
                        if (nextCell == null || nextCell.Address == firstAddress)
                        {
                            break;
                        }

                        results.Add(nextCell);
                        foundCell = nextCell;
                    }

                    LogUtility.LogDebug($"Found {results.Count} cells containing: {searchText}");
                    return results;
                }
                catch (Exception ex)
                {
                    ExceptionHelper.LogDetailedException(ex, $"FindCellsWithText: {searchText}");
                    return results;
                }
            }
        }

        /// <summary>
        /// Sets column widths efficiently
        /// </summary>
        public static void AutoFitColumns(Range range, double? maxWidth = null)
        {
            using (new LogUtility.LogScope("AutoFitColumns"))
            {
                try
                {
                    if (range == null)
                    {
                        LogUtility.LogWarn("Range is null");
                        return;
                    }

                    LogUtility.LogDebug("Auto-fitting columns");

                    range.Columns.AutoFit();

                    if (maxWidth.HasValue)
                    {
                        foreach (Range column in range.Columns)
                        {
                            if ((double)column.ColumnWidth > maxWidth.Value)
                            {
                                column.ColumnWidth = maxWidth.Value;
                            }
                            ExcelComHelper.SafeRelease(column, "Column");
                        }
                    }

                    LogUtility.LogDebug("Columns auto-fitted successfully");
                }
                catch (Exception ex)
                {
                    ExceptionHelper.LogDetailedException(ex, "AutoFitColumns");
                }
            }
        }

        /// <summary>
        /// Applies filters to a range
        /// </summary>
        public static void ApplyAutoFilter(Range range, int? columnIndex = null, string criteria = null)
        {
            using (new LogUtility.LogScope("ApplyAutoFilter"))
            {
                try
                {
                    if (range == null)
                    {
                        LogUtility.LogWarn("Range is null");
                        return;
                    }

                    LogUtility.LogDebug($"Applying auto filter - Column: {columnIndex}, Criteria: {criteria}");

                    if (columnIndex.HasValue && !string.IsNullOrEmpty(criteria))
                    {
                        range.AutoFilter(columnIndex.Value, criteria);
                    }
                    else
                    {
                        range.AutoFilter();
                    }

                    LogUtility.LogDebug("Auto filter applied successfully");
                }
                catch (Exception ex)
                {
                    ExceptionHelper.LogDetailedException(ex, "ApplyAutoFilter");
                }
            }
        }

        /// <summary>
        /// Freezes panes at specified location
        /// </summary>
        public static void FreezePanes(Worksheet worksheet, int row, int column)
        {
            using (new LogUtility.LogScope($"FreezePanes(row={row}, col={column})"))
            {
                try
                {
                    if (worksheet == null)
                    {
                        LogUtility.LogWarn("Worksheet is null");
                        return;
                    }

                    LogUtility.LogDebug($"Freezing panes at row {row}, column {column}");

                    var cell = worksheet.Cells[row, column] as Range;
                    if (cell != null)
                    {
                        cell.Select();
                        worksheet.Application.ActiveWindow.FreezePanes = true;
                        
                        ExcelComHelper.SafeRelease(cell, "Freeze cell");
                        LogUtility.LogDebug("Panes frozen successfully");
                    }
                }
                catch (Exception ex)
                {
                    ExceptionHelper.LogDetailedException(ex, "FreezePanes");
                }
            }
        }

        /// <summary>
        /// Applies formatting to header row
        /// </summary>
        public static void FormatHeaderRow(Range headerRange, bool bold = true, string backgroundColor = null)
        {
            using (new LogUtility.LogScope("FormatHeaderRow"))
            {
                try
                {
                    if (headerRange == null)
                    {
                        LogUtility.LogWarn("Header range is null");
                        return;
                    }

                    LogUtility.LogDebug("Formatting header row");

                    if (bold)
                    {
                        headerRange.Font.Bold = true;
                    }

                    if (!string.IsNullOrEmpty(backgroundColor))
                    {
                        headerRange.Interior.Color = System.Drawing.ColorTranslator.FromHtml(backgroundColor);
                    }

                    headerRange.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                    headerRange.VerticalAlignment = XlVAlign.xlVAlignCenter;

                    LogUtility.LogDebug("Header row formatted successfully");
                }
                catch (Exception ex)
                {
                    ExceptionHelper.LogDetailedException(ex, "FormatHeaderRow");
                }
            }
        }

        /// <summary>
        /// Converts column letter to index (A=1, B=2, etc.)
        /// </summary>
        public static int ColumnLetterToIndex(string columnLetter)
        {
            if (string.IsNullOrEmpty(columnLetter))
            {
                LogUtility.LogWarn("Column letter is null or empty");
                return -1;
            }

            try
            {
                int index = 0;
                columnLetter = columnLetter.ToUpperInvariant();

                foreach (char c in columnLetter)
                {
                    index = index * 26 + (c - 'A' + 1);
                }

                LogUtility.LogDebug($"Column '{columnLetter}' converted to index: {index}");
                return index;
            }
            catch (Exception ex)
            {
                ExceptionHelper.LogDetailedException(ex, $"ColumnLetterToIndex: {columnLetter}");
                return -1;
            }
        }

        /// <summary>
        /// Converts column index to letter (1=A, 2=B, etc.)
        /// </summary>
        public static string ColumnIndexToLetter(int columnIndex)
        {
            using (new LogUtility.LogScope($"ColumnIndexToLetter({columnIndex})"))
            {
                try
                {
                    if (columnIndex <= 0)
                    {
                        LogUtility.LogWarn($"Invalid column index: {columnIndex}");
                        return string.Empty;
                    }

                    string columnLetter = string.Empty;
                    int temp = columnIndex;

                    while (temp > 0)
                    {
                        temp--;
                        columnLetter = (char)('A' + (temp % 26)) + columnLetter;
                        temp /= 26;
                    }

                    LogUtility.LogDebug($"Column index {columnIndex} converted to: {columnLetter}");
                    return columnLetter;
                }
                catch (Exception ex)
                {
                    ExceptionHelper.LogDetailedException(ex, $"ColumnIndexToLetter: {columnIndex}");
                    return string.Empty;
                }
            }
        }
    }
}
