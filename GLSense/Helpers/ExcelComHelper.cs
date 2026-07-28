using GLSense.Utilities;
using Microsoft.Office.Interop.Excel;
using System;
using System.Runtime.InteropServices;

namespace GLSense.Helpers
{
    /// <summary>
    /// Helper for safe Excel COM object manipulation and disposal
    /// </summary>
    public static class ExcelComHelper
    {
        /// <summary>
        /// Safely releases a COM object with logging
        /// </summary>
        public static void SafeRelease(object comObject, string objectDescription = "COM object")
        {
            if (comObject == null) return;

            try
            {
                if (Marshal.IsComObject(comObject))
                {
                    LogUtility.LogDebug($"Releasing {objectDescription}");
                    Marshal.FinalReleaseComObject(comObject);
                }
            }
            catch (Exception ex)
            {
                ExceptionHelper.LogDetailedException(ex, $"Failed to release {objectDescription}");
            }
        }

        /// <summary>
        /// Safely gets the address of a range with error handling
        /// </summary>
        public static string GetRangeAddress(
            Range range,
            bool rowAbsolute = true,
            bool columnAbsolute = true,
            XlReferenceStyle referenceStyle = XlReferenceStyle.xlA1,
            bool external = false)
        {
            try
            {
                if (range == null) return string.Empty;

                return range.Address[rowAbsolute, columnAbsolute, referenceStyle, external];
            }
            catch (Exception ex)
            {
                ExceptionHelper.LogDetailedException(ex, "GetRangeAddress");
                return string.Empty;
            }
        }

        /// <summary>
        /// Safely gets the worksheet name from a range
        /// </summary>
        public static string GetWorksheetName(Range range)
        {
            try
            {
                if (range == null) return string.Empty;

                var worksheet = range.Worksheet;
                if (worksheet == null) return string.Empty;

                string name = worksheet.Name;
                LogUtility.LogDebug($"Retrieved worksheet name: {name}");

                return name;
            }
            catch (Exception ex)
            {
                ExceptionHelper.LogDetailedException(ex, "GetWorksheetName");
                return string.Empty;
            }
        }

        /// <summary>
        /// Safely gets the formula from a range
        /// </summary>
        public static string GetFormula(Range range)
        {
            try
            {
                if (range == null) return string.Empty;

                object formula = range.Formula;
                return formula?.ToString() ?? string.Empty;
            }
            catch (Exception ex)
            {
                ExceptionHelper.LogDetailedException(ex, "GetFormula");
                return string.Empty;
            }
        }

        /// <summary>
        /// Safely checks if a range has a formula
        /// </summary>
        public static bool HasFormula(Range range)
        {
            try
            {
                if (range == null) return false;

                return range.HasFormula is bool hasFormula && hasFormula;
            }
            catch (Exception ex)
            {
                ExceptionHelper.LogDetailedException(ex, "HasFormula");
                return false;
            }
        }

        /// <summary>
        /// Safely gets cell value with logging
        /// </summary>
        public static object GetCellValue(Range range)
        {
            try
            {
                if (range == null) return null;

                object value = range.Value2;
                LogUtility.LogDebug($"Cell value retrieved: {value ?? "(null)"}");

                return value;
            }
            catch (Exception ex)
            {
                ExceptionHelper.LogDetailedException(ex, "GetCellValue");
                return null;
            }
        }

        /// <summary>
        /// Safely sets cell value with logging
        /// </summary>
        public static void SetCellValue(Range range, object value, string description = "")
        {
            try
            {
                if (range == null)
                {
                    LogUtility.LogWarn($"Cannot set value - range is null: {description}");
                    return;
                }

                LogUtility.LogDebug($"Setting cell value{(string.IsNullOrEmpty(description) ? "" : $" ({description})")}: {value ?? "(null)"}");

                range.Value2 = value;
            }
            catch (Exception ex)
            {
                ExceptionHelper.LogDetailedException(ex, $"SetCellValue{(string.IsNullOrEmpty(description) ? "" : $": {description}")}");
            }
        }

        /// <summary>
        /// Safely clears a range with logging
        /// </summary>
        public static void ClearRange(Range range, string description = "")
        {
            try
            {
                if (range == null)
                {
                    LogUtility.LogWarn($"Cannot clear - range is null: {description}");
                    return;
                }

                LogUtility.LogDebug($"Clearing range{(string.IsNullOrEmpty(description) ? "" : $": {description}")}");

                range.Clear();
            }
            catch (Exception ex)
            {
                ExceptionHelper.LogDetailedException(ex, $"ClearRange{(string.IsNullOrEmpty(description) ? "" : $": {description}")}");
            }
        }

        /// <summary>
        /// Safely sets number format with logging
        /// </summary>
        public static void SetNumberFormat(Range range, string format, string description = "")
        {
            try
            {
                if (range == null)
                {
                    LogUtility.LogWarn($"Cannot set format - range is null: {description}");
                    return;
                }

                LogUtility.LogDebug($"Setting number format{(string.IsNullOrEmpty(description) ? "" : $" ({description})")}: {format}");

                range.NumberFormat = format;
            }
            catch (Exception ex)
            {
                ExceptionHelper.LogDetailedException(ex, $"SetNumberFormat{(string.IsNullOrEmpty(description) ? "" : $": {description}")}");
            }
        }

        /// <summary>
        /// Checks if Excel application is alive and accessible
        /// </summary>
        public static bool IsExcelAppAlive(Application excelApp)
        {
            try
            {
                if (excelApp == null) return false;

                _ = excelApp.Hwnd;
                return true;
            }
            catch (Exception ex)
            {
                LogUtility.LogDebug($"Excel app check failed: {ex.Message}");
                return false;
            }
        }
    }
}
