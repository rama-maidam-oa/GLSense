// ExcelExternalRef.cs in GLSense.Addin.Core
// Ported from GLSense\Helpers\ExcelExternalRef.cs (FinalWorkingCode).
// Only change: resolves the Excel Application via ServiceLocator.ExcelApp instead of
// the old AppState.Instance.ExcelApp static, to match this project's service-locator
// pattern for accessing the host context.
using GLSense.Addin.Core.Infrastructure;
using GLSense.Addin.Core.Models;
using System;
using Excel = Microsoft.Office.Interop.Excel;

namespace GLSense.Addin.Core.Helpers
{
    public static class ExcelExternalRef
    {
        /// <summary>
        /// Builds a fully-qualified external address from a Range, e.g. "[Book1.xlsx]Sheet 1!$A$1:$B$10".
        /// </summary>
        public static string BuildExternalAddress(Excel.Range rng)
        {
            if (rng == null)
            {
                ServiceLocator.Logger?.LogWarn("ExcelExternalRef.BuildExternalAddress: called with a null Range.");
                throw new ArgumentNullException(nameof(rng));
            }
            // Returns: [WorkbookName]SheetName!$A$1:$B$10 (sheet name quoted automatically if needed)

            string address = rng.Address[
                RowAbsolute: true,
                ColumnAbsolute: true,
                ReferenceStyle: Excel.XlReferenceStyle.xlA1,
                External: true
            ];

            ServiceLocator.Logger?.LogDebug($"ExcelExternalRef.BuildExternalAddress: built '{address}'");
            return address;
        }

        /// <summary>
        /// Parsed representation of an external address string.
        /// </summary>
        public sealed class ExternalParts
        {
            public string Workbook { get; set; }   // e.g., "Book1.xlsx"
            public string Worksheet { get; set; }  // e.g., "Sheet 1" (unquoted)
            public string Address { get; set; }    // e.g., "$A$1:$B$10"
            public string FullPrefix { get; set; } // e.g., "C:\\Path\\[Book1.xlsx]Sheet 1"
        }

        /// <summary>
        /// Parses strings like "[Book1.xlsx]Sheet1!$A$1:$B$10" or "'[Book1.xlsx]Sheet 1'!$A$1".
        /// Handles quoted sheet names and doubled apostrophes.
        /// </summary>
        public static ExternalParts ParseExternalAddress(string external)
        {
            if (string.IsNullOrEmpty(external))
            {
                ServiceLocator.Logger?.LogWarn("ExcelExternalRef.ParseExternalAddress: called with a null/empty address.");
                throw new ArgumentException("External address is empty.", nameof(external));
            }

            // 1) Split at the first '!' to separate the left token and the address.
            int exclPos = external.IndexOf('!');
            if (exclPos < 0)
            {
                ServiceLocator.Logger?.LogWarn($"ExcelExternalRef.ParseExternalAddress: not a valid external address (missing '!'): {external}");
                throw new FormatException("Not a valid external address (missing '!'): " + external);
            }

            string left = external.Substring(0, exclPos);           // e.g. [Book1]Sheet1 OR '[Book1]Sheet 1' OR 'C:\Path\[WB.xlsx]Sheet 1'
            string addr = external.Substring(exclPos + 1);          // e.g. $J$4 or $A$1:$B$10

            // 2) If quoted, strip outer quotes and unescape doubled apostrophes.
            if (left.Length >= 2 && left[0] == '\'' && left[left.Length - 1] == '\'')
            {
                left = left.Substring(1, left.Length - 2).Replace("''", "'");
            }

            // 3) Within the left token, locate the workbook brackets: "...[Workbook]Sheet..."
            // There may be a path prefix before the bracket: "C:\Path\[Workbook.xlsx]Sheet Name"
            int bracketOpen = left.LastIndexOf('['); // use last in case of a path containing '[' earlier
            int bracketClose = (bracketOpen >= 0) ? left.IndexOf(']', bracketOpen + 1) : -1;

            string workbook = null;
            string worksheet;

            if (bracketOpen >= 0 && bracketClose > bracketOpen)
            {
                workbook = left.Substring(bracketOpen + 1, bracketClose - bracketOpen - 1); // inside [...]
                worksheet = left.Substring(bracketClose + 1);                               // after ]
            }
            else
            {
                // No workbook brackets found; treat entire left as worksheet.
                worksheet = left;
            }

            // Trim whitespace around worksheet (Excel may not add, but safe)
            if (worksheet != null) worksheet = worksheet.Trim();

            return new ExternalParts
            {
                Workbook = workbook,     // null if not present
                Worksheet = worksheet,
                Address = addr,
                FullPrefix = left         // useful if you want the full token including path
            };
        }

        /// <summary>
        /// Resolves an external address string back to an Excel.Range using the provided Application.
        /// Works when the workbook is already open in the same Excel instance.
        /// </summary>
        public static ExternalResolveResult ResolveRangeWithContext(string externalAddress)
        {
            ServiceLocator.Logger?.LogDebug($"ExcelExternalRef.ResolveRangeWithContext: resolving '{externalAddress}'");

            if (ServiceLocator.ExcelApp == null)
            {
                ServiceLocator.Logger?.LogWarn("ExcelExternalRef.ResolveRangeWithContext: Excel application is not available.");
                throw new InvalidOperationException("Excel application is not available.");
            }
            Excel.Application app = ServiceLocator.ExcelApp;
            var parts = ParseExternalAddress(externalAddress);

            // Find workbook by name among open workbooks
            Excel.Workbook wb = null;
            foreach (Excel.Workbook w in app.Workbooks)
            {
                if (!string.IsNullOrEmpty(parts.Workbook) &&
                    string.Equals(w.Name, parts.Workbook, StringComparison.OrdinalIgnoreCase))
                {
                    wb = w;
                    break;
                }
            }
            if (wb == null)
            {
                ServiceLocator.Logger?.LogWarn($"ExcelExternalRef.ResolveRangeWithContext: workbook '{parts.Workbook}' not found in the current Excel instance.");
                throw new InvalidOperationException($"Workbook '{parts.Workbook}' not found in the current Excel instance.");
            }

            if (wb.Worksheets[parts.Worksheet] is not Excel.Worksheet ws)
            {
                ServiceLocator.Logger?.LogWarn($"ExcelExternalRef.ResolveRangeWithContext: worksheet '{parts.Worksheet}' not found in workbook '{wb.Name}'.");
                throw new InvalidOperationException($"Worksheet '{parts.Worksheet}' not found in workbook '{wb.Name}'.");
            }

            if (ws.Range[parts.Address] is not Excel.Range rng)
            {
                ServiceLocator.Logger?.LogWarn($"ExcelExternalRef.ResolveRangeWithContext: address '{parts.Address}' not found on worksheet '{parts.Worksheet}'.");
                throw new InvalidOperationException($"Address '{parts.Address}' not found on worksheet '{parts.Worksheet}'.");
            }

            ServiceLocator.Logger?.LogDebug($"ExcelExternalRef.ResolveRangeWithContext: resolved to workbook '{wb.Name}', worksheet '{ws.Name}', address '{parts.Address}'");
            return new ExternalResolveResult
            {
                Workbook = wb,
                Worksheet = ws,
                Range = rng
            };
        }
    }
}
