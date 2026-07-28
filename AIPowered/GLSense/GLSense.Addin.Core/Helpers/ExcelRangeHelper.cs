// ExcelRangeHelper.cs in GLSense.Addin.Core
// Verbatim port of GLSense\Helpers\ExcelRangeHelper.cs (FinalWorkingCode) - purely
// string/regex logic, no Excel COM calls, no statics from the old project. No changes
// needed beyond the namespace.
using GLSense.Addin.Core.Infrastructure;
using System;
using System.Text.RegularExpressions;

namespace GLSense.Addin.Core.Helpers
{
    /// <summary>
    /// Helper for validating Excel A1-style references/ranges.
    /// Applies ordered signals: $, !, [ ] ' : INDIRECT then validates via regex + bounds.
    /// </summary>
    public static class ExcelRangeHelper
    {
        // Comprehensive A1 regex:
        // - Optional [Workbook]
        // - Optional Sheet! (supports quoted names with doubled single quotes: '')
        // - First address (A1 / $A$1 / mixed)
        // - Optional :Second address
        private static readonly Regex A1AddressOrRangeRegex = new(
            @"^" +
            @"(?:(\[[^\]]+\]))?" +                         // [Workbook]
            @"(?:(?:[A-Za-z0-9_]+|'(?:''|[^'])+')!)?" +    // Sheet! or 'Sheet Name'!
            @"(?:\$?[A-Za-z]{1,3}\$?\d+)" +                // First address (A1/$A$1/mixed)
            @"(?::\$?[A-Za-z]{1,3}\$?\d+)?" +              // Optional :Second address
            @"$",
            RegexOptions.Compiled | RegexOptions.CultureInvariant);

        // Detect INDIRECT( ... ) ignoring case and whitespace; allows leading/trailing spaces
        private static readonly Regex IndirectCallRegex = new(
            @"^\s*INDIRECT\s*\(", RegexOptions.Compiled | RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);

        // Extract a simple quoted string argument from INDIRECT("...") allowing doubled quotes inside Excel strings
        private static readonly Regex IndirectQuotedArgRegex = new(
            // Start: optional spaces, INDIRECT( optional spaces, then a quoted string
            @"^\s*INDIRECT\s*\(\s*""(?<arg>(?:[^""]|"""")*)""\s*\)\s*$",
            RegexOptions.Compiled | RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);

        /// <summary>
        /// Returns true if <paramref name="input"/> is a valid Excel A1 reference or range,
        /// following ordered signal rules:
        /// 1) '$' → validate
        /// 2) '!' → validate
        /// 3) any of [ ] ' : → validate
        /// 4) INDIRECT(...) → validate (simple quoted argument), or via resolver when provided
        /// Otherwise returns false (treat as arbitrary string).
        /// </summary>
        /// <param name="input">Candidate string (e.g., "$A$1", "Sheet1!A1:B2", "INDIRECT("'Sheet1'!A1")").</param>
        /// <param name="indirectResolver">
        /// Optional resolver for dynamic INDIRECT expressions (concats, references).
        /// Given the raw input (e.g., "INDIRECT(A1&"":B2"")"), it should return
        /// a resolved address string if possible; otherwise null.
        /// </param>
        public static bool IsRealRange(object input, Func<string, string> indirectResolver = null)
        {
            if (input == null || string.IsNullOrWhiteSpace(input.ToString())) return false;

            var s = input.ToString().Trim();

            // Direct Excel range patterns
            if (HasExcelSyntax(s)) return IsValidA1RefOrRange(s);

            // INDIRECT function patterns
            return IndirectCallRegex.IsMatch(s) && EvaluateIndirectPattern(s, indirectResolver);
        }

        private static bool HasExcelSyntax(string s) =>
            s.IndexOf('$') >= 0 || s.IndexOf('!') >= 0 ||
            s.IndexOf('[') >= 0 || s.IndexOf(']') >= 0 ||
            s.IndexOf('\'') >= 0 || s.IndexOf(':') >= 0;

        private static bool EvaluateIndirectPattern(string s, Func<string, string> indirectResolver)
        {
            var match = IndirectQuotedArgRegex.Match(s);
            if (match.Success)
            {
                var inner = UnescapeExcelString(match.Groups["arg"].Value).Trim();
                return IsValidA1RefOrRange(inner);
            }

            return indirectResolver != null && TryResolveIndirect(s, indirectResolver);
        }

        private static bool TryResolveIndirect(string s, Func<string, string> resolver)
        {
            try
            {
                var resolved = resolver(s)?.Trim();
                return !string.IsNullOrWhiteSpace(resolved) && IsValidA1RefOrRange(resolved);
            }
            catch (Exception ex)
            {
                ServiceLocator.Logger?.LogDebug($"ExcelRangeHelper.TryResolveIndirect: indirectResolver threw for '{s}' - {ex.Message}");
                return false;
            }
        }

        /// <summary>
        /// Regex screen + bounds checks; validates single address or a colon-separated range.
        /// </summary>
        private static bool IsValidA1RefOrRange(string s)
        {
            if (!A1AddressOrRangeRegex.IsMatch(s))
                return false;

            // Strip workbook/sheet qualifiers by taking substring after last '!'
            var addressPart = s;
            int bang = s.LastIndexOf('!');
            if (bang >= 0 && bang + 1 < s.Length)
                addressPart = s.Substring(bang + 1);

            var parts = addressPart.Split(':');
            if (parts.Length == 1)
                return IsValidSingleAddress(parts[0]);
            if (parts.Length == 2)
                return IsValidSingleAddress(parts[0]) && IsValidSingleAddress(parts[1]);

            // More than one ':' → invalid
            return false;
        }

        /// <summary>
        /// Validates a single address token (e.g., "A1", "$A$1", "A$1", "$A1").
        /// Ensures column letters ≤ XFD (1..16384) and row ≤ 1,048,576.
        /// </summary>
        private static bool IsValidSingleAddress(string addr)
        {
            if (string.IsNullOrWhiteSpace(addr)) return false;
            addr = addr.Trim();

            var m = Regex.Match(addr,
                @"^\$?([A-Za-z]{1,3})\$(\d+)$|^\$?([A-Za-z]{1,3})(\d+)$",
                RegexOptions.CultureInvariant);

            if (!m.Success) return false;

            string letters = !string.IsNullOrEmpty(m.Groups[1].Value) ? m.Groups[1].Value : m.Groups[3].Value;
            string digits = !string.IsNullOrEmpty(m.Groups[2].Value) ? m.Groups[2].Value : m.Groups[4].Value;

            letters = letters.ToUpperInvariant();
            if (!int.TryParse(digits, out int row)) return false;

            // Excel row bounds: 1..1,048,576
            if (row < 1 || row > 1_048_576) return false;

            // Excel column bounds: A..XFD → 1..16384
            int col = ExcelColumnToNumber(letters);
            return col >= 1 && col <= 16_384;
        }

        /// <summary>
        /// Converts Excel column letters (A..Z, AA..ZZ, AAA..ZZZ) to a 1-based index.
        /// Returns -1 if non A–Z characters occur.
        /// </summary>
        private static int ExcelColumnToNumber(string col)
        {
            int n = 0;
            foreach (char ch in col)
            {
                if (ch < 'A' || ch > 'Z') return -1;
                n = n * 26 + (ch - 'A' + 1);
            }
            return n;
        }

        /// <summary>
        /// Unescapes an Excel-quoted string by turning doubled quotes "" into ".
        /// </summary>
        private static string UnescapeExcelString(string s)
        {
            return s?.Replace(@"""""", @"""");
        }
    }
}
