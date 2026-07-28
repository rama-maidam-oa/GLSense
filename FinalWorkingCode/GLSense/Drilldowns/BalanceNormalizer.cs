using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Text.Json.Nodes;
using System.Threading.Tasks;
using GLSense.Utilities;

namespace GLSense.Drilldowns
{
#nullable enable
    public static class BalanceNormalizer
    {
        private const string FunctionName = "GLSENSE_GETBALANCE";

        public static JsonArray Normalize(JsonArray raw)
        {
            LogUtility.LogDebug($"BalanceNormalizer.Normalize started. Input item count: {raw?.Count ?? 0}");
            var output = new JsonArray();

            foreach (var itemNode in raw.OfType<JsonObject>())
            {
                string sheet = itemNode["excelSheet"]?.GetValue<string>() ?? "";
                string cell = itemNode["excelCell"]?.GetValue<string>() ?? "";
                string formulaText = itemNode["inputFormula"]?.GetValue<string>() ?? "";

                // 1) Extract all UDF calls in order from the formula string
                var udfCalls = ExtractUdfCalls(formulaText, FunctionName);

                // 2) Read and order all balanceValue* properties
                var balances = ExtractOrderedBalances(itemNode);

                // 3) Pair them by index and emit rows
                int count = Math.Max(udfCalls.Count, balances.Count);
                for (int i = 0; i < count; i++)
                {
                    var row = new JsonObject
                    {
                        ["excelSheet"] = sheet,
                        ["excelCell"] = cell,
                        ["inputFormula"] = i < udfCalls.Count ? udfCalls[i] : JsonValue.Create<string>(null),
                        ["balanceValue"] = i < balances.Count ? balances[i] : JsonValue.Create<string>(null),
                        ["cache"] = itemNode["cache"]?.GetValue<bool>() ?? false
                    };

                    output.Add(row);
                }
            }

            LogUtility.LogDebug($"BalanceNormalizer.Normalize completed. Output row count: {output.Count}");
            return output;
        }

        // Extract all GLSENSE_GETBALANCE(...) occurrences (with optional @ and _XLL.)
        private static List<string> ExtractUdfCalls(string formulaText, string udfName)
        {
            var result = new List<string>();

            if (string.IsNullOrWhiteSpace(formulaText))
                return result;

            int currentIndex = 0;
            int formulaLength = formulaText.Length;

            while (currentIndex < formulaLength)
            {
                // Look for the function name
                int nameIndex = formulaText.IndexOf(udfName, currentIndex, StringComparison.OrdinalIgnoreCase);
                if (nameIndex < 0)
                    break;

                // Find the actual start of the function call (including possible @ and _XLL. prefix)
                int callStart = FindCallStart(formulaText, nameIndex);

                // Find the end of the function call (matching parentheses)
                int argsStart = formulaText.IndexOf('(', nameIndex + udfName.Length);
                if (argsStart < 0)
                {
                    currentIndex = nameIndex + udfName.Length;
                    continue;
                }

                int callEnd = FindMatchingClosingParenthesis(formulaText, argsStart);
                if (callEnd > argsStart)
                {
                    // Extract the full function call
                    string fullCall = formulaText.Substring(callStart, callEnd - callStart + 1);
                    result.Add(fullCall);

                    // Move currentIndex to after this function call
                    currentIndex = callEnd + 1;
                }
                else
                {
                    currentIndex = nameIndex + udfName.Length;
                }
            }

            return result;
        }

        private static int FindCallStart(string text, int functionNameIndex)
        {
            int start = functionNameIndex;

            // Check for optional @ prefix
            if (start > 0 && text[start - 1] == '@')
            {
                start--;
            }

            // Check for optional _XLL. prefix
            if (start >= "_XLL.".Length &&
                text.Substring(start - "_XLL.".Length, "_XLL.".Length).Equals("_XLL.", StringComparison.OrdinalIgnoreCase))
            {
                start -= "_XLL.".Length;
            }

            return start;
        }

        private static int FindMatchingClosingParenthesis(string text, int openParenIndex)
        {
            int depth = 1;
            bool inQuotes = false;
            int index = openParenIndex + 1;

            while (index < text.Length && depth > 0)
            {
                char current = text[index];

                // Handle quoted strings
                if (current == '"' && (index == 0 || text[index - 1] != '\\'))
                {
                    inQuotes = !inQuotes;
                    index++;
                    continue;
                }

                if (!inQuotes)
                {
                    if (current == '(')
                    {
                        depth++;
                    }
                    else if (current == ')')
                    {
                        depth--;
                        if (depth == 0)
                        {
                            return index;
                        }
                    }
                }
                index++;
            }

            return -1; // No matching closing parenthesis found
        }

        // Extract balanceValue, balanceValue1, balanceValue2, ... ordered by suffix
        private static List<string?> ExtractOrderedBalances(JsonObject item)
        {
            var values = new List<(int Index, string? Value)>();

            foreach (var prop in item)
            {
                string propName = prop.Key;

                if (!propName.StartsWith("balanceValue", StringComparison.OrdinalIgnoreCase))
                    continue;

                string suffix = propName.Substring("balanceValue".Length);
                int idx = 0;

                if (!string.IsNullOrEmpty(suffix))
                {
                    if (!int.TryParse(suffix, NumberStyles.Integer, CultureInfo.InvariantCulture, out idx))
                        continue;
                }

                string? value = prop.Value?.GetValue<string>();
                values.Add((idx, value));
            }

            return values
                .OrderBy(v => v.Index)
                .Select(v => v.Value)
                .ToList();
        }
    }
#nullable disable
}