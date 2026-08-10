using GLSense.Utilities;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.Linq;
using System.Text.Json.Nodes;
using System.Text.RegularExpressions;

namespace GLSense.Drilldowns
{
#nullable enable
    /// <summary>
    /// Generates Excel-ready formulas and pre-calculated values for optimal Excel updating.
    /// Cells with no references can be updated directly with values.
    /// Cells with references get formulas that Excel can evaluate.
    /// </summary>
    public static class ExcelFormulaGenerator
    {
        private const string FunctionName = "GLSENSE_GETBALANCE";

        // Regular expression to identify Excel cell references (e.g., A1, $B$2, J9, $K$10)
        private static readonly Regex CellReferenceRegex = new Regex(
            @"\$?[A-Za-z]+\$?\d+",
            RegexOptions.Compiled
        );

        /// <summary>
        /// Main entry point - processes raw JSON and generates Excel-ready output
        /// </summary>
        /// <param name="raw">Raw JSON array containing formula data</param>
        /// <returns>JsonArray with Excel-ready data including formula, value, and reference info</returns>
        public static JsonArray GenerateExcelOutput(JsonArray raw)
        {
            var output = new JsonArray();
            var processedCells = new HashSet<string>();

            LogUtility.LogDebug($"ExcelFormulaGenerator.GenerateExcelOutput started. Input item count: {raw?.Count ?? 0}");

            using (new LogUtility.LogScope("=== Excel Formula Generator Started ===\n"))
            {
                foreach (var itemNode in raw.OfType<JsonObject>())
                {
                    try
                    {
                        var excelItem = ProcessFormulaItem(itemNode);
                        if (excelItem != null)
                        {
                            output.Add(excelItem);

                            string cellKey = $"{excelItem["excelSheet"]}!{excelItem["excelCell"]}";
                            processedCells.Add(cellKey);

                            // Log the item
                            bool hasRefs = excelItem["hasReferences"]?.GetValue<bool>() ?? false;
                            string value = excelItem["value"]?.GetValue<string>() ?? "N/A";
                            string formula = excelItem["formula"]?.GetValue<string>() ?? "N/A";

                            LogUtility.LogDebug($"{cellKey}: hasReferences={hasRefs}, value={value}");
                            if (hasRefs)
                            {
                                LogUtility.LogDebug($"  Formula: {formula}");
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        string sheet = itemNode["excelSheet"]?.GetValue<string>() ?? "unknown";
                        string cell = itemNode["excelCell"]?.GetValue<string>() ?? "unknown";
                        LogUtility.LogException(ex, $"ExcelFormulaGenerator.GenerateExcelOutput: error processing item {sheet}!{cell}");
                    }
                }

                LogUtility.LogDebug($"\n=== Processing Complete: {processedCells.Count} items generated ===");
            }
            return output;
        }

        /// <summary>
        /// Processes a single formula item
        /// </summary>
        private static JsonObject? ProcessFormulaItem(JsonObject itemNode)
        {
            // Extract basic information
            string sheet = itemNode["excelSheet"]?.GetValue<string>() ?? "";
            string cell = itemNode["excelCell"]?.GetValue<string>() ?? "";
            string originalFormula = itemNode["inputFormula"]?.GetValue<string>() ?? "";
            bool cache = itemNode["cache"]?.GetValue<bool>() ?? false;

            if (string.IsNullOrWhiteSpace(originalFormula))
            {
                // If no formula, try to get direct value
                string? directValue = ExtractDirectValue(itemNode);
                return CreateOutputItem(sheet, cell, directValue ?? string.Empty, "", false, cache);
            }

            // Step 1: Extract all GLSENSE_GETBALANCE calls
            var udfCalls = ExtractAllUdfCalls(originalFormula, FunctionName);

            // Step 2: Extract all balance values (balanceValue, balanceValue1, etc.)
            var balanceValues = ExtractAllBalanceValues(itemNode);

            // Step 3: Generate Excel formula by replacing UDF calls with values
            string formulaWithValues = GenerateFormula(originalFormula, udfCalls, balanceValues);

            // Step 4: Identify any cell references in the formula
            var cellReferences = FindCellReferences(formulaWithValues);
            bool hasReferences = cellReferences.Count > 0;

            // Step 5: Preserve the JSON value instead of replacing it with a calculated default.
            string calculatedValue = ExtractDirectValue(itemNode)
                ?? (balanceValues.Count > 0 ? balanceValues[0] : string.Empty);

            // Step 6: Create the output item
            return CreateOutputItem(
                sheet,
                cell,
                calculatedValue,
                formulaWithValues,
                hasReferences,
                cache,
                cellReferences
            );
        }

        /// <summary>
        /// Creates the output JSON item with all required fields
        /// </summary>
        private static JsonObject CreateOutputItem(
            string sheet,
            string cell,
            string value,
            string formula,
            bool hasReferences,
            bool cache,
            List<string>? references = null)
        {
            var item = new JsonObject
            {
                ["excelSheet"] = sheet,
                ["excelCell"] = cell,
                ["value"] = value,
                ["formula"] = formula.StartsWith("=") ? formula : "=" + formula,
                ["hasReferences"] = hasReferences,
                ["cache"] = cache
            };

            // Add references list if there are any
            if (references != null && references.Count > 0)
            {
                item["references"] = new JsonArray(references.Select(r => JsonValue.Create(r)).ToArray());
            }

            return item;
        }

        /// <summary>
        /// Extracts direct value if the cell has no formula
        /// </summary>
        private static string? ExtractDirectValue(JsonObject itemNode)
        {
            if (itemNode.TryGetPropertyValue("value", out var rawValue))
            {
                string? value = ExtractJsonValueAsString(rawValue);
                if (!string.IsNullOrWhiteSpace(value))
                    return value;
            }

            // Check if there's a balanceValue directly
            var balances = ExtractAllBalanceValues(itemNode);
            if (balances.Count > 0)
            {
                return balances[0];
            }

            // Fall back to any primitive property value without coercing to 0
            foreach (var prop in itemNode)
            {
                string? value = ExtractJsonValueAsString(prop.Value);
                if (!string.IsNullOrWhiteSpace(value))
                    return value;
            }

            return null;
        }

        private static string? ExtractJsonValueAsString(JsonNode? node)
        {
            if (node == null)
                return null;

            if (node is JsonValue jsonValue)
            {
                if (jsonValue.TryGetValue<string>(out var stringValue))
                    return stringValue;
                if (jsonValue.TryGetValue<double>(out var doubleValue))
                    return doubleValue.ToString(CultureInfo.InvariantCulture);
                if (jsonValue.TryGetValue<decimal>(out var decimalValue))
                    return decimalValue.ToString(CultureInfo.InvariantCulture);
                if (jsonValue.TryGetValue<long>(out var longValue))
                    return longValue.ToString(CultureInfo.InvariantCulture);
                if (jsonValue.TryGetValue<int>(out var intValue))
                    return intValue.ToString(CultureInfo.InvariantCulture);
                if (jsonValue.TryGetValue<bool>(out var boolValue))
                    return boolValue ? "True" : "False";
            }

            return node.ToJsonString().Trim('"');
        }

        /// <summary>
        /// Calculates the final value for cells with no references
        /// </summary>
        private static string CalculateValue(string formulaWithValues)
        {
            try
            {
                // Remove the = sign if present
                string expression = formulaWithValues.StartsWith("=")
                    ? formulaWithValues.Substring(1)
                    : formulaWithValues;

                // Clean and evaluate
                expression = CleanExpression(expression);
                double result = EvaluateArithmeticExpression(expression);
                return FormatNumericResult(result);
            }
            catch (Exception ex)
            {
                LogUtility.LogWarn($"ExcelFormulaGenerator.CalculateValue: failed to evaluate expression \"{formulaWithValues}\", defaulting to 0: {ex.Message}");
                return "0";
            }
        }

        /// <summary>
        /// Generates Excel formula by replacing UDF calls with their cached values
        /// </summary>
        private static string GenerateFormula(string originalFormula, List<UdfCall> udfCalls, List<string> balanceValues)
        {
            string formula = originalFormula;

            // Replace each UDF call with its corresponding value
            for (int i = 0; i < udfCalls.Count; i++)
            {
                if (i < balanceValues.Count)
                {
                    string value = ConvertToDecimalString(balanceValues[i]);
                    formula = formula.Replace(udfCalls[i].FullCall, value);
                }
            }

            // Clean up any remaining artifacts
            formula = CleanFormula(formula);

            return formula;
        }

        /// <summary>
        /// Extracts all GLSENSE_GETBALANCE calls from a formula
        /// </summary>
        private static List<UdfCall> ExtractAllUdfCalls(string formulaText, string udfName)
        {
            var result = new List<UdfCall>();

            if (string.IsNullOrWhiteSpace(formulaText))
                return result;

            int currentIndex = 0;
            while (currentIndex < formulaText.Length)
            {
                int nameIndex = formulaText.IndexOf(udfName, currentIndex, StringComparison.OrdinalIgnoreCase);
                if (nameIndex < 0)
                    break;

                int callStart = FindCallStart(formulaText, nameIndex);
                int openParenIndex = formulaText.IndexOf('(', nameIndex + udfName.Length);

                if (openParenIndex < 0)
                {
                    currentIndex = nameIndex + udfName.Length;
                    continue;
                }

                int closeParenIndex = FindMatchingClosingParenthesis(formulaText, openParenIndex);
                if (closeParenIndex > openParenIndex)
                {
                    string fullCall = formulaText.Substring(callStart, closeParenIndex - callStart + 1);

                    result.Add(new UdfCall
                    {
                        FullCall = fullCall,
                        StartIndex = callStart,
                        EndIndex = closeParenIndex,
                        NameIndex = nameIndex
                    });

                    currentIndex = closeParenIndex + 1;
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

            if (start > 0 && text[start - 1] == '@')
                start--;

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

                if (current == '"' && (index == 0 || text[index - 1] != '\\'))
                {
                    inQuotes = !inQuotes;
                    index++;
                    continue;
                }

                if (!inQuotes)
                {
                    if (current == '(')
                        depth++;
                    else if (current == ')')
                    {
                        depth--;
                        if (depth == 0)
                            return index;
                    }
                }
                index++;
            }

            return -1;
        }

        /// <summary>
        /// Extracts all balance values (balanceValue, balanceValue1, balanceValue2, etc.)
        /// </summary>
        private static List<string> ExtractAllBalanceValues(JsonObject item)
        {
            var values = new List<(int Index, string Value)>();

            foreach (var prop in item)
            {
                string propName = prop.Key;

                if (!propName.StartsWith("balanceValue", StringComparison.OrdinalIgnoreCase))
                    continue;

                string suffix = propName.Substring("balanceValue".Length);
                int idx = 0;

                if (string.IsNullOrEmpty(suffix))
                {
                    idx = 0;
                }
                else if (int.TryParse(suffix, NumberStyles.Integer, CultureInfo.InvariantCulture, out idx))
                {
                    // Valid suffix number
                }
                else
                {
                    continue;
                }

                string? value = null;
                if (prop.Value is JsonValue jsonValue)
                {
                    if (jsonValue.TryGetValue<string>(out var stringValue))
                        value = stringValue;
                    else if (jsonValue.TryGetValue<double>(out var doubleValue))
                        value = doubleValue.ToString(CultureInfo.InvariantCulture);
                }

                if (value != null)
                {
                    values.Add((idx, value));
                }
            }

            return values.OrderBy(v => v.Index).Select(v => v.Value).ToList();
        }

        /// <summary>
        /// Finds all cell references in a formula
        /// </summary>
        private static List<string> FindCellReferences(string formula)
        {
            var matches = CellReferenceRegex.Matches(formula);
            return matches.Cast<Match>()
                         .Select(m => m.Value)
                         .Where(r => !r.Contains("GLSENSE_GETBALANCE")) // Exclude function names
                         .Distinct()
                         .ToList();
        }

        private static string ConvertToDecimalString(string value)
        {
            if (string.IsNullOrEmpty(value)) return "0";

            try
            {
                if (value.Contains('E') || value.Contains('e'))
                {
                    double doubleValue = double.Parse(value, NumberStyles.Any, CultureInfo.InvariantCulture);
                    return doubleValue.ToString(CultureInfo.InvariantCulture);
                }
                return value;
            }
            catch (Exception ex)
            {
                LogUtility.LogWarn($"ExcelFormulaGenerator.ConvertToDecimalString: failed to parse \"{value}\" as a number, returning it unchanged: {ex.Message}");
                return value;
            }
        }

        private static string CleanFormula(string formula)
        {
            string cleaned = formula;

            // Remove Excel function prefixes
            cleaned = cleaned.Replace("_XLL.", "");
            cleaned = cleaned.Replace("@", "");

            // Ensure formula starts with =
            if (!cleaned.StartsWith("="))
            {
                cleaned = "=" + cleaned;
            }

            return cleaned;
        }

        private static string CleanExpression(string expression)
        {
            // Remove Excel artifacts for calculation
            string cleaned = expression
                .Replace("=", "")
                .Replace("@", "")
                .Replace("_XLL.", "");

            // Remove quoted strings
            cleaned = Regex.Replace(cleaned, "\"[^\"]*\"", "");

            return cleaned;
        }

        private static double EvaluateArithmeticExpression(string expression)
        {
            if (string.IsNullOrWhiteSpace(expression))
                return 0;

            // Remove all whitespace
            expression = new string(expression.Where(c => !char.IsWhiteSpace(c)).ToArray());

            // Handle parentheses recursively
            while (expression.Contains('('))
            {
                int lastOpenParen = expression.LastIndexOf('(');
                int closeParen = expression.IndexOf(')', lastOpenParen);

                if (closeParen > lastOpenParen)
                {
                    string subExpr = expression.Substring(lastOpenParen + 1, closeParen - lastOpenParen - 1);
                    double subResult = EvaluateSimpleExpression(subExpr);

                    expression = expression.Substring(0, lastOpenParen) +
                                subResult.ToString(CultureInfo.InvariantCulture) +
                                expression.Substring(closeParen + 1);
                }
            }

            return EvaluateSimpleExpression(expression);
        }

        private static double EvaluateSimpleExpression(string expression)
        {
            // Handle scientific notation
            expression = ConvertScientificNotation(expression);

            // Tokenize
            var tokens = TokenizeExpression(expression);

            if (tokens.Count == 0) return 0;
            if (tokens.Count == 1) return double.Parse(tokens[0], CultureInfo.InvariantCulture);

            // Handle * and / first
            for (int i = 1; i < tokens.Count - 1; i += 2)
            {
                if (tokens[i] == "*" || tokens[i] == "/")
                {
                    double left = double.Parse(tokens[i - 1], CultureInfo.InvariantCulture);
                    double right = double.Parse(tokens[i + 1], CultureInfo.InvariantCulture);

                    double result = tokens[i] == "*" ? left * right : left / right;

                    tokens[i - 1] = result.ToString(CultureInfo.InvariantCulture);
                    tokens.RemoveRange(i, 2);
                    i -= 2;
                }
            }

            // Handle + and -
            double finalResult = double.Parse(tokens[0], CultureInfo.InvariantCulture);
            for (int i = 1; i < tokens.Count - 1; i += 2)
            {
                double nextValue = double.Parse(tokens[i + 1], CultureInfo.InvariantCulture);

                if (tokens[i] == "+")
                    finalResult += nextValue;
                else if (tokens[i] == "-")
                    finalResult -= nextValue;
            }

            return finalResult;
        }

        private static List<string> TokenizeExpression(string expression)
        {
            var tokens = new List<string>();
            string currentNumber = "";

            for (int i = 0; i < expression.Length; i++)
            {
                char c = expression[i];

                if (char.IsDigit(c) || c == '.' || c == 'E' || c == 'e' ||
                    (c == '-' && (i == 0 || IsOperator(expression[i - 1].ToString()))))
                {
                    currentNumber += c;
                }
                else if (IsOperator(c.ToString()) || c == '(' || c == ')')
                {
                    if (!string.IsNullOrEmpty(currentNumber))
                    {
                        tokens.Add(currentNumber);
                        currentNumber = "";
                    }
                    tokens.Add(c.ToString());
                }
            }

            if (!string.IsNullOrEmpty(currentNumber))
                tokens.Add(currentNumber);

            return tokens;
        }

        private static string ConvertScientificNotation(string expression)
        {
            var regex = new Regex(@"(\d+\.?\d*[Ee][+-]?\d+)");
            return regex.Replace(expression, m =>
            {
                double value = double.Parse(m.Value, NumberStyles.Any, CultureInfo.InvariantCulture);
                return value.ToString(CultureInfo.InvariantCulture);
            });
        }

        private static string FormatNumericResult(double value)
        {
            if (Math.Abs(value) > 1e10 || (Math.Abs(value) < 1e-10 && value != 0))
            {
                return value.ToString("E7", CultureInfo.InvariantCulture);
            }
            return value.ToString("G15", CultureInfo.InvariantCulture);
        }

        private static bool IsOperator(string token)
        {
            return token == "+" || token == "-" || token == "*" || token == "/";
        }

        private class UdfCall
        {
            public string FullCall { get; set; } = string.Empty;
            public int StartIndex { get; set; }
            public int EndIndex { get; set; }
            public int NameIndex { get; set; }
        }
    }
#nullable restore
}