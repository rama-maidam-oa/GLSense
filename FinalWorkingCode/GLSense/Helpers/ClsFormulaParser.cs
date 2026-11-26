using GLSense.Utilities;
using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;

namespace GLSense.Helpers
{
    public class ClsFormulaParser
    {
        public string Formula { get; set; }
        public ClsFormulaParser(string formula)
        {
            Formula = formula;
        }
        // Function to extract function calls from the formula
        public List<string> ExtractFunctions()
        {
            var functions = new List<string>();
            string regexPattern = @"@?\w+\s*\((?:[^()]*|\((?:[^()]*|\([^()]*\))*\))*\)";
            var matches = Regex.Matches(Formula, regexPattern);

            foreach (Match match in matches)
            {
                functions.Add(match.Value);
            }

            return functions;
        }
        // Function to extract arguments from a single function call
        public static List<string> ExtractArguments(string functionCall)
        {
            var args = new List<string>();
            try
            {
                int startIdx = functionCall.IndexOf("(") + 1;
                int endIdx = functionCall.LastIndexOf(")");
                if (startIdx > 0 && endIdx > startIdx)
                {
                    string argString = functionCall.Substring(startIdx, endIdx - startIdx);
                    // A quoted literal may contain ) and , (e.g. "R&D (US), Inc") - OISR-22864
                    string regexPattern = @"""([^""]|(""""))*""(?:&[^,]*)?|[^,]+";

                    args = Regex.Matches(argString, regexPattern)
                                .Cast<Match>()
                                .Select(m => m.Value.Trim())
                                .ToList();
                }
            }
            catch (Exception ex)
            {
                LogError(ex);
                return new List<string>();
            }
            return args;
        }
        public static List<string> ExtractArguments_WithValues(string functionCall)
        {
            try
            {
                var args = ExtractArguments(functionCall);
                if (args == null) return new List<string>();

                for (int i = 0; i < args.Count; i++)
                {
                    if (!IsQuotedLiteral(args[i]) && args[i].Contains("$"))
                    {
                        args[i] = ValueFromReference(args[i]);
                    }
                    // Joins a period range "Jan-24"&"~"&"Mar-24"; a plain quoted value keeps its "&" ("R&D~") - OISR-22901
                    else if (!IsQuotedLiteral(args[i]) && args[i].Contains("&") && args[i].Contains("~"))
                    {
                        args[i] = args[i].Replace("&", "");
                    }
                }

                return args;
            }
            catch (Exception ex)
            {
                LogError(ex);
                return new List<string>();
            }
        }

        public List<string> FormulaArgs()
        {
            var args = new List<string>();
            try
            {
                // Extract content inside the parentheses
                int startIdx = Formula.IndexOf("(");
                int endIdx = Formula.LastIndexOf(")");

                // Ensure valid parentheses exist
                if (startIdx < 0 || endIdx < 0 || endIdx <= startIdx)
                {
                    return args; // Return empty list if parentheses are incorrect
                }

                string insideArgs = Formula.Substring(startIdx + 1, endIdx - startIdx - 1);

                // Use a parser that respects nested parentheses
                args = SplitFunctionArguments(insideArgs);
            }
            catch (Exception ex)
            {
                LogError(ex);
            }

            return args;
        }
        // Parses function arguments while respecting nested parentheses and quoted strings
        private static List<string> SplitFunctionArguments(string input)
        {
            var args = new List<string>();
            var sb = new StringBuilder();
            bool inQuotes = false;
            int openParens = 0;

            for (int i = 0; i < input.Length; i++)
            {
                char ch = input[i];

                // Toggle quotes tracking (ignore commas inside quotes)
                if (ch == '\"')
                {
                    inQuotes = !inQuotes;
                }

                // Track parentheses depth (ignore commas inside nested functions)
                if (ch == '(' && !inQuotes)
                {
                    openParens++;
                }
                else if (ch == ')' && !inQuotes)
                {
                    openParens--;
                }

                // Split arguments only on commas that are outside parentheses and quotes
                if (ch == ',' && openParens == 0 && !inQuotes)
                {
                    args.Add(sb.ToString().Trim());
                    sb.Clear();
                }
                else
                {
                    sb.Append(ch);
                }
            }

            // Add the last argument
            if (sb.Length > 0)
            {
                args.Add(sb.ToString().Trim());
            }

            return args;
        }
        public List<string> FormulaArgs_Values()
        {
            try
            {
                // Use Regex to split by commas but respect double quotes
                var arguments = FormulaArgs();
                if (arguments == null) return new List<string>();

                for (int i = 0; i < arguments.Count; i++)
                {
                    if (!IsQuotedLiteral(arguments[i]) && arguments[i].Contains("$"))
                    {
                        arguments[i] = ValueFromReference(arguments[i]);
                    }
                    // Joins a period range "Jan-24"&"~"&"Mar-24"; a plain quoted value keeps its "&" ("R&D~") - OISR-22901
                    else if (!IsQuotedLiteral(arguments[i]) && arguments[i].Contains("&") && arguments[i].Contains("~"))
                    {
                        arguments[i] = arguments[i].Replace("&", "");
                    }
                }

                return arguments;
            }
            catch (Exception ex)
            {
                LogError(ex);
                return new List<string>();
            }
        }

        public string Formula_Values()
        {
            try
            {
                var arguments = FormulaArgs();
                if (!HasArguments(arguments))
                {
                    return string.Empty;
                }

                var processedArguments = ProcessArguments(arguments);
                return FormatFormula(processedArguments);
            }
            catch (Exception ex)
            {
                LogError(ex);
                return string.Empty;
            }
        }

        private static bool HasArguments(List<string> arguments)
        {
            return arguments != null && arguments.Count > 0;
        }

        // GLSense_GetBalance: 11 fixed arguments (sign ... JE Category), segments from index 11.
        private const int FirstSegmentIndex = 11;

        // OISR-22901: only a real ";" list in a segment argument (all segments in one argument) is expanded, in place.
        // Before, any $ reference in the last two arguments took this branch and the rebuild kept only the first 10
        // arguments, so the reference's value landed in JE Category and the other segments were dropped.
        private static List<string> ProcessArguments(List<string> arguments)
        {
            for (int i = Math.Max(FirstSegmentIndex, arguments.Count - 2); i < arguments.Count; i++)
            {
                var cleanArg = CleanQuotes(arguments[i]);
                if (string.IsNullOrEmpty(cleanArg)) continue;

                var splitValues = TrySplitArguments(cleanArg, IsQuotedLiteral(arguments[i]));
                if (splitValues.Length > 1)
                {
                    return BuildArgumentsWithSplitValues(arguments, i, splitValues);
                }
            }

            return ResolveAllReferences(arguments);
        }

        private static string CleanQuotes(string input)
        {
            return input.Replace("\"", "");
        }

        private static string[] TrySplitArguments(string argument, bool isLiteral)
        {
            if (argument.Contains(";") && (isLiteral || !argument.Contains("$")))
            {
                return argument.Split(';');
            }

            if (!isLiteral && argument.Contains("$"))
            {
                var resolvedValue = ValueFromReference(argument);
                if (!string.IsNullOrEmpty(resolvedValue))
                {
                    return resolvedValue.Split(';');
                }
            }

            return Array.Empty<string>();
        }

        private static List<string> BuildArgumentsWithSplitValues(List<string> originalArgs, int splitIndex, string[] splitValues)
        {
            var result = new List<string>();

            // Arguments before the list, references resolved
            result.AddRange(originalArgs.Take(splitIndex).Select(ResolveArgument));

            // Add split values without quotes
            result.AddRange(splitValues.Select(CleanQuotes));

            // Arguments after the list (the other of the last two), references resolved
            result.AddRange(originalArgs.Skip(splitIndex + 1).Select(ResolveArgument));

            // Add empty string as required
            result.Add(string.Empty);

            return FormatArguments(result);
        }

        private static List<string> ResolveAllReferences(List<string> arguments)
        {
            var processedArgs = arguments.Select(ResolveArgument).ToList();
            return FormatArguments(processedArgs);
        }

        private static string ResolveArgument(string argument)
        {
            // A plain quoted value is used as written: its "&" is part of the value ("R&D") - OISR-22901
            if (IsQuotedLiteral(argument))
                return argument;

            // A reference: the cell's value as it is (ValueFromReference already joins $A$1&"~"&$B$1).
            if (argument.Contains("$"))
                return ValueFromReference(argument);

            // "Jan-24"&"~"&"Mar-24": the "&"s become quotes, which FormatArguments then removes
            return argument.Replace("&", "\"");
        }

        // A whole-argument Excel text constant ("..." with "" escapes) is a value, never a cell
        // reference, even if it contains $ (e.g. segment value "11$9") - OISR-22860
        private static bool IsQuotedLiteral(string arg) =>
            arg != null && Regex.IsMatch(arg.Trim(), "^\"(?:[^\"]|\"\")*\"$");

        private static List<string> FormatArguments(List<string> arguments)
        {
            return arguments.Select(arg => $"\"{CleanQuotes(arg)}\"").ToList();
        }

        private string FormatFormula(List<string> arguments)
        {
            var openParenIndex = Formula.IndexOf("(");
            var functionPart = Formula.Substring(0, openParenIndex + 1);
            var argumentsPart = string.Join(",", arguments);

            return $"{functionPart}{argumentsPart})";
        }

        // Method to count the number of arguments in the formula
        public int Formula_ArgsCount()
        {
            try
            {
                // Fetch the list of arguments
                var arguments = FormulaArgs();
                if (arguments == null || arguments.Count == 0)
                {
                    return 0; // Return 0 if no arguments are provided
                }

                int argCount = arguments.Count; // Start with the base count of arguments

                // Process the last two arguments (or fewer if the list is shorter)
                for (int i = Math.Max(0, arguments.Count - 2); i < arguments.Count; i++)
                {
                    string argStr = arguments[i].Replace('\"'.ToString(), ""); // Remove quotes
                    if (string.IsNullOrEmpty(argStr)) continue; // Skip empty arguments

                    // Handle semicolon-delimited strings
                    if (argStr.Contains(";") && !argStr.Contains("$"))
                    {
                        var splitArgs = argStr.Split(';'); // Split by ';'
                        argCount += splitArgs.Length - 1; // Add the additional parts to the count
                    }
                    // Handle reference strings containing '$'
                    else if (argStr.Contains("$"))
                    {
                        string resolvedValue = ValueFromReference(argStr);
                        if (!string.IsNullOrEmpty(resolvedValue) && resolvedValue.Contains(";"))
                        {
                            var splitArgs = resolvedValue.Split(';'); // Split the resolved value
                            argCount += splitArgs.Length - 1; // Add the additional parts
                        }
                    }
                }

                return argCount; // Return the final computed argument count
            }
            catch (Exception ex)
            {
                LogError(ex); // Log the error if something goes wrong
                return 0; // Return 0 in case of an error
            }
        }

        public string Formula_Correction(int expectedArgumentCount, int insertAfterIndex)
        {
            try
            {
                var arguments = FormulaArgs();
                int missingArguments = expectedArgumentCount - arguments.Count;
                if (missingArguments > 0)
                {
                    for (int i = 1; i <= missingArguments; i++)
                    {
                        arguments.Insert(insertAfterIndex + 1, "\"\"");
                    }
                }
                string correctedArgs = string.Join(",", arguments);
                return $"{Formula.Substring(0, Formula.IndexOf("(") + 1)}{correctedArgs})";
            }
            catch (Exception ex)
            {
                LogError(ex);
                return string.Empty;
            }
        }

        private static string ValueFromReference(string strAddress)
        {
            if (strAddress.Contains("&"))
            {
                strAddress = Regex.Replace(strAddress, @"(&)""~""(&)", "\"~\"");
            }

            try
            {
                if (strAddress.Contains(";"))
                {
                    return string.Join(";", strAddress.Split(';').Select(part => ProcessSection(part)));
                }
                else if (strAddress.Contains(","))
                {
                    return string.Join(",", strAddress.Split(',').Select(part => ProcessSection(part)));
                }
                else if (strAddress.Contains("~"))
                {
                    var parts = strAddress.TrimEnd('~').Split('~');
                    return string.Join("~", parts.Select(p => ProcessAddressPart(p)));
                }
                else if (strAddress.Contains("|"))
                {
                    var parts = strAddress.Split('|');
                    return string.Join("|", parts.Select(p => ProcessAddressPart(p)));
                }
                else
                {
                    return RangeVal(strAddress);
                }
            }
            catch (Exception ex)
            {
                LogError(ex);
                return string.Empty;
            }
        }

        private static string ProcessSection(string section)
        {
            if (section.Contains(","))
            {
                return string.Join(",", section.Split(',').Select(part => ProcessSubsection(part)));
            }
            else
            {
                return ProcessSubsection(section);
            }
        }

        private static string ProcessSubsection(string subsection)
        {
            if (subsection.Contains("|"))
            {
                var parts = subsection.Split('|');
                return string.Join("|", parts.Select(p => ProcessAddressPart(p)));
            }
            else
            {
                return ProcessAddressPart(subsection);
            }
        }

        private static string ProcessAddressPart(string part)
        {
            string result;

            if (string.IsNullOrEmpty(part))
            {
                return "\"";
            }
            else if (part.Contains("$"))
            {
                result = RangeVal(part);
            }
            else
            {
                result = part.Trim();
            }

            return result;
        }

        private static string RangeVal(string refAddress)
        {
            if (AppState.Instance.ExcelApp == null)
            {
                return "\"";
            }

            try
            {
                string rngAddress = string.Empty;
                rngAddress = ReplaceInDirects(refAddress.Replace('\"'.ToString(), ""));

                if (ExcelRangeHelper.IsRealRange(rngAddress))
                {
                    Range rng = AppState.Instance.ExcelApp.Range[rngAddress];
                    if (rng != null && rng.Value != null)
                    {
                        try
                        {
                            return "\"" + (rng.Value?.ToString().Trim() ?? "") + "\"";
                        }
                        catch (Exception ex)
                        {
                            LogUtility.LogWarn($"ClsFormulaParser.RangeVal: Trim() failed on range value (ignored): {ex.Message}");
                            return "\"" + (rng.Value?.ToString() ?? "") + "\"";
                        }
                    }
                }
                return "\"";
            }
            catch (Exception ex)
            {
                LogError(ex);
                return "\"";
            }
        }

        private static string ReplaceInDirects(string indirectExpression)
        {
            try
            {
                if (string.IsNullOrEmpty(indirectExpression))
                {
                    return string.Empty;
                }

                if (indirectExpression.IndexOf("INDIRECT", StringComparison.OrdinalIgnoreCase) == -1)
                {
                    return indirectExpression;
                }

                // Define a regex pattern to capture the innermost INDIRECT argument
                string pattern = @"INDIRECT\((.*)\)";
                var regex = new Regex(pattern, RegexOptions.IgnoreCase);
                string result = indirectExpression.Trim();

                // Process until all INDIRECT calls are resolved
                while (regex.IsMatch(result))
                {
                    Match match = regex.Match(result);
                    if (match.Success)
                    {
                        result = match.Groups[1].Value.Trim(); // Extract the argument within INDIRECT()

                        // Remove surrounding quotes if present
                        if (result.StartsWith("\"") && result.EndsWith("\""))
                        {
                            result = result.Substring(1, result.Length - 2);
                        }
                    }
                    else
                    {
                        break;
                    }
                }

                return result;
            }
            catch (Exception ex)
            {
                LogError(ex);
                return string.Empty;
            }
        }

        // You'll need to implement these methods or ensure they exist elsewhere
        private static void LogError(Exception ex, [System.Runtime.CompilerServices.CallerMemberName] string context = "")
        {
            LogUtility.LogException(ex, $"ClsFormulaParser.{context}");
        }
    }
}
