// ClsFormulaParser.cs in GLSense.Addin.Core
// Port of GLSense\Helpers\ClsFormulaParser.cs (FinalWorkingCode).
// Changes: AppState.Instance.ExcelApp (old static) -> ServiceLocator.ExcelApp;
// LogUtility.LogException -> ServiceLocator.Logger.LogException. ExcelRangeHelper is
// same-namespace (also ported, see ExcelRangeHelper.cs in this folder) so no
// qualification change needed there.
using GLSense.Addin.Core.Infrastructure;
using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Text;
using System.Text.RegularExpressions;

namespace GLSense.Addin.Core.Helpers
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
                    string regexPattern = @"""([^""\)]|(""""))*""(?:&[^,]*)?|[^,]+";

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
                    if (args[i].Contains("$") || args[i].IndexOf("$") >= 0)
                    {
                        args[i] = ValueFromReference(args[i]);
                    }
                    else if (args[i].Contains("&") && args[i].Contains("~"))
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
                    if (arguments[i].Contains("$") || arguments[i].IndexOf("$") >= 0)
                    {
                        arguments[i] = ValueFromReference(arguments[i]);
                    }
                    else if (arguments[i].Contains("&") && arguments[i].Contains("~"))
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

        private static List<string> ProcessArguments(List<string> arguments)
        {
            var lastTwoArgs = ExtractLastTwoArguments(arguments);
            var splitValues = ProcessLastArguments(lastTwoArgs);

            return splitValues.Any()
                ? BuildArgumentsWithSplitValues(arguments, splitValues)
                : ResolveAllReferences(arguments);
        }

        private static List<string> ExtractLastTwoArguments(List<string> arguments)
        {
            var startIndex = Math.Max(0, arguments.Count - 2);
            return arguments.Skip(startIndex).ToList();
        }

        private static string[] ProcessLastArguments(List<string> lastTwoArgs)
        {
            foreach (var arg in lastTwoArgs)
            {
                var cleanArg = CleanQuotes(arg);
                if (string.IsNullOrEmpty(cleanArg)) continue;

                var splitValues = TrySplitArguments(cleanArg);
                if (splitValues != null && splitValues.Length > 0)
                {
                    return splitValues;
                }
            }

            return Array.Empty<string>();
        }

        private static string CleanQuotes(string input)
        {
            return input.Replace("\"", "");
        }

        private static string[] TrySplitArguments(string argument)
        {
            if (argument.Contains(";") && !argument.Contains("$"))
            {
                return argument.Split(';');
            }

            if (argument.Contains("$"))
            {
                var resolvedValue = ValueFromReference(argument);
                if (!string.IsNullOrEmpty(resolvedValue))
                {
                    return resolvedValue.Split(';');
                }
            }

            return Array.Empty<string>();
        }

        private static List<string> BuildArgumentsWithSplitValues(List<string> originalArgs, string[] splitValues)
        {
            var result = new List<string>();

            // Add first 10 arguments or all if less than 10
            var countToTake = Math.Min(10, originalArgs.Count);
            result.AddRange(originalArgs.Take(countToTake));

            // Add split values without quotes
            result.AddRange(splitValues.Select(CleanQuotes));

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
            var resolvedArg = argument.Contains("$")
                ? ValueFromReference(argument)
                : argument;

            return resolvedArg.Replace("&", "\"");
        }

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
            if (ServiceLocator.ExcelApp == null)
            {
                return "\"";
            }

            try
            {
                string rngAddress = string.Empty;
                rngAddress = ReplaceInDirects(refAddress.Replace('\"'.ToString(), ""));

                if (ExcelRangeHelper.IsRealRange(rngAddress))
                {
                    Range rng = ServiceLocator.ExcelApp.Range[rngAddress];
                    if (rng != null && rng.Value != null)
                    {
                        try
                        {
                            return "\"" + (rng.Value?.ToString().Trim() ?? "") + "\"";
                        }
                        catch (Exception)
                        {
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

        private static void LogError(Exception ex, [CallerMemberName] string callerName = "")
        {
            // CallerMemberName automatically captures the calling method (ExtractArguments,
            // FormulaArgs, RangeVal, etc.) without needing to touch every call site below -
            // gives each of this hot-path parser's many catch blocks a distinct, useful
            // context string in the log for root-causing bad-formula parsing failures.
            ServiceLocator.Logger?.LogException(ex, $"ClsFormulaParser.{callerName}");
        }
    }
}
