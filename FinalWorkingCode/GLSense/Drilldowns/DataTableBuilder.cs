using GLSense.Helpers;
using GLSense.Utilities;
using System;
using System.Data;
using System.Globalization;
using System.Linq;
using System.Text.Json;
using System.Text.Json.Nodes;

namespace GLSense.Drilldowns
{
#nullable enable
    public static class DataTableBuilder
    {
        private const string blValue = "balanceValue";
        private const string xlSht = "excelSheet";
        private const string xlCell = "excelCell";
        private const string xlFormula = "inputFormula";
        private const string xlFormulaKey = "formulaKey";
        private const string xlCache = "cache";
        private static readonly string Formuladelimiter = "~!~";
        public static DataTable ToDataTable(JsonArray normalized)
        {
            LogUtility.LogDebug($"DataTableBuilder.ToDataTable started. Input row count: {normalized?.Count ?? 0}");
            var dt = new DataTable();
            dt.Columns.Add(xlSht, typeof(string));
            dt.Columns.Add(xlCell, typeof(string));
            dt.Columns.Add(xlFormula, typeof(string));
            dt.Columns.Add(xlFormulaKey, typeof(string));
            dt.Columns.Add(blValue, typeof(string));
            dt.Columns.Add(xlCache, typeof(bool));

            foreach (var objNode in normalized.OfType<JsonObject>())
            {
                var row = dt.NewRow();

                row[xlSht] = objNode[xlSht]?.GetValue<string>() ?? "";
                row[xlCell] = objNode[xlCell]?.GetValue<string>() ?? "";

                string formula = objNode[xlFormula]?.GetValue<string>() ?? "";
                row[xlFormula] = formula;

                string formulaKey = CreateFormulaKey(formula);
                row[xlFormulaKey] = formulaKey;

                row[xlCache] = objNode[xlCache]?.GetValue<bool>() ?? false;

                // Handle balanceValue: number or string, preserve textual form
                JsonNode? token = objNode[blValue];
                row[blValue] = ToInvariantString(token);

                dt.Rows.Add(row);
            }

            LogUtility.LogDebug($"DataTableBuilder.ToDataTable completed. Output row count: {dt.Rows.Count}");
            return dt;
        }

        /// <summary>
        /// Returns a culture-invariant string for numeric tokens and the original
        /// string for non-numeric tokens. Null → empty string.
        /// </summary>
        private static string ToInvariantString(JsonNode? token)
        {
            if (token == null ||
                token.GetValueKind() == JsonValueKind.Null ||
                token.GetValueKind() == JsonValueKind.Undefined)
            {
                return string.Empty;
            }

            switch (token.GetValueKind())
            {
                case JsonValueKind.Number:
                    // Try integer first
                    if (token is JsonValue jsonValue && jsonValue.TryGetValue(out long longValue))
                    {
                        return longValue.ToString(CultureInfo.InvariantCulture);
                    }

                    // Then double
                    if (token is JsonValue jsonValue2 && jsonValue2.TryGetValue(out double doubleValue))
                    {
                        return doubleValue.ToString("G17", CultureInfo.InvariantCulture);
                    }

                    // Fallback
                    return token.ToJsonString();

                case JsonValueKind.String:
                    return token.GetValue<string>() ?? string.Empty;

                case JsonValueKind.True:
                    return "true";
                case JsonValueKind.False:
                    return "false";

                default:
                    // Arrays, objects, etc.
                    return token.ToJsonString();
            }
        }

        /// <summary>
        /// Creates a compressed formula key based on arguments
        /// </summary>
        private static string CreateFormulaKey(string formulaKey)
        {
            try
            {
                var arguments = CommonFunctions.MultiFormulaValues(formulaKey, "Arguments");
                if (arguments == null)
                    return string.Empty;

                if (arguments.Count > 10 && arguments[11].Contains(";"))
                {
                    var combinedSegments = arguments[11]
                        .Split(new[] { ';' }, StringSplitOptions.None)
                        .ToList();

                    arguments.RemoveAt(11);
                    arguments.InsertRange(11, combinedSegments);
                }

                // Pad to expected argument count (31)
                while (arguments.Count < 31)
                    arguments.Add(string.Empty);

                // Remove quotes from arguments
                for (int j = 0; j < arguments.Count; j++)
                {
                    arguments[j] = arguments[j].Replace("\"", string.Empty);
                }

                // Join and compress
                string joined = string.Join(Formuladelimiter, arguments);
                string compressedFormula = CompressionHelper.CompressString(joined);

                return compressedFormula;
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex, $"Exception encountered while creating formula compressed key for formula {formulaKey}");
                return string.Empty;
            }
        }

    }
#nullable disable
}
