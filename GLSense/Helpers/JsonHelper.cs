using System;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.Encodings.Web;
using System.Text.Json;
using System.Text.Json.Serialization;

namespace GLSense.Helpers
{
#nullable enable
    /// <summary>
    /// Global reusable JSON configuration
    /// Used for Serialize / Deserialize operations
    /// </summary>
    public static class JsonGlobals
    {
        public static JsonSerializerOptions Options { get; }

        static JsonGlobals()
        {
            var opts = new JsonSerializerOptions
            {
                PropertyNameCaseInsensitive = true,
                PropertyNamingPolicy = null,
                DefaultIgnoreCondition = JsonIgnoreCondition.WhenWritingNull,
                ReadCommentHandling = JsonCommentHandling.Skip,
                WriteIndented = false,
                AllowTrailingCommas = true,
                Encoder = JavaScriptEncoder.UnsafeRelaxedJsonEscaping,
                NumberHandling =
                    JsonNumberHandling.AllowReadingFromString |
                    JsonNumberHandling.AllowNamedFloatingPointLiterals
            };

            // Add custom converter for flexible string handling
            opts.Converters.Add(new FlexibleStringConverter());

#if NET8_0_OR_GREATER
            opts.MakeReadOnly();
#endif

            Options = opts;
        }
    }

    /// <summary>
    /// Custom JSON converter that flexibly handles string conversions
    /// Converts numbers, booleans, and null values to strings
    /// </summary>
    // ComVisible(false): this assembly is [assembly: ComVisible(true)] overall (needed
    // for the Add-in Express ribbon/ExcelFunctions COM surface), but this is an internal
    // System.Text.Json helper, never touched via COM. Since it derives from the generic
    // JsonConverter<string>, the type library exporter can't build a COM "class interface"
    // for it and warns about it on every build - opting this type out of the COM surface
    // (matching the same pattern already used on Models/ViewModels in this codebase)
    // removes the warning with no functional effect.
    [ComVisible(false)]
    public class FlexibleStringConverter : JsonConverter<string>
    {
        public override string? Read(ref Utf8JsonReader reader, Type typeToConvert, JsonSerializerOptions options)
        {
            return reader.TokenType switch
            {
                JsonTokenType.String => reader.GetString(),
                JsonTokenType.Number => ReadNumberAsString(ref reader),
                JsonTokenType.True => "true",
                JsonTokenType.False => "false",
                JsonTokenType.Null => null,
                _ => reader.GetString() // Fallback
            };
        }

        private static string ReadNumberAsString(ref Utf8JsonReader reader)
        {
            // Try to get as different numeric types and convert to string
            if (reader.TryGetInt64(out long longValue))
                return longValue.ToString();

            if (reader.TryGetInt32(out int intValue))
                return intValue.ToString();

            if (reader.TryGetDouble(out double doubleValue))
                return doubleValue.ToString(System.Globalization.CultureInfo.InvariantCulture);

            if (reader.TryGetDecimal(out decimal decimalValue))
                return decimalValue.ToString(System.Globalization.CultureInfo.InvariantCulture);

            // Fallback: read as string using the raw text span
            try
            {
                var span = reader.ValueSpan;
                if (!span.IsEmpty)
                {
                    // Convert ReadOnlySpan<byte> to byte[] then to string
                    byte[] byteArray = span.ToArray();
                    return Encoding.UTF8.GetString(byteArray);
                }
            }
            catch
            {
                // If conversion fails, try alternative
            }

            // Ultimate fallback - try to get as string
            try
            {
                return reader.GetString() ?? "0";
            }
            catch
            {
                return "0";
            }
        }

        public override void Write(Utf8JsonWriter writer, string? value, JsonSerializerOptions options)
        {
            if (value is null)
                writer.WriteNullValue();
            else
                writer.WriteStringValue(value);
        }
    }

    /// <summary>
    /// Safe JSON parsing helpers for dynamic API responses
    /// </summary>
    public static class JsonHelper
    {
        public static bool TryGetProperty(
            JsonElement element,
            string propertyName,
            out JsonElement value)
        {
            if (element.ValueKind != JsonValueKind.Object)
            {
                value = default;
                return false;
            }

            var match = element
                .EnumerateObject()
                .FirstOrDefault(p =>
                    string.Equals(p.Name, propertyName,
                        StringComparison.OrdinalIgnoreCase));

            if (!match.Equals(default(JsonProperty)))
            {
                value = match.Value;
                return true;
            }

            value = default;
            return false;
        }

        public static bool TryGetDouble(JsonElement element, out double result)
        {
            result = 0;

            switch (element.ValueKind)
            {
                case JsonValueKind.Number:
                    return element.TryGetDouble(out result);

                case JsonValueKind.String:
                    return double.TryParse(
                        element.GetString(),
                        System.Globalization.NumberStyles.Any,
                        System.Globalization.CultureInfo.InvariantCulture,
                        out result);

                default:
                    return false;
            }
        }

        public static bool TryGetString(JsonElement element, out string? result)
        {
            result = null;

            if (element.ValueKind == JsonValueKind.String)
            {
                result = element.GetString();
                return !string.IsNullOrWhiteSpace(result);
            }

            // Also handle numbers as strings
            if (element.ValueKind == JsonValueKind.Number)
            {
                result = GetNumberAsString(element);
                return !string.IsNullOrWhiteSpace(result);
            }

            return false;
        }

        private static string GetNumberAsString(JsonElement element)
        {
            if (element.TryGetInt64(out long longValue))
                return longValue.ToString();

            if (element.TryGetInt32(out int intValue))
                return intValue.ToString();

            if (element.TryGetDouble(out double doubleValue))
                return doubleValue.ToString(System.Globalization.CultureInfo.InvariantCulture);

            if (element.TryGetDecimal(out decimal decimalValue))
                return decimalValue.ToString(System.Globalization.CultureInfo.InvariantCulture);

            return element.ToString() ?? "0";
        }

        public static string GetStringSafe(JsonElement element, string propertyName)
        {
            if (element.ValueKind != JsonValueKind.Object)
                return string.Empty;

            var property = element
                .EnumerateObject()
                .FirstOrDefault(p =>
                    string.Equals(p.Name,
                                  propertyName,
                                  StringComparison.OrdinalIgnoreCase));

            if (!property.Equals(default(JsonProperty)))
            {
                var value = property.Value;

                return value.ValueKind switch
                {
                    JsonValueKind.String => value.GetString() ?? string.Empty,
                    JsonValueKind.Number => GetNumberAsString(value),
                    JsonValueKind.True => "true",
                    JsonValueKind.False => "false",
                    JsonValueKind.Null => string.Empty,
                    _ => value.ToString() ?? string.Empty
                };
            }

            return string.Empty;
        }

        /// <summary>
        /// Safely gets an integer from a JsonElement, handling both number and string values
        /// </summary>
        public static bool TryGetInt(JsonElement element, out int result)
        {
            result = 0;

            switch (element.ValueKind)
            {
                case JsonValueKind.Number:
                    return element.TryGetInt32(out result);

                case JsonValueKind.String:
                    return int.TryParse(
                        element.GetString(),
                        System.Globalization.NumberStyles.Any,
                        System.Globalization.CultureInfo.InvariantCulture,
                        out result);

                default:
                    return false;
            }
        }

        /// <summary>
        /// Safely gets a long from a JsonElement, handling both number and string values
        /// </summary>
        public static bool TryGetLong(JsonElement element, out long result)
        {
            result = 0;

            switch (element.ValueKind)
            {
                case JsonValueKind.Number:
                    return element.TryGetInt64(out result);

                case JsonValueKind.String:
                    return long.TryParse(
                        element.GetString(),
                        System.Globalization.NumberStyles.Any,
                        System.Globalization.CultureInfo.InvariantCulture,
                        out result);

                default:
                    return false;
            }
        }
    }
#nullable restore
}