// JsonGlobals.cs in GLSense.Addin.Core
// Port of GLSense\Helpers\JsonHelper.cs (FinalWorkingCode) - only JsonGlobals +
// FlexibleStringConverter are ported here (both fully self-contained, no statics to
// re-point). The static JsonHelper class (TryGetProperty/TryGetDouble/etc.) from the
// same original file is NOT ported - nothing in the currently-migrated code calls it
// yet; add it here (unchanged) the first time some group actually needs it.
using GLSense.Addin.Core.Infrastructure;
using System;
using System.Text;
using System.Text.Encodings.Web;
using System.Text.Json;
using System.Text.Json.Serialization;

namespace GLSense.Addin.Core.Helpers
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
            catch (Exception ex)
            {
                // If conversion fails, try alternative
                ServiceLocator.Logger?.LogDebug($"FlexibleStringConverter.ReadNumberAsString: raw span conversion failed, trying next fallback - {ex.Message}");
            }

            // Ultimate fallback - try to get as string
            try
            {
                return reader.GetString() ?? "0";
            }
            catch (Exception ex)
            {
                ServiceLocator.Logger?.LogWarn($"FlexibleStringConverter.ReadNumberAsString: could not convert JSON number token to string by any known method, defaulting to \"0\" - {ex.Message}");
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
#nullable restore
}
