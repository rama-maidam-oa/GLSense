using GLSense.Models;
using GLSense.Utilities;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.Json;
using System.Threading.Tasks;

namespace GLSense.Helpers
{
    public static class ApiResponseHelper
    {
        private static readonly string[] KnownPayloadNames = new[]
        {
            "records",
            "record",
            "preferences",
            "data",
            "items",
            "value",
            "values",
            "result"
        };

        // =====================================================
        // PUBLIC ENTRY POINT
        // =====================================================
        public static ApiResult<T> Parse<T>(string rawResponse, JsonSerializerOptions options)
        {
            if (string.IsNullOrWhiteSpace(rawResponse))
                return ApiResult<T>.Failure("Empty response from server.");

            if (!IsLikelyJson(rawResponse))
            {
                LogUtility.LogWarn($"ApiResponseHelper | Non-JSON response: {rawResponse}");
                string apiMessage = string.Empty;

                if (rawResponse.IndexOf("<!doctype html>", StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    apiMessage = "Invalid response format. Received HTML instead of JSON.";
                }
                else if (rawResponse.IndexOf("InternalServerError", StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    apiMessage = "Server encountered an error. Please try again later.";
                }
                else
                {
                    apiMessage = "Received non-JSON response from server.";
                }

                return ApiResult<T>.Failure(apiMessage);
            }

            try
            {
                using var doc = JsonDocument.Parse(rawResponse);
                var root = doc.RootElement;

                // -------------------------------------------------
                // CASE 1: Root is ARRAY → direct payload
                // -------------------------------------------------
                if (root.ValueKind == JsonValueKind.Array)
                {
                    return DeserializePayload<T>(root);
                }

                // -------------------------------------------------
                // CASE 2: Root is OBJECT
                // -------------------------------------------------
                if (root.ValueKind == JsonValueKind.Object)
                {
                    // If contains status → treat as wrapped response
                    if (ContainsProperty(root, "status"))
                    {
                        return HandleWrappedResponse<T>(root);
                    }

                    // No status → assume root is direct payload
                    return DeserializePayload<T>(root);
                }

                return ApiResult<T>.Failure("Unsupported JSON format.");
            }
            catch (JsonException ex)
            {
                LogUtility.LogException(ex, "Invalid JSON received.");
                LogUtility.LogRawJson("ApiResponseHelper.Parse", rawResponse);
                return ApiResult<T>.Failure("Invalid JSON response.");
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex, "Unexpected parsing error.");
                return ApiResult<T>.Failure("Unexpected response format.");
            }
        }

        // =====================================================
        // WRAPPED RESPONSE HANDLER
        // =====================================================
        private static ApiResult<T> HandleWrappedResponse<T>(JsonElement root)
        {
            string status = GetStringSafe(root, "status");
            string message = GetStringSafe(root, "message");

            if (string.IsNullOrWhiteSpace(message))
                message = GetStringSafe(root, "msg");

            bool success = string.Equals(
                status,
                AppConstants.Success,
                StringComparison.OrdinalIgnoreCase);

            if (!success)
            {
                return ApiResult<T>.Failure(
                    string.IsNullOrWhiteSpace(message)
                        ? "Server returned failure status."
                        : message);
            }

            // If T is JsonElement → return whole root
            if (typeof(T) == typeof(JsonElement))
            {
                object clone = root.Clone();
                return ApiResult<T>.Success((T)clone);
            }

            // For collection types, try to get the records specifically
            if (IsCollectionType(typeof(T)))
            {
                // Try to find the "records" property specifically
                var recordsProp = root.EnumerateObject()
                    .FirstOrDefault(p => string.Equals(p.Name, "records", StringComparison.OrdinalIgnoreCase));

                if (!recordsProp.Equals(default(JsonProperty)) && recordsProp.Value.ValueKind == JsonValueKind.Array)
                {
                    return DeserializePayload<T>(recordsProp.Value);
                }
            }

            // First, try deserializing the entire wrapped object to T, using the root
            // exactly as-is (see DeserializeExact for why this must NOT go through
            // GetAppropriatePayload's auto-unwrap detection).
            var wholeResult = DeserializeExact<T>(root);
            if (wholeResult.IsSuccess)
            {
                return wholeResult;
            }

            // Auto-detect payload as fallback
            var payload = DetectPayload(root);
            if (!payload.Equals(root)) // Only if we detected something different
            {
                return DeserializePayload<T>(payload);
            }

            // Final fallback: try to find any array in the response
            foreach (var prop in root.EnumerateObject())
            {
                if (prop.Value.ValueKind == JsonValueKind.Array)
                {
                    var arrayResult = DeserializePayload<T>(prop.Value);
                    if (arrayResult.IsSuccess)
                    {
                        return arrayResult;
                    }
                }
            }

            return ApiResult<T>.Failure($"Could not find suitable payload for type {typeof(T).Name}");
        }

        // =====================================================
        // AUTO PAYLOAD DETECTION
        // =====================================================
        private static JsonElement DetectPayload(JsonElement root)
        {
            // 1️⃣ Known names
            foreach (var name in KnownPayloadNames)
            {
                var prop = root.EnumerateObject()
                    .FirstOrDefault(p =>
                        string.Equals(p.Name, name,
                            StringComparison.OrdinalIgnoreCase));

                if (!prop.Equals(default(JsonProperty)))
                    return prop.Value;
            }

            // 2️⃣ First non-meta object/array
            foreach (var prop in root.EnumerateObject())
            {
                if (IsMetaProperty(prop.Name))
                    continue;

                if (prop.Value.ValueKind == JsonValueKind.Array ||
                    prop.Value.ValueKind == JsonValueKind.Object)
                    return prop.Value;
            }

            // 3️⃣ Fallback → entire root
            return root;
        }

        // =====================================================
        // DESERIALIZATION CORE
        // =====================================================

        // Deserializes the given element exactly as-is, WITHOUT GetAppropriatePayload's
        // "detect a more specific nested payload" auto-unwrapping.
        //
        // Why this exists: HandleWrappedResponse's "first, try the entire wrapped
        // object" step used to call DeserializePayload<T>(root), which internally runs
        // every element through GetAppropriatePayload -> DetectPayload. DetectPayload
        // unconditionally prefers a handful of "known payload names" (records, record,
        // preferences, data, items, value, values, result) if the root object happens
        // to contain one - which is exactly right when the real data is nested one
        // level deeper (e.g. {"status":"success","records":[...]}), but exactly wrong
        // when T's own properties ARE the root's properties (e.g. UserConfigResponse,
        // whose "Preferences" property maps to a root-level "preferences" key that is
        // ALSO in that known-names list). For that case, DetectPayload substituted the
        // *inner* "preferences" object for the root before deserialization ever ran, so
        // JsonSerializer bound it against UserConfigResponse's Message/Status/Preferences
        // shape and matched none of them - producing a non-null UserConfigResponse with
        // every property left null (not an exception, so it read as "success").
        // Deserializing the untouched root here first fixes that without touching the
        // DetectPayload-based fallback chain that other response shapes still rely on.
        private static ApiResult<T> DeserializeExact<T>(JsonElement element)
        {
            try
            {
                var result = JsonSerializer.Deserialize<T>(element.GetRawText(), JsonGlobals.Options);

                if (object.Equals(result, default(T)))
                    return ApiResult<T>.Failure("Failed to deserialize response.");

                return ApiResult<T>.Success(result);
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex, "Exact payload deserialization failed.");
                LogUtility.LogRawJson($"ApiResponseHelper.DeserializeExact<{typeof(T).Name}>", element.GetRawText());
                return ApiResult<T>.Failure("Invalid payload format.");
            }
        }

        private static ApiResult<T> DeserializePayload<T>(JsonElement element)
        {
            try
            {
                JsonElement payloadElement = GetAppropriatePayload<T>(element);

                var result = JsonSerializer.Deserialize<T>(
                    payloadElement.GetRawText(),
                    JsonGlobals.Options);

                if (object.Equals(result, default(T)))
                    return ApiResult<T>.Failure("Failed to deserialize response.");

                return ApiResult<T>.Success(result);
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex, "Payload deserialization failed.");
                LogUtility.LogRawJson($"ApiResponseHelper.DeserializePayload<{typeof(T).Name}>", element.GetRawText());
                return ApiResult<T>.Failure("Invalid payload format.");
            }
        }

        private static JsonElement GetAppropriatePayload<T>(JsonElement element)
        {
            bool isCollection = IsCollectionType(typeof(T));

            // If the element is already the right type, use it
            if ((isCollection && element.ValueKind == JsonValueKind.Array) ||
                (!isCollection && element.ValueKind == JsonValueKind.Object))
            {
                // But check if there's a more specific nested payload
                if (element.ValueKind == JsonValueKind.Object)
                {
                    var detected = DetectPayload(element);
                    if ((isCollection && detected.ValueKind == JsonValueKind.Array) ||
                        (!isCollection && detected.ValueKind == JsonValueKind.Object))
                    {
                        return detected;
                    }
                }
                return element;
            }

            // Try to detect the right payload
            var detectedPayload = DetectPayload(element);

            // If detected payload matches what we expect, use it
            if ((isCollection && detectedPayload.ValueKind == JsonValueKind.Array) ||
                (!isCollection && detectedPayload.ValueKind == JsonValueKind.Object))
            {
                return detectedPayload;
            }

            // For collections, if we found an object but need an array, 
            // check if the object has an array property
            if (isCollection && detectedPayload.ValueKind == JsonValueKind.Object)
            {
                foreach (var prop in detectedPayload.EnumerateObject())
                {
                    if (prop.Value.ValueKind == JsonValueKind.Array)
                    {
                        return prop.Value;
                    }
                }
            }

            // For single objects, if we found an array but need an object,
            // try to use the first element if it's an object
            if (!isCollection && detectedPayload.ValueKind == JsonValueKind.Array)
            {
                var arrayEnumerator = detectedPayload.EnumerateArray();
                if (arrayEnumerator.Any() && arrayEnumerator.First().ValueKind == JsonValueKind.Object)
                {
                    return arrayEnumerator.First();
                }
            }

            // Fallback to the original element
            return element;
        }

        // =====================================================
        // TYPE DETECTION
        // =====================================================
        private static bool IsCollectionType(Type type)
        {
            if (type == typeof(string))
                return false;

            if (type.IsArray)
                return true;

            return typeof(IEnumerable).IsAssignableFrom(type);
        }

        // =====================================================
        // HELPERS
        // =====================================================
        private static bool IsLikelyJson(string input)
        {
            input = input.TrimStart();
            return input.StartsWith("{") || input.StartsWith("[");
        }

        private static bool ContainsProperty(JsonElement element, string name)
        {
            return element.EnumerateObject()
                .Any(p => string.Equals(p.Name,
                                        name,
                                        StringComparison.OrdinalIgnoreCase));
        }

        private static bool IsMetaProperty(string name)
        {
            return string.Equals(name, "status", StringComparison.OrdinalIgnoreCase)
                || string.Equals(name, "msg", StringComparison.OrdinalIgnoreCase)
                || string.Equals(name, "message", StringComparison.OrdinalIgnoreCase)
                || string.Equals(name, "redirectURL", StringComparison.OrdinalIgnoreCase)
                || string.Equals(name, "domain", StringComparison.OrdinalIgnoreCase)
                || string.Equals(name, "contextPath", StringComparison.OrdinalIgnoreCase);
        }

        private static string GetStringSafe(JsonElement element, string propertyName)
        {
            var prop = element.EnumerateObject()
                .FirstOrDefault(p =>
                    string.Equals(p.Name,
                                  propertyName,
                                  StringComparison.OrdinalIgnoreCase));

            if (prop.Equals(default(JsonProperty)))
                return string.Empty;

            return prop.Value.ValueKind == JsonValueKind.String
                ? prop.Value.GetString() ?? string.Empty
                : prop.Value.ToString();
        }
    }
}
