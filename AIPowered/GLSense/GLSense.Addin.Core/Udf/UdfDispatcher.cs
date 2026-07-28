// UdfDispatcher.cs in GLSense.Addin.Core
// Task #16 (Wire UDFs to ExecuteUdf) - the business-logic side of the 16 GLSense_* Excel
// UDFs. GLSense.Addin.Core.AddinEntry.ExecuteUdf(string functionName, object[] args)
// delegates straight into UdfDispatcher.Execute(functionName, args).
//
// Port of GLSense\GLSenseExcelFunctions.cs (FinalWorkingCode) - specifically:
//   - The shared helper methods (lines ~137-1022 of the old file): FormulaCacheString/
//     BuildFunctionCacheKey/SafeValue/ResolveLedger/PModel/LoadPeriodsForLedger/
//     GetSegmentByDirection + its own helpers/GetBalanceParameters/numeric-extraction
//     helpers/ParseAndCacheBalanceResponse/TryGetRecordsNode/CacheResult/HandleCachedResult/
//     UpdateCache/GetBatchResult/HandleBatchCalculation/ExecuteApiCallWithTimeoutAsync/
//     ExecuteSingleRefreshAsync/GetResultValue/CleanSegmentValue/GetSegmentValueSetId/
//     ParseSegmentDFFResponse/ExtractStringValue/ParseDailyRateResponse/
//     FormatConversionDate/GetPeriods/ParsePeriodDate/FindPeriodName/CalculateOffsetPeriod/
//     GetSegmentString/NormalizeCombinedSegment/SplitCombinedSegments.
//   - The 16 UDF bodies themselves (old lines ~1025-2057), minus everything already handled
//     host-side (ValidateInputs/[ExcelParam] mandatory-parameter checks, GetCellCallerAddress,
//     the ADXExcelAsyncCallObject plumbing for the 3 async UDFs).
//
// Namespace/service re-pointings vs. the original (see the porting brief for the full
// mapping table):
//   - AppState.Instance.*                       -> unchanged (AppState lives in this
//                                                   assembly's root namespace too).
//   - LogUtility.Log*                           -> GLSense.Addin.Core.Infrastructure.
//                                                   ServiceLocator.Logger?.Log*
//   - GLSense.Service.ServiceLocator.*DataService -> GLSense.Addin.Core.Services.
//                                                     DataServiceLocator.*DataService
//   - ApiHelper/ApiResponseHelper/JsonGlobals/CompressionHelper -> already-ported
//     GLSense.Addin.Core.Helpers.* equivalents.
//   - AddinExpress.MSO.ADXExcelError.xlErrorGettingData / xlErrorNull -> the plain string
//     sentinels GLSense.Contracts.UdfSentinels.XlErrorGettingData / XlErrorNull (this
//     project has zero AddinExpress.MSO reference - only the host does).
//   - System.Reflection.Missing can never appear in `args` here - the host's ValidateInputs
//     (reflection-based, tied to the host wrapper method's own [ExcelParam] attributes)
//     already substituted any omitted optional Excel parameter with its default value
//     before crossing the AppDomain boundary, or returned a mandatory-parameter error
//     directly to Excel without even calling ExecuteUdf. The one deliberate exception is the
//     optional "Ledger" parameter (every function except GLSense_GetBalance, whose
//     LedgerName is a distinct, mandatory parameter): the host leaves it unresolved (raw
//     string, "" if omitted) specifically so ResolveLedger here can consult AppState -
//     that's why every handler below calls ResolveLedger(...) as its first step.
using GLSense.Addin.Core.Caching;
using GLSense.Addin.Core.Helpers;
using GLSense.Addin.Core.Infrastructure;
using GLSense.Addin.Core.Models;
using GLSense.Addin.Core.Services;
using GLSense.Contracts;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Globalization;
using System.Linq;
using System.Net.Http;
using System.Text.Json;
using System.Threading;
using System.Threading.Tasks;

namespace GLSense.Addin.Core.Udf
{
    /// <summary>
    /// Entry point AddinEntry.ExecuteUdf delegates to. Fully ADX-free - never references
    /// AddinExpress.MSO or System.Reflection.Missing.
    /// </summary>
    internal static class UdfDispatcher
    {
        private const string GLFormulaDelimiter = "~!~";
        private const string GLNotLoggedIn = "Not Logged In...";
        private const string GLClickToRefresh = "Click Refresh...";
        private const string GLDateFormat = "dd-MM-yyyy";

        internal static object Execute(string functionName, object[] args)
        {
            try
            {
                switch (functionName)
                {
                    case "GLSense_GetPeriodByDate":
                        return HandleGetPeriodByDate(args);
                    case "GLSense_GetPeriod":
                        return HandleGetPeriod(args);
                    case "GLSense_GetPeriodByYear":
                        return HandleGetPeriodByYear(args);
                    case "GLSense_GetPeriodStart":
                        return HandleGetPeriodStart(args);
                    case "GLSense_GetPeriodEnd":
                        return HandleGetPeriodEnd(args);
                    case "GLSense_GetPeriodNum":
                        return HandleGetPeriodNum(args);
                    case "GLSense_GetPeriodQuarter":
                        return HandleGetPeriodQuarter(args);
                    case "GLSense_GetPeriodYear":
                        return HandleGetPeriodYear(args);
                    case "GLSense_GetSegmentDesc":
                        return HandleGetSegmentDesc(args);
                    case "GLSense_GetSegmentEnabledFlag":
                        return HandleGetSegmentEnabledFlag(args);
                    case "GLSense_GetSegmentSummaryFlag":
                        return HandleGetSegmentSummaryFlag(args);
                    case "GLSense_GetNextSegment":
                        return GetSegmentByDirection(args, Direction.Next);
                    case "GLSense_GetPreviousSegment":
                        return GetSegmentByDirection(args, Direction.Previous);
                    case "GLSense_GetSegmentDFF":
                        return HandleGetSegmentDFF(args);
                    case "GLSense_GetAccountType":
                        return HandleGetAccountType(args);
                    case "GLSense_GetDailyRate":
                        return HandleGetDailyRate(args);
                    case "GLSense_GetBalance":
                        return HandleGetBalance(args);
                    default:
                        ServiceLocator.Logger?.LogWarn($"UdfDispatcher.Execute: unknown function '{functionName}'");
                        return UdfSentinels.XlErrorGettingData;
                }
            }
            catch (Exception ex)
            {
                ServiceLocator.Logger?.LogException(ex, $"UdfDispatcher.Execute('{functionName}'): unhandled error");
                return UdfSentinels.XlErrorGettingData;
            }
        }

        #region Shared helpers

        /// <summary>
        /// Resolves the "Ledger" parameter the host deliberately left unresolved: falls back
        /// to the currently-selected ledger when omitted/empty. Must be the first thing every
        /// handler (except GLSense_GetBalance, whose LedgerName is unrelated/mandatory) does -
        /// this must run BEFORE FormulaCacheString so the cache key reflects the resolved
        /// ledger, matching the old monolith's ordering.
        /// </summary>
        internal static string ResolveLedger(object ledgerRaw)
        {
            string ledger = ledgerRaw as string ?? ledgerRaw?.ToString();
            return string.IsNullOrEmpty(ledger)
                ? (AppState.Instance.SelectedLedger?.LedgerName ?? string.Empty)
                : ledger;
        }

        private static object HandleNotLoggedIn(string formulaKey)
        {
            try
            {
                if (FormulaCacheManager.Instance.TryGetValue(formulaKey, out var entry))
                {
                    return entry.Value;
                }
                return GLNotLoggedIn;
            }
            catch (Exception ex)
            {
                ServiceLocator.Logger?.LogException(ex, $"HandleNotLoggedIn(formulaKey='{formulaKey}')");
                return GLNotLoggedIn;
            }
        }

        private static string FormulaCacheString(object[] args)
        {
            var paramArgs = args.Select(SafeValue).ToArray();
            return BuildFunctionCacheKey(paramArgs);
        }

        private static string BuildFunctionCacheKey(params object[] parameters)
        {
            var key = string.Join(GLFormulaDelimiter, parameters.Select(SafeValue));
            return CompressionHelper.CompressString(key);
        }

        private static string SafeValue(object value)
        {
            if (value == null) return string.Empty;
            return value.ToString().Trim();
        }

        /// <summary>Local re-implementation of the boolean coercion the host's own ToBool
        /// performs - needed here only for GetFilter's NextParent/NextChild args, which cross
        /// the boundary raw (see GetSegmentByDirection). No ADX/Missing awareness needed -
        /// those never reach this side.</summary>
        private static bool ToBool(object value, bool defaultValue = false, double tolerance = 1e-6)
        {
            if (value == null) return defaultValue;
            if (value is bool b) return b;

            if (value is double d)
            {
                if (Math.Abs(d - 1.0) < tolerance) return true;
                if (Math.Abs(d - 0.0) < tolerance) return false;
                return defaultValue;
            }

            if (value is string s)
            {
                if (string.IsNullOrWhiteSpace(s)) return defaultValue;

                s = s.Trim();
                if (s.Equals("TRUE", StringComparison.OrdinalIgnoreCase)) return true;
                if (s.Equals("FALSE", StringComparison.OrdinalIgnoreCase)) return false;
                if (s.Equals("1")) return true;
                if (s.Equals("0")) return false;
            }

            return bool.TryParse(value.ToString(), out var parsed) ? parsed : defaultValue;
        }

        private static List<PeriodModel> PModel(string lName)
        {
            try
            {
                List<PeriodModel> periods = LoadPeriodsForLedger(lName);
                return periods ?? new List<PeriodModel>();
            }
            catch (Exception ex)
            {
                ServiceLocator.Logger?.LogException(ex, $"PModel(lName='{lName}')");
                return new List<PeriodModel>();
            }
        }

        private static List<PeriodModel> LoadPeriodsForLedger(string ledgerName)
        {
            try
            {
                var dataService = DataServiceLocator.PeriodDataService;
                return new List<PeriodModel>(dataService.GetPeriodsForLedger(ledgerName, allowRemoteFetch: false));
            }
            catch (Exception ex)
            {
                ServiceLocator.Logger?.LogException(ex, $"Failed to load periods for ledger '{ledgerName}'");
                return new List<PeriodModel>();
            }
        }

        private enum Direction { Next, Previous }

        // args = [SegmentValue, SegmentName, NextParent, NextChild, Ledger]
        private static object GetSegmentByDirection(object[] args, Direction direction)
        {
            try
            {
                string ledgerValue = ResolveLedger(args[4]);
                object[] parameters = new object[] { args[0], args[1], args[2], args[3], ledgerValue };

                string directionFuncName = direction == Direction.Next ? "GLSense_GetNextSegment" : "GLSense_GetPreviousSegment";
                ServiceLocator.Logger?.LogDebug($"{directionFuncName}: segmentValue='{parameters[0]}' segmentName='{parameters[1]}' nextParent='{parameters[2]}' nextChild='{parameters[3]}' ledger='{ledgerValue}'");

                string formulaCompressed = FormulaCacheString(parameters);

                if (!AppState.Instance.IsLoginCompleted)
                    return GetCachedFormulaResultOrError(formulaCompressed);

                string resolvedSegmentName = ResolveSegmentName(parameters[1], ledgerValue);
                if (string.IsNullOrEmpty(resolvedSegmentName))
                    return LogAndReturnError($"Unable to resolve segment name from '{parameters[1]}'", direction, formulaCompressed);

                var segmentsForName = GetFilteredSegments(parameters);
                if (!segmentsForName.Any())
                    return LogAndReturnError($"No segments found for '{resolvedSegmentName}'", direction, formulaCompressed);

                int currentIndex = FindCurrentSegmentIndex(segmentsForName, parameters[0].ToString().Trim());
                if (currentIndex == -1)
                    return LogAndReturnError($"Segment '{parameters[0].ToString().Trim()}' not found", direction, formulaCompressed);

                string output = FindMatchingSegment(segmentsForName, currentIndex, GetFilter(parameters), direction);

                ServiceLocator.Logger?.LogDebug($"{directionFuncName}: resolved output='{output}'");
                UpdateFormulaCache(formulaCompressed, output);
                return output;
            }
            catch (Exception ex)
            {
                string methodName = direction == Direction.Next ? "GLSense_GetNextSegment" : "GLSense_GetPreviousSegment";
                ServiceLocator.Logger?.LogException(ex, $"{methodName} Function : Unexpected error");
                return UdfSentinels.XlErrorGettingData;
            }
        }

        private static void UpdateFormulaCache(string formulaCompressed, object output)
        {
            try
            {
                CachedFormulaHelper.Store(formulaCompressed, output, DateTime.UtcNow);
            }
            catch (Exception ex)
            {
                ServiceLocator.Logger?.LogException(ex, "Failed to update formula cache");
            }
        }

        private static object GetCachedFormulaResultOrError(string formulaCompressed)
        {
            try
            {
                if (FormulaCacheManager.Instance.TryGetValue(formulaCompressed, out var entry))
                {
                    return GetResultValue(entry.Value);
                }

                return UdfSentinels.XlErrorGettingData;
            }
            catch (Exception ex)
            {
                ServiceLocator.Logger?.LogException(ex, $"Formula: {CompressionHelper.DecompressString(formulaCompressed)}");
                return UdfSentinels.XlErrorGettingData;
            }
        }

        private static string ResolveSegmentName(object segmentNameObj, string ledgerName) =>
            DataServiceLocator.SegmentDataService.ResolveSegmentName(segmentNameObj, ledgerName);

        private static List<SegmentValueModel> GetFilteredSegments(object[] parameters)
        {
            string ledgerName = parameters[parameters.Length - 1].ToString();
            var segmentValues = LoadSegmentValues(ledgerName);
            if (segmentValues == null || !segmentValues.Any())
                return new List<SegmentValueModel>();

            string resolvedSegmentName = ResolveSegmentName(parameters[1], ledgerName);
            return segmentValues
                .Where(sv => sv.SegmentName.Equals(resolvedSegmentName, StringComparison.OrdinalIgnoreCase))
                .OrderBy(sv => sv.SegmentValue)
                .ToList();
        }

        private static ObservableCollection<SegmentValueModel> LoadSegmentValues(string ledgerName)
        {
            try
            {
                var task = Task.Run(() => DataServiceLocator.SegmentDataService.GetSegmentValues(ledgerName));

                if (task.Wait(TimeSpan.FromSeconds(180)))
                {
                    return task.Result;
                }
                throw new TimeoutException("Timeout loading segment values from service");
            }
            catch (Exception ex)
            {
                ServiceLocator.Logger?.LogException(ex, "Failed to load segment values");
                return new ObservableCollection<SegmentValueModel>();
            }
        }

        private static int FindCurrentSegmentIndex(List<SegmentValueModel> segments, string segmentValueStr)
        {
            var currentSegment = segments.FirstOrDefault(sv =>
                sv.SegmentValue.Equals(segmentValueStr, StringComparison.OrdinalIgnoreCase));
            return currentSegment != null ? segments.IndexOf(currentSegment) : -1;
        }

        private static Func<SegmentValueModel, bool> GetFilter(object[] parameters)
        {
            bool nextParent = ToBool(parameters[2], false);
            bool nextChild = ToBool(parameters[3], false);

            if (nextParent && !nextChild)
                return sv => sv.SummaryFlag == "Y";
            if (!nextParent && nextChild)
                return sv => sv.SummaryFlag != "Y";
            return sv => true;
        }

        private static string FindMatchingSegment(List<SegmentValueModel> segments, int currentIndex,
            Func<SegmentValueModel, bool> filter, Direction direction)
        {
            IEnumerable<SegmentValueModel> candidateSegments = direction == Direction.Next
                ? segments.Skip(currentIndex + 1)
                : segments.Take(currentIndex).Reverse();

            var matchingSegment = candidateSegments.FirstOrDefault(filter);
            return matchingSegment?.SegmentValue ?? UdfSentinels.XlErrorGettingData;
        }

        private static object LogAndReturnError(string message, Direction direction, string formulaCompressed)
        {
            string methodName = direction == Direction.Next ? "GLSense_GetNextSegment" : "GLSense_GetPreviousSegment";
            ServiceLocator.Logger?.LogDebug($"{methodName} Function : {message}");
            return GetCachedFormulaResultOrError(formulaCompressed);
        }

        private static List<string> GetBalanceParameters(object[] parameters)
        {
            return parameters.Select(p => p?.ToString() ?? string.Empty).ToList();
        }

        private static object ExtractNumericValue(JsonElement element)
        {
            return element.ValueKind switch
            {
                JsonValueKind.Number => GetNumberValue(element),
                JsonValueKind.String => ParseStringNumber(element.GetString()),
                _ => element.GetRawText()
            };
        }

        private static object GetNumberValue(JsonElement element)
        {
            if (element.TryGetDecimal(out decimal dec)) return dec;
            if (element.TryGetDouble(out double dbl)) return dbl;
            return element.GetInt64();
        }

        private static object ParseStringNumber(string numberString)
        {
            return double.TryParse(numberString, NumberStyles.Any, CultureInfo.InvariantCulture, out double parsed)
                ? parsed
                : numberString;
        }

        private static object ParseAndCacheBalanceResponse(string apiResponse, string formulaCompressed)
        {
            if (string.IsNullOrWhiteSpace(apiResponse))
                return "#Err: Empty response";

            var result = ApiResponseHelper.Parse<JsonElement>(apiResponse, JsonGlobals.Options);

            if (!result.IsSuccess)
                return "#Err: Response(Status: Fail)";

            JsonElement root = result.Value;

            if (!TryGetRecordsNode(root, out JsonElement recordsElem))
                recordsElem = root;

            object value = ExtractNumericValue(recordsElem);

            CacheResult(formulaCompressed, value);

            if (value is double or decimal or int or long)
            {
                return value switch
                {
                    double d => d,
                    decimal m => (double)m,
                    int i => i,
                    long l => l,
                    _ => "#Err: Invalid numeric value"
                };
            }
            return value;
        }

        private static bool TryGetRecordsNode(JsonElement root, out JsonElement recordsNode)
        {
            var recordProp = root.EnumerateObject()
                .FirstOrDefault(prop => string.Equals(prop.Name, "records", StringComparison.OrdinalIgnoreCase));

            if (recordProp.Value.ValueKind != JsonValueKind.Undefined)
            {
                recordsNode = recordProp.Value;
                return true;
            }

            recordsNode = default;
            return false;
        }

        private static void CacheResult(string formulaCompressed, object result)
        {
            string cacheValue = result switch
            {
                double d => d.ToString(CultureInfo.InvariantCulture),
                decimal m => ((double)m).ToString(CultureInfo.InvariantCulture),
                int i => i.ToString(CultureInfo.InvariantCulture),
                long l => l.ToString(CultureInfo.InvariantCulture),
                _ => string.Empty
            };

            CachedFormulaHelper.Store(formulaCompressed, cacheValue, DateTime.UtcNow);
        }

        private static object HandleCachedResult(string formulaCompressed)
        {
            try
            {
                ServiceLocator.Logger?.LogDebug($"Formula: {CompressionHelper.DecompressString(formulaCompressed)}");

                if (FormulaCacheManager.Instance.ContainsKey(formulaCompressed))
                {
                    return GetResultValue(FormulaCacheManager.Instance.GetValue(formulaCompressed));
                }

                return GLClickToRefresh;
            }
            catch (Exception ex)
            {
                ServiceLocator.Logger?.LogException(ex, $"Formula: {CompressionHelper.DecompressString(formulaCompressed)}");
                return GLClickToRefresh;
            }
        }

        private static void UpdateCache(string formulaCompressed, object result)
        {
            string resultString = result?.ToString() ?? string.Empty;
            CachedFormulaHelper.Store(formulaCompressed, resultString, DateTime.UtcNow);
        }

        private static object GetBatchResult(string formulaCompressed)
        {
            if (AppState.Instance.PreComputedBalances?.Count == 0)
                return GLClickToRefresh;

            if (AppState.Instance.PreComputedBalances?.TryGetValue(formulaCompressed, out var cachedValue) == true)
            {
                return cachedValue;
            }

            return GLClickToRefresh;
        }

        private static object HandleBatchCalculation(string formulaCompressed)
        {
            try
            {
                object result = GetBatchResult(formulaCompressed);
                UpdateCache(formulaCompressed, result);
                return result;
            }
            catch (Exception ex)
            {
                ServiceLocator.Logger?.LogException(ex, $"HandleBatchCalculation(formulaCompressed='{formulaCompressed}')");
                return GLClickToRefresh;
            }
        }

        private static async Task<string> ExecuteApiCallWithTimeoutAsync(string apiUrl, string jsonPayload)
        {
            using var cts = new CancellationTokenSource(TimeSpan.FromSeconds(300));

            try
            {
                string response = await ApiHelper.ServerAPI(apiUrl, "JSON", jsonPayload, "POST", cts.Token);
                return response;
            }
            catch (OperationCanceledException)
            {
                ServiceLocator.Logger?.LogWarn($"API call timed out after 300s: {apiUrl}");
                return null;
            }
            catch (HttpRequestException ex)
            {
                ServiceLocator.Logger?.LogError($"HTTP/network error: {ex.Message}");
                return null;
            }
            catch (Exception ex)
            {
                ServiceLocator.Logger?.LogException(ex, "Unexpected error in API call");
                return null;
            }
        }

        private static async Task<object> ExecuteSingleRefreshAsync(
            string formulaCompressed,
            List<string> balanceParameters,
            string xllR1C1)
        {
            try
            {
                var balanceDto = BalanceDto.CreateFromXllParameters(xllR1C1, balanceParameters);
                string jsonPayload = JsonSerializer.Serialize(balanceDto, JsonGlobals.Options);
                string apiUrl = $"{AppState.Instance.LoginUrl}{AppConstants.RestSecure}balance?cubeId={AppState.Instance.SelectedCube.CubeId}";

                string apiResponse = await ExecuteApiCallWithTimeoutAsync(apiUrl, jsonPayload);

                return apiResponse != null
                    ? ParseAndCacheBalanceResponse(apiResponse, formulaCompressed)
                    : GLClickToRefresh;
            }
            catch (Exception ex)
            {
                ServiceLocator.Logger?.LogException(ex, $"ExecuteSingleRefreshAsync(formulaCompressed='{formulaCompressed}')");
                return GLClickToRefresh;
            }
        }

        private static object GetResultValue(object obj)
        {
            try
            {
                if (obj is CachedFormulaEntry cachedEntry)
                {
                    obj = cachedEntry.Value;
                }

                if (obj == null || string.IsNullOrEmpty(obj.ToString()))
                    return obj ?? "#Null Exception";

                if (obj is double d)
                {
                    return d;
                }

                if (double.TryParse(obj.ToString().Trim(), NumberStyles.Any, CultureInfo.InvariantCulture, out double numericValue))
                {
                    return numericValue;
                }

                return obj;
            }
            catch (Exception ex)
            {
                ServiceLocator.Logger?.LogException(ex, $"GetResultValue|{obj}");
                return obj;
            }
        }

        private static string CleanSegmentValue(string segmentValue)
        {
            if (string.IsNullOrEmpty(segmentValue)) return segmentValue;

            if (segmentValue.StartsWith("--"))
            {
                segmentValue = segmentValue.Substring(2);
            }
            if (segmentValue.Contains("~"))
            {
                segmentValue = segmentValue.Replace("~", "").Trim();
            }

            return segmentValue.Trim();
        }

        private static long GetSegmentValueSetId(string segmentName, string ledgerName)
        {
            try
            {
                var segments = DataServiceLocator.SegmentDataService.GetSegments(ledgerName);
                var segment = segments?.FirstOrDefault(s =>
                    s.SegmentName.Equals(segmentName, StringComparison.OrdinalIgnoreCase));

                return segment?.SegmentValueSetId ?? -1;
            }
            catch (Exception ex)
            {
                ServiceLocator.Logger?.LogException(ex, $"Failed to get segment value set ID for '{segmentName}'");
                return -1;
            }
        }

        // GLSense_GetAccountType now receives the segment's 1-based dropdown position
        // directly from the formula, so the ApplicationColumnName-regex lookup this used to
        // require (GetSegmentSequenceIndex, removed) is no longer needed - just parse the
        // number. Excel passes an un-quoted numeric literal as a boxed double, so this
        // accepts double/int/numeric-string defensively.
        private static bool TryParseSegmentIndex(object value, out int index)
        {
            index = -1;
            switch (value)
            {
                case null:
                    return false;
                case double d:
                    index = (int)Math.Round(d);
                    return true;
                case int i:
                    index = i;
                    return true;
                default:
                    return int.TryParse(value.ToString().Trim(), NumberStyles.Any, CultureInfo.InvariantCulture, out index);
            }
        }

        private static string ParseSegmentsApiResponse(string apiResponse)
        {
            if (string.IsNullOrWhiteSpace(apiResponse))
                return UdfSentinels.XlErrorGettingData;

            ServiceLocator.Logger?.LogDebug(apiResponse);

            var result = ApiResponseHelper.Parse<JsonElement>(apiResponse, JsonGlobals.Options);

            if (!result.IsSuccess)
                return string.IsNullOrEmpty(result.ErrorMessage) ? UdfSentinels.XlErrorGettingData : result.ErrorMessage;

            JsonElement root = result.Value;

            if (!TryGetRecordsNode(root, out JsonElement recordsElem))
                recordsElem = root;

            string output = ExtractStringValue(recordsElem);

            return string.IsNullOrWhiteSpace(output) ? "#Empty Response" : output;
        }

        private static string ExtractStringValue(JsonElement element)
        {
            switch (element.ValueKind)
            {
                case JsonValueKind.String:
                    return element.GetString();

                case JsonValueKind.Number:
                    return element.ToString(); // preserves precision

                case JsonValueKind.True:
                case JsonValueKind.False:
                    return element.GetBoolean().ToString();

                case JsonValueKind.Array:
                    if (element.GetArrayLength() == 0)
                        return null;
                    return ExtractStringValue(element[0]);

                case JsonValueKind.Object:
                    foreach (var prop in element.EnumerateObject())
                    {
                        var val = prop.Value;
                        if (val.ValueKind == JsonValueKind.String ||
                            val.ValueKind == JsonValueKind.Number ||
                            val.ValueKind == JsonValueKind.True ||
                            val.ValueKind == JsonValueKind.False)
                        {
                            return val.ToString();
                        }

                        var nested = ExtractStringValue(val);
                        if (!string.IsNullOrWhiteSpace(nested))
                            return nested;
                    }
                    return null;

                case JsonValueKind.Null:
                case JsonValueKind.Undefined:
                default:
                    return null;
            }
        }

        private static object ParseDailyRateResponse(string apiResponse)
        {
            try
            {
                if (string.IsNullOrWhiteSpace(apiResponse))
                    return UdfSentinels.XlErrorGettingData;

                if (!apiResponse.TrimStart().StartsWith("{") &&
                    !apiResponse.TrimStart().StartsWith("["))
                {
                    ServiceLocator.Logger?.LogWarn($"DailyRate API returned non-JSON: {apiResponse}");
                    return apiResponse;
                }

                var result = ApiResponseHelper.Parse<List<DailyRateRecord>>(apiResponse, JsonGlobals.Options);

                if (!result.IsSuccess)
                {
                    ServiceLocator.Logger?.LogWarn($"DailyRate parse failed: {result.ErrorMessage}");
                    return string.IsNullOrEmpty(result.ErrorMessage) ? UdfSentinels.XlErrorGettingData : (object)result.ErrorMessage;
                }

                var records = result.Value;

                if (records == null || records.Count == 0)
                    return UdfSentinels.XlErrorGettingData;

                var first = records[0];

                if (first?.CONVERSION_RATE is double rate)
                    return rate;

                return UdfSentinels.XlErrorGettingData;
            }
            catch (Exception ex)
            {
                ServiceLocator.Logger?.LogException(ex, "Failed to parse DailyRate API response.");
                return UdfSentinels.XlErrorGettingData;
            }
        }

        private static string FormatConversionDate(object conversionDate)
        {
            try
            {
                if (conversionDate == null)
                    return string.Empty;

                if (conversionDate is double oaDate)
                {
                    return DateTime.FromOADate(oaDate).ToString(GLDateFormat);
                }

                string dateString = conversionDate.ToString();

                if (DateTime.TryParseExact(dateString, GLDateFormat, CultureInfo.InvariantCulture,
                    DateTimeStyles.None, out DateTime _))
                {
                    return dateString;
                }

                if (DateTime.TryParse(dateString, CultureInfo.InvariantCulture, DateTimeStyles.None,
                    out DateTime parsedDate))
                {
                    return parsedDate.ToString(GLDateFormat, CultureInfo.InvariantCulture);
                }

                string[] excelFormats = new[] {
                    "MM/dd/yyyy", "M/d/yyyy", "dd/MM/yyyy", "d/M/yyyy",
                    "yyyy-MM-dd", "MM-dd-yyyy", GLDateFormat
                };

                if (DateTime.TryParseExact(dateString, excelFormats,
                    CultureInfo.InvariantCulture, DateTimeStyles.None, out DateTime parsedDate3))
                {
                    return parsedDate3.ToString(GLDateFormat);
                }

                ServiceLocator.Logger?.LogDebug($"Date parsing failed for: {dateString}");
                return dateString;
            }
            catch (Exception ex)
            {
                ServiceLocator.Logger?.LogException(ex, $"Error in FormatConversionDate for: {conversionDate}");
                return conversionDate?.ToString() ?? string.Empty;
            }
        }

        private static List<PeriodModel> GetPeriods(string ledgerName)
        {
            return PModel(ledgerName) ?? new List<PeriodModel>();
        }

        private static DateTime ParsePeriodDate(object dateObj)
        {
            try
            {
                if (dateObj == null)
                    return default;

                if (dateObj is double pDate)
                    return DateTime.FromOADate(pDate);

                string dateString = dateObj.ToString()?.Trim() ?? string.Empty;
                if (string.IsNullOrWhiteSpace(dateString))
                    return default;

                if (dateString.Length > 0 && dateString[0] == '\'')
                    dateString = dateString.Substring(1).Trim();

                if (double.TryParse(dateString, NumberStyles.Any, CultureInfo.InvariantCulture, out double oaDate))
                {
                    try
                    {
                        return DateTime.FromOADate(oaDate);
                    }
                    catch (Exception oaEx)
                    {
                        // Continue to string parsing below - this is an expected fallback
                        // (the numeric value wasn't a valid OADate), not a hard failure, so
                        // log at Debug rather than as a full exception dump.
                        ServiceLocator.Logger?.LogDebug($"ParsePeriodDate: '{dateString}' failed FromOADate parse ({oaEx.Message}); falling back to string parsing.");
                    }
                }

                if (DateTime.TryParse(dateString, CultureInfo.InvariantCulture, DateTimeStyles.AllowWhiteSpaces, out DateTime parsedDate))
                    return parsedDate;

                string[] formats = new[]
                {
                    "dd-MMM-yyyy", "d-MMM-yyyy", "dd/MM/yyyy", "d/M/yyyy",
                    "MM/dd/yyyy", "M/d/yyyy", "dd-MM-yyyy", "d-M-yyyy",
                    "yyyy-MM-dd", "yyyy/MM/dd", "dd MMM yyyy", "d MMM yyyy",
                    "dd.MM.yyyy", "d.M.yyyy"
                };

                if (DateTime.TryParseExact(dateString, formats, CultureInfo.InvariantCulture, DateTimeStyles.AllowWhiteSpaces, out parsedDate))
                    return parsedDate;

                return Convert.ToDateTime(dateString, CultureInfo.CurrentCulture);
            }
            catch (Exception ex)
            {
                ServiceLocator.Logger?.LogException(ex, $"ParsePeriodDate: Error converting date {dateObj}");
                return default;
            }
        }

        private static string FindPeriodName(IEnumerable<PeriodModel> periods, DateTime date)
        {
            return periods.FirstOrDefault(p => p.StartDate <= date && p.EndDate >= date)?.PeriodName ?? string.Empty;
        }

        private static string CalculateOffsetPeriod(List<PeriodModel> periods, string basePeriod, int offset)
        {
            try
            {
                var orderedPeriods = periods.OrderBy(p => p.StartDate).ToList();
                int baseIndex = orderedPeriods.FindIndex(p => p.PeriodName == basePeriod);
                int newIndex = baseIndex + offset;

                return newIndex >= 0 && newIndex < orderedPeriods.Count
                    ? orderedPeriods[newIndex].PeriodName ?? UdfSentinels.XlErrorGettingData
                    : UdfSentinels.XlErrorGettingData;
            }
            catch (Exception ex)
            {
                ServiceLocator.Logger?.LogException(ex, $"CalculateOffsetPeriod(basePeriod='{basePeriod}', offset={offset})");
                return UdfSentinels.XlErrorGettingData;
            }
        }

        /// <summary>Null/bounds-safe args[index]?.ToString() - used purely for entry-log
        /// summaries so a short/malformed args array can never throw while just trying to
        /// log diagnostics.</summary>
        private static string SafeArg(object[] args, int index)
        {
            if (args == null || index < 0 || index >= args.Length || args[index] == null)
                return string.Empty;
            return args[index].ToString();
        }

        private static string GetSegmentString(object seg)
        {
            return seg switch
            {
                null => string.Empty,
                "" => string.Empty,
                _ => seg.ToString() ?? string.Empty
            };
        }

        private static string NormalizeCombinedSegment(string segmentPiece)
        {
            if (string.IsNullOrWhiteSpace(segmentPiece))
                return string.Empty;

            string normalized = segmentPiece.Trim().Trim('"');
            return CleanSegmentValue(normalized);
        }

        private static string[] SplitCombinedSegments(string combinedSegment)
        {
            return combinedSegment
                .Split(new[] { ';' }, StringSplitOptions.None)
                .Select(NormalizeCombinedSegment)
                .ToArray();
        }

        #endregion

        #region UDF handlers

        // args = [PeriodDate, Ledger, offset]
        private static object HandleGetPeriodByDate(object[] args)
        {
            try
            {
                string ledgerValue = ResolveLedger(args[1]);
                object[] parameters = new object[] { args[0], ledgerValue, args[2] ?? 0 };

                ServiceLocator.Logger?.LogDebug($"GLSense_GetPeriodByDate: periodDate='{parameters[0]}' ledger='{ledgerValue}' offset='{parameters[2]}'");

                string formulaCompressed = FormulaCacheString(parameters);

                if (!AppState.Instance.IsLoginCompleted)
                    return GetCachedFormulaResultOrError(formulaCompressed);

                var periods = GetPeriods(ledgerValue);
                if (periods == null || !periods.Any())
                {
                    ServiceLocator.Logger?.LogDebug($"GLSense_GetPeriodByDate: no periods loaded for ledger '{ledgerValue}' - returning error sentinel.");
                    return UdfSentinels.XlErrorGettingData;
                }

                DateTime periodDate = ParsePeriodDate(parameters[0]);
                if (periodDate == default)
                {
                    ServiceLocator.Logger?.LogDebug($"GLSense_GetPeriodByDate: could not parse periodDate '{parameters[0]}' - returning error sentinel.");
                    return UdfSentinels.XlErrorGettingData;
                }

                string selectedPeriod = FindPeriodName(periods, periodDate);
                if (string.IsNullOrEmpty(selectedPeriod))
                {
                    ServiceLocator.Logger?.LogDebug($"GLSense_GetPeriodByDate: no period found containing date '{periodDate}' for ledger '{ledgerValue}'.");
                    return GetCachedFormulaResultOrError(formulaCompressed);
                }

                string output = CalculateOffsetPeriod(periods, selectedPeriod, Convert.ToInt32(parameters[2]));
                ServiceLocator.Logger?.LogDebug($"GLSense_GetPeriodByDate: resolved output='{output}' (basePeriod='{selectedPeriod}')");
                UpdateFormulaCache(formulaCompressed, output);
                return output;
            }
            catch (Exception ex)
            {
                ServiceLocator.Logger?.LogException(ex, "GLSense_GetPeriodByDate Function : Unexpected error");
                return UdfSentinels.XlErrorGettingData;
            }
        }

        // args = [Period, offset, Ledger]
        private static object HandleGetPeriod(object[] args)
        {
            try
            {
                string ledgerValue = ResolveLedger(args[2]);
                object[] parameters = new object[] { args[0], args[1] ?? 0, ledgerValue };

                ServiceLocator.Logger?.LogDebug($"GLSense_GetPeriod: period='{parameters[0]}' offset='{parameters[1]}' ledger='{ledgerValue}'");

                string formulaCompressed = FormulaCacheString(parameters);

                if (!AppState.Instance.IsLoginCompleted)
                    return GetCachedFormulaResultOrError(formulaCompressed);

                var periods = PModel(ledgerValue);
                if (periods == null || !periods.Any())
                    return GetCachedFormulaResultOrError(formulaCompressed);

                string periodName = parameters[0].ToString();
                int selectedOffset = Convert.ToInt32(parameters[1]);

                var basePeriod = periods.FirstOrDefault(p => p.PeriodName == periodName);
                if (basePeriod == null)
                {
                    ServiceLocator.Logger?.LogDebug($"GLSense_GetPeriod Function : Period '{periodName}' not found for ledger \"{ledgerValue}\"");
                    return GetCachedFormulaResultOrError(formulaCompressed);
                }

                string output;
                try
                {
                    var orderedPeriods = periods.OrderBy(p => p.StartDate).ToList();
                    int baseIndex = orderedPeriods.FindIndex(p => p.PeriodName == periodName);
                    int newIndex = baseIndex + selectedOffset;

                    output = newIndex >= 0 && newIndex < orderedPeriods.Count
                        ? orderedPeriods[newIndex].PeriodName ?? UdfSentinels.XlErrorGettingData
                        : UdfSentinels.XlErrorGettingData;
                }
                catch (Exception ex)
                {
                    ServiceLocator.Logger?.LogException(ex, $"GLSense_GetPeriod: offset calculation failed for periodName='{periodName}', offset={selectedOffset}");
                    output = UdfSentinels.XlErrorGettingData;
                }

                ServiceLocator.Logger?.LogDebug($"GLSense_GetPeriod: resolved output='{output}'");
                UpdateFormulaCache(formulaCompressed, output);
                return output;
            }
            catch (Exception ex)
            {
                ServiceLocator.Logger?.LogException(ex, "GLSense_GetPeriod Function : Unexpected error");
                return UdfSentinels.XlErrorGettingData;
            }
        }

        // args = [PeriodYear, PeriodNum, Ledger]
        private static object HandleGetPeriodByYear(object[] args)
        {
            try
            {
                string ledgerValue = ResolveLedger(args[2]);
                object[] parameters = new object[] { args[0], args[1], ledgerValue };

                ServiceLocator.Logger?.LogDebug($"GLSense_GetPeriodByYear: periodYear='{parameters[0]}' periodNum='{parameters[1]}' ledger='{ledgerValue}'");

                string formulaCompressed = FormulaCacheString(parameters);

                if (!AppState.Instance.IsLoginCompleted)
                    return GetCachedFormulaResultOrError(formulaCompressed);

                var periods = PModel(ledgerValue);
                if (periods == null || !periods.Any())
                    return GetCachedFormulaResultOrError(formulaCompressed);

                int periodYear = Convert.ToInt32(parameters[0]);
                int periodNum = Convert.ToInt32(parameters[1]);

                var period = periods.FirstOrDefault(p => p.PeriodYear == periodYear && p.PeriodNum == periodNum);
                string output = period?.PeriodName ?? UdfSentinels.XlErrorGettingData;

                ServiceLocator.Logger?.LogDebug($"GLSense_GetPeriodByYear: resolved output='{output}'");
                UpdateFormulaCache(formulaCompressed, output);
                return output;
            }
            catch (Exception ex)
            {
                ServiceLocator.Logger?.LogException(ex, "GLSense_GetPeriodByYear Function : Unexpected error");
                return UdfSentinels.XlErrorGettingData;
            }
        }

        // args = [Period, Ledger, adjacentPeriodsBool]
        private static object HandleGetPeriodStart(object[] args)
        {
            try
            {
                string ledgerValue = ResolveLedger(args[1]);
                object[] parameters = new object[] { args[0], ledgerValue, args[2] };

                ServiceLocator.Logger?.LogDebug($"GLSense_GetPeriodStart: period='{parameters[0]}' ledger='{ledgerValue}' includeAdjacent='{parameters[2]}'");

                string formulaCompressed = FormulaCacheString(parameters);

                if (!AppState.Instance.IsLoginCompleted)
                    return GetCachedFormulaResultOrError(formulaCompressed);

                var periods = PModel(ledgerValue);
                if (periods == null || !periods.Any())
                    return GetCachedFormulaResultOrError(formulaCompressed);

                string periodName = parameters[0].ToString();

                var basePeriod = periods.FirstOrDefault(p => p.PeriodName == periodName);
                if (basePeriod == null)
                {
                    ServiceLocator.Logger?.LogDebug($"GLSense_GetPeriodStart Function : Period '{periodName}' not found for ledger \"{ledgerValue}\"");
                    return GetCachedFormulaResultOrError(formulaCompressed);
                }

                int targetYear = basePeriod.PeriodYear;
                var yearPeriods = periods.Where(p => p.PeriodYear == targetYear);

                if (!ToBool(parameters[2], false))
                {
                    yearPeriods = yearPeriods.Where(p => p.AdjustmentPeriodFlag != "Y");
                }

                var firstPeriod = yearPeriods.OrderBy(p => p.PeriodNum).FirstOrDefault();
                string output = firstPeriod?.PeriodName ?? UdfSentinels.XlErrorGettingData;

                ServiceLocator.Logger?.LogDebug($"GLSense_GetPeriodStart: resolved output='{output}'");
                UpdateFormulaCache(formulaCompressed, output);
                return output;
            }
            catch (Exception ex)
            {
                ServiceLocator.Logger?.LogException(ex, "GLSense_GetPeriodStart Function : Unexpected error");
                return UdfSentinels.XlErrorGettingData;
            }
        }

        // args = [Period, Ledger, adjacentPeriodsBool]
        private static object HandleGetPeriodEnd(object[] args)
        {
            try
            {
                string ledgerValue = ResolveLedger(args[1]);
                object[] parameters = new object[] { args[0], ledgerValue, args[2] };

                ServiceLocator.Logger?.LogDebug($"GLSense_GetPeriodEnd: period='{parameters[0]}' ledger='{ledgerValue}' includeAdjacent='{parameters[2]}'");

                string formulaCompressed = FormulaCacheString(parameters);

                if (!AppState.Instance.IsLoginCompleted)
                    return GetCachedFormulaResultOrError(formulaCompressed);

                var periods = PModel(ledgerValue);
                if (periods == null || !periods.Any())
                    return GetCachedFormulaResultOrError(formulaCompressed);

                string periodName = parameters[0].ToString();

                var basePeriod = periods.FirstOrDefault(p => p.PeriodName == periodName);
                if (basePeriod == null)
                {
                    ServiceLocator.Logger?.LogDebug($"GLSense_GetPeriodEnd Function : Period '{periodName}' not found for ledger \"{ledgerValue}\"");
                    return GetCachedFormulaResultOrError(formulaCompressed);
                }

                int targetYear = basePeriod.PeriodYear;
                var yearPeriods = periods.Where(p => p.PeriodYear == targetYear);

                if (!ToBool(parameters[2], false))
                {
                    yearPeriods = yearPeriods.Where(p => p.AdjustmentPeriodFlag != "Y");
                }

                var lastPeriod = yearPeriods.OrderByDescending(p => p.PeriodNum).FirstOrDefault();
                string output = lastPeriod?.PeriodName ?? UdfSentinels.XlErrorGettingData;

                ServiceLocator.Logger?.LogDebug($"GLSense_GetPeriodEnd: resolved output='{output}'");
                UpdateFormulaCache(formulaCompressed, output);
                return output;
            }
            catch (Exception ex)
            {
                ServiceLocator.Logger?.LogException(ex, "GLSense_GetPeriodEnd Function : Unexpected error");
                return UdfSentinels.XlErrorGettingData;
            }
        }

        // args = [Period, Ledger]
        private static object HandleGetPeriodNum(object[] args)
        {
            try
            {
                string ledgerValue = ResolveLedger(args[1]);
                object[] parameters = new object[] { args[0], ledgerValue };

                ServiceLocator.Logger?.LogDebug($"GLSense_GetPeriodNum: period='{parameters[0]}' ledger='{ledgerValue}'");

                string formulaCompressed = FormulaCacheString(parameters);

                if (!AppState.Instance.IsLoginCompleted)
                    return GetCachedFormulaResultOrError(formulaCompressed);

                var periods = PModel(ledgerValue);
                if (periods == null || !periods.Any())
                    return GetCachedFormulaResultOrError(formulaCompressed);

                string periodName = parameters[0].ToString();
                var period = periods.FirstOrDefault(p => p.PeriodName == periodName);

                object output = period != null
                    ? (object)period.PeriodNum
                    : GetCachedFormulaResultOrError(formulaCompressed);

                ServiceLocator.Logger?.LogDebug($"GLSense_GetPeriodNum: resolved output='{output}'");
                UpdateFormulaCache(formulaCompressed, output);
                return output;
            }
            catch (Exception ex)
            {
                ServiceLocator.Logger?.LogException(ex, "GLSense_GetPeriodNum Function : Unexpected error");
                return UdfSentinels.XlErrorGettingData;
            }
        }

        // args = [Period, Ledger]
        private static object HandleGetPeriodQuarter(object[] args)
        {
            try
            {
                string ledgerValue = ResolveLedger(args[1]);
                object[] parameters = new object[] { args[0], ledgerValue };

                ServiceLocator.Logger?.LogDebug($"GLSense_GetPeriodQuarter: period='{parameters[0]}' ledger='{ledgerValue}'");

                string formulaCompressed = FormulaCacheString(parameters);

                if (!AppState.Instance.IsLoginCompleted)
                    return GetCachedFormulaResultOrError(formulaCompressed);

                var periods = PModel(ledgerValue);
                if (periods == null || !periods.Any())
                    return GetCachedFormulaResultOrError(formulaCompressed);

                string periodName = parameters[0].ToString();
                var period = periods.FirstOrDefault(p => p.PeriodName == periodName);

                object output = period != null
                    ? (object)period.QuarterNum
                    : GetCachedFormulaResultOrError(formulaCompressed);

                ServiceLocator.Logger?.LogDebug($"GLSense_GetPeriodQuarter: resolved output='{output}'");
                UpdateFormulaCache(formulaCompressed, output);
                return output;
            }
            catch (Exception ex)
            {
                ServiceLocator.Logger?.LogException(ex, "GLSense_GetPeriodQuarter Function : Unexpected error");
                return UdfSentinels.XlErrorGettingData;
            }
        }

        // args = [Period, Ledger]
        private static object HandleGetPeriodYear(object[] args)
        {
            try
            {
                string ledgerValue = ResolveLedger(args[1]);
                object[] parameters = new object[] { args[0], ledgerValue };

                ServiceLocator.Logger?.LogDebug($"GLSense_GetPeriodYear: period='{parameters[0]}' ledger='{ledgerValue}'");

                string formulaCompressed = FormulaCacheString(parameters);

                if (!AppState.Instance.IsLoginCompleted)
                    return GetCachedFormulaResultOrError(formulaCompressed);

                var periods = PModel(ledgerValue);
                if (periods == null || !periods.Any())
                    return GetCachedFormulaResultOrError(formulaCompressed);

                string periodName = parameters[0].ToString();
                var period = periods.FirstOrDefault(p => p.PeriodName == periodName);

                object output = period != null
                    ? (object)period.PeriodYear
                    : GetCachedFormulaResultOrError(formulaCompressed);

                ServiceLocator.Logger?.LogDebug($"GLSense_GetPeriodYear: resolved output='{output}'");
                UpdateFormulaCache(formulaCompressed, output);
                return output;
            }
            catch (Exception ex)
            {
                ServiceLocator.Logger?.LogException(ex, "GLSense_GetPeriodYear Function : Unexpected error");
                return UdfSentinels.XlErrorGettingData;
            }
        }

        // args = [SegmentValue, SegmentName, includeIdBool, Ledger]
        private static object HandleGetSegmentDesc(object[] args)
        {
            try
            {
                string ledgerValue = ResolveLedger(args[3]);
                object[] parameters = new object[] { args[0], args[1], args[2], ledgerValue };

                ServiceLocator.Logger?.LogDebug($"GLSense_GetSegmentDesc: segmentValue='{parameters[0]}' segmentName='{parameters[1]}' includeId='{parameters[2]}' ledger='{ledgerValue}'");

                string formulaCompressed = FormulaCacheString(parameters);

                if (!AppState.Instance.IsLoginCompleted)
                    return GetCachedFormulaResultOrError(formulaCompressed);

                string resolvedSegmentName = DataServiceLocator.SegmentDataService.ResolveSegmentName(parameters[1], ledgerValue);
                if (string.IsNullOrEmpty(resolvedSegmentName))
                {
                    ServiceLocator.Logger?.LogDebug($"GLSense_GetSegmentDesc Function : Unable to resolve segment name from '{parameters[1]}'");
                    return GetCachedFormulaResultOrError(formulaCompressed);
                }

                var segmentValues = LoadSegmentValues(ledgerValue);
                if (segmentValues == null || !segmentValues.Any())
                {
                    ServiceLocator.Logger?.LogDebug("GLSense_GetSegmentDesc Function : Unable to get segment values");
                    return GetCachedFormulaResultOrError(formulaCompressed);
                }

                string segmentValueStr = CleanSegmentValue(parameters[0].ToString());

                var segmentValue = segmentValues.FirstOrDefault(sv =>
                    sv.SegmentName.Equals(resolvedSegmentName, StringComparison.OrdinalIgnoreCase) &&
                    sv.SegmentValue.Equals(segmentValueStr, StringComparison.OrdinalIgnoreCase));

                string output;
                if (segmentValue != null)
                {
                    output = ToBool(parameters[2], false)
                        ? $"{segmentValue.SegmentValue} - {segmentValue.Description}"
                        : segmentValue.Description;
                }
                else
                {
                    output = GetCachedFormulaResultOrError(formulaCompressed).ToString();
                }

                ServiceLocator.Logger?.LogDebug($"GLSense_GetSegmentDesc: resolved output='{output}'");
                UpdateFormulaCache(formulaCompressed, output);
                return output;
            }
            catch (Exception ex)
            {
                ServiceLocator.Logger?.LogException(ex, "GLSense_GetSegmentDesc Function : Unexpected error");
                return UdfSentinels.XlErrorGettingData;
            }
        }

        // args = [SegmentValue, SegmentName, Ledger]
        private static object HandleGetSegmentEnabledFlag(object[] args)
        {
            try
            {
                string ledgerValue = ResolveLedger(args[2]);
                object[] parameters = new object[] { args[0], args[1], ledgerValue };

                ServiceLocator.Logger?.LogDebug($"GLSense_GetSegmentEnabledFlag: segmentValue='{parameters[0]}' segmentName='{parameters[1]}' ledger='{ledgerValue}'");

                string formulaCompressed = FormulaCacheString(parameters);

                if (!AppState.Instance.IsLoginCompleted)
                    return GetCachedFormulaResultOrError(formulaCompressed);

                string resolvedSegmentName = DataServiceLocator.SegmentDataService.ResolveSegmentName(parameters[1], ledgerValue);
                if (string.IsNullOrEmpty(resolvedSegmentName))
                {
                    ServiceLocator.Logger?.LogDebug($"GLSense_GetSegmentEnabledFlag Function : Unable to resolve segment name from '{parameters[1]}'");
                    return GetCachedFormulaResultOrError(formulaCompressed);
                }

                var segmentValues = LoadSegmentValues(ledgerValue);
                if (segmentValues == null || !segmentValues.Any())
                {
                    ServiceLocator.Logger?.LogDebug("GLSense_GetSegmentEnabledFlag Function : Unable to get segment values");
                    return GetCachedFormulaResultOrError(formulaCompressed);
                }

                string segmentValueStr = CleanSegmentValue(parameters[0].ToString());

                var segmentValue = segmentValues.FirstOrDefault(sv =>
                    sv.SegmentName.Equals(resolvedSegmentName, StringComparison.OrdinalIgnoreCase) &&
                    sv.SegmentValue.Equals(segmentValueStr, StringComparison.OrdinalIgnoreCase) &&
                    !sv.SummaryFlag.Equals("RG", StringComparison.OrdinalIgnoreCase));

                string output = segmentValue?.EnabledFlag ?? GetCachedFormulaResultOrError(formulaCompressed).ToString();

                ServiceLocator.Logger?.LogDebug($"GLSense_GetSegmentEnabledFlag: resolved output='{output}'");
                UpdateFormulaCache(formulaCompressed, output);
                return output;
            }
            catch (Exception ex)
            {
                ServiceLocator.Logger?.LogException(ex, "GLSense_GetSegmentEnabledFlag Function : Unexpected error");
                return UdfSentinels.XlErrorGettingData;
            }
        }

        // args = [SegmentValue, SegmentName, Ledger]
        private static object HandleGetSegmentSummaryFlag(object[] args)
        {
            try
            {
                string ledgerValue = ResolveLedger(args[2]);
                object[] parameters = new object[] { args[0], args[1], ledgerValue };

                ServiceLocator.Logger?.LogDebug($"GLSense_GetSegmentSummaryFlag: segmentValue='{parameters[0]}' segmentName='{parameters[1]}' ledger='{ledgerValue}'");

                string formulaCompressed = FormulaCacheString(parameters);

                if (!AppState.Instance.IsLoginCompleted)
                    return GetCachedFormulaResultOrError(formulaCompressed);

                string resolvedSegmentName = DataServiceLocator.SegmentDataService.ResolveSegmentName(parameters[1], ledgerValue);
                if (string.IsNullOrEmpty(resolvedSegmentName))
                {
                    ServiceLocator.Logger?.LogDebug($"GLSense_GetSegmentSummaryFlag Function : Unable to resolve segment name from '{parameters[1]}'");
                    return GetCachedFormulaResultOrError(formulaCompressed);
                }

                var segmentValues = LoadSegmentValues(ledgerValue);
                if (segmentValues == null || !segmentValues.Any())
                {
                    ServiceLocator.Logger?.LogDebug("GLSense_GetSegmentSummaryFlag Function : Unable to get segment values");
                    return GetCachedFormulaResultOrError(formulaCompressed);
                }

                string segmentValueStr = CleanSegmentValue(parameters[0].ToString());

                var segmentValue = segmentValues.FirstOrDefault(sv =>
                    sv.SegmentName.Equals(resolvedSegmentName, StringComparison.OrdinalIgnoreCase) &&
                    sv.SegmentValue.Equals(segmentValueStr, StringComparison.OrdinalIgnoreCase) &&
                    !sv.SummaryFlag.Equals("RG", StringComparison.OrdinalIgnoreCase));

                string output = segmentValue?.SummaryFlag ?? GetCachedFormulaResultOrError(formulaCompressed).ToString();

                ServiceLocator.Logger?.LogDebug($"GLSense_GetSegmentSummaryFlag: resolved output='{output}'");
                UpdateFormulaCache(formulaCompressed, output);
                return output;
            }
            catch (Exception ex)
            {
                ServiceLocator.Logger?.LogException(ex, "GLSense_GetSegmentSummaryFlag Function : Unexpected error");
                return UdfSentinels.XlErrorGettingData;
            }
        }

        // args = [SegmentValue, SegmentName, Attribute, Ledger] - async, but ExecuteUdf is
        // called from inside the host's own Task.Run, so this blocks synchronously via
        // .GetAwaiter().GetResult() (same convention as BalanceDtoModel.cs's EnsureLedgerDataLoaded).
        private static object HandleGetSegmentDFF(object[] args)
        {
            string ledgerValue = ResolveLedger(args[3]);
            object[] parameters = new object[] { args[0], args[1], args[2], ledgerValue };

            ServiceLocator.Logger?.LogDebug($"GLSense_GetSegmentDFF: segmentValue='{parameters[0]}' segmentName='{parameters[1]}' attribute='{parameters[2]}' ledger='{ledgerValue}'");

            string formulaCompressed = FormulaCacheString(parameters);

            if (!AppState.Instance.IsLoginCompleted)
                return HandleNotLoggedIn(formulaCompressed);

            try
            {
                string segmentValueStr = parameters[0]?.ToString() ?? string.Empty;
                string attributeStr = parameters[2]?.ToString() ?? string.Empty;

                segmentValueStr = CleanSegmentValue(segmentValueStr);
                if (string.IsNullOrEmpty(segmentValueStr) || segmentValueStr == "null")
                {
                    return string.Empty;
                }

                string resolvedSegmentName = ResolveSegmentName(parameters[1], ledgerValue);
                if (string.IsNullOrEmpty(resolvedSegmentName))
                {
                    ServiceLocator.Logger?.LogDebug($"GLSense_GetSegmentDFF : Unable to resolve segment name from '{parameters[1]}'");
                    return UdfSentinels.XlErrorGettingData;
                }

                long segmentValueSetId = GetSegmentValueSetId(resolvedSegmentName, ledgerValue);
                if (segmentValueSetId < 0)
                {
                    ServiceLocator.Logger?.LogDebug($"GLSense_GetSegmentDFF : Segment value set ID not found for segment '{resolvedSegmentName}'");
                    return UdfSentinels.XlErrorGettingData;
                }

                var requestData = new SegmentDff
                {
                    segmentValue = segmentValueStr,
                    segmentValueSetId = segmentValueSetId,
                    attributeName = attributeStr
                };

                string jsonPayload = JsonSerializer.Serialize(requestData, JsonGlobals.Options);
                string apiUrl = $"{AppState.Instance.LoginUrl}{AppConstants.RestSecure}segment-dff-value" +
                                $"?cubeId={AppState.Instance.SelectedCube.CubeId}";

                ServiceLocator.Logger?.LogDebug(apiUrl);
                ServiceLocator.Logger?.LogDebug(jsonPayload);

                using var cts = new CancellationTokenSource(TimeSpan.FromSeconds(300));
                string apiResponse = ApiHelper.ServerAPI(apiUrl, "JSON", jsonPayload, "POST", cts.Token)
                    .GetAwaiter().GetResult();

                if (string.IsNullOrWhiteSpace(apiResponse))
                {
                    return UdfSentinels.XlErrorGettingData;
                }

                string output = ParseSegmentsApiResponse(apiResponse) ?? string.Empty;
                ServiceLocator.Logger?.LogDebug($"GLSense_GetSegmentDFF: resolved output='{output}'");
                UpdateFormulaCache(formulaCompressed, output);
                return output;
            }
            catch (OperationCanceledException)
            {
                ServiceLocator.Logger?.LogWarn("GLSense_GetSegmentDFF cancelled/timeout");
                return UdfSentinels.XlErrorGettingData;
            }
            catch (Exception ex)
            {
                ServiceLocator.Logger?.LogException(ex, "GLSense_GetSegmentDFF : Unexpected error");
                return UdfSentinels.XlErrorGettingData;
            }
        }

        // args = [SegmentValue, SegmentIndex, Ledger] (Ledger unresolved) - async, same
        // blocking-call convention as HandleGetSegmentDFF above. Port of FinalWorkingCode's
        // GLSense_GetAccountType (GLSenseExcelFunctions.cs): GET .../account-type, reusing the
        // same ParseSegmentsApiResponse/TryGetRecordsNode/ExtractStringValue string-parsing
        // pipeline as GLSense_GetSegmentDFF, since the API response is the same
        // {"status","msg","records"} shape with "records" holding the plain string result
        // (e.g. {"msg":"default","records":"Assets","status":"success"}). SegmentIndex is the
        // segment's 1-based position in the Segment dropdown, sent straight through as the
        // API's segmentNumber - no more resolving a segment NAME via
        // ResolveSegmentName/GetSegmentSequenceIndex (removed).
        private static object HandleGetAccountType(object[] args)
        {
            string ledgerValue = ResolveLedger(args[2]);
            object[] parameters = new object[] { args[0], args[1], ledgerValue };

            ServiceLocator.Logger?.LogDebug($"GLSense_GetAccountType: segmentValue='{parameters[0]}' segmentIndex='{parameters[1]}' ledger='{ledgerValue}'");

            string formulaCompressed = FormulaCacheString(parameters);

            if (!AppState.Instance.IsLoginCompleted)
                return HandleNotLoggedIn(formulaCompressed);

            try
            {
                string segmentValueStr = parameters[0]?.ToString() ?? string.Empty;

                segmentValueStr = CleanSegmentValue(segmentValueStr);
                if (string.IsNullOrEmpty(segmentValueStr) || segmentValueStr == "null")
                {
                    return string.Empty;
                }

                if (!TryParseSegmentIndex(parameters[1], out int segmentSequence) || segmentSequence < 1)
                {
                    ServiceLocator.Logger?.LogDebug($"GLSense_GetAccountType : Invalid segment index '{parameters[1]}'");
                    return UdfSentinels.XlErrorGettingData;
                }

                var ledgerId = AppState.Instance.SelectedCube.GetLedgerIdByName(ledgerValue);
                if (!ledgerId.HasValue)
                {
                    ServiceLocator.Logger?.LogDebug($"GLSense_GetAccountType : Ledger '{ledgerValue}' not found in selected cube");
                    return UdfSentinels.XlErrorGettingData;
                }

                string apiUrl =
                    $"{AppState.Instance.LoginUrl}{AppConstants.RestSecure}account-type" +
                    $"?cubeId={AppState.Instance.SelectedCube.CubeId}" +
                    $"&segmentValue={segmentValueStr}" +
                    $"&segmentNumber={segmentSequence}" +
                    $"&ledgerId={ledgerId}";

                ServiceLocator.Logger?.LogDebug(apiUrl);

                using var cts = new CancellationTokenSource(TimeSpan.FromSeconds(300));
                string apiResponse = ApiHelper.ServerAPI(apiUrl, "Form", string.Empty, "GET", cts.Token)
                    .GetAwaiter().GetResult();

                if (string.IsNullOrWhiteSpace(apiResponse))
                {
                    return UdfSentinels.XlErrorGettingData;
                }

                string output = ParseSegmentsApiResponse(apiResponse) ?? string.Empty;
                ServiceLocator.Logger?.LogDebug($"GLSense_GetAccountType: resolved output='{output}'");
                UpdateFormulaCache(formulaCompressed, output);
                return output;
            }
            catch (OperationCanceledException)
            {
                ServiceLocator.Logger?.LogWarn("GLSense_GetAccountType cancelled/timeout");
                return UdfSentinels.XlErrorGettingData;
            }
            catch (Exception ex)
            {
                ServiceLocator.Logger?.LogException(ex, "GLSense_GetAccountType : Unexpected error");
                return UdfSentinels.XlErrorGettingData;
            }
        }

        // args = [FromCurrency, ToCurrency, ConversionType, ConversionDate] - async, same
        // blocking-call convention as HandleGetSegmentDFF above.
        private static object HandleGetDailyRate(object[] args)
        {
            object[] parameters = args;
            ServiceLocator.Logger?.LogDebug($"GLSense_GetDailyRate: fromCurrency='{SafeArg(parameters, 0)}' toCurrency='{SafeArg(parameters, 1)}' conversionType='{SafeArg(parameters, 2)}' conversionDate='{SafeArg(parameters, 3)}'");
            string formulaCompressed = FormulaCacheString(parameters);

            if (!AppState.Instance.IsLoginCompleted)
                return HandleNotLoggedIn(formulaCompressed);

            try
            {
                var requestData = new DailyRateQuery
                {
                    fromCurrency = parameters[0]?.ToString(),
                    toCurrency = parameters[1]?.ToString(),
                    conversionType = parameters[2]?.ToString(),
                    conversionDate = FormatConversionDate(parameters[3])
                };

                string jsonPayload = JsonSerializer.Serialize(requestData, JsonGlobals.Options);
                string apiUrl = $"{AppState.Instance.LoginUrl}{AppConstants.RestSecure}gldaily-rates?cubeId={AppState.Instance.SelectedCube.CubeId}";

                ServiceLocator.Logger?.LogDebug(apiUrl);
                ServiceLocator.Logger?.LogDebug(jsonPayload);

                string apiResponse = ExecuteApiCallWithTimeoutAsync(apiUrl, jsonPayload).GetAwaiter().GetResult();

                if (string.IsNullOrWhiteSpace(apiResponse))
                {
                    return UdfSentinels.XlErrorGettingData;
                }

                object output = ParseDailyRateResponse(apiResponse);
                ServiceLocator.Logger?.LogDebug($"GLSense_GetDailyRate: resolved output='{output}'");
                UpdateFormulaCache(formulaCompressed, output);
                return output;
            }
            catch (OperationCanceledException)
            {
                ServiceLocator.Logger?.LogWarn($"GLSense_GetDailyRate timeout");
                return UdfSentinels.XlErrorGettingData;
            }
            catch (Exception ex)
            {
                ServiceLocator.Logger?.LogException(ex, "GLSense_GetDailyRate background error");
                return UdfSentinels.XlErrorGettingData;
            }
        }

        // args = [ChangeSign, LedgerName, Activity, Period, BalanceType, CurrencyCode,
        //         TranslatedFlag, ActualFlag, BudorEncName, JESource, JECategory,
        //         Seg1..Seg20, XLLR1C1] (32 elements: 11 required + 20 segments + 1 trailing
        // XLLR1C1 computed host-side). LedgerName here is a distinct, mandatory raw value -
        // unrelated to the optional "Ledger" ResolveLedger resolution used by every other
        // function above.
        private static object HandleGetBalance(object[] args)
        {
            // Old monolith's short-circuit: checked BEFORE building/validating parameters.
            // AppState now lives here, so the check moves here too, first thing.
            if (AppState.Instance.ResetFormulas)
            {
                ServiceLocator.Logger?.LogDebug("GLSense_GetBalance: AppState.ResetFormulas is set - short-circuiting to 'Click Refresh...' without calling the API.");
                return GLClickToRefresh;
            }

            try
            {
                string xllR1C1 = args.Length > 31 ? (args[31]?.ToString() ?? string.Empty) : string.Empty;

                var segmentValues = new string[20];
                for (int i = 0; i < 20; i++)
                {
                    segmentValues[i] = GetSegmentString(args[11 + i]);
                }

                if (!string.IsNullOrWhiteSpace(segmentValues[0]) && segmentValues[0].Contains(";"))
                {
                    var combinedSegments = SplitCombinedSegments(segmentValues[0]);
                    for (int i = 0; i < segmentValues.Length && i < combinedSegments.Length; i++)
                    {
                        segmentValues[i] = combinedSegments[i];
                    }
                }

                var parametersList = new List<object>
                {
                    args[0], args[1], args[2], args[3], args[4],
                    args[5], args[6], args[7], args[8], args[9], args[10]
                };
                parametersList.AddRange(segmentValues);
                object[] parameters = parametersList.ToArray();

                string formulaCompressed = FormulaCacheString(parameters);

                ServiceLocator.Logger?.LogDebug($"GLSense_GetBalance: ledger='{args[1]}' activity='{args[2]}' period='{args[3]}' balanceType='{args[4]}' currency='{args[5]}' segmentsSet={segmentValues.Count(s => !string.IsNullOrWhiteSpace(s))} singleRefresh={AppState.Instance.SingleRefresh} batchCalc={AppState.Instance.StartBatchCalc}");

                if (!AppState.Instance.IsLoginCompleted)
                {
                    ServiceLocator.Logger?.LogDebug("GLSense_GetBalance: not logged in - returning cached/'Not Logged In' result.");
                    return HandleNotLoggedIn(formulaCompressed);
                }

                List<string> balanceParameters = GetBalanceParameters(parameters);

                object rawResult;
                if (AppState.Instance.SingleRefresh && !AppState.Instance.StartBatchCalc)
                {
                    ServiceLocator.Logger?.LogDebug("GLSense_GetBalance: dispatching single-refresh (live API) balance call.");
                    rawResult = ExecuteSingleRefreshAsync(formulaCompressed, balanceParameters, xllR1C1).GetAwaiter().GetResult();
                }
                else if (AppState.Instance.StartBatchCalc)
                {
                    ServiceLocator.Logger?.LogDebug("GLSense_GetBalance: dispatching batch-calc precomputed lookup.");
                    rawResult = HandleBatchCalculation(formulaCompressed);
                }
                else
                {
                    ServiceLocator.Logger?.LogDebug("GLSense_GetBalance: dispatching cached-result lookup (no refresh in progress).");
                    rawResult = HandleCachedResult(formulaCompressed);
                }

                object finalResult = rawResult switch
                {
                    double d => d,
                    decimal m => (double)m,
                    int i => (double)i,
                    long l => (double)l,
                    float f => (double)f,
                    null => UdfSentinels.XlErrorNull,
                    string s when s.StartsWith("#Err:") => s,
                    _ => rawResult ?? GLClickToRefresh
                };

                ServiceLocator.Logger?.LogDebug($"GLSense_GetBalance: resolved output='{finalResult}' (raw='{rawResult}')");
                return finalResult;
            }
            catch (Exception ex)
            {
                ServiceLocator.Logger?.LogException(ex, "Error in GLSense_GetBalance");
                return GLClickToRefresh;
            }
        }

        #endregion
    }
}
