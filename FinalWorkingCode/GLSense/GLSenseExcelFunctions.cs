using AddinExpress.MSO;
using GLSense.Caching;
using GLSense.Helpers;
using GLSense.Models;
using GLSense.Service;
using GLSense.Utilities;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Data;
using System.Globalization;
using System.Linq;
using System.Net.Http;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text.Json;
using System.Threading;
using System.Threading.Tasks;
using System.Web.UI.WebControls;

namespace GLSense
{
#nullable enable
    /// <summary>
    ///   Add-in Express XLL Add-in Module
    /// </summary>
    [ComVisible(true)]
    public partial class GLSenseExcelFunctions : AddinExpress.MSO.ADXXLLModule
    {
        public GLSenseExcelFunctions()
        {
            InitializeComponent();
            // Please add any initialization code to the OnInitialize event handler
        }



        #region Add-in Express automatic code

        // Required by Add-in Express - do not modify
        // the methods within this region

        public override System.ComponentModel.IContainer GetContainer()
        {
            if (components == null)
                components = new System.ComponentModel.Container();
            return components;
        }

        [ComRegisterFunctionAttribute]
        public static void RegisterXLL(Type t)
        {
            AddinExpress.MSO.ADXXLLModule.RegisterXLLInternal(t);
        }

        [ComUnregisterFunctionAttribute]
        public static void UnregisterXLL(Type t)
        {
            AddinExpress.MSO.ADXXLLModule.UnregisterXLLInternal(t);
        }

        #endregion

        public static new GLSenseExcelFunctions CurrentInstance
        {
            get
            {
                return (GLSenseExcelFunctions)(AddinExpress.MSO.ADXXLLModule.CurrentInstance
                    ?? throw new InvalidOperationException("GLSenseExcelFunctions instance is not available."));
            }
        }

        #region Define your UDFs in this section

        /// <summary>
        /// The container for user-defined functions (UDFs). Every UDF is a public static (Public Shared in VB.NET) method that returns a value of any base type: string, double, integer.
        /// </summary>
        internal static class XLLContainer
        {
            /// <summary>
            /// Required by Add-in Express. Please do not modify this method.
            /// </summary>
            internal static GLSenseExcelFunctions Module
            {
                get
                {
                    return (GLSenseExcelFunctions)(AddinExpress.MSO.ADXXLLModule.CurrentInstance
                        ?? throw new InvalidOperationException("GLSenseExcelFunctions instance is not available."));
                }
            }

            private static readonly string GLFormuladelimiter = "~!~";
            private static readonly string GLMissingLedger = "#Error: Missing Ledger";
            private static readonly string GLMissingPeriod = "#Error: Missing Period";
            private static readonly string GLMissingSegmentName = "#Error: Missing SegmentName";
            private static readonly string GLMissingSegmentValue = "#Error: Missing SegmentValue";
            private static readonly string GLNotLoggedIn = "Not Logged In...";
            private static readonly string GLClicktoRefresh = "Click Refresh...";
            private static readonly string GLDateFormat = "dd-MM-yyyy";

            private static string? XLLR1C1;

            //GLSense Custom excel functions helper methods

            // Custom attribute to define Excel function parameters
            [AttributeUsage(AttributeTargets.Parameter, AllowMultiple = false, Inherited = false)]
            public class ExcelParamAttribute : Attribute
            {
                public bool IsMandatory { get; }
                public object? DefaultValueKey { get; }

                public ExcelParamAttribute(bool isMandatory = false, object? defaultValueKey = null)
                {
                    IsMandatory = isMandatory;
                    DefaultValueKey = defaultValueKey;
                }
            }

            // Global defaults as static readonly fields with unique keys
            public static class ExcelDefaults
            {
                public static readonly object DefaultOffset = 0;
                public static readonly object DefaultAdjacentPeriods = true;
                public static readonly object DefaultIncludeId = false;
                public static readonly object DefaultNextParent = false;
                public static readonly object DefaultNextChild = false;

                public static readonly string DefaultLedgerName =
                    AppState.Instance.SelectedLedger?.LedgerName ?? string.Empty;

                public static readonly object GLDefaultText = "";
            }
            #region GLSense Excel Function Helpers
            private static object HandleNotLoggedIn(string FormulaKey)
            {
                try
                {
                    if (FormulaCacheManager.Instance.TryGetValue(FormulaKey, out var entry))
                    {
                        return entry.Value;
                    }
                    return GLNotLoggedIn;
                }
                catch (Exception ex)
                {
                    LogUtility.LogDebug($"HandleNotLoggedIn: could not read cached value for key ({ex.Message}).");
                    return GLNotLoggedIn;
                }
            }
            private static string FormulaCacheString(object[] args)
            {

                var paramArga = args
                    .Select(SafeValue) // SafeValue should return a string or object
                    .ToArray();

                return BuildFunctionCacheKey(paramArga);

            }
            private static string BuildFunctionCacheKey(params object[] parameters)
            {
                var key = string.Join(GLFormuladelimiter, parameters.Select(SafeValue));
                return CompressionHelper.CompressString(key);
            }
            private static string SafeValue(object value)
            {
                if (value == null || value is System.Reflection.Missing)
                {
                    return ""; // Replace missing values with an empty string
                }
                return value.ToString().Trim(); // Convert to string and trim spaces
            }
            private static string GetMandatoryError(string paramName)
            {
                string lowerName = paramName?.ToLower() ?? string.Empty;

                if (lowerName == "ledger") return GLMissingLedger;
                if (lowerName == "period" || lowerName == "perioddate" ||
                    lowerName == "periodyear" || lowerName == "periodnum")
                    return GLMissingPeriod;
                if (lowerName == "segmentname") return GLMissingSegmentName;
                if (lowerName == "segmentvalue") return GLMissingSegmentValue;

                return "#Error: Missing Parameter";
            }
            private static object? GetDefaultValue(string? key)
            {
                if (string.IsNullOrEmpty(key))
                    return null;

                return key switch
                {
                    "DefaultOffset" => ExcelDefaults.DefaultOffset,
                    "DefaultAdjacentPeriods" => ExcelDefaults.DefaultAdjacentPeriods,
                    "DefaultIncludeId" => ExcelDefaults.DefaultIncludeId,
                    "DefaultNextParent" => ExcelDefaults.DefaultNextParent,
                    "DefaultNextChild" => ExcelDefaults.DefaultNextChild,
                    "DefaultLedgerName" => ExcelDefaults.DefaultLedgerName,
                    _ => null
                };
            }

            // Resolve ledger parameter to a concrete ledger name string.
            // Uses supplied Ledger if non-empty, otherwise falls back to ExcelDefaults.
            private static string ResolveLedger(object? Ledger)
            {
                return string.IsNullOrEmpty(Ledger as string)
                    ? AppState.Instance.SelectedLedger?.LedgerName ?? string.Empty
                    : (Ledger as string) ?? string.Empty;
            }
            private static bool ToBool(object? value, bool defaultValue = false, double tolerance = 1e-6)
            {
                if (value == null || value is System.Reflection.Missing || value is AddinExpress.MSO.ADXExcelError)
                    return defaultValue;

                if (value is bool b)
                    return b;

                if (value is double d)
                {
                    // Treat values close to 1 as true, close to 0 as false
                    if (Math.Abs(d - 1.0) < tolerance) return true;
                    if (Math.Abs(d - 0.0) < tolerance) return false;
                    return defaultValue;
                }

                if (value is string s)
                {
                    if (string.IsNullOrWhiteSpace(s))
                        return defaultValue;

                    s = s.Trim();
                    if (s.Equals("TRUE", StringComparison.OrdinalIgnoreCase))
                        return true;
                    if (s.Equals("FALSE", StringComparison.OrdinalIgnoreCase))
                        return false;
                    if (s.Equals("1"))
                        return true;
                    if (s.Equals("0"))
                        return false;
                }

                // Fallback: try parsing
                if (bool.TryParse(value.ToString(), out var parsed))
                    return parsed;

                return defaultValue;
            }
            private static object? ValidateInputs(object[] parameters)
            {
                var method = new System.Diagnostics.StackTrace().GetFrame(1)?.GetMethod();
                if (method == null || method.GetParameters().Length < parameters.Length)
                    return "#Err: Invalid method parameters";

                for (int i = 0; i < parameters.Length; i++)
                {
                    var param = parameters[i];
                    if (param is System.Reflection.Missing || param == null)
                    {
                        var paramInfo = method.GetParameters()[i];
                        var attr = paramInfo.GetCustomAttribute<ExcelParamAttribute>();

                        if (attr?.IsMandatory == true)
                            return GetMandatoryError(paramInfo.Name!);

                        // Fixed: Consistent null checking pattern
                        if (attr?.DefaultValueKey != null &&
                            !string.IsNullOrEmpty(attr.DefaultValueKey.ToString()))
                        {
                            var defaultValue = GetDefaultValue(attr.DefaultValueKey.ToString()!);
                            parameters[i] = defaultValue ?? string.Empty;
                        }
                    }
                }

                return null;
            }
            private static List<PeriodModel> PModel(string lName)
            {
                try
                {
                    // Get periods using the new service
                    List<PeriodModel> periods = LoadPeriodsForLedger(lName);
                    if (periods == null)
                    {
                        return new List<PeriodModel>();
                    }
                    else
                    {
                        return periods;
                    }
                }
                catch (Exception ex)
                {
                    LogUtility.LogException(ex);
                    return new List<PeriodModel>();
                }
            }
            private static List<PeriodModel> LoadPeriodsForLedger(string ledgerName)
            {
                try
                {
                    var dataService = ServiceLocator.PeriodDataService;
                    return new List<PeriodModel>(dataService.GetPeriodsForLedger(ledgerName, allowRemoteFetch: false));
                }
                catch (Exception ex)
                {
                    LogUtility.LogException(ex, $"Failed to load periods for ledger '{ledgerName}'");
                    return new List<PeriodModel>();
                }
            }
            private enum Direction { Next, Previous }

            private static object GetSegmentByDirection(object SegmentValue, object SegmentName, object NextParent, object NextChild, Direction direction, object? Ledger)
            {
                LogUtility.LogDebug($"GetSegmentByDirection ({direction}) invoked. SegmentValue={SegmentValue}, SegmentName={SegmentName}, NextParent={NextParent}, NextChild={NextChild}, Ledger={Ledger}");

                var ledgerValue = ResolveLedger(Ledger);

                object[] parameters = new object[] { SegmentValue, SegmentName, NextParent, NextChild, ledgerValue };
                object? result = ValidateInputs(parameters);
                if (result != null) return result;

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

                LogUtility.LogDebug($"GetSegmentByDirection ({direction}) result: {output}");
                UpdateFormulaCache(formulaCompressed, output);
                return output;
            }
            private static void UpdateFormulaCache(string formulaCompressed, object output)
            {
                try
                {
                    CachedFormulaHelper.Store(formulaCompressed, output, DateTime.UtcNow);
                }
                catch (Exception ex)
                {
                    LogUtility.LogException(ex, "Failed to update formula cache");
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

                    return ADXExcelError.xlErrorGettingData;
                }
                catch (Exception ex)
                {
                    LogUtility.LogException(ex, $"Formula: {CompressionHelper.DecompressString(formulaCompressed)}");
                    return ADXExcelError.xlErrorGettingData;
                }
            }
            private static string ResolveSegmentName(object segmentNameObj, object ledgerName) =>
                       ServiceLocator.SegmentDataService.ResolveSegmentName(segmentNameObj, ledgerName.ToString());

            private static List<SegmentValueModel> GetFilteredSegments(object[] parameters)
            {
                string ledgerName = parameters[parameters.Count() - 1].ToString();
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
                    var task = Task.Run(() =>
                    {
                        var dataService = ServiceLocator.SegmentDataService;
                        return dataService.GetSegmentValues(ledgerName);
                    });

                    if (task.Wait(TimeSpan.FromSeconds(180)))
                    {
                        return task.Result;
                    }
                    else
                    {
                        throw new TimeoutException("Timeout loading segment values from service");
                    }
                }
                catch (Exception ex)
                {
                    LogUtility.LogException(ex, "Failed to load segment values");
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
                bool nextParent = ToBool(parameters[2],false);
                bool nextChild = ToBool(parameters[3],false);

                if (nextParent && !nextChild)
                    return sv => sv.SummaryFlag == "Y";
                else if (!nextParent && nextChild)
                    return sv => sv.SummaryFlag != "Y";
                else
                    return sv => true;
            }
            private static string FindMatchingSegment(List<SegmentValueModel> segments, int currentIndex,
                Func<SegmentValueModel, bool> filter, Direction direction)
            {
                IEnumerable<SegmentValueModel> candidateSegments;

                if (direction == Direction.Next)
                    candidateSegments = segments.Skip(currentIndex + 1);
                else
                    candidateSegments = segments.Take(currentIndex).Reverse();

                var matchingSegment = candidateSegments.FirstOrDefault(filter);
                return matchingSegment?.SegmentValue ?? ADXExcelError.xlErrorGettingData.ToString();
            }

            private static object LogAndReturnError(string message, Direction direction, string formulaCompressed)
            {
                string methodName = direction == Direction.Next ? "GLSense_GetNextSegment" : "GLSense_GetPreviousSegment";
                LogUtility.LogDebug($"{methodName} Function : {message}");
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
                    JsonValueKind.Number =>
                    GetNumberValue(element),
                    JsonValueKind.String =>
                        // Safe null-forgiving since GetString() after ValueKind.String check
                        ParseStringNumber(element.GetString()!)!,

                    _ => element.GetRawText()
                };
            }
            private static object GetNumberValue(JsonElement element)
            {
                // Try decimal first (full precision)
                if (element.TryGetDecimal(out decimal dec))
                    return dec;

                // Then double
                if (element.TryGetDouble(out double dbl))
                    return dbl;

                // Fallback to int64
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

                // ---------------------------------------------
                // Auto-detect records node
                // ---------------------------------------------


                if (!TryGetRecordsNode(root, out JsonElement recordsElem))
                {
                    // If no "records" property, assume root itself is records
                    recordsElem = root;
                }

                object value = ExtractNumericValue(recordsElem);

                CacheResult(formulaCompressed, value);

                // Cache successful numeric values only
                if (value is double or decimal or int or long)
                {
                    // Excel-compatible return
                    return value switch
                    {
                        double d => d,
                        decimal m => (double)m,
                        int i => i,
                        long l => l,
                        _ => "#Err: Invalid numeric value"
                    };

                }
                else
                    return value;
            }

            private static bool TryGetRecordsNode(
                    JsonElement root,
                    out JsonElement recordsNode)
            {
                var recordProp = root.EnumerateObject()
                    .FirstOrDefault(prop => string.Equals(prop.Name,
                                                         "records",
                                                         StringComparison.OrdinalIgnoreCase));

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
                if (FormulaCacheManager.Instance == null) return;

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
                    object? sObj = null;

                    LogUtility.LogDebug($"Formula: {CompressionHelper.DecompressString(formulaCompressed)}");

                    if (FormulaCacheManager.Instance.ContainsKey(formulaCompressed))
                    {
                        sObj = GetResultValue(FormulaCacheManager.Instance.GetValue(formulaCompressed));
                        return sObj;
                    }
                    else
                    {
                        return GLClicktoRefresh;
                    }
                }
                catch (Exception ex)
                {
                    LogUtility.LogException(ex, $"Formula: {CompressionHelper.DecompressString(formulaCompressed)}");
                    return GLClicktoRefresh;
                }
            }
            private static void UpdateCache(string formulaCompressed, object result)
            {
                if (FormulaCacheManager.Instance == null) return;

                string resultString = result?.ToString() ?? string.Empty;
                CachedFormulaHelper.Store(formulaCompressed, resultString, DateTime.UtcNow);
            }
            private static object GetBatchResult(string formulaCompressed)
            {

                if (AppState.Instance.PreComputedBalances?.Count == 0)
                    return GLClicktoRefresh;

                //Removed the condition for the cell and sheet check as batch calculation now handles all cells together

                if (AppState.Instance.PreComputedBalances?.TryGetValue(formulaCompressed, out var cachedValue) == true)
                {
                    return cachedValue;
                }

                return GLClicktoRefresh;
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
                    LogUtility.LogException(ex);
                    return GLClicktoRefresh;
                }
            }
            private static async Task<string?> ExecuteApiCallWithTimeoutAsync(
                string apiUrl,
                string jsonPayload)
            {
                using var cts = new CancellationTokenSource(TimeSpan.FromSeconds(300));

                try
                {
                    // Assuming ApiHelper.ServerAPI is already async (returns Task<string>)
                    string response = await ApiHelper.ServerAPI(
                        apiUrl,
                        "JSON",
                        jsonPayload,
                        "POST",
                        cts.Token);

                    return response;
                }
                catch (OperationCanceledException)
                {
                    LogUtility.LogWarn($"API call timed out after 30s: {apiUrl}");
                    return null;
                }
                catch (HttpRequestException ex)
                {
                    LogUtility.LogError($"HTTP/network error: {ex.Message}");
                    return null;
                }
                catch (Exception ex)
                {
                    LogUtility.LogException(ex, "Unexpected error in API call");
                    return null;
                }
            }
            private static async Task<object> ExecuteSingleRefreshAsync(
                string formulaCompressed,
                List<string> balanceParameters)
            {
                try
                {
                    var balanceDto = BalanceDto.CreateFromXllParameters(XLLR1C1, balanceParameters);
                    string jsonPayload = JsonSerializer.Serialize(balanceDto, JsonGlobals.Options);
                    string apiUrl = $"{AppState.Instance.LoginUrl}{AppConstants.RestSecure}balance?cubeId={AppState.Instance.SelectedCube.CubeId}";

                    string? apiResponse = await ExecuteApiCallWithTimeoutAsync(apiUrl, jsonPayload);

                    return apiResponse != null
                        ? ParseAndCacheBalanceResponse(apiResponse, formulaCompressed)
                        : GLClicktoRefresh;
                }
                catch (Exception ex)
                {
                    LogUtility.LogException(ex);
                    return GLClicktoRefresh;
                }
            }
            private static void GetCellCallerAddress()
            {
                try
                {
                    if (Module.CallWorksheetFunction(AddinExpress.MSO.ADXExcelWorksheetFunction.Caller) is not AddinExpress.MSO.ADXExcelRef caller)
                    {
                        LogUtility.LogWarn("GetCellCallerAddress: Caller returned null");
                        return;
                    }

                    int rowFirst = caller.RowFirst + 1;
                    int columnFirst = caller.ColumnFirst + 1;

                    XLLR1C1 = "R" + rowFirst.ToString() + "C" + columnFirst.ToString();
                }
                catch (Exception ex)
                {
                    LogUtility.LogException(ex);
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

                    // Handle Excel double directly (most common case)
                    if (obj is double d)
                    {
                        return d;
                    }

                    // Parse string values (handles "40645271.100")
                    if (double.TryParse(obj.ToString().Trim(), NumberStyles.Any, CultureInfo.InvariantCulture, out double numericValue))
                    {
                        return numericValue;
                    }

                    return obj; // Non-numeric strings, dates, etc.
                }
                catch (Exception ex)
                {
                    LogUtility.LogException(ex, $"GetResultValue|{obj}");
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
                    var segments = ServiceLocator.SegmentDataService.GetSegments(ledgerName);
                    var segment = segments?.FirstOrDefault(s =>
                        s.SegmentName.Equals(segmentName, StringComparison.OrdinalIgnoreCase));

                    return segment?.SegmentValueSetId ?? -1;
                }
                catch (Exception ex)
                {
                    LogUtility.LogException(ex, $"Failed to get segment value set ID for '{segmentName}'");
                    return -1;
                }
            }
            // GLSense_GetAccountType now receives the segment's 1-based dropdown position
            // directly from the formula (e.g. GLSense_GetAccountType("1000", 3, "Vision Ops")),
            // so the ApplicationColumnName-regex lookup this used to require
            // (GetSegmentSequenceIndex, removed) is no longer needed - just parse the number.
            // Excel passes an un-quoted numeric literal as a boxed double, so this accepts
            // double/int/numeric-string defensively.
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
            private static string? ParseSegmentsApiResponse(string apiResponse)
            {
                if (string.IsNullOrWhiteSpace(apiResponse))
                    return ADXExcelError.xlErrorGettingData.ToString();

                LogUtility.LogDebug(apiResponse);

                var result =
                    ApiResponseHelper.Parse<JsonElement>(apiResponse, JsonGlobals.Options);

                if (!result.IsSuccess)
                    return result.ErrorMessage
                           ?? ADXExcelError.xlErrorGettingData.ToString();

                JsonElement root = result.Value;

                if (!TryGetRecordsNode(root, out JsonElement recordsElem))
                    recordsElem = root;

                string? output = ExtractStringValue(recordsElem);

                return string.IsNullOrWhiteSpace(output)
                    ? "#Empty Response"
                    : output;
            }

            private static string? ExtractStringValue(JsonElement element)
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
                        // If array, try first element
                        if (element.GetArrayLength() == 0)
                            return null;

                        return ExtractStringValue(element[0]);

                    case JsonValueKind.Object:
                        foreach (var val in
                        // Try to find first primitive property
                        from prop in element.EnumerateObject()
                        let val = prop.Value
                        select val)
                        {
                            if (val.ValueKind == JsonValueKind.String ||
                                                        val.ValueKind == JsonValueKind.Number ||
                                                        val.ValueKind == JsonValueKind.True ||
                                                        val.ValueKind == JsonValueKind.False)
                            {
                                return val.ToString();
                            }
                            // Recursive search (nested object/array)
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
                        return ADXExcelError.xlErrorGettingData;

                    // If not JSON → API returned plain error text
                    if (!apiResponse.TrimStart().StartsWith("{") &&
                        !apiResponse.TrimStart().StartsWith("["))
                    {
                        LogUtility.LogWarn($"DailyRate API returned non-JSON: {apiResponse}");
                        return apiResponse;
                    }

                    // Use new enterprise parser
                    var result = ApiResponseHelper.Parse<List<DailyRateRecord>>(apiResponse, JsonGlobals.Options);

                    if (!result.IsSuccess)
                    {
                        LogUtility.LogWarn($"DailyRate parse failed: {result.ErrorMessage}");
                        return (object?)result.ErrorMessage ?? ADXExcelError.xlErrorGettingData;
                    }

                    var records = result.Value;

                    if (records == null || records.Count == 0)
                        return ADXExcelError.xlErrorGettingData;

                    var first = records[0];

                    if (first?.CONVERSION_RATE is double rate)
                        return rate;

                    return ADXExcelError.xlErrorGettingData;
                }
                catch (Exception ex)
                {
                    LogUtility.LogException(ex, "Failed to parse DailyRate API response.");
                    return ADXExcelError.xlErrorGettingData;
                }
            }

            private static string FormatConversionDate(object conversionDate)
            {
                try
                {
                    if (conversionDate == null)
                        return string.Empty;

                    // Handle Excel date numbers
                    if (conversionDate is double oaDate)
                    {
                        return DateTime.FromOADate(oaDate).ToString(GLDateFormat);
                    }

                    string dateString = conversionDate.ToString();

                    // If it's already in "dd-MM-yyyy" format, return as-is
                    if (DateTime.TryParseExact(dateString, GLDateFormat, CultureInfo.InvariantCulture,
                        DateTimeStyles.None, out DateTime _))
                    {
                        return dateString; // Already in correct format
                    }

                    // Try parsing as a regular date with culture-aware provider
                    if (DateTime.TryParse(dateString, CultureInfo.InvariantCulture, DateTimeStyles.None,
                        out DateTime parsedDate))
                    {
                        return parsedDate.ToString(GLDateFormat, CultureInfo.InvariantCulture);
                    }

                    // If parsing fails, try common Excel date formats
                    string[] excelFormats = new[] {
                            "MM/dd/yyyy", "M/d/yyyy", "dd/MM/yyyy", "d/M/yyyy",
                            "yyyy-MM-dd", "MM-dd-yyyy", GLDateFormat
                        };

                    if (DateTime.TryParseExact(dateString, excelFormats,
                        CultureInfo.InvariantCulture, DateTimeStyles.None, out DateTime parsedDate3))
                    {
                        return parsedDate3.ToString(GLDateFormat);
                    }

                    // Last resort: return original string and let API handle validation
                    LogUtility.LogDebug($"Date parsing failed for: {dateString}");
                    return dateString;
                }
                catch (Exception ex)
                {
                    LogUtility.LogException(ex, $"Error in FormatConversionDate for: {conversionDate}");
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
                        catch
                        {
                            // Continue to string parsing below.
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
                    LogUtility.LogException(ex, $"GLSense_GetPeriodByDate: Error converting date {dateObj}");
                    return default;
                }
            }

            private static string FindPeriodName(IEnumerable<PeriodModel> periods, DateTime date)
            {
                // Calendar-day containment, not instant containment: a period's stored EndDate is
                // midnight of its last day (not the last instant of that day), so comparing full
                // DateTime values would wrongly exclude the last day for any date carrying a
                // time-of-day component. Comparing .Date matches GLPeriodByDateModel's own
                // (already-correct) period lookup and is what "date falls within the period" means.
                return periods.FirstOrDefault(p => p.StartDate.Date <= date.Date && p.EndDate.Date >= date.Date)?.PeriodName ?? string.Empty;
            }

            private static string CalculateOffsetPeriod(List<PeriodModel> periods, string basePeriod, int offset)
            {
                try
                {
                    var orderedPeriods = periods.OrderBy(p => p.StartDate).ToList();
                    int baseIndex = orderedPeriods.FindIndex(p => p.PeriodName == basePeriod);
                    int newIndex = baseIndex + offset;

                    return newIndex >= 0 && newIndex < orderedPeriods.Count
                        ? orderedPeriods[newIndex].PeriodName ?? ADXExcelError.xlErrorGettingData.ToString()
                        : ADXExcelError.xlErrorGettingData.ToString();
                }
                catch (Exception ex)
                {
                    LogUtility.LogException(ex);
                    return ADXExcelError.xlErrorGettingData.ToString();
                }
            }

            #endregion GLSense Excel Function Helpers

            #region GLSense Excel Functions
            // Get Period by Date Function
            public static object GLSense_GetPeriodByDate([ExcelParam(true)] object PeriodDate, [ExcelParam(false, "DefaultLedgerName")] object? Ledger = null, [ExcelParam(false, "DefaultOffset")] object? offset = null)
            {
                try
                {
                    LogUtility.LogDebug($"GLSense_GetPeriodByDate invoked. PeriodDate={PeriodDate}, Ledger={Ledger}, offset={offset}");
                    var ledgerValue = ResolveLedger(Ledger);
                    var parameters = new object[] { PeriodDate, ledgerValue, offset ?? 0 };

                    var validationResult = ValidateInputs(parameters);
                    if (validationResult != null) return validationResult;

                    string formulaCompressed = FormulaCacheString(parameters);

                    if (!AppState.Instance.IsLoginCompleted)
                    {
                        return GetCachedFormulaResultOrError(formulaCompressed);
                    }

                    var periods = GetPeriods(ledgerValue);
                    if (periods == null || !periods.Any())
                        return ADXExcelError.xlErrorGettingData;

                    DateTime periodDate = ParsePeriodDate(parameters[0]);
                    if (periodDate == default)  // Invalid date parsing returns default(DateTime)
                        return ADXExcelError.xlErrorGettingData;

                    string selectedPeriod = FindPeriodName(periods, periodDate);
                    if (string.IsNullOrEmpty(selectedPeriod))
                        return GetCachedFormulaResultOrError(formulaCompressed);

                    string output = CalculateOffsetPeriod(periods, selectedPeriod, Convert.ToInt32(parameters[2]));
                    LogUtility.LogDebug($"GLSense_GetPeriodByDate result: {output}");
                    UpdateFormulaCache(formulaCompressed, output);
                    return output;
                }
                catch (Exception ex)
                {
                    LogUtility.LogException(ex, "GLSense_GetPeriodByDate Function : Unexpected error");
                    return ADXExcelError.xlErrorGettingData;
                }
            }
            public static object GLSense_GetPeriod([ExcelParam(true)] object Period, [ExcelParam(false, "DefaultOffset")] object? offset = null, [ExcelParam(false, "DefaultLedgerName")] object? Ledger = null)
            {
                try
                {
                    LogUtility.LogDebug($"GLSense_GetPeriod invoked. Period={Period}, offset={offset}, Ledger={Ledger}");
                    // Input validation

                    var ledgerValue = ResolveLedger(Ledger);
                    object[] parameters = new object[] { Period, offset ?? 0, ledgerValue };

                    // 2. Validate AND set defaults in-place
                    object? result = ValidateInputs(parameters);
                    if (result != null) return result;

                    string formulaCompressed = FormulaCacheString(parameters);

                    // Check cache first (if logged out)
                    if (!AppState.Instance.IsLoginCompleted)
                    {
                        return GetCachedFormulaResultOrError(formulaCompressed);
                    }
                        
                    // Get periods using the new service
                    var periods = PModel(ledgerValue);
                    if (periods == null || !periods.Any())
                    {
                        return GetCachedFormulaResultOrError(formulaCompressed);
                    }

                    string periodName = parameters[0].ToString();
                    string output = string.Empty;
                    int selectedOffset = Convert.ToInt32(parameters[1]);

                    // Find the period by name
                    var basePeriod = periods.FirstOrDefault(p => p.PeriodName == periodName);
                    if (basePeriod == null)
                    {
                        LogUtility.LogDebug($"GLSense_GetPeriod Function : Period '{periodName}' not found for ledger \"{ledgerValue}\"");
                        return GetCachedFormulaResultOrError(formulaCompressed);
                    }

                    try
                    {
                        var orderedPeriods = periods.OrderBy(p => p.StartDate).ToList();
                        int baseIndex = orderedPeriods.FindIndex(p => p.PeriodName == periodName);
                        int newIndex = baseIndex + selectedOffset;

                        if (newIndex >= 0 && newIndex < orderedPeriods.Count)
                        {
                            output = orderedPeriods[newIndex].PeriodName ?? ADXExcelError.xlErrorGettingData.ToString();
                        }
                        else
                        {
                            output = ADXExcelError.xlErrorGettingData.ToString();
                        }
                    }
                    catch (Exception ex)
                    {
                        LogUtility.LogException(ex);
                        output = ADXExcelError.xlErrorGettingData.ToString();
                    }

                    // Update cache
                    LogUtility.LogDebug($"GLSense_GetPeriod result: {output}");
                    UpdateFormulaCache(formulaCompressed, output);

                    return output;
                }
                catch (Exception ex)
                {
                    LogUtility.LogException(ex, "GLSense_GetPeriod Function : Unexpected error");
                    return ADXExcelError.xlErrorGettingData;
                }
            }
            public static object GLSense_GetPeriodByYear([ExcelParam(true)] object PeriodYear, [ExcelParam(true)] object PeriodNum, [ExcelParam(false, "DefaultLedgerName")] object? Ledger = null)
            {
                try
                {
                    LogUtility.LogDebug($"GLSense_GetPeriodByYear invoked. PeriodYear={PeriodYear}, PeriodNum={PeriodNum}, Ledger={Ledger}");
                    // Input validation

                    var ledgerValue = ResolveLedger(Ledger);
                    object[] parameters = new object[] { PeriodYear, PeriodNum, ledgerValue };

                    // 2. Validate AND set defaults in-place
                    object? result = ValidateInputs(parameters);
                    if (result != null) return result;

                    string formulaCompressed = FormulaCacheString(parameters);

                    // Check cache first (if logged out)
                    if (!AppState.Instance.IsLoginCompleted)
                    {
                        return GetCachedFormulaResultOrError(formulaCompressed);
                    }

                    // Get periods using the new service
                    var periods = PModel(ledgerValue);
                    if (periods == null || !periods.Any())
                    {
                        return GetCachedFormulaResultOrError(formulaCompressed);
                    }

                    int periodYear = Convert.ToInt32(parameters[0]);
                    int periodNum = Convert.ToInt32(parameters[1]);

                    // Find period by year and number
                    var period = periods.FirstOrDefault(p => p.PeriodYear == periodYear && p.PeriodNum == periodNum);
                    string output = period?.PeriodName ?? ADXExcelError.xlErrorGettingData.ToString();

                    // Update cache
                    LogUtility.LogDebug($"GLSense_GetPeriodByYear result: {output}");
                    UpdateFormulaCache(formulaCompressed, output);

                    return output;
                }
                catch (Exception ex)
                {
                    LogUtility.LogException(ex, "GLSense_GetPeriodByYear Function : Unexpected error");
                    return ADXExcelError.xlErrorGettingData;
                }
            }
            public static object GLSense_GetPeriodStart([ExcelParam(true)] object Period, [ExcelParam(false, "DefaultLedgerName")] object? Ledger = null, [ExcelParam(false, "DefaultAdjacentPeriods")] object? AdjacentPeriods = null)
            {
                try
                {
                    LogUtility.LogDebug($"GLSense_GetPeriodStart invoked. Period={Period}, Ledger={Ledger}, AdjacentPeriods={AdjacentPeriods}");
                    // Input validation

                    bool adjacentPeriodsBool = ToBool(AdjacentPeriods, false);

                    var ledgerValue = ResolveLedger(Ledger);
                    object[] parameters = new object[] { Period, ledgerValue, adjacentPeriodsBool };

                    // 2. Validate AND set defaults in-place
                    object? result = ValidateInputs(parameters);
                    if (result != null) return result;

                    string formulaCompressed = FormulaCacheString(parameters);

                    // Check cache first (if logged out)
                    if (!AppState.Instance.IsLoginCompleted)
                        return GetCachedFormulaResultOrError(formulaCompressed);

                    // Get periods using the new service
                    var periods = PModel(ledgerValue);
                    if (periods == null || !periods.Any())
                    {
                        return GetCachedFormulaResultOrError(formulaCompressed);
                    }

                    string periodName = parameters[0].ToString();
                    string output = string.Empty;

                    // Find the base period to get the year
                    var basePeriod = periods.FirstOrDefault(p => p.PeriodName == periodName);
                    if (basePeriod == null)
                    {
                        LogUtility.LogDebug($"GLSense_GetPeriodStart Function : Period '{periodName}' not found for ledger \"{ledgerValue}\"");
                        return GetCachedFormulaResultOrError(formulaCompressed);
                    }

                    int targetYear = basePeriod.PeriodYear;

                    // Get all periods for the target year
                    var yearPeriods = periods.Where(p => p.PeriodYear == targetYear);

                    // Filter by adjacent flag if needed
                    if (!(bool)parameters[2])
                    {
                        yearPeriods = yearPeriods.Where(p => p.AdjustmentPeriodFlag != "Y");
                    }

                    // Get the period with minimum period number
                    var firstPeriod = yearPeriods.OrderBy(p => p.PeriodNum).FirstOrDefault();
                    output = firstPeriod?.PeriodName ?? ADXExcelError.xlErrorGettingData.ToString();

                    // Update cache
                    LogUtility.LogDebug($"GLSense_GetPeriodStart result: {output}");
                    UpdateFormulaCache(formulaCompressed, output);

                    return output;
                }
                catch (Exception ex)
                {
                    LogUtility.LogException(ex, "GLSense_GetPeriodStart Function : Unexpected error");
                    return ADXExcelError.xlErrorGettingData;
                }
            }
            public static object GLSense_GetPeriodEnd([ExcelParam(true)] object Period, [ExcelParam(false, "DefaultLedgerName")] object? Ledger = null, [ExcelParam(false, "DefaultAdjacentPeriods")] object? AdjacentPeriods = null)
            {
                try
                {
                    LogUtility.LogDebug($"GLSense_GetPeriodEnd invoked. Period={Period}, Ledger={Ledger}, AdjacentPeriods={AdjacentPeriods}");
                    // Input validation

                    bool adjacentPeriodsBool = ToBool(AdjacentPeriods, false);

                    var ledgerValue = ResolveLedger(Ledger);
                    object[] parameters = new object[] { Period, ledgerValue, adjacentPeriodsBool };

                    // 2. Validate AND set defaults in-place
                    object? result = ValidateInputs(parameters);
                    if (result != null) return result;

                    string formulaCompressed = FormulaCacheString(parameters);

                    // Check cache first (if logged out)
                    if (!AppState.Instance.IsLoginCompleted)
                        return GetCachedFormulaResultOrError(formulaCompressed);

                    // Get periods using the new service
                    var periods = PModel(ledgerValue);
                    if (periods == null || !periods.Any())
                    {
                        return GetCachedFormulaResultOrError(formulaCompressed);
                    }

                    string periodName = parameters[0].ToString();
                    string output = string.Empty;

                    // Find the base period to get the year
                    var basePeriod = periods.FirstOrDefault(p => p.PeriodName == periodName);
                    if (basePeriod == null)
                    {
                        LogUtility.LogDebug($"GLSense_GetPeriodEnd Function : Period '{periodName}' not found for ledger \"{ledgerValue}\"");
                        return GetCachedFormulaResultOrError(formulaCompressed);
                    }

                    int targetYear = basePeriod.PeriodYear;

                    // Get all periods for the target year
                    var yearPeriods = periods.Where(p => p.PeriodYear == targetYear);

                    // Filter by adjacent flag if needed
                    if (!(bool)parameters[2])
                    {
                        yearPeriods = yearPeriods.Where(p => p.AdjustmentPeriodFlag != "Y");
                    }

                    // Get the period with maximum period number
                    var lastPeriod = yearPeriods.OrderByDescending(p => p.PeriodNum).FirstOrDefault();
                    output = lastPeriod?.PeriodName ?? ADXExcelError.xlErrorGettingData.ToString();

                    // Update cache
                    LogUtility.LogDebug($"GLSense_GetPeriodEnd result: {output}");
                    UpdateFormulaCache(formulaCompressed, output);

                    return output;
                }
                catch (Exception ex)
                {
                    LogUtility.LogException(ex, "GLSense_GetPeriodEnd Function : Unexpected error");
                    return ADXExcelError.xlErrorGettingData;
                }
            }
            public static object GLSense_GetPeriodNum([ExcelParam(true)] object Period, [ExcelParam(false, "DefaultLedgerName")] object? Ledger = null)
            {
                try
                {
                    LogUtility.LogDebug($"GLSense_GetPeriodNum invoked. Period={Period}, Ledger={Ledger}");

                    // Input validation

                    var ledgerValue = ResolveLedger(Ledger);
                    object[] parameters = new object[] { Period, ledgerValue };

                    // 2. Validate AND set defaults in-place
                    object? result = ValidateInputs(parameters);
                    if (result != null) return result;

                    string formulaCompressed = FormulaCacheString(parameters);

                    // Check cache first (if logged out)
                    if (!AppState.Instance.IsLoginCompleted)
                        return GetCachedFormulaResultOrError(formulaCompressed);

                    // Get periods using the new service
                    var periods = PModel(ledgerValue);
                    if (periods == null || !periods.Any())
                    {
                        return GetCachedFormulaResultOrError(formulaCompressed);
                    }

                    string periodName = parameters[0].ToString();

                    // Find the period and return its number
                    var period = periods.FirstOrDefault(p => p.PeriodName == periodName);

                    object output = period?.PeriodNum != null
                        ? (object)period.PeriodNum
                        : GetCachedFormulaResultOrError(formulaCompressed);

                    // Update cache
                    LogUtility.LogDebug($"GLSense_GetPeriodNum result: {output}");
                    UpdateFormulaCache(formulaCompressed, output);

                    return output;
                }
                catch (Exception ex)
                {
                    LogUtility.LogException(ex, "GLSense_GetPeriodNum Function : Unexpected error");
                    return ADXExcelError.xlErrorGettingData;
                }
            }
            public static object GLSense_GetPeriodQuarter([ExcelParam(true)] object Period, [ExcelParam(false, "DefaultLedgerName")] object? Ledger = null)
            {
                try
                {
                    LogUtility.LogDebug($"GLSense_GetPeriodQuarter invoked. Period={Period}, Ledger={Ledger}");
                    // Input validation

                    var ledgerValue = ResolveLedger(Ledger);
                    object[] parameters = new object[] { Period, ledgerValue };

                    // 2. Validate AND set defaults in-place
                    object? result = ValidateInputs(parameters);
                    if (result != null) return result;

                    string formulaCompressed = FormulaCacheString(parameters);

                    // Check cache first (if logged out)
                    if (!AppState.Instance.IsLoginCompleted)
                        return GetCachedFormulaResultOrError(formulaCompressed);

                    // Get periods using the new service
                    var periods = PModel(ledgerValue);
                    if (periods == null || !periods.Any())
                    {
                        return GetCachedFormulaResultOrError(formulaCompressed);
                    }

                    string periodName = parameters[0].ToString();

                    // Find the period and return its quarter
                    var period = periods.FirstOrDefault(p => p.PeriodName == periodName);

                    object output = period?.QuarterNum != null
                        ? (object)period.QuarterNum
                        : GetCachedFormulaResultOrError(formulaCompressed);

                    // Update cache
                    LogUtility.LogDebug($"GLSense_GetPeriodQuarter result: {output}");
                    UpdateFormulaCache(formulaCompressed, output);

                    return output;
                }
                catch (Exception ex)
                {
                    LogUtility.LogException(ex, "GLSense_GetPeriodQuarter Function : Unexpected error");
                    return ADXExcelError.xlErrorGettingData;
                }
            }
            public static object GLSense_GetPeriodYear([ExcelParam(true)] object Period, [ExcelParam(false, "DefaultLedgerName")] object? Ledger = null)
            {
                try
                {
                    LogUtility.LogDebug($"GLSense_GetPeriodYear invoked. Period={Period}, Ledger={Ledger}");
                    // Input validation

                    var ledgerValue = ResolveLedger(Ledger);
                    object[] parameters = new object[] { Period, ledgerValue };

                    // 2. Validate AND set defaults in-place
                    object? result = ValidateInputs(parameters);
                    if (result != null) return result;

                    string formulaCompressed = FormulaCacheString(parameters);

                    // Check cache first (if logged out)
                    if (!AppState.Instance.IsLoginCompleted)
                        return GetCachedFormulaResultOrError(formulaCompressed);

                    // Get periods using the new service
                    var periods = PModel(ledgerValue);
                    if (periods == null || !periods.Any())
                    {
                        return GetCachedFormulaResultOrError(formulaCompressed);
                    }

                    string periodName = parameters[0].ToString();

                    // Find the period and return its year
                    var period = periods.FirstOrDefault(p => p.PeriodName == periodName);
                    object output = period?.PeriodYear != null
                        ? (object)period.PeriodYear
                        : GetCachedFormulaResultOrError(formulaCompressed);

                    // Update cache
                    LogUtility.LogDebug($"GLSense_GetPeriodYear result: {output}");
                    UpdateFormulaCache(formulaCompressed, output);

                    return output;
                }
                catch (Exception ex)
                {
                    LogUtility.LogException(ex, "GLSense_GetPeriodYear Function : Unexpected error");
                    return ADXExcelError.xlErrorGettingData;
                }
            }
            public static object GLSense_GetSegmentDesc([ExcelParam(true)] object SegmentValue, [ExcelParam(true)] object SegmentName, [ExcelParam(false, "DefaultIncludeId")] object? IncludeId = null, [ExcelParam(false, "DefaultLedgerName")] object? Ledger = null)
            {
                try
                {
                    LogUtility.LogDebug($"GLSense_GetSegmentDesc invoked. SegmentValue={SegmentValue}, SegmentName={SegmentName}, IncludeId={IncludeId}, Ledger={Ledger}");
                     bool includeIdBool = ToBool(IncludeId, false);
                    // Input validation

                    var ledgerValue = ResolveLedger(Ledger);
                    object[] parameters = new object[] { SegmentValue, SegmentName, includeIdBool, ledgerValue };

                    // 2. Validate AND set defaults in-place
                    object? result = ValidateInputs(parameters);
                    if (result != null) return result;

                    string formulaCompressed = FormulaCacheString(parameters);

                    // Check cache first (if logged out)
                    if (!AppState.Instance.IsLoginCompleted)
                        return GetCachedFormulaResultOrError(formulaCompressed);

                    string resolvedSegmentName = ServiceLocator.SegmentDataService.ResolveSegmentName(parameters[1], ledgerValue);
                    if (string.IsNullOrEmpty(resolvedSegmentName))
                    {
                        LogUtility.LogDebug($"GLSense_GetSegmentDesc Function : Unable to resolve segment name from '{parameters[1]}'");
                        return GetCachedFormulaResultOrError(formulaCompressed);
                    }

                    // Get segment values
                    var segmentValues = LoadSegmentValues(ledgerValue);
                    if (segmentValues == null || !segmentValues.Any())
                    {
                        LogUtility.LogDebug($"GLSense_GetSegmentDesc Function : Unable to get segment values");
                        return GetCachedFormulaResultOrError(formulaCompressed);
                    }

                    string segmentValueStr = CleanSegmentValue(parameters[0].ToString());

                    // Find the segment value
                    var segmentValue = segmentValues.FirstOrDefault(sv =>
                        sv.SegmentName.Equals(resolvedSegmentName, StringComparison.OrdinalIgnoreCase) &&
                        sv.SegmentValue.Equals(segmentValueStr, StringComparison.OrdinalIgnoreCase));

                    string output;
                    if (segmentValue != null)
                    {
                        output = (bool)parameters[2] ?
                            $"{segmentValue.SegmentValue} - {segmentValue.Description}" :
                            segmentValue.Description;
                    }
                    else
                    {
                        output = GetCachedFormulaResultOrError(formulaCompressed).ToString();
                    }

                    // Update cache
                    LogUtility.LogDebug($"GLSense_GetSegmentDesc result: {output}");
                    UpdateFormulaCache(formulaCompressed, output);

                    return output;
                }
                catch (Exception ex)
                {
                    LogUtility.LogException(ex, "GLSense_GetSegmentDesc Function : Unexpected error");
                    return ADXExcelError.xlErrorGettingData;
                }
            }
            public static object GLSense_GetSegmentEnabledFlag([ExcelParam(true)] object SegmentValue, [ExcelParam(true)] object SegmentName, [ExcelParam(false, "DefaultLedgerName")] object? Ledger = null)
            {
                try
                {
                    LogUtility.LogDebug($"GLSense_GetSegmentEnabledFlag invoked. SegmentValue={SegmentValue}, SegmentName={SegmentName}, Ledger={Ledger}");
                    // Input validation

                    var ledgerValue = ResolveLedger(Ledger);
                    object[] parameters = new object[] { SegmentValue, SegmentName, ledgerValue };

                    // 2. Validate AND set defaults in-place
                    object? result = ValidateInputs(parameters);
                    if (result != null) return result;

                    string formulaCompressed = FormulaCacheString(parameters);

                    // Check cache first (if logged out)
                    if (!AppState.Instance.IsLoginCompleted)
                        return GetCachedFormulaResultOrError(formulaCompressed);

                    // Resolve segment name (handles both string and sequence)
                    string resolvedSegmentName = ServiceLocator.SegmentDataService.ResolveSegmentName(parameters[1], ledgerValue);
                    if (string.IsNullOrEmpty(resolvedSegmentName))
                    {
                        LogUtility.LogDebug($"GLSense_GetSegmentEnabledFlag Function : Unable to resolve segment name from '{parameters[1]}'");
                        return GetCachedFormulaResultOrError(formulaCompressed);
                    }

                    // Get segment values
                    var segmentValues = LoadSegmentValues(ledgerValue);
                    if (segmentValues == null || !segmentValues.Any())
                    {
                        LogUtility.LogDebug($"GLSense_GetSegmentEnabledFlag Function : Unable to get segment values");
                        return GetCachedFormulaResultOrError(formulaCompressed);
                    }

                    string segmentValueStr = CleanSegmentValue(parameters[0].ToString());

                    // Find the segment value
                    var segmentValue = segmentValues.FirstOrDefault(sv =>
                        sv.SegmentName.Equals(resolvedSegmentName, StringComparison.OrdinalIgnoreCase) &&
                        sv.SegmentValue.Equals(segmentValueStr, StringComparison.OrdinalIgnoreCase) &&
                        !sv.SummaryFlag.Equals("RG", StringComparison.OrdinalIgnoreCase));

                    string output = segmentValue?.EnabledFlag ?? GetCachedFormulaResultOrError(formulaCompressed).ToString();

                    // Update cache
                    LogUtility.LogDebug($"GLSense_GetSegmentEnabledFlag result: {output}");
                    UpdateFormulaCache(formulaCompressed, output);

                    return output;
                }
                catch (Exception ex)
                {
                    LogUtility.LogException(ex, "GLSense_GetSegmentEnabledFlag Function : Unexpected error");
                    return ADXExcelError.xlErrorGettingData;
                }
            }
            public static object GLSense_GetSegmentSummaryFlag([ExcelParam(true)] object SegmentValue, [ExcelParam(true)] object SegmentName, [ExcelParam(false, "DefaultLedgerName")] object? Ledger = null)
            {
                try
                {
                    LogUtility.LogDebug($"GLSense_GetSegmentSummaryFlag invoked. SegmentValue={SegmentValue}, SegmentName={SegmentName}, Ledger={Ledger}");
                    // Input validation

                    var ledgerValue = ResolveLedger(Ledger);
                    object[] parameters = new object[] { SegmentValue, SegmentName, ledgerValue };

                    // 2. Validate AND set defaults in-place
                    object? result = ValidateInputs(parameters);
                    if (result != null) return result;

                    string formulaCompressed = FormulaCacheString(parameters);

                    // Check cache first (if logged out)
                    if (!AppState.Instance.IsLoginCompleted)
                        return GetCachedFormulaResultOrError(formulaCompressed);

                    // Resolve segment name (handles both string and sequence)
                    string resolvedSegmentName = ServiceLocator.SegmentDataService.ResolveSegmentName(parameters[1], ledgerValue);
                    if (string.IsNullOrEmpty(resolvedSegmentName))
                    {
                        LogUtility.LogDebug($"GLSense_GetSegmentSummaryFlag Function : Unable to resolve segment name from '{parameters[1]}'");
                        return GetCachedFormulaResultOrError(formulaCompressed);
                    }

                    // Get segment values
                    var segmentValues = LoadSegmentValues(ledgerValue);
                    if (segmentValues == null || !segmentValues.Any())
                    {
                        LogUtility.LogDebug($"GLSense_GetSegmentSummaryFlag Function : Unable to get segment values");
                        return GetCachedFormulaResultOrError(formulaCompressed);
                    }

                    string segmentValueStr = CleanSegmentValue(parameters[0].ToString());

                    // Find the segment value
                    var segmentValue = segmentValues.FirstOrDefault(sv =>
                        sv.SegmentName.Equals(resolvedSegmentName, StringComparison.OrdinalIgnoreCase) &&
                        sv.SegmentValue.Equals(segmentValueStr, StringComparison.OrdinalIgnoreCase) &&
                        !sv.SummaryFlag.Equals("RG", StringComparison.OrdinalIgnoreCase));

                    string output = segmentValue?.SummaryFlag ?? GetCachedFormulaResultOrError(formulaCompressed).ToString();

                    // Update cache
                    LogUtility.LogDebug($"GLSense_GetSegmentSummaryFlag result: {output}");
                    UpdateFormulaCache(formulaCompressed, output);

                    return output;
                }
                catch (Exception ex)
                {
                    LogUtility.LogException(ex, "GLSense_GetSegmentSummaryFlag Function : Unexpected error");
                    return ADXExcelError.xlErrorGettingData;
                }
            }
            public static object GLSense_GetNextSegment([ExcelParam(true)] object SegmentValue, [ExcelParam(true)] object SegmentName, [ExcelParam(true)] object NextParent, [ExcelParam(true)] object NextChild, [ExcelParam(false, "DefaultLedgerName")] object? Ledger = null)
            {
                return GetSegmentByDirection(SegmentValue, SegmentName, NextParent, NextChild, Direction.Next, Ledger);
            }
            public static object GLSense_GetPreviousSegment([ExcelParam(true)] object SegmentValue, [ExcelParam(true)] object SegmentName, [ExcelParam(true)] object NextParent, [ExcelParam(true)] object NextChild, [ExcelParam(false, "DefaultLedgerName")] object? Ledger = null)
            {
                return GetSegmentByDirection(SegmentValue, SegmentName, NextParent, NextChild, Direction.Previous, Ledger);
            }
            public static void GLSense_GetAccountType(
                [ExcelParam(true)] object SegmentValue,
                [ExcelParam(true)] object SegmentIndex,
                [ExcelParam(false, "DefaultLedgerName")] object? Ledger,
                ADXExcelAsyncCallObject asyncCallObject)
            {
                LogUtility.LogDebug($"GLSense_GetAccountType invoked. SegmentValue={SegmentValue}, SegmentIndex={SegmentIndex}, Ledger={Ledger}");
                // Input validation
                var ledgerValue = ResolveLedger(Ledger);
                object[] parameters = new object[] { SegmentValue, SegmentIndex, ledgerValue };

                // 2. Validate AND set defaults in-place
                object? earlyResult = ValidateInputs(parameters);
                if (earlyResult != null)
                {
                    asyncCallObject.ReturnResult(earlyResult ?? string.Empty);
                    return;
                }

                string formulaCompressed = FormulaCacheString(parameters);

                // Cache hit when not logged in
                if (!AppState.Instance.IsLoginCompleted)
                {
                    asyncCallObject.ReturnResult(HandleNotLoggedIn(formulaCompressed));
                    return;
                }

                try
                {

                    // Quick checks before background
                    string segmentValueStr = parameters[0]?.ToString() ?? string.Empty;

                    segmentValueStr = CleanSegmentValue(segmentValueStr);
                    if (string.IsNullOrEmpty(segmentValueStr) || segmentValueStr == "null")
                    {
                        asyncCallObject.ReturnResult(string.Empty);
                        return;
                    }

                    // The formula now carries the segment's 1-based position directly
                    // (e.g. GLSense_GetAccountType("1000", 3, "Vision Ops")), picked from
                    // the Segment dropdown in the picker window - it is no longer a segment
                    // NAME that needs resolving via ResolveSegmentName/GetSegmentSequenceIndex
                    // (that ApplicationColumnName-regex helper has been removed entirely).
                    // Excel passes an un-quoted numeric literal as a boxed double, so parse
                    // defensively (double, int, or numeric string all accepted).
                    if (!TryParseSegmentIndex(parameters[1], out int segmentSequence) || segmentSequence < 1)
                    {
                        LogUtility.LogDebug($"GLSense_GetAccountType : Invalid segment index '{parameters[1]}'");
                        asyncCallObject.ReturnResult(ADXExcelError.xlErrorGettingData);
                        return;
                    }

                    var ledgerId = AppState.Instance.SelectedCube.GetLedgerIdByName(ledgerValue);
                    if (!ledgerId.HasValue)
                    {
                        LogUtility.LogDebug($"GLSense_GetAccountType : Ledger '{ledgerValue}' not found in selected cube");
                        asyncCallObject.ReturnResult(ADXExcelError.xlErrorGettingData);
                        return;
                    }

                    // ────────────────────────────────────────────────
                    // At this point: we need the real API call → go async
                    // ────────────────────────────────────────────────

                    var worker = new BackgroundWorker();
                    worker.DoWork += (sender, e) =>
                    {
                        var callObj = (ADXExcelAsyncCallObject)e.Argument;

                        try
                        {

                            string apiUrl =
                                $"{AppState.Instance.LoginUrl}{AppConstants.RestSecure}account-type" +
                                $"?cubeId={AppState.Instance.SelectedCube.CubeId}" +
                                $"&segmentValue={segmentValueStr}" +
                                $"&segmentNumber={segmentSequence}" +
                                $"&ledgerId={ledgerId}";

                            LogUtility.LogDebug(apiUrl);

                            using var cts = new CancellationTokenSource(TimeSpan.FromSeconds(300));

                            // Your API call (synchronous inside background is fine)
                            string apiResponse = ApiHelper.ServerAPI(
                                apiUrl, "Form", string.Empty, "GET", cts.Token)
                                .GetAwaiter().GetResult();   // or better: make ApiHelper truly async & await

                            if (string.IsNullOrWhiteSpace(apiResponse))
                            {
                                callObj.ReturnResult(ADXExcelError.xlErrorGettingData);
                                return;
                            }

                            string output = ParseSegmentsApiResponse(apiResponse) ?? string.Empty;

                            LogUtility.LogDebug($"GLSense_GetAccountType result: {output}");

                            // Update cache from background (assuming it's thread-safe)
                            UpdateFormulaCache(formulaCompressed, output);

                            // Final success
                            callObj.ReturnResult(output);
                        }
                        catch (OperationCanceledException)
                        {
                            LogUtility.LogWarn("GLSense_GetAccountType cancelled/timeout");
                            callObj.ReturnResult(ADXExcelError.xlErrorGettingData);
                        }
                        catch (Exception ex)
                        {
                            LogUtility.LogException(ex, "GLSense_GetAccountType : Unexpected error in background");
                            callObj.ReturnResult(ADXExcelError.xlErrorGettingData);
                        }
                    };

                    // Start background – pass the async handle
                    worker.RunWorkerAsync(asyncCallObject);
                }
                catch (AggregateException ex) when (ex.InnerException is OperationCanceledException)
                {
                    LogUtility.LogWarn($"GLSense_GetAccountType cancelled/timeout: {ex.InnerException?.Message}");
                    asyncCallObject.ReturnResult(ADXExcelError.xlErrorGettingData);
                }
                catch (Exception ex)
                {
                    LogUtility.LogException(ex, "GLSense_GetAccountType Function : Unexpected error");
                    asyncCallObject.ReturnResult(ADXExcelError.xlErrorGettingData);
                }
            }
            public static void GLSense_GetSegmentDFF(
                [ExcelParam(true)] object SegmentValue,
                [ExcelParam(true)] object SegmentName,
                [ExcelParam(true)] object Attribute,
                [ExcelParam(false, "DefaultLedgerName")] object? Ledger,
                ADXExcelAsyncCallObject asyncCallObject)
            {
                LogUtility.LogDebug($"GLSense_GetSegmentDFF invoked. SegmentValue={SegmentValue}, SegmentName={SegmentName}, Attribute={Attribute}, Ledger={Ledger}");
                // Input validation
                var ledgerValue = ResolveLedger(Ledger);
                object[] parameters = new object[] { SegmentValue, SegmentName, Attribute, ledgerValue };

                // 2. Validate AND set defaults in-place
                object? earlyResult = ValidateInputs(parameters);
                if (earlyResult != null)
                {
                    asyncCallObject.ReturnResult(earlyResult ?? string.Empty);
                    return;
                }

                string formulaCompressed = FormulaCacheString(parameters);

                // Cache hit when not logged in
                if (!AppState.Instance.IsLoginCompleted)
                {
                    asyncCallObject.ReturnResult(HandleNotLoggedIn(formulaCompressed));
                    return;
                }

                try
                {

                    // Quick checks before background
                    string segmentValueStr = parameters[0]?.ToString() ?? string.Empty;
                    string attributeStr = parameters[2]?.ToString() ?? string.Empty;

                    segmentValueStr = CleanSegmentValue(segmentValueStr);
                    if (string.IsNullOrEmpty(segmentValueStr) || segmentValueStr == "null")
                    {
                        asyncCallObject.ReturnResult(string.Empty);
                        return;
                    }

                    string resolvedSegmentName = ServiceLocator.SegmentDataService.ResolveSegmentName(parameters[1], ledgerValue);
                    if (string.IsNullOrEmpty(resolvedSegmentName))
                    {
                        LogUtility.LogDebug($"GLSense_GetSegmentDFF : Unable to resolve segment name from '{parameters[1]}'");
                        asyncCallObject.ReturnResult(ADXExcelError.xlErrorGettingData);
                        return;
                    }

                    long segmentValueSetId = GetSegmentValueSetId(resolvedSegmentName, ledgerValue);
                    if (segmentValueSetId < 0)
                    {
                        LogUtility.LogDebug($"GLSense_GetSegmentDFF : Segment value set ID not found for segment '{resolvedSegmentName}'");
                        asyncCallObject.ReturnResult(ADXExcelError.xlErrorGettingData);
                        return;
                    }

                    // ────────────────────────────────────────────────
                    // At this point: we need the real API call → go async
                    // ────────────────────────────────────────────────

                    var worker = new BackgroundWorker();
                    worker.DoWork += (sender, e) =>
                    {
                        var callObj = (ADXExcelAsyncCallObject)e.Argument;

                        try
                        {
                            var requestData = new SegmentDff
                            {
                                segmentValue = segmentValueStr,
                                segmentValueSetId = segmentValueSetId,
                                attributeName = attributeStr
                            };

                            string jsonPayload = JsonSerializer.Serialize(requestData, JsonGlobals.Options);

                            string apiUrl = $"{AppState.Instance.LoginUrl}{AppConstants.RestSecure}segment-dff-value" +
                                            $"?cubeId={AppState.Instance.SelectedCube.CubeId}";

                            LogUtility.LogDebug(apiUrl);
                            LogUtility.LogDebug(jsonPayload);

                            using var cts = new CancellationTokenSource(TimeSpan.FromSeconds(300));

                            // Your API call (synchronous inside background is fine)
                            string apiResponse = ApiHelper.ServerAPI(
                                apiUrl, "JSON", jsonPayload, "POST", cts.Token)
                                .GetAwaiter().GetResult();   // or better: make ApiHelper truly async & await

                            if (string.IsNullOrWhiteSpace(apiResponse))
                            {
                                callObj.ReturnResult(ADXExcelError.xlErrorGettingData);
                                return;
                            }

                            string output = ParseSegmentsApiResponse(apiResponse) ?? string.Empty;

                            LogUtility.LogDebug($"GLSense_GetSegmentDFF result: {output}");

                            // Update cache from background (assuming it's thread-safe)
                            UpdateFormulaCache(formulaCompressed, output);

                            // Final success
                            callObj.ReturnResult(output);
                        }
                        catch (OperationCanceledException)
                        {
                            LogUtility.LogWarn("GLSense_GetSegmentDFF cancelled/timeout");
                            callObj.ReturnResult(ADXExcelError.xlErrorGettingData);
                        }
                        catch (Exception ex)
                        {
                            LogUtility.LogException(ex, "GLSense_GetSegmentDFF : Unexpected error in background");
                            callObj.ReturnResult(ADXExcelError.xlErrorGettingData);
                        }
                    };

                    // Start background – pass the async handle
                    worker.RunWorkerAsync(asyncCallObject);
                }
                catch (AggregateException ex) when (ex.InnerException is OperationCanceledException)
                {
                    LogUtility.LogWarn($"GLSense_GetSegmentDFF cancelled/timeout: {ex.InnerException?.Message}");
                    asyncCallObject.ReturnResult(ADXExcelError.xlErrorGettingData);
                }
                catch (Exception ex)
                {
                    LogUtility.LogException(ex, "GLSense_GetSegmentDFF Function : Unexpected error");
                    asyncCallObject.ReturnResult(ADXExcelError.xlErrorGettingData);
                }
            }
            public static void GLSense_GetDailyRate(
                [ExcelParam(true)] object FromCurrency,
                [ExcelParam(true)] object ToCurrency,
                [ExcelParam(true)] object ConversionType,
                [ExcelParam(true)] object ConversionDate,
                ADXExcelAsyncCallObject asyncCallObject)   // ← must be LAST
            {
                try
                {
                    LogUtility.LogDebug($"GLSense_GetDailyRate invoked. FromCurrency={FromCurrency}, ToCurrency={ToCurrency}, ConversionType={ConversionType}, ConversionDate={ConversionDate}");
                    object[] parameters = new object[] { FromCurrency, ToCurrency, ConversionType, ConversionDate };

                    // Quick sync validation & early exits
                    object? earlyResult = ValidateInputs(parameters);
                    if (earlyResult != null)
                    {
                        asyncCallObject.ReturnResult(earlyResult!);   // or earlyResult ?? someDefault
                        return;
                    }

                    string formulaCompressed = FormulaCacheString(parameters);

                    if (!AppState.Instance.IsLoginCompleted)
                    {
                        asyncCallObject.ReturnResult(HandleNotLoggedIn(formulaCompressed));
                        return;
                    }

                    // Prepare request data (quick sync part)
                    var requestData = new DailyRateQuery
                    {
                        fromCurrency = parameters[0]?.ToString(),
                        toCurrency = parameters[1]?.ToString(),
                        conversionType = parameters[2]?.ToString(),
                        conversionDate = FormatConversionDate(parameters[3])
                    };

                    string jsonPayload = JsonSerializer.Serialize(requestData, JsonGlobals.Options);
                    string apiUrl = $"{AppState.Instance.LoginUrl}{AppConstants.RestSecure}gldaily-rates?cubeId={AppState.Instance.SelectedCube.CubeId}";

                    LogUtility.LogDebug(apiUrl);
                    LogUtility.LogDebug(jsonPayload);

                    // ────────────────────────────────────────────────
                    // Slow part → background
                    // ────────────────────────────────────────────────
                    Task.Run(async () =>
                    {
                        try
                        {
                            using var cts = new CancellationTokenSource(TimeSpan.FromSeconds(300));

                            // Assuming ApiHelper.ServerAPI is async (returns Task<string>)
                            string apiResponse = await ApiHelper.ServerAPI(
                                apiUrl, "JSON", jsonPayload, "POST", cts.Token);

                            if (string.IsNullOrWhiteSpace(apiResponse))
                            {
                                asyncCallObject.ReturnResult(ADXExcelError.xlErrorGettingData);
                                return;
                            }

                            object output = ParseDailyRateResponse(apiResponse);
                            LogUtility.LogDebug($"GLSense_GetDailyRate result: {output}");
                            UpdateFormulaCache(formulaCompressed, output);

                            asyncCallObject.ReturnResult(output);
                        }
                        catch (OperationCanceledException)
                        {
                            LogUtility.LogWarn($"GLSense_GetDailyRate timeout: {apiUrl}");
                            asyncCallObject.ReturnResult(ADXExcelError.xlErrorGettingData);
                        }
                        catch (Exception ex)
                        {
                            LogUtility.LogException(ex, "GLSense_GetDailyRate background error");
                            asyncCallObject.ReturnResult(ADXExcelError.xlErrorGettingData);
                        }
                    });
                }
                catch (Exception ex)
                {
                    LogUtility.LogException(ex, "GLSense_GetDailyRate sync error");
                    asyncCallObject.ReturnResult(ADXExcelError.xlErrorGettingData);
                }
            }

#pragma warning disable S107   //This is used to suppress the warning of using 31 paramters in the function below
            private static string GetSegmentString(object? seg)
            {
                string segStr = seg switch
                {
                    System.Reflection.Missing _ => "",                  // omitted in formula
                    null => "",                  // rare, but safe
                    "" => "",                  // explicitly empty
                    _ => seg.ToString() ?? "" // any other value → string
                };

                return segStr;
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
             public static void GLSense_GetBalance(
                     // 1. All REQUIRED parameters (no default values) come FIRST
                     [ExcelParam(true)] object ChangeSign,
                     [ExcelParam(true)] object LedgerName,
                     [ExcelParam(true)] object Activity,
                     [ExcelParam(true)] object Period,
                     [ExcelParam(true)] object BalanceType,
                     [ExcelParam(true)] object CurrencyCode,
                     [ExcelParam(true)] object TranslatedFlag,
                     [ExcelParam(true)] object ActualFlag,
                     [ExcelParam(true)] object BudorEncName,
                     [ExcelParam(true)] object JESource,
                     [ExcelParam(true)] object JECategory,

                        // 2. All OPTIONAL parameters come AFTER all required ones
                        [ExcelParam(false, "SEGMENT1")] object? Seg1,
                        [ExcelParam(false, "SEGMENT2")] object? Seg2,
                        [ExcelParam(false, "SEGMENT3")] object? Seg3,
                        [ExcelParam(false, "SEGMENT4")] object? Seg4,
                        [ExcelParam(false, "SEGMENT5")] object? Seg5,
                        [ExcelParam(false, "SEGMENT6")] object? Seg6,
                        [ExcelParam(false, "SEGMENT7")] object? Seg7,
                        [ExcelParam(false, "SEGMENT8")] object? Seg8,
                        [ExcelParam(false, "SEGMENT9")] object? Seg9,
                        [ExcelParam(false, "SEGMENT10")] object? Seg10,
                        [ExcelParam(false, "SEGMENT11")] object? Seg11,
                        [ExcelParam(false, "SEGMENT12")] object? Seg12,
                        [ExcelParam(false, "SEGMENT13")] object? Seg13,
                        [ExcelParam(false, "SEGMENT14")] object? Seg14,
                        [ExcelParam(false, "SEGMENT15")] object? Seg15,
                        [ExcelParam(false, "SEGMENT16")] object? Seg16,
                        [ExcelParam(false, "SEGMENT17")] object? Seg17,
                        [ExcelParam(false, "SEGMENT18")] object? Seg18,
                        [ExcelParam(false, "SEGMENT19")] object? Seg19,
                        [ExcelParam(false, "SEGMENT20")] object? Seg20,

                    // 3. ADXExcelAsyncCallObject MUST be last – and REQUIRED (no default!)
                    ADXExcelAsyncCallObject asyncCallObject)
            {
                try
                {
                    LogUtility.LogDebug($"GLSense_GetBalance invoked. LedgerName={LedgerName}, Activity={Activity}, Period={Period}, BalanceType={BalanceType}, CurrencyCode={CurrencyCode}, TranslatedFlag={TranslatedFlag}, ActualFlag={ActualFlag}, BudorEncName={BudorEncName}, JESource={JESource}, JECategory={JECategory}, ChangeSign={ChangeSign}");

                    if (AppState.Instance.ResetFormulas)
                    {
                        LogUtility.LogDebug("GLSense_GetBalance: ResetFormulas is active, returning click-to-refresh placeholder.");
                        asyncCallObject.ReturnResult(GLClicktoRefresh);
                        return;
                    }

                    string seg1Str = GetSegmentString(Seg1);
                    string seg2Str = GetSegmentString(Seg2);
                    string seg3Str = GetSegmentString(Seg3);
                    string seg4Str = GetSegmentString(Seg4);
                    string seg5Str = GetSegmentString(Seg5);
                    string seg6Str = GetSegmentString(Seg6);
                    string seg7Str = GetSegmentString(Seg7);
                    string seg8Str = GetSegmentString(Seg8);
                    string seg9Str = GetSegmentString(Seg9);
                    string seg10Str = GetSegmentString(Seg10);
                    string seg11Str = GetSegmentString(Seg11);
                    string seg12Str = GetSegmentString(Seg12);
                    string seg13Str = GetSegmentString(Seg13);
                    string seg14Str = GetSegmentString(Seg14);
                    string seg15Str = GetSegmentString(Seg15);
                    string seg16Str = GetSegmentString(Seg16);
                    string seg17Str = GetSegmentString(Seg17);
                    string seg18Str = GetSegmentString(Seg18);
                    string seg19Str = GetSegmentString(Seg19);
                    string seg20Str = GetSegmentString(Seg20);


                    var segmentValues = new[]
                    {
                        seg1Str, seg2Str, seg3Str, seg4Str, seg5Str, seg6Str, seg7Str, seg8Str, seg9Str, seg10Str,
                        seg11Str, seg12Str, seg13Str, seg14Str, seg15Str, seg16Str, seg17Str, seg18Str, seg19Str, seg20Str
                    };

                    if (!string.IsNullOrWhiteSpace(segmentValues[0]) && segmentValues[0].Contains(";"))
                    {
                        var combinedSegments = SplitCombinedSegments(segmentValues[0]);
                        for (int i = 0; i < segmentValues.Length && i < combinedSegments.Length; i++)
                        {
                            segmentValues[i] = combinedSegments[i];
                        }
                    }

                    // Summarize non-empty segments only, to avoid dumping all 20 slots on every call
                    string segSummary = string.Join(", ", segmentValues
                        .Select((v, i) => (v, i))
                        .Where(t => !string.IsNullOrWhiteSpace(t.v))
                        .Select(t => $"Seg{t.i + 1}={t.v}"));
                    LogUtility.LogDebug($"GLSense_GetBalance segments: {(string.IsNullOrEmpty(segSummary) ? "<none>" : segSummary)}");

                    var parametersList = new List<object>
                    {
                        ChangeSign, LedgerName, Activity, Period, BalanceType,
                        CurrencyCode, TranslatedFlag, ActualFlag, BudorEncName, JESource, JECategory
                    };

                    parametersList.AddRange(segmentValues);
                    object[] parameters = parametersList.ToArray();

                    // 2. Validate AND set defaults in-place
                    object? validationResult = ValidateInputs(parameters);
                    if (validationResult != null)
                    {
                        asyncCallObject.ReturnResult(validationResult!);   // or ?? fallback if needed
                        return;
                    }

                    string formulaCompressed = FormulaCacheString(parameters);
                    // Check cache first (if logged out)
                    if (!AppState.Instance.IsLoginCompleted)
                    {
                        asyncCallObject.ReturnResult(HandleNotLoggedIn(formulaCompressed));
                        return;
                    }

                    GetCellCallerAddress();

                    List<string> balanceParameters = GetBalanceParameters(parameters);

                    // ────────────────────────────────────────────────
                    // Potentially slow part (cache miss → API) → background
                    // ────────────────────────────────────────────────
                    Task.Run(async () =>
                    {
                        try
                        {
                            object rawResult;

                            if (AppState.Instance.SingleRefresh && !AppState.Instance.StartBatchCalc)
                            {
                                // Await the async version
                                rawResult = await ExecuteSingleRefreshAsync(formulaCompressed, balanceParameters);
                            }
                            else if (AppState.Instance.StartBatchCalc)
                            {
                                // Assuming batch is sync and reasonably fast
                                // If batch is also long-running → make it async too
                                rawResult = HandleBatchCalculation(formulaCompressed);
                            }
                            else
                            {
                                // Cache read – usually very fast
                                rawResult = HandleCachedResult(formulaCompressed);
                            }

                            // Your number → double conversion logic
                            object finalResult = rawResult switch
                            {
                                double d => d,
                                decimal m => (double)m,
                                int i => (double)i,
                                long l => (double)l,
                                float f => (double)f,
                                null => ADXExcelError.xlErrorNull.ToString(),
                                string s when s.StartsWith("#Err:") => s,
                                _ => rawResult ?? GLClicktoRefresh
                            };

                            LogUtility.LogDebug($"GLSense_GetBalance result: {finalResult}");
                            asyncCallObject.ReturnResult(finalResult);
                        }
                        catch (Exception ex)
                        {
                            LogUtility.LogException(ex, "Error in GLSense_GetBalance background");
                            asyncCallObject.ReturnResult(GLClicktoRefresh);
                        }
                    });
                }

                catch (Exception ex)
                {
                    LogUtility.LogException(ex, "GLSense_GetBalance Function : Unexpected error");
                    asyncCallObject.ReturnResult(GLClicktoRefresh);
                }
            }

#pragma warning restore S107   //This is used to suppress the warning of using 31 paramters in the function below
            #endregion GLSense Excel Functions
        }
        #endregion
    }
#nullable restore
}