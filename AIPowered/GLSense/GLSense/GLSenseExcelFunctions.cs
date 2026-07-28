using System;
using System.Reflection;
using System.Runtime.InteropServices;
using AddinExpress.MSO;
using GLSense.Contracts;

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

        // ═══════════════════════════════════════════════════════════════════════════════
        // Task #16 (Wire UDFs to ExecuteUdf) - architecture notes
        // ═══════════════════════════════════════════════════════════════════════════════
        // These 16 functions are thin, host-only wrappers around
        // GlobalsEx.Addin?.ExecuteUdf(functionName, args) (GLSense.Addin.Core.AddinEntry -
        // hot-reloadable, holds ALL the real business logic: formula-result caching via
        // FormulaCacheManager/SQLite, period/segment lookups via DataServiceLocator, REST
        // calls via ApiHelper). Each wrapper keeps host-only/ADX-only pieces here and crosses
        // the AppDomain boundary with only primitives (string/double/bool/null), per the
        // established convention used by every other ribbon/event dispatch in this migration:
        //
        //   - ValidateInputs(object[]): uses `new StackTrace().GetFrame(1)?.GetMethod()` to
        //     read the CALLING method's own [ExcelParam] attributes, so it only works when
        //     called directly from the actual wrapper method below (frame 1 = the wrapper).
        //     It fills in default values / returns a mandatory-parameter error string in place,
        //     exactly as the old monolith's GLSenseExcelFunctions.cs did. This must stay here -
        //     it would silently break (StackTrace would point at ExecuteUdf's method info
        //     instead) if moved into Addin.Core.
        //
        //   - GetCellCallerAddress(): uses the ADX-only
        //     Module.CallWorksheetFunction(ADXExcelWorksheetFunction.Caller) API (only
        //     available inside this ADXXLLModule) to get GLSense_GetBalance's own cell address
        //     in R1C1 form. Addin.Core has no ADX reference at all, so this value is computed
        //     here and passed across as an extra trailing argument.
        //
        //   - Ledger-name resolution ("use the currently selected ledger if the Ledger
        //     parameter was omitted") is NOT done here anymore (the old monolith's
        //     ResolveLedger(...) read AppState.Instance.SelectedLedger, and AppState now lives
        //     in GLSense.Addin.Core - the host must never reference that assembly directly, or
        //     hot-reload isolation breaks). ValidateInputs still fills an omitted/Missing
        //     Ledger parameter with an empty string (see GetDefaultValue("DefaultLedgerName")
        //     below); Addin.Core's ExecuteUdf resolves "" to the actual selected ledger name
        //     itself, using its own copy of ResolveLedger(object) against its own AppState.
        //     This is a deliberate, documented split of a single old helper method across the
        //     boundary - functionally identical to the old behavior.
        //
        //   - Wherever the old code returned AddinExpress.MSO.ADXExcelError directly,
        //     Addin.Core (which cannot reference AddinExpress.MSO) returns one of the plain
        //     string sentinels in GLSense.Contracts.UdfSentinels instead. TranslateSentinel(...)
        //     below converts it back before the value reaches Excel.
        //
        // Per-function argument-array shapes (documented on each wrapper) intentionally mirror
        // the old monolith's own internal `parameters` object[] arrays, so Addin.Core's
        // ExecuteUdf switch can reuse the exact same positional logic the old UDF bodies used.
        // ═══════════════════════════════════════════════════════════════════════════════

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

            #region Cross-AppDomain dispatch helpers

            private static object CrossToExecuteUdf(string functionName, object[] args)
            {
                var addin = GlobalsEx.Addin;
                if (addin == null)
                {
                    // Mid hot-reload swap, or add-in not yet initialized. There is no sensible
                    // cached value to fall back to from the host, so surface the same "can't
                    // get data right now" error Excel already understands from the old code.
                    GlobalsEx.Context?.Logger?.LogWarn($"{functionName}: GlobalsEx.Addin is null (add-in not loaded) - returning xlErrorGettingData.");
                    return AddinExpress.MSO.ADXExcelError.xlErrorGettingData;
                }

                try
                {
                    object result = addin.ExecuteUdf(functionName, args);
                    object translated = TranslateSentinel(result);
                    GlobalsEx.Context?.Logger?.LogDebug($"{functionName}: ExecuteUdf returned \'{result}\' (translated: \'{translated}\').");
                    return translated;
                }
                catch (Exception ex)
                {
                    GlobalsEx.Context?.Logger?.LogException(ex, $"{functionName}: ExecuteUdf threw (host wrapper)");
                    return AddinExpress.MSO.ADXExcelError.xlErrorGettingData;
                }
            }

            private static object TranslateSentinel(object result)
            {
                if (result is string s)
                {
                    if (s == UdfSentinels.XlErrorGettingData)
                        return AddinExpress.MSO.ADXExcelError.xlErrorGettingData;
                    if (s == UdfSentinels.XlErrorNull)
                        return AddinExpress.MSO.ADXExcelError.xlErrorNull;
                }
                return result;
            }

            #endregion

            #region GLSense Excel Function Helpers (host-only: reflection / ADX-Module-specific)

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

            // Global defaults as static readonly fields with unique keys.
            // NOTE: unlike the old monolith's ExcelDefaults, there is deliberately no
            // DefaultLedgerName field here anymore - see the "Ledger-name resolution" note
            // above. GetDefaultValue("DefaultLedgerName") below returns "" instead, and
            // Addin.Core resolves the actual selected-ledger fallback on its side.
            public static class ExcelDefaults
            {
                public static readonly object DefaultOffset = 0;
                public static readonly object DefaultAdjacentPeriods = true;
                public static readonly object DefaultIncludeId = false;
                public static readonly object DefaultNextParent = false;
                public static readonly object DefaultNextChild = false;
                public static readonly object GLDefaultText = "";
            }

            private static string GetMandatoryError(string paramName)
            {
                string lowerName = paramName?.ToLower() ?? string.Empty;

                if (lowerName == "ledger") return "#Error: Missing Ledger";
                if (lowerName == "period" || lowerName == "perioddate" ||
                    lowerName == "periodyear" || lowerName == "periodnum")
                    return "#Error: Missing Period";
                if (lowerName == "segmentname") return "#Error: Missing SegmentName";
                if (lowerName == "segmentvalue") return "#Error: Missing SegmentValue";

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
                    // Resolved Addin.Core-side (AppState.Instance.SelectedLedger) - see notes above.
                    "DefaultLedgerName" => string.Empty,
                    _ => null
                };
            }

            private static bool ToBool(object? value, bool defaultValue = false, double tolerance = 1e-6)
            {
                if (value == null || value is System.Reflection.Missing || value is AddinExpress.MSO.ADXExcelError)
                    return defaultValue;

                if (value is bool b)
                    return b;

                if (value is double d)
                {
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

                if (bool.TryParse(value.ToString(), out var parsed))
                    return parsed;

                return defaultValue;
            }

            /// <summary>
            /// Ported verbatim from the old monolith. IMPORTANT: relies on
            /// `new StackTrace().GetFrame(1)` to read the CALLING method's own [ExcelParam]
            /// attributes - only call this directly from one of the 16 UDF wrapper methods
            /// below, never from a helper (the frame index would then be wrong).
            /// </summary>
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

            private static string? XLLR1C1;

            /// <summary>
            /// Ported verbatim from the old monolith. ADX-only (Module.CallWorksheetFunction)
            /// - only meaningful for GLSense_GetBalance, which needs its own cell address to
            /// build the REST payload's "excelCell" field.
            /// </summary>
            private static void GetCellCallerAddress()
            {
                try
                {
                    if (Module.CallWorksheetFunction(AddinExpress.MSO.ADXExcelWorksheetFunction.Caller) is not AddinExpress.MSO.ADXExcelRef caller)
                    {
                        GlobalsEx.Context?.Logger?.LogWarn("GetCellCallerAddress: Caller returned null");
                        return;
                    }

                    int rowFirst = caller.RowFirst + 1;
                    int columnFirst = caller.ColumnFirst + 1;

                    XLLR1C1 = "R" + rowFirst.ToString() + "C" + columnFirst.ToString();
                }
                catch (Exception ex)
                {
                    GlobalsEx.Context?.Logger?.LogException(ex, "GetCellCallerAddress");
                }
            }

            #endregion

            #region GLSense Excel Functions

            // args = [PeriodDate, Ledger, offset] (Ledger unresolved - see notes above)
            public static object GLSense_GetPeriodByDate(
                [ExcelParam(true)] object PeriodDate,
                [ExcelParam(false, "DefaultLedgerName")] object? Ledger = null,
                [ExcelParam(false, "DefaultOffset")] object? offset = null)
            {
                try
                {
                    GlobalsEx.Context?.Logger?.LogDebug($"GLSense_GetPeriodByDate called (PeriodDate={PeriodDate}, Ledger={Ledger}, offset={offset})");
                    object[] parameters = new object[] { PeriodDate, Ledger!, offset! };
                    object? validationResult = ValidateInputs(parameters);
                    if (validationResult != null) return validationResult;

                    return CrossToExecuteUdf("GLSense_GetPeriodByDate", parameters);
                }
                catch (Exception ex)
                {
                    GlobalsEx.Context?.Logger?.LogException(ex, "GLSense_GetPeriodByDate Function : Unexpected error (host wrapper)");
                    return AddinExpress.MSO.ADXExcelError.xlErrorGettingData;
                }
            }

            // args = [Period, offset, Ledger] (Ledger unresolved)
            public static object GLSense_GetPeriod(
                [ExcelParam(true)] object Period,
                [ExcelParam(false, "DefaultOffset")] object? offset = null,
                [ExcelParam(false, "DefaultLedgerName")] object? Ledger = null)
            {
                try
                {
                    GlobalsEx.Context?.Logger?.LogDebug($"GLSense_GetPeriod called (Period={Period}, offset={offset}, Ledger={Ledger})");
                    object[] parameters = new object[] { Period, offset!, Ledger! };
                    object? validationResult = ValidateInputs(parameters);
                    if (validationResult != null) return validationResult;

                    return CrossToExecuteUdf("GLSense_GetPeriod", parameters);
                }
                catch (Exception ex)
                {
                    GlobalsEx.Context?.Logger?.LogException(ex, "GLSense_GetPeriod Function : Unexpected error (host wrapper)");
                    return AddinExpress.MSO.ADXExcelError.xlErrorGettingData;
                }
            }

            // args = [PeriodYear, PeriodNum, Ledger] (Ledger unresolved)
            public static object GLSense_GetPeriodByYear(
                [ExcelParam(true)] object PeriodYear,
                [ExcelParam(true)] object PeriodNum,
                [ExcelParam(false, "DefaultLedgerName")] object? Ledger = null)
            {
                try
                {
                    GlobalsEx.Context?.Logger?.LogDebug($"GLSense_GetPeriodByYear called (PeriodYear={PeriodYear}, PeriodNum={PeriodNum}, Ledger={Ledger})");
                    object[] parameters = new object[] { PeriodYear, PeriodNum, Ledger! };
                    object? validationResult = ValidateInputs(parameters);
                    if (validationResult != null) return validationResult;

                    return CrossToExecuteUdf("GLSense_GetPeriodByYear", parameters);
                }
                catch (Exception ex)
                {
                    GlobalsEx.Context?.Logger?.LogException(ex, "GLSense_GetPeriodByYear Function : Unexpected error (host wrapper)");
                    return AddinExpress.MSO.ADXExcelError.xlErrorGettingData;
                }
            }

            // args = [Period, Ledger, adjacentPeriodsBool] (Ledger unresolved)
            public static object GLSense_GetPeriodStart(
                [ExcelParam(true)] object Period,
                [ExcelParam(false, "DefaultLedgerName")] object? Ledger = null,
                [ExcelParam(false, "DefaultAdjacentPeriods")] object? AdjacentPeriods = null)
            {
                try
                {
                    GlobalsEx.Context?.Logger?.LogDebug($"GLSense_GetPeriodStart called (Period={Period}, Ledger={Ledger}, AdjacentPeriods={AdjacentPeriods})");
                    bool adjacentPeriodsBool = ToBool(AdjacentPeriods, false);
                    object[] parameters = new object[] { Period, Ledger!, adjacentPeriodsBool };
                    object? validationResult = ValidateInputs(parameters);
                    if (validationResult != null) return validationResult;

                    return CrossToExecuteUdf("GLSense_GetPeriodStart", parameters);
                }
                catch (Exception ex)
                {
                    GlobalsEx.Context?.Logger?.LogException(ex, "GLSense_GetPeriodStart Function : Unexpected error (host wrapper)");
                    return AddinExpress.MSO.ADXExcelError.xlErrorGettingData;
                }
            }

            // args = [Period, Ledger, adjacentPeriodsBool] (Ledger unresolved)
            public static object GLSense_GetPeriodEnd(
                [ExcelParam(true)] object Period,
                [ExcelParam(false, "DefaultLedgerName")] object? Ledger = null,
                [ExcelParam(false, "DefaultAdjacentPeriods")] object? AdjacentPeriods = null)
            {
                try
                {
                    GlobalsEx.Context?.Logger?.LogDebug($"GLSense_GetPeriodEnd called (Period={Period}, Ledger={Ledger}, AdjacentPeriods={AdjacentPeriods})");
                    bool adjacentPeriodsBool = ToBool(AdjacentPeriods, false);
                    object[] parameters = new object[] { Period, Ledger!, adjacentPeriodsBool };
                    object? validationResult = ValidateInputs(parameters);
                    if (validationResult != null) return validationResult;

                    return CrossToExecuteUdf("GLSense_GetPeriodEnd", parameters);
                }
                catch (Exception ex)
                {
                    GlobalsEx.Context?.Logger?.LogException(ex, "GLSense_GetPeriodEnd Function : Unexpected error (host wrapper)");
                    return AddinExpress.MSO.ADXExcelError.xlErrorGettingData;
                }
            }

            // args = [Period, Ledger] (Ledger unresolved)
            public static object GLSense_GetPeriodNum(
                [ExcelParam(true)] object Period,
                [ExcelParam(false, "DefaultLedgerName")] object? Ledger = null)
            {
                try
                {
                    GlobalsEx.Context?.Logger?.LogDebug($"GLSense_GetPeriodNum called (Period={Period}, Ledger={Ledger})");
                    object[] parameters = new object[] { Period, Ledger! };
                    object? validationResult = ValidateInputs(parameters);
                    if (validationResult != null) return validationResult;

                    return CrossToExecuteUdf("GLSense_GetPeriodNum", parameters);
                }
                catch (Exception ex)
                {
                    GlobalsEx.Context?.Logger?.LogException(ex, "GLSense_GetPeriodNum Function : Unexpected error (host wrapper)");
                    return AddinExpress.MSO.ADXExcelError.xlErrorGettingData;
                }
            }

            // args = [Period, Ledger] (Ledger unresolved)
            public static object GLSense_GetPeriodQuarter(
                [ExcelParam(true)] object Period,
                [ExcelParam(false, "DefaultLedgerName")] object? Ledger = null)
            {
                try
                {
                    GlobalsEx.Context?.Logger?.LogDebug($"GLSense_GetPeriodQuarter called (Period={Period}, Ledger={Ledger})");
                    object[] parameters = new object[] { Period, Ledger! };
                    object? validationResult = ValidateInputs(parameters);
                    if (validationResult != null) return validationResult;

                    return CrossToExecuteUdf("GLSense_GetPeriodQuarter", parameters);
                }
                catch (Exception ex)
                {
                    GlobalsEx.Context?.Logger?.LogException(ex, "GLSense_GetPeriodQuarter Function : Unexpected error (host wrapper)");
                    return AddinExpress.MSO.ADXExcelError.xlErrorGettingData;
                }
            }

            // args = [Period, Ledger] (Ledger unresolved)
            public static object GLSense_GetPeriodYear(
                [ExcelParam(true)] object Period,
                [ExcelParam(false, "DefaultLedgerName")] object? Ledger = null)
            {
                try
                {
                    GlobalsEx.Context?.Logger?.LogDebug($"GLSense_GetPeriodYear called (Period={Period}, Ledger={Ledger})");
                    object[] parameters = new object[] { Period, Ledger! };
                    object? validationResult = ValidateInputs(parameters);
                    if (validationResult != null) return validationResult;

                    return CrossToExecuteUdf("GLSense_GetPeriodYear", parameters);
                }
                catch (Exception ex)
                {
                    GlobalsEx.Context?.Logger?.LogException(ex, "GLSense_GetPeriodYear Function : Unexpected error (host wrapper)");
                    return AddinExpress.MSO.ADXExcelError.xlErrorGettingData;
                }
            }

            // args = [SegmentValue, SegmentName, includeIdBool, Ledger] (Ledger unresolved)
            public static object GLSense_GetSegmentDesc(
                [ExcelParam(true)] object SegmentValue,
                [ExcelParam(true)] object SegmentName,
                [ExcelParam(false, "DefaultIncludeId")] object? IncludeId = null,
                [ExcelParam(false, "DefaultLedgerName")] object? Ledger = null)
            {
                try
                {
                    GlobalsEx.Context?.Logger?.LogDebug($"GLSense_GetSegmentDesc called (SegmentValue={SegmentValue}, SegmentName={SegmentName}, IncludeId={IncludeId}, Ledger={Ledger})");
                    bool includeIdBool = ToBool(IncludeId, false);
                    object[] parameters = new object[] { SegmentValue, SegmentName, includeIdBool, Ledger! };
                    object? validationResult = ValidateInputs(parameters);
                    if (validationResult != null) return validationResult;

                    return CrossToExecuteUdf("GLSense_GetSegmentDesc", parameters);
                }
                catch (Exception ex)
                {
                    GlobalsEx.Context?.Logger?.LogException(ex, "GLSense_GetSegmentDesc Function : Unexpected error (host wrapper)");
                    return AddinExpress.MSO.ADXExcelError.xlErrorGettingData;
                }
            }

            // args = [SegmentValue, SegmentName, Ledger] (Ledger unresolved)
            public static object GLSense_GetSegmentEnabledFlag(
                [ExcelParam(true)] object SegmentValue,
                [ExcelParam(true)] object SegmentName,
                [ExcelParam(false, "DefaultLedgerName")] object? Ledger = null)
            {
                try
                {
                    GlobalsEx.Context?.Logger?.LogDebug($"GLSense_GetSegmentEnabledFlag called (SegmentValue={SegmentValue}, SegmentName={SegmentName}, Ledger={Ledger})");
                    object[] parameters = new object[] { SegmentValue, SegmentName, Ledger! };
                    object? validationResult = ValidateInputs(parameters);
                    if (validationResult != null) return validationResult;

                    return CrossToExecuteUdf("GLSense_GetSegmentEnabledFlag", parameters);
                }
                catch (Exception ex)
                {
                    GlobalsEx.Context?.Logger?.LogException(ex, "GLSense_GetSegmentEnabledFlag Function : Unexpected error (host wrapper)");
                    return AddinExpress.MSO.ADXExcelError.xlErrorGettingData;
                }
            }

            // args = [SegmentValue, SegmentName, Ledger] (Ledger unresolved)
            public static object GLSense_GetSegmentSummaryFlag(
                [ExcelParam(true)] object SegmentValue,
                [ExcelParam(true)] object SegmentName,
                [ExcelParam(false, "DefaultLedgerName")] object? Ledger = null)
            {
                try
                {
                    GlobalsEx.Context?.Logger?.LogDebug($"GLSense_GetSegmentSummaryFlag called (SegmentValue={SegmentValue}, SegmentName={SegmentName}, Ledger={Ledger})");
                    object[] parameters = new object[] { SegmentValue, SegmentName, Ledger! };
                    object? validationResult = ValidateInputs(parameters);
                    if (validationResult != null) return validationResult;

                    return CrossToExecuteUdf("GLSense_GetSegmentSummaryFlag", parameters);
                }
                catch (Exception ex)
                {
                    GlobalsEx.Context?.Logger?.LogException(ex, "GLSense_GetSegmentSummaryFlag Function : Unexpected error (host wrapper)");
                    return AddinExpress.MSO.ADXExcelError.xlErrorGettingData;
                }
            }

            // args = [SegmentValue, SegmentName, NextParent, NextChild, Ledger] (Ledger unresolved)
            // Addin.Core distinguishes Next/Previous by functionName.
            public static object GLSense_GetNextSegment(
                [ExcelParam(true)] object SegmentValue,
                [ExcelParam(true)] object SegmentName,
                [ExcelParam(true)] object NextParent,
                [ExcelParam(true)] object NextChild,
                [ExcelParam(false, "DefaultLedgerName")] object? Ledger = null)
            {
                try
                {
                    GlobalsEx.Context?.Logger?.LogDebug($"GLSense_GetNextSegment called (SegmentValue={SegmentValue}, SegmentName={SegmentName}, NextParent={NextParent}, NextChild={NextChild}, Ledger={Ledger})");
                    object[] parameters = new object[] { SegmentValue, SegmentName, NextParent, NextChild, Ledger! };
                    object? validationResult = ValidateInputs(parameters);
                    if (validationResult != null) return validationResult;

                    return CrossToExecuteUdf("GLSense_GetNextSegment", parameters);
                }
                catch (Exception ex)
                {
                    GlobalsEx.Context?.Logger?.LogException(ex, "GLSense_GetNextSegment Function : Unexpected error (host wrapper)");
                    return AddinExpress.MSO.ADXExcelError.xlErrorGettingData;
                }
            }

            // args = [SegmentValue, SegmentName, NextParent, NextChild, Ledger] (Ledger unresolved)
            public static object GLSense_GetPreviousSegment(
                [ExcelParam(true)] object SegmentValue,
                [ExcelParam(true)] object SegmentName,
                [ExcelParam(true)] object NextParent,
                [ExcelParam(true)] object NextChild,
                [ExcelParam(false, "DefaultLedgerName")] object? Ledger = null)
            {
                try
                {
                    GlobalsEx.Context?.Logger?.LogDebug($"GLSense_GetPreviousSegment called (SegmentValue={SegmentValue}, SegmentName={SegmentName}, NextParent={NextParent}, NextChild={NextChild}, Ledger={Ledger})");
                    object[] parameters = new object[] { SegmentValue, SegmentName, NextParent, NextChild, Ledger! };
                    object? validationResult = ValidateInputs(parameters);
                    if (validationResult != null) return validationResult;

                    return CrossToExecuteUdf("GLSense_GetPreviousSegment", parameters);
                }
                catch (Exception ex)
                {
                    GlobalsEx.Context?.Logger?.LogException(ex, "GLSense_GetPreviousSegment Function : Unexpected error (host wrapper)");
                    return AddinExpress.MSO.ADXExcelError.xlErrorGettingData;
                }
            }

            // ─── Async UDFs (return void, take a trailing required ADXExcelAsyncCallObject) ───
            // The real work runs on a background thread (Task.Run) so the ExecuteUdf call
            // (which blocks synchronously across the AppDomain boundary while it does REST
            // calls / SQLite reads) never blocks Excel's UI thread. asyncCallObject.ReturnResult
            // is safe to call from any thread - ADX marshals it back into Excel itself (this
            // was already true of the old monolith's own BackgroundWorker/Task.Run usage).

            // args = [SegmentValue, SegmentName, Attribute, Ledger] (Ledger unresolved)
            public static void GLSense_GetSegmentDFF(
                [ExcelParam(true)] object SegmentValue,
                [ExcelParam(true)] object SegmentName,
                [ExcelParam(true)] object Attribute,
                [ExcelParam(false, "DefaultLedgerName")] object? Ledger,
                ADXExcelAsyncCallObject asyncCallObject)
            {
                try
                {
                    GlobalsEx.Context?.Logger?.LogDebug($"GLSense_GetSegmentDFF called (SegmentValue={SegmentValue}, SegmentName={SegmentName}, Attribute={Attribute}, Ledger={Ledger})");
                    object[] parameters = new object[] { SegmentValue, SegmentName, Attribute, Ledger! };
                    object? validationResult = ValidateInputs(parameters);
                    if (validationResult != null)
                    {
                        asyncCallObject.ReturnResult(validationResult);
                        return;
                    }

                    System.Threading.Tasks.Task.Run(() =>
                    {
                        try
                        {
                            object result = CrossToExecuteUdf("GLSense_GetSegmentDFF", parameters);
                            asyncCallObject.ReturnResult(result);
                        }
                        catch (Exception ex)
                        {
                            GlobalsEx.Context?.Logger?.LogException(ex, "GLSense_GetSegmentDFF background (host wrapper)");
                            asyncCallObject.ReturnResult(AddinExpress.MSO.ADXExcelError.xlErrorGettingData);
                        }
                    });
                }
                catch (Exception ex)
                {
                    GlobalsEx.Context?.Logger?.LogException(ex, "GLSense_GetSegmentDFF Function : Unexpected error (host wrapper)");
                    asyncCallObject.ReturnResult(AddinExpress.MSO.ADXExcelError.xlErrorGettingData);
                }
            }

            // args = [SegmentValue, SegmentName, Ledger] (Ledger unresolved)
            public static void GLSense_GetAccountType(
                [ExcelParam(true)] object SegmentValue,
                [ExcelParam(true)] object SegmentIndex,
                [ExcelParam(false, "DefaultLedgerName")] object? Ledger,
                ADXExcelAsyncCallObject asyncCallObject)
            {
                try
                {
                    GlobalsEx.Context?.Logger?.LogDebug($"GLSense_GetAccountType called (SegmentValue={SegmentValue}, SegmentIndex={SegmentIndex}, Ledger={Ledger})");
                    object[] parameters = new object[] { SegmentValue, SegmentIndex, Ledger! };
                    object? validationResult = ValidateInputs(parameters);
                    if (validationResult != null)
                    {
                        asyncCallObject.ReturnResult(validationResult);
                        return;
                    }

                    System.Threading.Tasks.Task.Run(() =>
                    {
                        try
                        {
                            object result = CrossToExecuteUdf("GLSense_GetAccountType", parameters);
                            asyncCallObject.ReturnResult(result);
                        }
                        catch (Exception ex)
                        {
                            GlobalsEx.Context?.Logger?.LogException(ex, "GLSense_GetAccountType background (host wrapper)");
                            asyncCallObject.ReturnResult(AddinExpress.MSO.ADXExcelError.xlErrorGettingData);
                        }
                    });
                }
                catch (Exception ex)
                {
                    GlobalsEx.Context?.Logger?.LogException(ex, "GLSense_GetAccountType Function : Unexpected error (host wrapper)");
                    asyncCallObject.ReturnResult(AddinExpress.MSO.ADXExcelError.xlErrorGettingData);
                }
            }

            // args = [FromCurrency, ToCurrency, ConversionType, ConversionDate]
            public static void GLSense_GetDailyRate(
                [ExcelParam(true)] object FromCurrency,
                [ExcelParam(true)] object ToCurrency,
                [ExcelParam(true)] object ConversionType,
                [ExcelParam(true)] object ConversionDate,
                ADXExcelAsyncCallObject asyncCallObject)   // must be LAST
            {
                try
                {
                    GlobalsEx.Context?.Logger?.LogDebug($"GLSense_GetDailyRate called (FromCurrency={FromCurrency}, ToCurrency={ToCurrency}, ConversionType={ConversionType}, ConversionDate={ConversionDate})");
                    object[] parameters = new object[] { FromCurrency, ToCurrency, ConversionType, ConversionDate };
                    object? validationResult = ValidateInputs(parameters);
                    if (validationResult != null)
                    {
                        asyncCallObject.ReturnResult(validationResult);
                        return;
                    }

                    System.Threading.Tasks.Task.Run(() =>
                    {
                        try
                        {
                            object result = CrossToExecuteUdf("GLSense_GetDailyRate", parameters);
                            asyncCallObject.ReturnResult(result);
                        }
                        catch (Exception ex)
                        {
                            GlobalsEx.Context?.Logger?.LogException(ex, "GLSense_GetDailyRate background (host wrapper)");
                            asyncCallObject.ReturnResult(AddinExpress.MSO.ADXExcelError.xlErrorGettingData);
                        }
                    });
                }
                catch (Exception ex)
                {
                    GlobalsEx.Context?.Logger?.LogException(ex, "GLSense_GetDailyRate Function : Unexpected error (host wrapper)");
                    asyncCallObject.ReturnResult(AddinExpress.MSO.ADXExcelError.xlErrorGettingData);
                }
            }

            // args = [ChangeSign, LedgerName, Activity, Period, BalanceType, CurrencyCode,
            //         TranslatedFlag, ActualFlag, BudorEncName, JESource, JECategory,
            //         Seg1..Seg20, XLLR1C1] (33 elements - LedgerName here IS a mandatory raw
            // value, unrelated to the optional-Ledger "DefaultLedgerName" resolution used by
            // the other 12 functions; XLLR1C1 is appended last, computed host-side).
            //
            // NOTE on AppState.Instance.ResetFormulas: the old monolith checked this flag
            // BEFORE building `parameters`/calling ValidateInputs, short-circuiting to
            // "Click Refresh..." immediately. AppState now lives in Addin.Core, so that check
            // has moved there too (first thing GLSense_GetBalance's ExecuteUdf handler does).
            // ValidateInputs (host-only, reflection-based) always runs first here regardless -
            // the only behavior difference is the rare case of an invalid formula while a
            // batch reset is also in progress, which now surfaces the validation error instead
            // of "Click Refresh...". Documented, deliberate, and harmless.
#pragma warning disable S107
            public static void GLSense_GetBalance(
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
                    ADXExcelAsyncCallObject asyncCallObject)   // must be LAST
            {
                try
                {
                    GlobalsEx.Context?.Logger?.LogDebug($"GLSense_GetBalance called (LedgerName={LedgerName}, Activity={Activity}, Period={Period}, BalanceType={BalanceType}, CurrencyCode={CurrencyCode})");
                    object[] parameters = new object[]
                    {
                        ChangeSign, LedgerName, Activity, Period, BalanceType,
                        CurrencyCode, TranslatedFlag, ActualFlag, BudorEncName, JESource, JECategory,
                        Seg1!, Seg2!, Seg3!, Seg4!, Seg5!, Seg6!, Seg7!, Seg8!, Seg9!, Seg10!,
                        Seg11!, Seg12!, Seg13!, Seg14!, Seg15!, Seg16!, Seg17!, Seg18!, Seg19!, Seg20!
                    };

                    object? validationResult = ValidateInputs(parameters);
                    if (validationResult != null)
                    {
                        asyncCallObject.ReturnResult(validationResult);
                        return;
                    }

                    GetCellCallerAddress();

                    object[] crossingArgs = new object[parameters.Length + 1];
                    Array.Copy(parameters, crossingArgs, parameters.Length);
                    crossingArgs[parameters.Length] = XLLR1C1 ?? string.Empty;

                    System.Threading.Tasks.Task.Run(() =>
                    {
                        try
                        {
                            object result = CrossToExecuteUdf("GLSense_GetBalance", crossingArgs);
                            asyncCallObject.ReturnResult(result);
                        }
                        catch (Exception ex)
                        {
                            GlobalsEx.Context?.Logger?.LogException(ex, "GLSense_GetBalance background (host wrapper)");
                            asyncCallObject.ReturnResult("Click Refresh...");
                        }
                    });
                }
                catch (Exception ex)
                {
                    GlobalsEx.Context?.Logger?.LogException(ex, "GLSense_GetBalance Function : Unexpected error (host wrapper)");
                    asyncCallObject.ReturnResult("Click Refresh...");
                }
            }
#pragma warning restore S107

            #endregion
        }

        #endregion
    }
}
