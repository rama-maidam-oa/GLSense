// UdfSentinels.cs in GLSense.Contracts
//
// Part of Task #16 (Wire UDFs / GLSenseExcelFunctions to ExecuteUdf).
//
// GLSense.Addin.Core.AddinEntry.ExecuteUdf(string functionName, object[] args) crosses the
// AppDomain boundary as a plain object return value. The old monolith's UDF bodies
// (GLSenseExcelFunctions.cs, FinalWorkingCode) frequently return the ADX-specific
// AddinExpress.MSO.ADXExcelError enum (xlErrorGettingData / xlErrorNull) directly to Excel.
// Addin.Core must NOT reference AddinExpress.MSO (it has no ADX dependency at all - only the
// never-reloaded host project references ADX types), so it cannot return that enum itself.
//
// Instead, Addin.Core's ExecuteUdf returns one of these plain string sentinels wherever the
// old code would have returned/set an ADXExcelError value, and the host's thin UDF wrappers
// (GLSenseExcelFunctions.cs) translate the sentinel back into the real
// AddinExpress.MSO.ADXExcelError enum value before handing the result to Excel (either as a
// direct return for the 13 synchronous UDFs, or via asyncCallObject.ReturnResult(...) for the
// 3 async UDFs: GLSense_GetSegmentDFF, GLSense_GetDailyRate, GLSense_GetBalance).
//
// Any ordinary string result that happens to collide with one of these exact sentinel values
// is not a realistic concern here - they're deliberately distinctive and never produced by
// the old code's normal formatted output (period names, segment values, balances, etc.).
namespace GLSense.Contracts
{
    public static class UdfSentinels
    {
        /// <summary>
        /// Maps to AddinExpress.MSO.ADXExcelError.xlErrorGettingData. Used wherever the old
        /// UDF bodies returned that enum value directly (cache misses while logged out,
        /// unresolvable segment/period lookups, etc.).
        /// </summary>
        public const string XlErrorGettingData = "#GLSENSE_XL_ERROR_GETTING_DATA#";

        /// <summary>
        /// Maps to AddinExpress.MSO.ADXExcelError.xlErrorNull. Used only by
        /// GLSense_GetBalance's "rawResult is null" case.
        /// </summary>
        public const string XlErrorNull = "#GLSENSE_XL_ERROR_NULL#";
    }
}
