// UdfModels.cs in GLSense.Addin.Core
// Port of GLSense\Models\AllModels.cs (FinalWorkingCode, ~lines 408-486) - SegmentDff /
// DailyRateQuery / DailyRateRecord. Small REST request/response POCOs used only by
// Udf\UdfDispatcher.cs's GLSense_GetSegmentDFF / GLSense_GetDailyRate handlers. Plain data
// holders, no logic changes vs. the original beyond the namespace.
namespace GLSense.Addin.Core.Models
{
    /// <summary>Request payload for the /rest/secure/finance/segment-dff-value endpoint.</summary>
    public class SegmentDff
    {
        public string attributeName { get; set; }
        public string segmentValue { get; set; }
        public long segmentValueSetId { get; set; }
    }

    /// <summary>Request payload for the /rest/secure/finance/gldaily-rates endpoint.</summary>
    public class DailyRateQuery
    {
        public string fromCurrency { get; set; }
        public string toCurrency { get; set; }
        public string conversionType { get; set; }
        public string conversionDate { get; set; }
    }

    /// <summary>One record of the gldaily-rates response.</summary>
    public class DailyRateRecord
    {
        public double? CONVERSION_RATE { get; set; }
    }
}
