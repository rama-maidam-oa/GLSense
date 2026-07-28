// LedgerQueryModels.cs in GLSense.Addin.Core
// Right-sized carve-out from GLSense\Models\AllModels.cs (FinalWorkingCode) - only the
// JSON DTOs needed to deserialize the ledger-setup-data API response consumed by
// CommonFunctions.FillResponsibilitiesAsync and persisted by
// Repositories\LedgerDataRepository.cs (both Group B / Cube-Ledger selection). No logic
// changes vs. the original - these are plain data holders.
using System.Text.Json.Serialization;

namespace GLSense.Addin.Core.Models
{
    public class LedgerQueryData
    {
        [JsonPropertyName("records")]
        public Records records { get; set; }

        [JsonPropertyName("status")]
        public string status { get; set; }

        [JsonPropertyName("msg")]
        public string msg { get; set; }
    }

    public class Records
    {
        [JsonPropertyName("journalsources")]
        public JESources[] journalsources { get; set; }

        [JsonPropertyName("journalcategories")]
        public JECategories[] journalcategories { get; set; }

        [JsonPropertyName("activity")]
        public string[] activity { get; set; }

        [JsonPropertyName("ledgers")]
        public Ledgers ledgers { get; set; }

        [JsonPropertyName("encumbrances")]
        public Encumbrance[] encumbrances { get; set; }

        [JsonPropertyName("currencies")]
        public string[] currencies { get; set; }
    }

    public class Period
    {
        [JsonPropertyName("periodName")]
        public string periodName { get; set; }

        [JsonPropertyName("periodYear")]
        public int periodYear { get; set; }

        [JsonPropertyName("periodNum")]
        public int periodNum { get; set; }

        [JsonPropertyName("quarterNum")]
        public int quarterNum { get; set; }

        [JsonPropertyName("periodSetName")]
        public string periodSetName { get; set; }

        [JsonPropertyName("periodType")]
        public string periodType { get; set; }

        [JsonPropertyName("startDate")]
        public long startDate { get; set; }

        [JsonPropertyName("endDate")]
        public long endDate { get; set; }

        [JsonPropertyName("adjustmentPeriodFlag")]
        public string adjustmentPeriodFlag { get; set; }
    }

    public class LedgerSegmentValue
    {
        [JsonPropertyName("segmentValue")]
        public string segmentValue { get; set; }

        [JsonPropertyName("description")]
        public string description { get; set; }

        [JsonPropertyName("summaryFlag")]
        public string summaryFlag { get; set; }

        [JsonPropertyName("enabledFlag")]
        public string enabledFlag { get; set; }

        [JsonPropertyName("segmentValueSetId")]
        public long segmentValueSetId { get; set; }
    }

    public class LedgerSegment
    {
        [JsonPropertyName("coaid")]
        public int coaid { get; set; }

        [JsonPropertyName("segmentName")]
        public string segmentName { get; set; }

        [JsonPropertyName("segmentValueSetId")]
        public long segmentValueSetId { get; set; }

        [JsonPropertyName("securityEnabledFlag")]
        public string securityEnabledFlag { get; set; }

        [JsonPropertyName("defaultType")]
        public string defaultType { get; set; }

        [JsonPropertyName("defaultValue")]
        public string defaultValue { get; set; }

        [JsonPropertyName("displaySize")]
        public int displaySize { get; set; }

        [JsonPropertyName("segmentDelimiter")]
        public string segmentDelimiter { get; set; }

        [JsonPropertyName("segmentValues")]
        public LedgerSegmentValue[] segmentValues { get; set; }

        [JsonPropertyName("applicationColumnName")]
        public string applicationColumnName { get; set; }
    }

    public class LedgerData
    {
        [JsonPropertyName("budgets")]
        public string[] budgets { get; set; }

        [JsonPropertyName("periods")]
        public Period[] periods { get; set; }

        [JsonPropertyName("segments")]
        public LedgerSegment[] segments { get; set; }
    }

    public class Ledgers
    {
        [JsonPropertyName("ledgerId")]
        public long ledgerId { get; set; }

        [JsonPropertyName("ledgerName")]
        public string ledgerName { get; set; }

        [JsonPropertyName("coaid")]
        public int coaid { get; set; }

        [JsonPropertyName("periodSetName")]
        public string periodSetName { get; set; }

        [JsonPropertyName("currencyCode")]
        public string currencyCode { get; set; }

        [JsonPropertyName("periodType")]
        public string periodType { get; set; }

        [JsonPropertyName("ledgerData")]
        public LedgerData ledgerData { get; set; }
    }

    public class Encumbrance
    {
        [JsonPropertyName("encumbrancTypeId")]
        public long encumbranceTypeId { get; set; }

        [JsonPropertyName("encumbranceType")]
        public string encumbranceType { get; set; }
    }

    public class JESources
    {
        [JsonPropertyName("jeSourceName")]
        public string jeSourceName { get; set; }

        [JsonPropertyName("sourceName")]
        public string sourceName { get; set; }
    }

    public class JECategories
    {
        [JsonPropertyName("jeCategoryName")]
        public string jeCategoryName { get; set; }

        [JsonPropertyName("categoryName")]
        public string categoryName { get; set; }
    }
}
