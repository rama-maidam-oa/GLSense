// AppConstants.cs in GLSense.Addin.Core
// Ported from GLSense\AppConstants.cs (FinalWorkingCode). Grow this file as more
// ribbon actions get migrated - only add constants when something actually needs them.
namespace GLSense.Addin.Core
{
    public static class AppConstants
    {
        // Fixed file names (never change)
        public const string RefreshFileName = "BulkRefresh_ORBIT_1234567890.xlsx";
        public const string RefreshZipFileName = "BulkRefresh_ORBIT.zip";
        public const string RefreshJsonFileName = "BulkRefresh_ORBIT.json";
        public const string FormatJsonFileName = "BulkRefresh_FormattedJSON.json";

        public const string UnauthorizedMessage = "Unauthorized";
        public const string ErrorPrefix = "Error";
        public const string OracleErrorPrefix = "ORA-";
        public const string value = "value";
        public const string glBal = "GLSense_GetBalance";
        public const string Success = "success";
        public const string General = "General";
        public const string Text = "@";
        public const string Status = "status";
        public const string Records = "records";
        public const string WebSecure = "/web/secure/";
        // Group H (GLSegmentValues hierarchy combo) addition - port of GLSense\
        // AppConstants.cs RestSecure (FinalWorkingCode). Used by SegmentSelectorViewModel.
        // HierarhyApiAsync to build the /rest/secure/finance/segment-hierarchy API URL.
        public const string RestSecure = "/rest/secure/finance/";
        public const string DrilldownSheetMarkerCellAddress = "XEZ5";

        // Default values
        public const int DefaultSegmentPickedIndex = -1;

        // Logging constants
        public const long LogMaxFileSizeBytes = 20 * 1024 * 1024; // 20 MB
        public const int LogMaxArchiveFiles = 30;

        // Theme and accent
        public const string GLTheme = "Light";
        public const string GLAccentHex = "#2E86AB";

        // Balance Configurator: balance type codes (GLConfiguratorViewModel uses the
        // equivalent hardcoded string literals directly today - these constants are
        // added here for parity with GLSense\AppConstants.cs (FinalWorkingCode) and are
        // not yet wired into GLConfiguratorViewModel.cs; the literal values already
        // match exactly, so this is a documentation/parity addition, not a behavior fix).
        public const string BalanceTypePTD = "PTD";
        public const string BalanceTypeYTD = "YTD";
        public const string BalanceTypeCTD = "CTD";
        public const string BalanceTypeJED = "JED";
        public const string BalanceTypeJEDP = "JEDP";
        public const string BalanceTypeJEDU = "JEDU";

        // Balance Configurator: activity codes (GLConfiguratorViewModel)
        public const string ActivityDR = "DR";
        public const string ActivityCR = "CR";
        public const string ActivityFlagActual = "Actual";

        // Balance Configurator: currency type names (GLConfiguratorViewModel)
        public const string CurrencyTypeTotal = "Total";
        public const string CurrencyTypeEntered = "Entered";

        // Balance Configurator: combined Actual+Encumbrance short code (distinct from
        // the "Actual+Encumbrance" long-form AE constant already private to
        // GLConfiguratorViewModel - this is the abbreviated "A+E" ShortName value).
        public const string ActualEncumbranceShort = "A+E";

        // Reflection property-name lookups (GetType().GetProperty(...) calls)
        public const string PropLedgerName = "LedgerName";
        public const string PropIsSelected = "IsSelected";
        public const string PropShortName = "ShortName";

        // Common date format used across pickers/converters/repositories
        public const string DateFormatIso = "yyyy-MM-dd";
    }
}
