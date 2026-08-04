// DrilldownType.cs in GLSense.Addin.Core
// Ported verbatim from GLSense\Common\DrilldownType.cs (FinalWorkingCode).
// Namespace changed from GLSense.Common -> GLSense.Addin.Core.Common.
//
// Added as a direct dependency of DDDatatoWorksheet.cs. See the note in
// EnumExtensions.cs about reconciling with the parallel DD_BL/DD_JL/DD_SL port.
//
// BLDD_SL/BLDD_UF added in a later pass (FinalWorkingCode's enum has 9 values, this one
// originally only had 7): the ribbon buttons/click handlers/drilldown-execution dispatch for
// "Balances Drilldown to Sub-Ledgers" and "Balances Drilldown to Unified" already exist and
// work in this project (AddinModule.cs's RibBalancesDDToSubLedger/RibBalancesDDToUnified ->
// Drilldowns\DD_JL.cs, using raw "BLDD_SL"/"BLDD_UF" ddType strings) - only the strongly-typed
// enum values were missing, needed by DDDatatoWorksheet.cs's GetLocalMetadataRecordsKey
// (DrilldownType -> local-metadata-override records key mapping).
using System.ComponentModel;

namespace GLSense.Addin.Core.Common
{
    /// <summary>
    /// Supported drilldown types. Annotated with descriptions for display text.
    /// </summary>
    public enum DrilldownType
    {
        [Description("Balances Drilldown")]
        BL,

        [Description("Journals Drilldown")]
        JL,

        [Description("SubLedgers Drilldown")]
        SL,

        [Description("Balances to Journals Drilldown")]
        BL_JL,

        [Description("Balances to Sub-Ledgers Drilldown")]
        BL_SL,

        [Description("Balances Drilldown to Sub-Ledgers Drilldown")]
        BLDD_SL,

        [Description("Balances Drilldown to Unified Drilldown")]
        BLDD_UF,

        [Description("Unified Drilldown")]
        UF,

        [Description("Custom Drilldown")]
        CM
    }
}
