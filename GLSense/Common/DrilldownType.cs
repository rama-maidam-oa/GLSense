using System.ComponentModel;

namespace GLSense.Common
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
