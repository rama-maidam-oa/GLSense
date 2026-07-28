// DrilldownType.cs in GLSense.Addin.Core
// Ported verbatim from GLSense\Common\DrilldownType.cs (FinalWorkingCode).
// Namespace changed from GLSense.Common -> GLSense.Addin.Core.Common.
//
// Added as a direct dependency of DDDatatoWorksheet.cs. See the note in
// EnumExtensions.cs about reconciling with the parallel DD_BL/DD_JL/DD_SL port.
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

        [Description("Unified Drilldown")]
        UF,

        [Description("Custom Drilldown")]
        CM
    }
}
