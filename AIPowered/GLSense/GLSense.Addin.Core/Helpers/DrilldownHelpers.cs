// DrilldownHelpers.cs in GLSense.Addin.Core
// Ported verbatim from GLSense\Helpers\DrilldownHelpers.cs (FinalWorkingCode).
// Namespace changed from GLSense.Helpers -> GLSense.Addin.Core.Helpers; the
// GLSense.Common using now points to GLSense.Addin.Core.Common (DrilldownType).
//
// Added as a direct dependency of DDDatatoWorksheet.cs. See the note in
// Common\EnumExtensions.cs about reconciling with the parallel DD_BL/DD_JL/DD_SL port.
using GLSense.Addin.Core.Common;
using System;

namespace GLSense.Addin.Core.Helpers
{
    public static class DrilldownHelpers
    {
        public static bool TryParse(string input, out DrilldownType result)
        {
            input = (input ?? string.Empty).Trim();
            return Enum.TryParse(input, ignoreCase: true, out result);
        }

        public static DrilldownType ParseOrDefault(string input, DrilldownType fallback = DrilldownType.BL)
        {
            return TryParse(input, out var dd) ? dd : fallback;
        }

    }
}
