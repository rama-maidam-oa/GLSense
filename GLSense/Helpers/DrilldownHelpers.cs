using GLSense.Common;
using System;

namespace GLSense.Helpers
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
