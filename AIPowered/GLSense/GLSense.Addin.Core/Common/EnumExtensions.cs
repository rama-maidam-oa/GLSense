// EnumExtensions.cs in GLSense.Addin.Core
// Ported verbatim from GLSense\Common\EnumExtensions.cs (FinalWorkingCode).
// No re-pointing needed - fully self-contained (no logging/service-locator/COM usage).
// Namespace changed from GLSense.Common -> GLSense.Addin.Core.Common to match this
// project's layout.
//
// Added as a direct dependency of DDDatatoWorksheet.cs (Drilldowns\DDDatatoWorksheet.cs),
// via DrilldownMetadata.GetDisplay(DrilldownType). Note: the sibling drilldown entry
// points (DD_BL.cs/DD_JL.cs/DD_SL.cs) reference the same DrilldownType/DrilldownMetadata/
// DrilldownHelpers types in the old codebase - if those files are being ported in
// parallel, reconcile so this file exists only once in the project.
using System;
using System.Collections.Concurrent;
using System.ComponentModel;
using System.Linq;

namespace GLSense.Addin.Core.Common
{
    public static class EnumExtensions
    {
        // Cache by (Type, Name) to avoid boxing Enum keys
        private static readonly ConcurrentDictionary<(Type type, string name), string> _descriptionCache = new();

        /// <summary>
        /// Gets the DescriptionAttribute value for an enum value, or the enum name if not present.
        /// </summary>
        public static string GetDescription(this Enum value)
        {
            var key = (value.GetType(), value.ToString());
            return _descriptionCache.GetOrAdd(key, _ =>
            {
                var member = value.GetType().GetMember(value.ToString());
                var attr = member.Length > 0
                    ? member[0].GetCustomAttributes(typeof(DescriptionAttribute), false)
                              .OfType<DescriptionAttribute>()
                              .FirstOrDefault()
                    : null;
                return attr?.Description ?? value.ToString();
            });
        }
    }
}
