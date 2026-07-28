using System;
using System.Collections.Concurrent;
using System.ComponentModel;
using System.Linq;

namespace GLSense.Common
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
