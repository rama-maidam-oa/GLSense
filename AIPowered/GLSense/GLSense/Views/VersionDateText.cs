// VersionDateText.cs in GLSense\Views
using System;
using System.Globalization;

namespace GLSense
{
    /// <summary>
    /// OISR-22872: version dates are stored as "yyyy-MM-ddTHH:mm:ss" (manifest.json /
    /// ReleaseHistory.json); windows show them like the About window's Build Date.
    /// Display only - comparisons and storage keep the raw value.
    /// </summary>
    internal static class VersionDateText
    {
        private static readonly string[] DateTimeFormats = { "yyyy-MM-ddTHH:mm:ss", "yyyy-MM-ddTHH:mm:ssZ" };
        private static readonly string[] DateOnlyFormats = { "dd-MMM-yyyy", "yyyy-MM-dd" };

        public static string Format(string raw)
        {
            string value = raw?.Trim();
            if (string.IsNullOrEmpty(value)) return string.Empty;

            if (DateTime.TryParseExact(value, DateTimeFormats, CultureInfo.InvariantCulture, DateTimeStyles.None, out var dt))
                return Format(dt);
            if (DateTime.TryParseExact(value, DateOnlyFormats, CultureInfo.InvariantCulture, DateTimeStyles.None, out var d))
                return d.ToString("dd-MMM-yyyy", CultureInfo.InvariantCulture);
            return value;
        }

        public static string Format(DateTime value) =>
            value.ToString("dd-MMM-yyyy hh:mm tt", CultureInfo.InvariantCulture);
    }
}
