using GLSense.Models;
using System;
using System.Globalization;
using System.Windows.Data;

namespace GLSense.Converters
{
    public class TitleRowValueConverter : IMultiValueConverter
    {
        public object Convert(object[] values, Type targetType, object parameter, CultureInfo culture)
        {
            var row = values[0] as ISegmentRow;
            if (row == null) return string.Empty;

            // Show 'Title' text for title rows, else show segmentValue
            var summaryFlagProp = row.GetType().GetProperty("SummaryFlag");
            var summaryFlag = summaryFlagProp?.GetValue(row)?.ToString() ?? string.Empty;

            if (string.Equals(summaryFlag, "RG", StringComparison.OrdinalIgnoreCase))
            {
                var titleProp = row.GetType().GetProperty("Title");
                return titleProp?.GetValue(row)?.ToString() ?? string.Empty;
            }

            var segValProp = row.GetType().GetProperty("SegmentValue");
            return segValProp?.GetValue(row)?.ToString() ?? string.Empty;
        }

        public object[] ConvertBack(object value, Type[] targetTypes, object parameter, CultureInfo culture)
        {
            throw new NotImplementedException();
        }
    }
}
