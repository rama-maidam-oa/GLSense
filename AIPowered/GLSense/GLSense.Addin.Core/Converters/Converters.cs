// GLSense.Addin.Core/Converters/Converters.cs
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Windows;
using System.Windows.Data;

namespace GLSense.Addin.Core.Converters
{
    public class EmptyStringToVisibilityConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (value == null || string.IsNullOrEmpty(value.ToString()))
                return Visibility.Collapsed;
            return Visibility.Visible;
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            throw new NotImplementedException();
        }
    }
    public class WidthPercentageConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            double actualWidth = 0;
            if (value != null && double.TryParse(value.ToString(), out actualWidth))
            {
                double percentage = 0.9; // default 90%
                if (parameter != null)
                    double.TryParse(parameter.ToString(), out percentage);
                return actualWidth * percentage;
            }
            return value;
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            throw new NotImplementedException();
        }
    }
    public class NullToVisibilityConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            return value == null ? System.Windows.Visibility.Collapsed : System.Windows.Visibility.Visible;
        }
        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            throw new NotImplementedException();
        }
    }

    public class NullableDateConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (value is DateTime date)
                return date.ToString("dd-MMM-yyyy");
            return string.Empty;
        }


        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            var input = value as string;
            if (string.IsNullOrWhiteSpace(input))
                return null;

            DateTime result;

            // ✅ Try parsing with multiple formats first
            string[] formats = {
                    "dd-MMM-yyyy", "d-MMM-yyyy", "dd/MM/yyyy", "d/M/yyyy",
                    "MM/dd/yyyy", "M/d/yyyy", "dd-MM-yyyy", "d-M-yyyy",
                    "yyyy-MM-dd", "yyyy/MM/dd", "dd MMM yyyy", "d MMM yyyy",
                    "dd.MM.yyyy", "d.M.yyyy"
                };

            if (DateTime.TryParseExact(input, formats, CultureInfo.InvariantCulture,
                                        DateTimeStyles.AllowWhiteSpaces, out result))
            {
                return result; // Parsed successfully
            }

            // ✅ Fallback: Try normal DateTime.Parse for other formats
            if (DateTime.TryParse(input, CultureInfo.InvariantCulture, DateTimeStyles.AllowWhiteSpaces, out result))
            {
                return result;
            }

            return null; // Invalid input
        }

    }
    public class BoolToVisibilityConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (value is bool boolValue)
            {
                return boolValue ? System.Windows.Visibility.Visible : System.Windows.Visibility.Collapsed;
            }
            return System.Windows.Visibility.Collapsed;
        }
        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (value is Visibility visibility)
                return visibility == Visibility.Visible;
            return false;
        }
    }
    // A general-purpose converter to enable/disable dependent fields
    // Example usage: ConverterParameter="ActualFlag:Encumbrance,ConverterParameter="ActualFlag:Budget,ConverterParameter="BalanceType:EndPeriod"
    public class FieldDependencyEnableConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (value == null)
                return false;
            string flag = System.Convert.ToString(value);

            if (string.IsNullOrWhiteSpace(flag))
                return false;
            if (parameter == null)
                return false;

            var paramStr = parameter.ToString().ToLowerInvariant();

            // Handle ActualFlag dependencies
            if (paramStr.Contains("actualflag"))
            {
                if (paramStr.Contains("budget"))
                    return flag.Equals("budget", StringComparison.OrdinalIgnoreCase) || flag.Equals("b", StringComparison.OrdinalIgnoreCase);
                if (paramStr.Contains("encumbrance"))
                    return flag.Equals("encumbrance", StringComparison.OrdinalIgnoreCase) || flag.Equals("e", StringComparison.OrdinalIgnoreCase) || flag.Equals("actual+encumbrance", StringComparison.OrdinalIgnoreCase) || flag.Equals("a+e", StringComparison.OrdinalIgnoreCase);
            }

            // Handle BalanceType → EndPeriod
            if (paramStr.Contains("balancetype") && paramStr.Contains("endperiod"))
                // Enable only if BalanceType is not blank
                return !string.IsNullOrWhiteSpace(flag) && flag.Equals("ctd", StringComparison.OrdinalIgnoreCase);

            return false;
        }
        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            return System.Windows.Data.Binding.DoNothing;
        }
    }
    public class JournalValidationConverter : IMultiValueConverter
    {
        public object Convert(object[] values, Type targetType, object parameter, CultureInfo culture)
        {
            if (values == null || values.Length < 3)
                return false;

            // values[0] = Activity
            // values[1] = BalanceType
            // values[2] = Currency Type
            // values[3] = Current Field Enable State

            string activity = values[0]?.ToString();
            string balanceType = values[1]?.ToString();
            string currencyType = values[2]?.ToString();

            bool baseEnableState = values.Length > 3 && values[3] is bool && (bool)values[3];

            if (!baseEnableState)
                return false;

            // Check if Activity is valid for Journals
            var validActivities = new[] { "Debit", "DR", "Credit", "CR", "Net" };
            bool isValidActivity = !string.IsNullOrEmpty(activity) &&
                validActivities.Any(valid => valid.Equals(activity, StringComparison.OrdinalIgnoreCase));

            // Check if Balance Type is valid for Journals
            // (kept in sync with GLConfiguratorViewModel.IsJournalValidationSatisfied())
            var validBalanceTypes = new[] { "PTD", "YTD", "CTD", "JED", "JEDP", "JEDU" };
            bool isValidBalanceType = !string.IsNullOrEmpty(balanceType) &&
                validBalanceTypes.Any(valid => valid.Equals(balanceType, StringComparison.OrdinalIgnoreCase));

            // Check if Currency Type is valid for Journals
            // (kept in sync with GLConfiguratorViewModel.IsJournalValidationSatisfied())
            var validCurrencyTypes = new[] { "E", "ENTERED", "TOTAL" };
            bool isValidCurrencyType = !string.IsNullOrEmpty(currencyType) &&
                validCurrencyTypes.Any(valid => valid.Equals(currencyType, StringComparison.OrdinalIgnoreCase));

            return isValidActivity && isValidBalanceType && isValidCurrencyType;
        }

        public object[] ConvertBack(object value, Type[] targetTypes, object parameter, CultureInfo culture)
        {
            throw new NotImplementedException();
        }
    }
    public class BudgetEncumbranceMultiConverter : IMultiValueConverter
    {
        public object Convert(object[] values, Type targetType, object parameter, CultureInfo culture)
        {
            // Validate inputs
            if (values == null || values.Length < 3)
                return false;

            // values[0] = Actual Flag value
            // values[1] = RefEdit value
            // values[2] = Current Field Enable State (IsComboEnabled or IsRefEnabled)

            string actualFlag = values[0]?.ToString();
            string refValue = values[1]?.ToString();

            bool baseEnableState = values.Length > 2 && values[2] is bool && (bool)values[2];

            if (!baseEnableState)
                return false;

            if (string.IsNullOrWhiteSpace(actualFlag))
                return false;

            // Get the parameter (should be "Budget" or "Encumbrance")
            string paramStr = parameter?.ToString().ToLowerInvariant();
            if (string.IsNullOrEmpty(paramStr))
                return false;

            // Check if Actual Flag matches the required type
            bool isCorrectActualFlag = false;

            if (paramStr.Contains("budget"))
            {
                isCorrectActualFlag = actualFlag.Equals("budget", StringComparison.OrdinalIgnoreCase) ||
                                     actualFlag.Equals("b", StringComparison.OrdinalIgnoreCase);
            }
            else if (paramStr.Contains("encumbrance"))
            {
                isCorrectActualFlag = actualFlag.Equals("encumbrance", StringComparison.OrdinalIgnoreCase) ||
                                     actualFlag.Equals("e", StringComparison.OrdinalIgnoreCase) ||
                                     actualFlag.Equals("actual+encumbrance", StringComparison.OrdinalIgnoreCase) ||
                                     actualFlag.Equals("a+e", StringComparison.OrdinalIgnoreCase);
            }
            else
            {
                return false;
            }

            // Enable ONLY IF:
            // 1. Base enable state is true
            // 2. Actual Flag matches (Budget or Encumbrance)
            // 3. RefEdit is empty or null (no reference selected)
            bool isRefEmpty = string.IsNullOrEmpty(refValue);

            return isCorrectActualFlag && isRefEmpty;
        }

        public object[] ConvertBack(object value, Type[] targetTypes, object parameter, CultureInfo culture)
        {
            return new object[] { Binding.DoNothing, Binding.DoNothing, Binding.DoNothing };
        }
    }
    public class MultiBooleanAndConverter : IMultiValueConverter
    {
        public object Convert(object[] values, Type targetType, object parameter, CultureInfo culture)
        {
            return values != null
                && values.Length > 0
                && values.All(v => v is bool && System.Convert.ToBoolean(v));
        }

        public object[] ConvertBack(object value, Type[] targetTypes, object parameter, CultureInfo culture)
        {
            throw new NotImplementedException();
        }
    }

    // Group H (Balance Configurator) addition - verbatim port of GLSense\Converters\
    // Converters.cs BalanceTypeDatePeriodConverter (FinalWorkingCode). Used by
    // GLBalanceConfigurator.xaml's Start/End Date DatePickers (enabled only for
    // JED/JEDP/JEDU balance types) and the Period combo (enabled for the inverse, via
    // ConverterParameter="Invert").
    public class BalanceTypeDatePeriodConverter : IMultiValueConverter
    {
        private static readonly HashSet<string> JEDTypes =
            new HashSet<string>(StringComparer.OrdinalIgnoreCase)
            {
                "JED",
                "JEDP",
                "JEDU"
            };

        public object Convert(object[] values, Type targetType, object parameter, CultureInfo culture)
        {
            if (values == null || values.Length < 1)
                return false;

            var balanceType = values[0]?.ToString();

            // Check if balance type is JED/JEDP/JEDU
            bool isJEDType = !string.IsNullOrWhiteSpace(balanceType) &&
                             JEDTypes.Contains(balanceType);

            // If parameter is "Invert", return opposite (for Period control)
            bool invert = parameter != null &&
                          parameter.ToString()?.Equals("Invert", StringComparison.OrdinalIgnoreCase) == true;

            bool result = invert ? !isJEDType : isJEDType;

            // If a second value is provided (expected to be a boolean like IsComboEnabled), combine it
            if (values.Length > 1 && values[1] is bool b)
            {
                return result && b;
            }

            return result;
        }

        public object[] ConvertBack(object value, Type[] targetTypes, object parameter, CultureInfo culture)
        {
            throw new NotImplementedException();
        }
    }

    // Group H (LOVs/Roller/Account dialogs) addition - verbatim port of GLSense\
    // Converters\TitleRowValueConverter.cs (FinalWorkingCode). Used by GLRollerGroups.xaml's
    // left DataGrid "Value" column: shows the ISegmentRow.Title text for grouped "RG"
    // title rows, else the SegmentDataRow.SegmentValue text. Re-pointed: GLSense.Models ->
    // GLSense.Addin.Core.Models (ISegmentRow now lives in Models\SegmentSelectorModels.cs).
    public class TitleRowValueConverter : IMultiValueConverter
    {
        public object Convert(object[] values, Type targetType, object parameter, CultureInfo culture)
        {
            var row = values[0] as GLSense.Addin.Core.Models.ISegmentRow;
            if (row == null) return string.Empty;

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
