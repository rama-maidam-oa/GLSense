// DatePickerExtensions.cs in GLSense.Addin.Core
// Verbatim port of GLSense\Extensions\DatePickerExtensions.cs (FinalWorkingCode) for
// Group C (Segment/Period pickers) - thin extension-method wrapper around
// Helpers.DatePickerTooltipHelper, used by GLGetPeriodByDate.xaml.cs and
// GLDailyRates.xaml.cs (dtpDate.SetupTooltip(...)).
using System;
using System.Windows.Controls;
using System.Windows.Threading;

namespace GLSense.Addin.Core.Extensions
{
    public static class DatePickerExtensions
    {
        public static void SetupTooltip(
            this DatePicker datePicker,
            string title,
            Dispatcher dispatcher = null,
            string dateFormat = "yyyy-MM-dd",
            string instructionText = "Click calendar icon to select/change date",
            Action<DatePicker> onDateChangedAction = null)
        {
            Helpers.DatePickerTooltipHelper.InitializeTooltip(
                datePicker,
                title,
                dispatcher,
                dateFormat,
                instructionText,
                onDateChangedAction
            );
        }

        public static void UpdateTooltip(this DatePicker datePicker)
        {
            Helpers.DatePickerTooltipHelper.UpdateTooltip(datePicker);
        }

        public static void CleanupTooltip(this DatePicker datePicker)
        {
            Helpers.DatePickerTooltipHelper.Cleanup(datePicker);
        }
    }
}
