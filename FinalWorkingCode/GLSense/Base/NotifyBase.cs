using System.ComponentModel;
using System.Runtime.CompilerServices;
using GLSense.Utilities;

namespace GLSense.Base
{
#nullable enable
    public class NotifyBase : INotifyPropertyChanged
    {
        public event PropertyChangedEventHandler? PropertyChanged;

        protected void OnPropertyChanged(string propertyName)
            => PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));

        protected bool SetField<T>(ref T field, T value, [CallerMemberName] string? prop = null)
        {
            if (System.Collections.Generic.EqualityComparer<T>.Default.Equals(field, value)) return false;
            field = value;
            if (prop is not null)
                OnPropertyChanged(prop);

            LogPropertyChange(prop, value);
            return true;
        }

        private void LogPropertyChange<T>(string? propertyName, T value)
        {
            if (!LogUtility.DebugMode || string.IsNullOrWhiteSpace(propertyName))
                return;

            string typeName = GetType().Name;
            string valueText = FormatValue(value);
            LogUtility.LogDebug($"[{typeName}] Property '{propertyName}' changed to '{valueText}'");
        }

        private static string FormatValue<T>(T value)
        {
            string text = value?.ToString() ?? "<null>";
            const int maxLength = 300;
            if (text.Length > maxLength)
            {
                text = text.Substring(0, maxLength) + "...";
            }
            return text;
        }
    }
#nullable disable
}
