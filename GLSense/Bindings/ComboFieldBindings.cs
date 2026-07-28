using System.ComponentModel;
using System.Threading.Tasks;

namespace GLSense.Bindings
{
    public class ComboFieldBindings : INotifyPropertyChanged
    {
        public enum FieldType { Ledger, Date, Offset, Period, PeriodYear, PeriodNum, Segment, SegmentValue, Attribute, Currency }

        public FieldType Type { get; set; }
        public IFieldDependencyProvider DependencyProvider { get; set; }
        public IFieldDependencyProviderNonAsync DependencyProviderNonAsync { get; set; }

        private string _comboText;
        public string ComboText
        {
            get => _comboText;
            set
            {
                if (_comboText != value)
                {
                    _comboText = value;
                    OnPropertyChanged(nameof(ComboText));

                    if (string.IsNullOrWhiteSpace(_comboText))
                    {
                        ComboValue = null;
                    }
                }
            }
        }

        private object _comboValue;
        public object ComboValue
        {
            get => _comboValue;
            set
            {
                if (!Equals(_comboValue, value))
                {
                    _comboValue = value;

                    // If we're clearing the Combo while RefEdit was the source
                    if (value == null && IsValueFromRefEdit)
                    {
                        IsValueFromRefEdit = false;  // Reset the flag
                    }

                    OnPropertyChanged(nameof(ComboValue));
                    OnPropertyChanged(nameof(IsComboEnabled));
                    OnPropertyChanged(nameof(IsRefEnabled));
                    // Call async version if available
                    if (DependencyProvider != null)
                    {
                        DependencyProvider.OnFieldDependencyChanged(this);
                    }
                    // Call non-async version if available
                    else
                    {
                        DependencyProviderNonAsync?.OnFieldDependencyChanged(this);
                    }
                }
            }
        }

        private string _refValue;
        public string RefValue
        {
            get => _refValue;
            set
            {
                if (_refValue != value)
                {
                    _refValue = value;
                    OnPropertyChanged(nameof(RefValue));
                    OnPropertyChanged(nameof(IsComboEnabled));
                    OnPropertyChanged(nameof(IsRefEnabled));
                    // Call async version if available
                    if (DependencyProvider != null)
                    {
                        DependencyProvider.OnRefEditTextChanged(this, value);
                    }
                    // Call non-async version if available
                    else
                    {
                        DependencyProviderNonAsync?.OnRefEditTextChanged(this, value);
                    }
                }
            }
        }

        public bool IsValueFromRefEdit { get; set; } = false;
        public bool IsComboEnabled => DependencyProvider?.IsComboEnabled(this)
                                      ?? DependencyProviderNonAsync?.IsComboEnabled(this)
                                      ?? true;

        public bool IsRefEnabled => DependencyProvider?.IsRefEnabled(this)
                                    ?? DependencyProviderNonAsync?.IsRefEnabled(this)
                                    ?? true;

        public void RefreshEnableState()
        {
            OnPropertyChanged(nameof(IsComboEnabled));
            OnPropertyChanged(nameof(IsRefEnabled));
        }

        public event PropertyChangedEventHandler PropertyChanged;
        protected void OnPropertyChanged(string name) => PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
    }

    public interface IFieldDependencyProvider
    {
        bool IsRefEnabled(ComboFieldBindings field);
        bool IsComboEnabled(ComboFieldBindings field);
        Task OnFieldDependencyChanged(ComboFieldBindings field);
        void OnRefEditTextChanged(ComboFieldBindings field, string newText);
    }
    public interface IFieldDependencyProviderNonAsync
    {
        bool IsRefEnabled(ComboFieldBindings field);
        bool IsComboEnabled(ComboFieldBindings field);
        void OnFieldDependencyChanged(ComboFieldBindings field);
        void OnRefEditTextChanged(ComboFieldBindings field, string newText);
    }
}
