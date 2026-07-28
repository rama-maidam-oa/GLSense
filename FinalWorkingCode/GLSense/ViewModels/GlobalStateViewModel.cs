using System.ComponentModel;

namespace GLSense.ViewModels
{
    public class GlobalStateViewModel : INotifyPropertyChanged
    {
        private string _referenceText = string.Empty;
        public string ReferenceText
        {
            get => _referenceText;
            set
            {
                if (_referenceText != value)
                {
                    _referenceText = value;
                    OnPropertyChanged(nameof(ReferenceText));
                }
            }
        }

        public static GlobalStateViewModel Instance { get; } = new GlobalStateViewModel();

        public event PropertyChangedEventHandler PropertyChanged;
        protected void OnPropertyChanged(string propertyName) =>
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
    }
}
