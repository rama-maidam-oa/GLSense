// GlobalStateViewModel.cs in GLSense.Addin.Core
// Verbatim port of GLSense\ViewModels\GlobalStateViewModel.cs (FinalWorkingCode) for
// Group C (Segment/Period pickers). Tiny singleton with a single ReferenceText property,
// shared by all 7 Group C views to show the active-cell address in their "Reference:"
// ExcelRefEditControl row. No old-project dependencies to re-point.
using System.ComponentModel;

namespace GLSense.Addin.Core.ViewModels
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
