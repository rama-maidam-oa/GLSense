using System.ComponentModel;
using System.Runtime.CompilerServices;

namespace GLSense.Models
{
#nullable enable
    /// <summary>
    /// One row in GLDrilldownDeleteCustomization's picker grid - a single DD type
    /// (records-key, e.g. "BALANCE") currently saved locally for the selected cube via
    /// Common\DrilldownMetadataXmlStore, plus how many column-metadata entries it holds.
    /// </summary>
    public class GLDrilldownTypeModel : INotifyPropertyChanged
    {
        public string DdType { get; set; } = string.Empty;

        public string DisplayName { get; set; } = string.Empty;

        public int RecordCount { get; set; }

        private bool _isSelected;
        public bool IsSelected
        {
            get => _isSelected;
            set
            {
                if (_isSelected != value)
                {
                    _isSelected = value;
                    OnPropertyChanged();
                }
            }
        }

        public event PropertyChangedEventHandler? PropertyChanged;

        protected virtual void OnPropertyChanged([CallerMemberName] string? propertyName = null)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
    }
#nullable disable
}
