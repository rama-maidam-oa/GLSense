// OnlineReleaseRow.cs in GLSense\Views
using System.ComponentModel;

namespace GLSense
{
    /// <summary>
    /// One row in GLReloadSourcePicker's Online-mode release list. Built by
    /// OnlineReleaseClassifier from a server VersionInfo entry plus the local
    /// ReleaseHistory.json catalog - see
    /// docs/superpowers/specs/2026-09-18-online-reload-multiversion-design.md.
    /// </summary>
    public class OnlineReleaseRow : INotifyPropertyChanged
    {
        private bool _isChecked;

        public string Version { get; set; }
        public string ReleaseDate { get; set; }
        public bool Mandatory { get; set; }
        public string FolderName { get; set; }
        public string FileName { get; set; }
        public string Checksum { get; set; }
        public string Notes { get; set; }

        public bool IsCurrentlyLoaded { get; set; }
        public bool IsAlreadyDownloaded { get; set; }

        /// <summary>Neither currently loaded nor already catalogued locally - the
        /// only state a checkbox can be enabled/checked for.</summary>
        public bool IsSelectable => !IsCurrentlyLoaded && !IsAlreadyDownloaded;

        public string StatusText =>
            IsCurrentlyLoaded ? "Currently loaded" :
            IsAlreadyDownloaded ? "Already Downloaded" :
            "New";

        public bool IsChecked
        {
            get => _isChecked;
            set
            {
                if (_isChecked == value) return;
                _isChecked = value;
                PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(IsChecked)));
            }
        }

        public event PropertyChangedEventHandler PropertyChanged;
    }
}
