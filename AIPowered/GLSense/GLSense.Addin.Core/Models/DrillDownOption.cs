// DrillDownOption.cs in GLSense.Addin.Core
// Port of GLSense\Models\AllModels.cs's DrillDownOption class (FinalWorkingCode) -
// the editable-DataGrid row model backing GLUserConfig's "DrillDowns" tab (Name/RunAsJob/
// CanEditRunAsJob/IncludeManualJournal/CanEditManualJournals/ShowManualJournalsColumn).
//
// Group E (Drilldowns) resolution: Utilities\UserConfig.cs originally deferred
// DrillDownSettings/DrillDownOption to "whenever Group E gets ported" - Group E (and every
// other group through H) is now fully ported, so that deferral was stale. This model is
// added here, alongside UserConfig.DrillDownSettings, specifically for Group I's
// GLUserConfig port (nothing else in this project references DrillDownOption yet).
//
// No NotifyBase involved: the original already implements INotifyPropertyChanged directly
// (it does NOT derive from GLSense.Base.NotifyBase), so this is a verbatim, namespace-only
// port - no base-class substitution needed (contrast with GenericLedgerModel in
// PeriodModels.cs, which did derive from NotifyBase and required that substitution).
using System.ComponentModel;

namespace GLSense.Addin.Core.Models
{
    public class DrillDownOption : INotifyPropertyChanged
    {
        public string Name { get; set; }

        private bool _runAsJob;
        public bool RunAsJob
        {
            get => _runAsJob;
            set
            {
                _runAsJob = value;
                OnPropertyChanged(nameof(RunAsJob));
            }
        }

        private bool _canEditRunAsJob = true;
        public bool CanEditRunAsJob
        {
            get => _canEditRunAsJob;
            set
            {
                _canEditRunAsJob = value;
                OnPropertyChanged(nameof(CanEditRunAsJob));
            }
        }

        // Manual Journals (SubLedger Drilldown only)
        private bool _includeManualJournal;
        public bool IncludeManualJournal
        {
            get => _includeManualJournal;
            set
            {
                _includeManualJournal = value;
                OnPropertyChanged(nameof(IncludeManualJournal));
            }
        }

        private bool _canEditManualJournals;
        public bool CanEditManualJournals
        {
            get => _canEditManualJournals;
            set
            {
                _canEditManualJournals = value;
                OnPropertyChanged(nameof(CanEditManualJournals));
            }
        }

        private bool _showManualJournalsColumn;
        public bool ShowManualJournalsColumn
        {
            get => _showManualJournalsColumn;
            set
            {
                _showManualJournalsColumn = value;
                OnPropertyChanged(nameof(ShowManualJournalsColumn));
            }
        }

        public event PropertyChangedEventHandler PropertyChanged;

        protected void OnPropertyChanged(string prop)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(prop));
        }
    }
}
