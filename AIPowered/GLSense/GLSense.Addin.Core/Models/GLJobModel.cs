// GLJobModel.cs in GLSense.Addin.Core
// Port of GLSense\Models\GLJobModel.cs (FinalWorkingCode) for Group E (Drilldowns) -
// backs the GLSubmittedJobsViewModel.Jobs collection shown in the GLJobsMonitor DataGrid
// (RibDrillJobs ribbon button). Verbatim port, no logic changes:
//   - CanDownload/DownloadIconKind/DownloadIcon/DownloadIconColor/DownloadBackgroundColor/
//     DownloadTooltip/StatusColor/PhaseColor are all computed straight off Phase/Status,
//     same thresholds and color values as the original.
//   - AppConstants.Success resolves automatically here since this namespace
//     (GLSense.Addin.Core.Models) nests under GLSense.Addin.Core, where AppConstants lives
//     (same resolution already relied on throughout Drilldowns\DD_BL.cs etc.).
using MahApps.Metro.IconPacks;
using System;
using System.ComponentModel;
using System.Runtime.CompilerServices;
using System.Windows.Media;

namespace GLSense.Addin.Core.Models
{
#nullable enable
    public class GLJobModel : INotifyPropertyChanged
    {
        private const string RunningStatus = "running";
        private const string PendingStatus = "pending";
        private const string CompletedStatus = "completed";
        private const string FailureStatus = "failure";
        private const string CancelledStatus = "cancelled";
        private bool _isSelected;

        public string? ProcessId { get; set; }

        public string? JobDescription { get; set; }

        public string? JobType { get; set; }

        private string? _name;
        public string? Name
        {
            get => _name;
            set
            {
                if (_name != value)
                {
                    _name = value;
                    OnPropertyChanged();
                    OnPropertyChanged(nameof(DownloadTooltip));
                }
            }
        }
        private string? _phase;
        public string? Phase
        {
            get => _phase;
            set
            {
                if (_phase != value)
                {
                    _phase = value;
                    OnPropertyChanged();
                    OnPropertyChanged(nameof(DownloadIcon));
                    OnPropertyChanged(nameof(DownloadIconKind));
                    OnPropertyChanged(nameof(DownloadTooltip));
                    OnPropertyChanged(nameof(DownloadIconColor));
                    OnPropertyChanged(nameof(DownloadBackgroundColor));
                }
            }
        }
        private string? _status;
        public string? Status
        {
            get => _status;
            set
            {
                if (_status != value)
                {
                    _status = value;
                    OnPropertyChanged();
                    OnPropertyChanged(nameof(DownloadIcon));
                    OnPropertyChanged(nameof(DownloadIconKind));
                    OnPropertyChanged(nameof(DownloadTooltip));
                    OnPropertyChanged(nameof(DownloadIconColor));
                    OnPropertyChanged(nameof(DownloadBackgroundColor));
                    OnPropertyChanged(nameof(CanDownload));
                }
            }
        }
        public DateTime CreatedDate { get; set; }
        public string? DateInfo { get; set; }
        public string? DrillType { get; set; } // "SS" for snapshot

        public bool IsSelected
        {
            get => _isSelected;
            set
            {
                _isSelected = value;
                OnPropertyChanged();
            }
        }

        public bool CanDownload
        {
            get
            {
                if (string.IsNullOrEmpty(Phase) || string.IsNullOrEmpty(Status))
                    return false;

                var _phase1 = Phase?.ToLower();
                var _status1 = Status?.ToLower();
                return (_phase1 == CompletedStatus || _phase1 == "complete") &&
                       _status1 == AppConstants.Success;
            }
        }

        // For display in UI
        public string? DisplayDate => CreatedDate > DateTime.MinValue ?
            CreatedDate.ToString("dd-MMM-yyyy hh:mm:ss tt") : DateInfo;

        // Excel icon property
        public PackIconFontAwesomeKind DownloadIconKind
        {
            get
            {
                if (CanDownload)
                    return PackIconFontAwesomeKind.FileExcelSolid; // Excel icon
                else if (Phase?.ToLower() == RunningStatus)
                    return PackIconFontAwesomeKind.RotateSolid; // Running/refresh icon
                else if (Phase?.ToLower() == PendingStatus)
                    return PackIconFontAwesomeKind.ClockSolid; // Clock/pending icon
                else if (Status?.ToLower() == FailureStatus)
                    return PackIconFontAwesomeKind.CircleXmarkSolid; // Failed/X icon
                else if (Status?.ToLower() == CancelledStatus)
                    return PackIconFontAwesomeKind.BanSolid; // Cancelled/ban icon
                else
                    return PackIconFontAwesomeKind.CircleInfoSolid; // Info icon for other cases
            }
        }
        public string DownloadIcon
        {
            get
            {
                if (CanDownload)
                    return "\U0001F4CA"; // Excel icon
                else if (Phase?.ToLower() == RunningStatus)
                    return "\U0001F504"; // Running icon
                else if (Phase?.ToLower() == PendingStatus)
                    return "⏳"; // Pending icon
                else if (Status?.ToLower() == FailureStatus)
                    return "❌"; // Failed icon
                else if (Status?.ToLower() == CancelledStatus)
                    return "\U0001F6AB"; // Cancelled icon
                else
                    return "ℹ️"; // Info icon for other cases
            }
        }
        // Color for the icon
        public SolidColorBrush DownloadIconColor
        {
            get
            {
                if (CanDownload)
                    return new SolidColorBrush(Color.FromRgb(33, 115, 70)); // Dark Green
                else if (Phase?.ToLower() == RunningStatus)
                    return new SolidColorBrush(Color.FromRgb(0, 123, 255)); // Blue
                else if (Phase?.ToLower() == PendingStatus)
                    return new SolidColorBrush(Color.FromRgb(255, 193, 7)); // Yellow/Orange
                else if (Status?.ToLower() == FailureStatus)
                    return new SolidColorBrush(Color.FromRgb(220, 53, 69)); // Red
                else if (Status?.ToLower() == CancelledStatus)
                    return new SolidColorBrush(Color.FromRgb(108, 117, 125)); // Gray
                else
                    return new SolidColorBrush(Color.FromRgb(108, 117, 125)); // Default Gray
            }
        }
        // Background color for the icon (circular background)
        public SolidColorBrush DownloadBackgroundColor
        {
            get
            {
                if (CanDownload)
                    return new SolidColorBrush(Color.FromRgb(232, 245, 233)); // Light Green
                else if (Phase?.ToLower() == RunningStatus)
                    return new SolidColorBrush(Color.FromRgb(230, 240, 255)); // Light Blue
                else if (Phase?.ToLower() == PendingStatus)
                    return new SolidColorBrush(Color.FromRgb(255, 248, 225)); // Light Yellow
                else if (Status?.ToLower() == FailureStatus)
                    return new SolidColorBrush(Color.FromRgb(255, 231, 233)); // Light Red
                else if (Status?.ToLower() == CancelledStatus)
                    return new SolidColorBrush(Color.FromRgb(233, 236, 239)); // Light Gray
                else
                    return new SolidColorBrush(Color.FromRgb(248, 249, 250)); // Very Light Gray
            }
        }
        // Tooltip with more details
        public string DownloadTooltip
        {
            get
            {
                if (CanDownload)
                    return $"✅ Completed successfully\n📊 Excel output available\n🕐 {DisplayDate}\n📝 {Name}";
                else if (Phase?.ToLower() == RunningStatus)
                    return $"🔄 Running...\n⏱️ Started: {DisplayDate}\n📝 {Name}";
                else if (Phase?.ToLower() == PendingStatus)
                    return $"⏳ Pending execution\n📅 Scheduled: {DisplayDate}\n📝 {Name}";
                else if (Status?.ToLower() == FailureStatus)
                    return $"❌ Job failed\n🚨 Phase: {Phase}\n🕐 {DisplayDate}\n📝 {Name}";
                else if (Status?.ToLower() == CancelledStatus)
                    return $"🚫 Job cancelled\n📅 {DisplayDate}\n📝 {Name}";
                else
                    return $"{Phase} - {Status}\n🕐 {DisplayDate}\n📝 {Name}";
            }
        }

        // Color coding based on status
        public SolidColorBrush StatusColor
        {
            get
            {
                if (string.IsNullOrEmpty(Status))
                    return new SolidColorBrush(Color.FromRgb(108, 117, 125));

                var statusLower = Status?.ToLower();
                if (statusLower == AppConstants.Success)
                    return new SolidColorBrush(Color.FromRgb(40, 167, 69)); // Green
                else if (statusLower == RunningStatus)
                    return new SolidColorBrush(Color.FromRgb(0, 123, 255));   // Blue
                else if (statusLower == FailureStatus)
                    return new SolidColorBrush(Color.FromRgb(220, 53, 69));    // Red
                else if (statusLower == PendingStatus)
                    return new SolidColorBrush(Color.FromRgb(255, 193, 7));   // Yellow
                else
                    return new SolidColorBrush(Color.FromRgb(108, 117, 125)); // Gray
            }
        }

        // Color coding based on phase
        public SolidColorBrush PhaseColor
        {
            get
            {
                if (string.IsNullOrEmpty(Phase))
                    return new SolidColorBrush(Color.FromRgb(108, 117, 125));

                var phaseLower = Phase?.ToLower();
                if (phaseLower == CompletedStatus || phaseLower == "complete")
                    return new SolidColorBrush(Color.FromRgb(40, 167, 69)); // Green
                else if (phaseLower == RunningStatus)
                    return new SolidColorBrush(Color.FromRgb(0, 123, 255));   // Blue
                else if (phaseLower == PendingStatus)
                    return new SolidColorBrush(Color.FromRgb(255, 193, 7));   // Yellow
                else
                    return new SolidColorBrush(Color.FromRgb(108, 117, 125)); // Gray
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
