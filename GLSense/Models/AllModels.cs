using GLSense.Base;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Text.Json;
using System.Text.Json.Serialization;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace GLSense.Models
{

    public class LedgerModel : INotifyPropertyChanged
    {
        public long LedgerId { get; set; }
        public long CubeId { get; set; }
        public string LedgerName { get; set; }
        public long CoaId { get; set; }
        public string PeriodSetName { get; set; }
        public string CurrencyCode { get; set; }
        public string LastRefreshedDate { get; set; }
        public string ADMRefreshedDate { get; set; }
        public string TimeZone { get; set; }

        private bool _hasWarnings;
        public bool HasWarnings
        {
            get => _hasWarnings;
            set
            {
                _hasWarnings = value;
                OnPropertyChanged(nameof(HasWarnings));
            }
        }

        public event PropertyChangedEventHandler PropertyChanged;
        protected virtual void OnPropertyChanged(string propertyName)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }

        // For multi-select
        public bool IsSelected { get; set; }
    }
    public class LedgerQueryData
    {
        [JsonPropertyName("records")]
        public Records records { get; set; }

        [JsonPropertyName("status")]
        public string status { get; set; }

        [JsonPropertyName("msg")]
        public string msg { get; set; }
    }

    public class Records
    {
        [JsonPropertyName("journalsources")]
        public JESources[] journalsources { get; set; }

        [JsonPropertyName("journalcategories")]
        public JECategories[] journalcategories { get; set; }

        [JsonPropertyName("activity")]
        public string[] activity { get; set; }

        [JsonPropertyName("ledgers")]
        public Ledgers ledgers { get; set; }

        [JsonPropertyName("encumbrances")]
        public Encumbrance[] encumbrances { get; set; }

        [JsonPropertyName("currencies")]
        public string[] currencies { get; set; }
    }

    public class Period
    {
        [JsonPropertyName("periodName")]
        public string periodName { get; set; }

        [JsonPropertyName("periodYear")]
        public int periodYear { get; set; }

        [JsonPropertyName("periodNum")]
        public int periodNum { get; set; }

        [JsonPropertyName("quarterNum")]
        public int quarterNum { get; set; }

        [JsonPropertyName("periodSetName")]
        public string periodSetName { get; set; }

        [JsonPropertyName("periodType")]
        public string periodType { get; set; }

        [JsonPropertyName("startDate")]
        public long startDate { get; set; }

        [JsonPropertyName("endDate")]
        public long endDate { get; set; }

        [JsonPropertyName("adjustmentPeriodFlag")]
        public string adjustmentPeriodFlag { get; set; }
    }

    public class LedgerSegmentValue
    {
        [JsonPropertyName("segmentValue")]
        public string segmentValue { get; set; }

        [JsonPropertyName("description")]
        public string description { get; set; }

        [JsonPropertyName("summaryFlag")]
        public string summaryFlag { get; set; }

        [JsonPropertyName("enabledFlag")]
        public string enabledFlag { get; set; }

        [JsonPropertyName("segmentValueSetId")]
        public long segmentValueSetId { get; set; }
    }

    public class LedgerSegment
    {
        [JsonPropertyName("coaid")]
        public int coaid { get; set; }

        [JsonPropertyName("segmentName")]
        public string segmentName { get; set; }

        [JsonPropertyName("segmentValueSetId")]
        public long segmentValueSetId { get; set; }

        [JsonPropertyName("securityEnabledFlag")]
        public string securityEnabledFlag { get; set; }

        [JsonPropertyName("defaultType")]
        public string defaultType { get; set; }

        [JsonPropertyName("defaultValue")]
        public string defaultValue { get; set; }

        [JsonPropertyName("displaySize")]
        public int displaySize { get; set; }

        [JsonPropertyName("segmentDelimiter")]
        public string segmentDelimiter { get; set; }

        [JsonPropertyName("segmentValues")]
        public LedgerSegmentValue[] segmentValues { get; set; }

        [JsonPropertyName("applicationColumnName")]
        public string applicationColumnName { get; set; }
    }

    public class LedgerData
    {
        [JsonPropertyName("budgets")]
        public string[] budgets { get; set; }

        [JsonPropertyName("periods")]
        public Period[] periods { get; set; }

        [JsonPropertyName("segments")]
        public LedgerSegment[] segments { get; set; }
    }

    public class Ledgers
    {
        [JsonPropertyName("ledgerId")]
        public long ledgerId { get; set; }

        [JsonPropertyName("ledgerName")]
        public string ledgerName { get; set; }

        [JsonPropertyName("coaid")]
        public int coaid { get; set; }

        [JsonPropertyName("periodSetName")]
        public string periodSetName { get; set; }

        [JsonPropertyName("currencyCode")]
        public string currencyCode { get; set; }

        [JsonPropertyName("periodType")]
        public string periodType { get; set; }

        [JsonPropertyName("ledgerData")]
        public LedgerData ledgerData { get; set; }
    }

    public class Encumbrance
    {
        [JsonPropertyName("encumbrancTypeId")]
        public long encumbranceTypeId { get; set; }

        [JsonPropertyName("encumbranceType")]
        public string encumbranceType { get; set; }
    }

    public class JESources
    {
        [JsonPropertyName("jeSourceName")]
        public string jeSourceName { get; set; }

        [JsonPropertyName("sourceName")]
        public string sourceName { get; set; }
    }

    public class JECategories
    {
        [JsonPropertyName("jeCategoryName")]
        public string jeCategoryName { get; set; }

        [JsonPropertyName("categoryName")]
        public string categoryName { get; set; }
    }
    public class SegmentSelectionModel
    {
        public string Value1 { get; set; }
        public string Value2 { get; set; }
        public string Segment { get; set; }
    }
    public class HierarchyRecord
    {
        public string parent { get; set; }
        public int lvl { get; set; }
        public string segmentValue { get; set; }
        public string description { get; set; }
    }

    public class HierarchyResponse
    {
        public string Message { get; set; }
        public string Status { get; set; }
        public List<HierarchyRecord> Records { get; set; }
    }
    public class SearchTypeModel
    {
        public string DisplayName { get; set; }
        public string Value { get; set; }

        public override string ToString()
        {
            return DisplayName;
        }
    }
    public class AttributeTypeModel
    {
        public string DisplayName { get; set; }
        public string Value { get; set; }

        public override string ToString()
        {
            return DisplayName;
        }
    }
    public class ScrollToTopMessage
    {
        public bool ScrollLeft { get; set; } = true;
        public bool ScrollRight { get; set; } = true;
        public string Trigger { get; set; } = "DataLoaded";
    }
    public class LovRow
    {
        public string Name { get; set; }          // E.g. "Account", "Activity", ...
        public int ItemsCount { get; set; }       // Number of choices
        public string Category { get; set; }      // "Segment", "Database", "Hardcoded"
    }
    public class BalanceTypeModel
    {
        public string Name { get; set; }
        public override string ToString()
        {
            return Name;
        }
    }
    public class BudgetModel : INotifyPropertyChanged
    {
        public long CubeId { get; set; }
        public long LedgerId { get; set; }
        public string BudgetName { get; set; }

        private bool _isSelected;
        public bool IsSelected
        {
            get => _isSelected;
            set
            {
                if (_isSelected != value)
                {
                    _isSelected = value;
                    OnPropertyChanged(nameof(IsSelected));
                }
            }
        }

        public override string ToString() => BudgetName;

        public event PropertyChangedEventHandler PropertyChanged;
        protected virtual void OnPropertyChanged(string name) =>
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
    }
    public class JournalSourceModel
    {
        public long CubeId { get; set; }
        public long LedgerId { get; set; }
        public string JeSourceName { get; set; }
        public string SourceName { get; set; }
    }

    public class JournalCategoryModel
    {
        public long CubeId { get; set; }
        public long LedgerId { get; set; }
        public string JeCategoryName { get; set; }
        public string CategoryName { get; set; }
    }
    public class PeriodModel
    {
        public long CubeId { get; set; }
        public long LedgerId { get; set; }
        public string PeriodName { get; set; }
        public int PeriodYear { get; set; }
        public int PeriodNum { get; set; }
        public int QuarterNum { get; set; }
        public string PeriodSetName { get; set; }
        public string PeriodType { get; set; }
        public DateTime StartDate { get; set; }
        public DateTime EndDate { get; set; }
        public string AdjustmentPeriodFlag { get; set; }
    }

    public class OffsetModel
    {
        public int OffsetValue { get; set; }
        public string DisplayName { get; set; }
    }

    public class CurrencyModel
    {
        public long CubeId { get; set; }
        public long LedgerId { get; set; }
        public string CurrencyCode { get; set; }
    }
    public interface ISegmentRow
    {
    }

    public class TitleRow : ISegmentRow
    {
        public string Title { get; set; }
        public string Description { get; set; }
        public string SummaryFlag { get; set; } = "RG"; // Indicate title row
    }

    public class SegmentDataRow : ISegmentRow
    {
        public string SegmentName { get; set; }
        public string SegmentValue { get; set; }
        public string Description { get; set; }
        public string SummaryFlag { get; set; }
        public long SegmentValueSetId { get; set; }
    }
    public class CheckableItem
    {
        public string DisplayText { get; set; } = string.Empty;
        public bool IsChecked { get; set; } = false;
    }
    public class OperationResult
    {
        public bool IsSuccess { get; set; }
        public string Message { get; set; }
        public Exception Exception { get; set; }
    }
    public class GenericLedgerModel : NotifyBase
    {
        public long LedgerId { get; set; }
        public long CubeId { get; set; }
        public string LedgerName { get; set; }
        public int CoaId { get; set; }
        public string PeriodSetName { get; set; }
        public string CurrencyCode { get; set; }

        private bool _isSelected;
        public bool IsSelected
        {
            get => _isSelected;
            set
            {
                if (_isSelected != value)
                {
                    _isSelected = value;
                    OnPropertyChanged(nameof(IsSelected));
                }
            }
        }

        public override string ToString() => LedgerName;
    }
    public class SegmentDff
    {
        public string attributeName { get; set; }
        public string segmentValue { get; set; }
        public long segmentValueSetId { get; set; }
    }

    public class UserConfigResponse
    {
        [JsonPropertyName("preferences")]
        public Preferences Preferences { get; set; }

        [JsonPropertyName("message")]
        public string Message { get; set; }

        [JsonPropertyName("status")]
        public string Status { get; set; }
    }
    public class Preferences
    {
        [JsonConverter(typeof(StringToBoolConverter))]
        [JsonPropertyName("validateCube")]
        public bool? ValidateCube { get; set; }

        [JsonConverter(typeof(StringToBoolConverter))]
        [JsonPropertyName("supressZeroBalDrilldown")]
        public bool? SupressZeroBalDrilldown { get; set; }
        [JsonConverter(typeof(StringToBoolConverter))]
        [JsonPropertyName("runSubLedgerDrilldownAsJob")]
        public bool? RunSubLedgerDrilldownAsJob { get; set; }

        [JsonConverter(typeof(StringToBoolConverter))]
        [JsonPropertyName("runBalDrilldownAsJob")]
        public bool? RunBalDrilldownAsJob { get; set; } 
        [JsonConverter(typeof(StringToBoolConverter))]
        [JsonPropertyName("runTotalDrilldownAsJob")]
        public bool? RunTotalDrilldownAsJob { get; set; }

        [JsonConverter(typeof(StringToIntConverter))]
        [JsonPropertyName("recordsPerPage")]
        public int? RecordsPerPage { get; set; }

        [JsonConverter(typeof(StringToIntConverter))]
        [JsonPropertyName("refreshCells")]
        public int? RefreshCells { get; set; }

        [JsonConverter(typeof(StringToBoolConverter))]
        [JsonPropertyName("runJournalDrilldownAsJob")]
        public bool? RunJournalDrilldownAsJob { get; set; }

        [JsonPropertyName("dataOption")]
        public string DataOption { get; set; }

        [JsonConverter(typeof(StringToBoolConverter))]
        [JsonPropertyName("journalRealTimeDataEnabled")]
        public bool? JournalRealTimeDataEnabled { get; set; }
        [JsonConverter(typeof(StringToBoolConverter))]
        [JsonPropertyName("includeManualJournal")]
        public bool? IncludeManualJournal { get; set; }
    }

    public class UserConfigResetResponse
    {
        public string msg { get; set; }
        public string status { get; set; }
        public string message { get; set; }
    }

    public class DailyRateQuery
    {
        public string fromCurrency { get; set; }
        public string toCurrency { get; set; }
        public string conversionType { get; set; }
        public string conversionDate { get; set; }
    }

    public class DailyRateRecord
    {
        public double? CONVERSION_RATE { get; set; }
    }

    public sealed class ExternalResolveResult
    {
        public Excel.Workbook Workbook { get; set; }
        public Excel.Worksheet Worksheet { get; set; }
        public Excel.Range Range { get; set; }
    }

    public class JournalAttachments
    {
        public long cubeId { get; set; }
        public long journalHeaderId { get; set; }
    }

    public class JrnalAttachRequest
    {
        public long cubeId { get; set; }
        public long[] fileIds { get; set; }
    }
    public class JournalAttachmentRecord
    {
        public string FILE_ID { get; set; }
        public string FILE_NAME { get; set; }
    }
    public class SegmentModel : INotifyPropertyChanged
    {
        private long _id;
        private long _cubeId;
        private long _ledgerId;
        private long _coaId;
        private string _segmentName;
        private long _segmentValueSetId;
        private string _securityEnabledFlag;
        private string _defaultType;
        private string _defaultValue;
        private int _displaySize;
        private string _segmentDelimiter;
        private string _applicationColumnName;
        private string _value;
        private string _reference;
        private bool _isVisible;
        private bool _isTextEnabled = true;
        private bool _isRefEditEnabled = true;
        private ObservableCollection<SegmentSelectionModel> _selectedValues = new ObservableCollection<SegmentSelectionModel>();

        public event PropertyChangedEventHandler PropertyChanged;

        protected virtual void OnPropertyChanged([CallerMemberName] string propertyName = null)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }

        protected bool SetProperty<T>(ref T field, T value, [CallerMemberName] string propertyName = null)
        {
            if (EqualityComparer<T>.Default.Equals(field, value)) return false;
            field = value;
            OnPropertyChanged(propertyName);
            return true;
        }


        public long Id
        {
            get => _id;
            set => SetProperty(ref _id, value);
        }
        public long CubeId
        {
            get => _cubeId;
            set => SetProperty(ref _cubeId, value);
        }
        public long LedgerId
        {
            get => _ledgerId;
            set => SetProperty(ref _ledgerId, value);
        }

        public long CoaId
        {
            get => _coaId;
            set => SetProperty(ref _coaId, value);
        }

        public string SegmentName
        {
            get => _segmentName;
            set => SetProperty(ref _segmentName, value);
        }

        public long SegmentValueSetId
        {
            get => _segmentValueSetId;
            set => SetProperty(ref _segmentValueSetId, value);
        }

        public string SecurityEnabledFlag
        {
            get => _securityEnabledFlag;
            set => SetProperty(ref _securityEnabledFlag, value);
        }

        public string DefaultType
        {
            get => _defaultType;
            set => SetProperty(ref _defaultType, value);
        }

        public string DefaultValue
        {
            get => _defaultValue;
            set => SetProperty(ref _defaultValue, value);
        }

        public int DisplaySize
        {
            get => _displaySize;
            set => SetProperty(ref _displaySize, value);
        }

        public string SegmentDelimiter
        {
            get => _segmentDelimiter;
            set => SetProperty(ref _segmentDelimiter, value);
        }

        public string ApplicationColumnName
        {
            get => _applicationColumnName;
            set => SetProperty(ref _applicationColumnName, value);
        }

        public string Value
        {
            get => _value;
            set => SetProperty(ref _value, value);
        }
        public string Reference
        {
            get => _reference;
            set => SetProperty(ref _reference, value);
        }
        public bool IsVisible
        {
            get => _isVisible;
            set => SetProperty(ref _isVisible, value);
        }
        public bool IsTextEnabled
        {
            get => _isTextEnabled;
            set => SetProperty(ref _isTextEnabled, value);
        }
        public bool IsRefEditEnabled
        {
            get => _isRefEditEnabled;
            set => SetProperty(ref _isRefEditEnabled, value);
        }

        public ObservableCollection<SegmentSelectionModel> SelectedValues
        {
            get => _selectedValues;
            set => SetProperty(ref _selectedValues, value);
        }

        // Helper methods
        public bool HasSelectedValues()
        {
            return _selectedValues != null && _selectedValues.Count > 0;
        }

        public void ClearSelectedValues()
        {
            if (_selectedValues != null)
            {
                _selectedValues.Clear();
                OnPropertyChanged(nameof(SelectedValues));
            }
        }

        public override string ToString()
        {
            return $"{SegmentName} ({ApplicationColumnName})";
        }
    }
    public class SegmentValueModel : INotifyPropertyChanged
    {
        // Basic fields

        public long Id { get; set; }
        public long CubeId { get; set; }
        public long LedgerId { get; set; }
        public string SegmentName { get; set; }
        public string SegmentValue { get; set; }
        public string Description { get; set; }
        public string EnabledFlag { get; set; }
        public long SegmentValueSetId { get; set; }
        public string ApplicationColumnName { get; set; }

        // hierarchy fields (optional when loading normal segments)
        public string Parent { get; set; }
        public int Level { get; set; }

        // Backing fields
        private string _summaryFlag;         // current persisted value (may be updated on save)
        private string _originalSummaryFlag; // original value loaded from DB
        private bool _isSummaryChecked;      // UI toggled value

        public event PropertyChangedEventHandler PropertyChanged;

        // summaryFlag property (keeps the persisted value)
        public string SummaryFlag
        {
            get => _summaryFlag;
            set
            {
                if (_summaryFlag != value)
                {
                    _summaryFlag = value;
                    // If original not yet captured, capture it now (first load)
                    if (string.IsNullOrEmpty(_originalSummaryFlag))
                    {
                        _originalSummaryFlag = value;
                    }
                    OnPropertyChanged();
                    OnPropertyChanged(nameof(IsSummaryAccount));
                    OnPropertyChanged(nameof(IsModified));
                }
            }
        }

        // Call this after constructing/loading the model from DB to initialize UI state
        public void MarkLoaded()
        {
            _originalSummaryFlag = _summaryFlag;
            _isSummaryChecked = string.Equals(_originalSummaryFlag, "Y", StringComparison.OrdinalIgnoreCase);
            OnPropertyChanged(nameof(IsSummaryChecked));
            OnPropertyChanged(nameof(IsSummaryAccount));
            OnPropertyChanged(nameof(IsModified));
        }

        // The value the user can toggle in the UI
        public bool IsSummaryChecked
        {
            get => _isSummaryChecked;
            set
            {
                if (_isSummaryChecked != value)
                {
                    _isSummaryChecked = value;
                    OnPropertyChanged();
                    OnPropertyChanged(nameof(IsModified));
                }
            }
        }

        // Read-only: is this a summary account *originally* (from DB) — used to color and enable checkbox
        public bool IsSummaryAccount => string.Equals(_originalSummaryFlag, "Y", StringComparison.OrdinalIgnoreCase);

        // Computed: did user change the checkbox relative to original DB value?
        public bool IsModified
        {
            get
            {
                bool original = string.Equals(_originalSummaryFlag, "Y", StringComparison.OrdinalIgnoreCase);
                return original != _isSummaryChecked;
            }
        }

        // Call this when you want to persist the current UI choice into the persisted field
        // (e.g. when saving to DB)

        public void AcceptChanges()
        {
            _summaryFlag = _isSummaryChecked ? "Y" : "N";
            _originalSummaryFlag = _summaryFlag;
            OnPropertyChanged(nameof(SummaryFlag));
            OnPropertyChanged(nameof(IsSummaryAccount));
            OnPropertyChanged(nameof(IsModified));
        }

        // Values to be displayed after selecting the hierarchy from hierarchy combo
        public string DisplaySegmentValue
        {
            get
            {
                string indent = Level > 0 ? new string(' ', Level * 2) : string.Empty;
                return $"{indent}{SegmentValue}";
            }
        }

        public string DisplayDescription => $"{SegmentValue} - {Description}";

        // Helper to raise PropertyChanged
        protected virtual void OnPropertyChanged([CallerMemberName] string propertyName = null)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
    }
    
    public class EncumbranceModel : NotifyBase
    {
        public long CubeId { get; set; }
        public long LedgerId { get; set; }
        public long EncumbranceTypeId { get; set; }
        public string EncumbranceType { get; set; }

        private bool _isSelected;
        public bool IsSelected
        {
            get => _isSelected;
            set
            {
                if (_isSelected != value)
                {
                    _isSelected = value;
                    OnPropertyChanged(nameof(IsSelected));
                }
            }
        }
        public override string ToString() => EncumbranceType;
    }
    public class CubeWindowResponse
    {
        public string Msg { get; set; }
        public List<CubeRecord> Records { get; set; }
        public string Status { get; set; }
    }
    public class CubeRecord
    {
        public long CubeId { get; set; }
        public string CubeName { get; set; }
        public string UserName { get; set; }
        public List<LedgerRecord> Ledgers { get; set; }
        public string LastRefreshedDate { get; set; }
        public bool BlazeEnabled { get; set; }
        public string ErpType { get; set; }
        public bool AdaptiveMemoryEnabled { get; set; }
        public string AdaptiveMemoryTableName { get; set; }
        public bool ViewBased { get; set; }

        public override string ToString()
        {
            return CubeName ?? string.Empty;
        }

        // Method to get LedgerID by ledger name
        public long? GetLedgerIdByName(string ledgerName)
        {
            if (string.IsNullOrEmpty(ledgerName) || Ledgers == null || !Ledgers.Any())
                return null;

            // Remove surrounding quotes if present
            string cleanLedgerName = ledgerName.Trim().Trim('"');

            // Also handle escaped quotes
            cleanLedgerName = cleanLedgerName.Replace("\\\"", "\"");

            var ledger = Ledgers.FirstOrDefault(l =>
                string.Equals(l.LedgerName?.Trim(), cleanLedgerName, StringComparison.OrdinalIgnoreCase));

            return ledger?.LedgerId;
        }

        public LedgerRecord GetLedgerByName(string ledgerName)
        {
            if (string.IsNullOrEmpty(ledgerName) || Ledgers == null || !Ledgers.Any())
                return null;

            return Ledgers.FirstOrDefault(l =>
                string.Equals(l.LedgerName?.Trim(), ledgerName.Trim(), StringComparison.OrdinalIgnoreCase));
        }

        // Method to get all ledger names
        public List<string> GetLedgerNames()
        {
            return Ledgers?.Select(l => l.LedgerName).ToList() ?? new List<string>();

        }
    }
    public class LedgerRecord
    {
        public long LedgerId { get; set; }
        public string LedgerName { get; set; }
        public long Coaid { get; set; }
        public string PeriodSetName { get; set; }
        public string CurrencyCode { get; set; }
        public string PeriodType { get; set; }
        public string LedgerData { get; set; }
    }
    public class CubeLedgerRecord
    {
        public long LedgerId { get; set; }
        public string LedgerName { get; set; }
        public string LastRefreshedDateInUTC { get; set; }
        public long LastRefreshedDateInMilliSecs { get; set; }
        public string LastRefreshedAdaptiveMemInUTC { get; set; }
        public long LastRefreshedAdaptiveMemDateInMilliSecs { get; set; }
        public string LastRefreshedSourceADMInUTC { get; set; }
        public long LastRefreshedSourceADMDateInMilliSecs { get; set; }
    }

    public class CubeLedgerResponse
    {
        public string Msg { get; set; }
        public List<CubeLedgerRecord> Records { get; set; }
        public string Status { get; set; }
    }

    public static class CubeCache
    {
        public static List<CubeRecord> AllCubes { get; set; }
        public static Dictionary<long, CubeValidationResult> Validations { get; set; } = new Dictionary<long, CubeValidationResult>();
    }

    public class LedgerValidationResult
    {
        public string LedgerName { get; set; }
        public bool IsValid { get; set; }
    }

    // Cube-level validation info
    public class CubeValidationResult
    {
        public long CubeId { get; set; }
        public string Message { get; set; }
        public string ErrorMessage { get; set; }
        public bool IsValidated { get; set; }
        public DateTime? ValidatedAt { get; set; }
        public bool IsInSync => !string.IsNullOrWhiteSpace(Message) && (Message.IndexOf("in sync", StringComparison.OrdinalIgnoreCase) >= 0);
        public bool NeedsConfirmation => IsValidated && !IsInSync;
        public List<LedgerValidationResult> Ledgers { get; set; } = new List<LedgerValidationResult>();
    }
    public class BroadcastMessage
    {
        public string MsgType { get; set; }
        public string Message { get; set; }
    }
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

        // New properties for Manual Journals
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
    [ComVisible(false)]
    [ClassInterface(ClassInterfaceType.None)]
    public class StringToBoolConverter : JsonConverter<bool>
    {
        public override bool Read(ref Utf8JsonReader reader, Type typeToConvert, JsonSerializerOptions options)
        {
            return reader.TokenType switch
            {
                JsonTokenType.True => true,
                JsonTokenType.False => false,
                JsonTokenType.String => reader.GetString()?.ToLower() switch
                {
                    "true" => true,
                    "false" => false,
                    "1" => true,
                    "0" => false,
                    "yes" => true,
                    "no" => false,
                    _ => LogAndThrow($"Cannot convert string '{reader.GetString()}' to bool")
                },
                _ => LogAndThrow($"Unexpected token type: {reader.TokenType}")
            };
        }

        private static bool LogAndThrow(string message)
        {
            GLSense.Utilities.LogUtility.LogError($"StringToBoolConverter.Read: {message}");
            throw new JsonException(message);
        }

        public override void Write(Utf8JsonWriter writer, bool value, JsonSerializerOptions options)
        {
            writer.WriteBooleanValue(value);
        }
    }

    // Handles server responses that send numeric fields as JSON strings (e.g.
    // "recordsPerPage":"1", "refreshCells":"1") instead of JSON numbers. Without this,
    // System.Text.Json throws on a plain int/int? property when the token is a string,
    // which was causing ApiResponseHelper.Parse<UserConfigResponse> to fail outright on
    // an otherwise well-formed /user-config preferences response.
    [ComVisible(false)]
    [ClassInterface(ClassInterfaceType.None)]
    public class StringToIntConverter : JsonConverter<int>
    {
        public override int Read(ref Utf8JsonReader reader, Type typeToConvert, JsonSerializerOptions options)
        {
            switch (reader.TokenType)
            {
                case JsonTokenType.Number:
                    return reader.GetInt32();
                case JsonTokenType.String:
                    var s = reader.GetString();
                    if (int.TryParse(s, out var result))
                        return result;
                    return LogAndThrow($"Cannot convert string '{s}' to int");
                default:
                    return LogAndThrow($"Unexpected token type: {reader.TokenType}");
            }
        }

        private static int LogAndThrow(string message)
        {
            GLSense.Utilities.LogUtility.LogError($"StringToIntConverter.Read: {message}");
            throw new JsonException(message);
        }

        public override void Write(Utf8JsonWriter writer, int value, JsonSerializerOptions options)
        {
            writer.WriteNumberValue(value);
        }
    }
}
