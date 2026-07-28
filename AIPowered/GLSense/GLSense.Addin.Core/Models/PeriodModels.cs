// PeriodModels.cs in GLSense.Addin.Core
// Right-sized carve-out from GLSense\Models\AllModels.cs (FinalWorkingCode) for Group C
// (Segment/Period pickers). Contains the plain data models the 7 Group C views/
// ViewModels need that were not already ported by Group B (SegmentModel already lives in
// Models\LedgerModel.cs).
//
// Models ported here:
//   - PeriodModel / CurrencyModel: verbatim, no logic.
//   - GenericLedgerModel: the old project derives this from a "NotifyBase" helper class
//     (GLSense\Base\NotifyBase.cs) that itself depends on GLSense.Utilities.LogUtility
//     (old project's static logger) for debug-mode property-change logging. Rather than
//     port NotifyBase (which would drag in a whole new base-class hierarchy for a single
//     model), this implements INotifyPropertyChanged directly - matching the style already
//     used by this project's own Models\LedgerModel.cs (LedgerModel/SegmentModel both
//     implement INotifyPropertyChanged directly, no NotifyBase equivalent exists here).
//   - AttributeTypeModel + static AttributeTypeService: verbatim port of the
//     AttributeTypeModel/AttributeTypeService half of GLSense\Service\SearchTypeService.cs
//     (lines 30-65) - the hardcoded ATTRIBUTE1..20 list used by GLSegmentFuncsViewModel's
//     "DFF" (descriptive flex field) mode.
//   - SegmentValueModel: verbatim port of GLSense\Models\AllModels.cs (~line 669+),
//     including the hierarchy/summary-flag bookkeeping (SummaryFlag/IsSummaryChecked/
//     IsSummaryAccount/IsModified/MarkLoaded/AcceptChanges) that GLSegmentFuncsViewModel's
//     GetAdjacentSegment relies on for its Next/Previous-segment parent/child logic.
//
// Group E (Drilldowns) addition - SearchTypeModel / SearchTypeService (port of
// GLSense\Service\SearchTypeService.cs lines 1-29, GLSense.Service namespace). A prior
// header note here deferred this to "Group H" as unneeded by Group C - that was
// premature: GLSubmittedJobsViewModel (ViewModels\GLSubmittedJobsViewModel.cs) needs it
// for its search-criteria combo. Moved here alongside AttributeTypeService, following the
// same "static service classes live next to their model in this project" convention
// already established for AttributeTypeModel/AttributeTypeService above (the old project
// kept both in GLSense.Service; this project has no equivalent Service namespace for
// these small static lookup lists).
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Runtime.CompilerServices;

namespace GLSense.Addin.Core.Models
{
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
        public System.DateTime StartDate { get; set; }
        public System.DateTime EndDate { get; set; }
        public string AdjustmentPeriodFlag { get; set; }
    }

    public class CurrencyModel
    {
        public long CubeId { get; set; }
        public long LedgerId { get; set; }
        public string CurrencyCode { get; set; }
    }

    public class GenericLedgerModel : INotifyPropertyChanged
    {
        public long LedgerId { get; set; }
        public long CubeId { get; set; }
        public string LedgerName { get; set; }
        public long CoaId { get; set; }
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

        public event PropertyChangedEventHandler PropertyChanged;
        protected virtual void OnPropertyChanged(string propertyName) =>
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
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

    public static class SearchTypeService
    {
        private static readonly ObservableCollection<SearchTypeModel> _searchTypes = new ObservableCollection<SearchTypeModel>
        {
            new SearchTypeModel { DisplayName = "Starts With", Value = "StartsWith" },
            new SearchTypeModel { DisplayName = "Does Not Start With", Value = "DoesNotStartWith" },
            new SearchTypeModel { DisplayName = "Ends With", Value = "EndsWith" },
            new SearchTypeModel { DisplayName = "Does Not End With", Value = "DoesNotEndWith" },
            new SearchTypeModel { DisplayName = "Contains", Value = "Contains" },
            new SearchTypeModel { DisplayName = "Not Contains", Value = "NotContains" },
            new SearchTypeModel { DisplayName = "Equals", Value = "Equals" },
            new SearchTypeModel { DisplayName = "Not Equals", Value = "NotEquals" },
        };

        public static ObservableCollection<SearchTypeModel> GetSearchTypes()
        {
            return _searchTypes;
        }

        public static SearchTypeModel GetDefaultSearchType()
        {
            return _searchTypes[4];
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

    public static class AttributeTypeService
    {
        private static readonly ObservableCollection<AttributeTypeModel> _attributeTypes = new ObservableCollection<AttributeTypeModel>
        {
            new AttributeTypeModel { DisplayName = "ATTRIBUTE1", Value = "ATTRIBUTE1" },
            new AttributeTypeModel { DisplayName = "ATTRIBUTE2", Value = "ATTRIBUTE2" },
            new AttributeTypeModel { DisplayName = "ATTRIBUTE3", Value = "ATTRIBUTE3" },
            new AttributeTypeModel { DisplayName = "ATTRIBUTE4", Value = "ATTRIBUTE4" },
            new AttributeTypeModel { DisplayName = "ATTRIBUTE5", Value = "ATTRIBUTE5" },
            new AttributeTypeModel { DisplayName = "ATTRIBUTE6", Value = "ATTRIBUTE6" },
            new AttributeTypeModel { DisplayName = "ATTRIBUTE7", Value = "ATTRIBUTE7" },
            new AttributeTypeModel { DisplayName = "ATTRIBUTE8", Value = "ATTRIBUTE8" },
            new AttributeTypeModel { DisplayName = "ATTRIBUTE9", Value = "ATTRIBUTE9" },
            new AttributeTypeModel { DisplayName = "ATTRIBUTE10", Value = "ATTRIBUTE10" },
            new AttributeTypeModel { DisplayName = "ATTRIBUTE11", Value = "ATTRIBUTE11" },
            new AttributeTypeModel { DisplayName = "ATTRIBUTE12", Value = "ATTRIBUTE12" },
            new AttributeTypeModel { DisplayName = "ATTRIBUTE13", Value = "ATTRIBUTE13" },
            new AttributeTypeModel { DisplayName = "ATTRIBUTE14", Value = "ATTRIBUTE14" },
            new AttributeTypeModel { DisplayName = "ATTRIBUTE15", Value = "ATTRIBUTE15" },
            new AttributeTypeModel { DisplayName = "ATTRIBUTE16", Value = "ATTRIBUTE16" },
            new AttributeTypeModel { DisplayName = "ATTRIBUTE17", Value = "ATTRIBUTE17" },
            new AttributeTypeModel { DisplayName = "ATTRIBUTE18", Value = "ATTRIBUTE18" },
            new AttributeTypeModel { DisplayName = "ATTRIBUTE19", Value = "ATTRIBUTE19" },
            new AttributeTypeModel { DisplayName = "ATTRIBUTE20", Value = "ATTRIBUTE20" },
        };

        public static ObservableCollection<AttributeTypeModel> GetAttributesType()
        {
            return _attributeTypes;
        }

        public static AttributeTypeModel GetDefaultAttributeType()
        {
            return _attributeTypes[0];
        }
    }

    // Group E (Drilldowns) addition - standalone port of GLSense\ViewModels\
    // GLConfiguratorViewModel.ActivityModel's DisplayName/ShortName parsing
    // (FinalWorkingCode, lines ~282-327). The original nested this inside
    // GLConfiguratorViewModel (Group H, not yet ported) and derived from NotifyBase;
    // DataRepository.GetActivities (Group E) only needs the plain data + the two
    // computed name properties, so this is a standalone POCO instead - same reasoning
    // GenericLedgerModel above used to avoid porting NotifyBase for one model.
    public class ActivityModel
    {
        public long CubeId { get; set; }
        public long LedgerId { get; set; }
        public string ActivityType { get; set; }

        public string DisplayName
        {
            get
            {
                if (string.IsNullOrWhiteSpace(ActivityType)) return string.Empty;
                var parts = ActivityType.Split(':');
                return parts.Length > 1 ? parts[1] : ActivityType;
            }
        }

        public string ShortName
        {
            get
            {
                if (string.IsNullOrWhiteSpace(ActivityType)) return string.Empty;
                var parts = ActivityType.Split(':');
                return parts[0];
            }
        }

        public override string ToString() => DisplayName;
    }

    // Group E (Drilldowns) addition - port of GLSense\Models\AllModels.cs
    // EncumbranceModel (FinalWorkingCode, ~line 784). Backs DataRepository.
    // GetEncumbrances, used by Models\BalanceDtoModel.cs to resolve
    // encumbranceTypeIdList from an encumbrance-name list embedded in a
    // GLSense_GetBalance(...) formula.
    //
    // Group H (Balance Configurator) addition - the old monolith's EncumbranceModel
    // (GLSense\Models\AllModels.cs) derives from NotifyBase and has a notifying
    // IsSelected property, used by GLConfiguratorViewModel/GLBalanceConfigurator for the
    // multi-select Encumbrances combo (SuggestAppendComboBox IsMultiSelect="True", same
    // checkbox-per-item pattern as GenericLedgerModel.IsSelected above). Adding
    // IsSelected + INotifyPropertyChanged here is purely additive - Models\BalanceDtoModel.cs
    // (Group E) only ever reads CubeId/LedgerId/EncumbranceTypeId/EncumbranceType, never
    // IsSelected, so this does not change that consumer's behavior.
    public class EncumbranceModel : INotifyPropertyChanged
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

        public event PropertyChangedEventHandler PropertyChanged;
        protected virtual void OnPropertyChanged(string propertyName) =>
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
    }

    // Group H (Balance Configurator) addition - verbatim ports of GLSense\Models\
    // AllModels.cs BudgetModel/JournalSourceModel/JournalCategoryModel (FinalWorkingCode).
    // Back DataRepository.GetBudgets/GetJournalSources/GetJournalCategories (added to
    // Repositories\DataRepository.cs alongside these). The old BudgetModel derived from
    // INotifyPropertyChanged with a notifying IsSelected property, but nothing in
    // GLConfiguratorViewModel ever reads/sets BudgetModel.IsSelected (Budgets is a
    // single-select combo, not a multi-select checkbox list like Ledgers/Encumbrances) -
    // per this project's "only add what's actually used" convention (see AppState.cs
    // header), IsSelected is intentionally omitted here.
    public class BudgetModel
    {
        public long CubeId { get; set; }
        public long LedgerId { get; set; }
        public string BudgetName { get; set; }

        public override string ToString() => BudgetName;
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

    // Port of GLSense.Models.AllModels.HierarchyRecord (FinalWorkingCode) - the shape
    // returned by the /rest/secure/finance/segment-hierarchy API and deserialized in
    // DataRepository.SaveHierarchyToCache (Group D). Field names are lowercase to match
    // the JSON property names exactly (System.Text.Json's default matching is
    // case-insensitive, but this mirrors the original verbatim).
    public class HierarchyRecord
    {
        public string parent { get; set; }
        public int lvl { get; set; }
        public string segmentValue { get; set; }
        public string description { get; set; }
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
            _isSummaryChecked = string.Equals(_originalSummaryFlag, "Y", System.StringComparison.OrdinalIgnoreCase);
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

        // Read-only: is this a summary account *originally* (from DB) - used to color and enable checkbox
        public bool IsSummaryAccount => string.Equals(_originalSummaryFlag, "Y", System.StringComparison.OrdinalIgnoreCase);

        // Computed: did user change the checkbox relative to original DB value?
        public bool IsModified
        {
            get
            {
                bool original = string.Equals(_originalSummaryFlag, "Y", System.StringComparison.OrdinalIgnoreCase);
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

        // Group H (GLSegmentValues) addition - verbatim port of GLSense\Models\AllModels.cs
        // SegmentValueModel.DisplaySegmentValue/DisplayDescription (FinalWorkingCode).
        // DisplaySegmentValue indents hierarchy-expanded rows by Level; DisplayDescription
        // backs the GLSegmentValues hierarchy combo's DisplayMemberPath.
        public string DisplaySegmentValue
        {
            get
            {
                string indent = Level > 0 ? new string(' ', Level * 2) : string.Empty;
                return $"{indent}{SegmentValue}";
            }
        }

        public string DisplayDescription => $"{SegmentValue} - {Description}";

        public override string ToString()
        {
            return SegmentValue;
        }

        protected void OnPropertyChanged([CallerMemberName] string propertyName = null) =>
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
    }
}
