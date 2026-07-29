using GLSense.Base;
using GLSense.Bindings;
using GLSense.Helpers;
using GLSense.Models;
using GLSense.Repositories;
using GLSense.Utilities;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Globalization;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Documents;
using System.Windows.Media;
using System.Windows.Threading;

namespace GLSense.ViewModels
{
#nullable enable
    public class GLConfiguratorViewModel : INotifyPropertyChanged
    {
        [ComVisible(false)]
        public delegate bool RefMatch<in T>(T model, string refText);
        public static GlobalStateViewModel GlobalState => GlobalStateViewModel.Instance;

        private const string Budget = "Budget";
        private const string Encumbrance = "Encumbrance";
        private const string AE = "Actual+Encumbrance";

        // ===== Check Sign, Zeroes and Text Factor =====
        private bool _isSignChecked;
        public bool IsSignChecked
        {
            get => _isSignChecked;
            set
            {
                _isSignChecked = value;
                OnPropertyChanged(nameof(IsSignChecked));
                OnFieldChanged(nameof(IsSignChecked));
            }
        }
        private void ValidateDates()
        {
            LogUtility.LogDebug($"GLConfiguratorViewModel.ValidateDates: StartDateSelected={StartDateSelected}, EndDateSelected={EndDateSelected}");
            if (StartDateSelected.HasValue && EndDateSelected.HasValue && StartDateSelected.Value > EndDateSelected.Value)
            {
                LogUtility.LogWarn("GLConfiguratorViewModel.ValidateDates: Start Date is greater than End Date; clearing End Date.");
                ShowWarningAction?.Invoke("Start Date cannot be greater than End Date. End Date has been cleared.");

                // Clear End Date
                EndDateSelected = null;
                return;
            }

            // If only End Date is set and Start Date is null, that's fine
            // If only Start Date is set and End Date is null, that's fine
            UpdateParameterSummary();
        }

        // Update StartDateSelected setter
        private DateTime? _startDateSelected;
        public DateTime? StartDateSelected
        {
            get => _startDateSelected;
            set
            {
                if (_startDateSelected != value)
                {
                    LogUtility.LogDebug($"GLConfiguratorViewModel.StartDateSelected: changing from {_startDateSelected} to {value}");
                    _startDateSelected = value;
                    OnPropertyChanged(nameof(StartDateSelected));

                    // Validate dates - this will clear EndDate if invalid
                    ValidateDates();

                    // Reflect selected date into StartDateField combo text (ISO format)
                    if (StartDateField != null)
                    {
                        if (_startDateSelected.HasValue)
                        {
                            var s = _startDateSelected.Value.ToString(AppConstants.DateFormatIso);
                            StartDateField.ComboText = s;
                            StartDateField.ComboValue = s;
                        }
                        else
                        {
                            StartDateField.ComboText = null;
                            StartDateField.ComboValue = null;
                        }
                        StartDateField.RefreshEnableState();
                    }
                }
            }
        }

        // Update EndDateSelected setter
        private DateTime? _endDateSelected;
        public DateTime? EndDateSelected
        {
            get => _endDateSelected;
            set
            {
                if (_endDateSelected != value)
                {
                    LogUtility.LogDebug($"GLConfiguratorViewModel.EndDateSelected: changing from {_endDateSelected} to {value}");
                    _endDateSelected = value;
                    OnPropertyChanged(nameof(EndDateSelected));

                    // Validate dates - this will clear EndDate if invalid
                    ValidateDates();

                    if (EndDateField != null)
                    {
                        if (_endDateSelected.HasValue)
                        {
                            var s = _endDateSelected.Value.ToString(AppConstants.DateFormatIso);
                            EndDateField.ComboText = s;
                            EndDateField.ComboValue = s;
                        }
                        else
                        {
                            EndDateField.ComboText = null;
                            EndDateField.ComboValue = null;
                        }
                        EndDateField.RefreshEnableState();
                    }
                }
            }
        }

        // Add a method to handle date changes from the UI
        public void OnDateChanged()
        {
            ValidateDates();
        }
        private void ProcessStartDate(FieldBinding field, string refText, string? rngValue)
        {
            LogUtility.LogDebug($"GLConfiguratorViewModel.ProcessStartDate: refText={refText}, rngValue={rngValue}");
            // If reference provided, keep it. If the referenced cell contains a date value, normalize to ISO format.
            field.RefValue = refText;
            if (!string.IsNullOrWhiteSpace(rngValue))
            {
                if (DateTime.TryParse(rngValue, CultureInfo.InvariantCulture, DateTimeStyles.None, out DateTime dt))
                {
                    var iso = dt.ToString(AppConstants.DateFormatIso, CultureInfo.InvariantCulture);
                    field.ComboValue = iso;
                    field.ComboText = iso;
                    // Also set the Selected date so the DatePicker reflects the referenced cell
                    try { _dispatcher.InvokeAsync(() => StartDateSelected = dt); }
                    catch (Exception ex)
                    {
                        LogUtility.LogException(ex, "GLConfiguratorViewModel.ProcessStartDate: failed to set StartDateSelected via dispatcher (non-fatal)");
                    }
                }
                else
                {
                    field.ComboValue = rngValue;
                    field.ComboText = rngValue;
                }
            }
            else
            {
                field.ComboValue = null;
                field.ComboText = null;
                try { _dispatcher.InvokeAsync(() => StartDateSelected = null); }
                catch (Exception ex)
                {
                    LogUtility.LogException(ex, "GLConfiguratorViewModel.ProcessStartDate: failed to clear StartDateSelected via dispatcher (non-fatal)");
                }
            }
            field.RefreshEnableState();
        }

        private void ProcessEndDate(FieldBinding field, string refText, string? rngValue)
        {
            LogUtility.LogDebug($"GLConfiguratorViewModel.ProcessEndDate: refText={refText}, rngValue={rngValue}");
            field.RefValue = refText;
            if (!string.IsNullOrWhiteSpace(rngValue))
            {
                if (DateTime.TryParse(rngValue, CultureInfo.InvariantCulture, DateTimeStyles.None, out DateTime dt))
                {
                    var iso = dt.ToString(AppConstants.DateFormatIso);
                    field.ComboValue = iso;
                    field.ComboText = iso;
                    // Also set the Selected date so the DatePicker reflects the referenced cell
                    try { _dispatcher.InvokeAsync(() => EndDateSelected = dt); }
                    catch (Exception ex)
                    {
                        LogUtility.LogException(ex, "GLConfiguratorViewModel.ProcessEndDate: failed to set EndDateSelected via dispatcher (non-fatal)");
                    }
                }
                else
                {
                    field.ComboValue = rngValue;
                    field.ComboText = rngValue;
                }
            }
            else
            {
                field.ComboValue = null;
                field.ComboText = null;
                try { _dispatcher.InvokeAsync(() => EndDateSelected = null); }
                catch (Exception ex)
                {
                    LogUtility.LogException(ex, "GLConfiguratorViewModel.ProcessEndDate: failed to clear EndDateSelected via dispatcher (non-fatal)");
                }
            }
            field.RefreshEnableState();
        }


        private bool _isZeroesChecked = true;
        public bool IsZeroesChecked
        {
            get => _isZeroesChecked;
            set
            {
                _isZeroesChecked = value;
                OnPropertyChanged(nameof(IsZeroesChecked));
            }
        }

        private string _factorText = string.Empty;
        public string FactorText
        {
            get => _factorText;
            set
            {
                _factorText = value;
                OnPropertyChanged(nameof(FactorText));
            }
        }

        private FlowDocument _parameterDisplayText = new();
        public FlowDocument ParameterDisplayText
        {
            get => _parameterDisplayText;
            set
            {
                _parameterDisplayText = value ?? new FlowDocument();
                OnPropertyChanged(nameof(ParameterDisplayText));
            }
        }

        // Collections (ItemsSource)
        public ObservableCollection<GenericLedgerModel> Ledgers { get; set; }
        public ObservableCollection<SegmentModel> ConfiguratorSegments { get; set; }
        public ObservableCollection<ActivityModel> Activities { get; set; }
        public ObservableCollection<BalanceTypeModel> BalanceTypes { get; set; }
        public ObservableCollection<CurrencyTypeModel> CurrencyTypes { get; set; }
        public ObservableCollection<PeriodModel> Periods { get; set; }
        public ObservableCollection<PeriodModel> EndPeriods { get; set; }
        public ObservableCollection<CurrencyModel> Currencies { get; set; }
        public ObservableCollection<ActualFlagsModel> ActualFlags { get; set; }
        public ObservableCollection<BudgetModel> Budgets { get; set; }
        public ObservableCollection<EncumbranceModel> Encumbrances { get; set; }
        public ObservableCollection<JournalSourceModel> JournalSources { get; set; }
        public ObservableCollection<JournalCategoryModel> JournalCategories { get; set; }

        public IEnumerable<EncumbranceModel> SelectedEncumbrances => Encumbrances.Where(e => e.IsSelected);
        public IEnumerable<GenericLedgerModel> SelectedLedgers => Ledgers.Where(l => l.IsSelected);

        // Row Bindings (one per line)
        public FieldBinding LedgerField { get; set; } = new FieldBinding();
        public FieldBinding ActivityField { get; set; } = new FieldBinding();
        public FieldBinding BalanceTypeField { get; set; } = new FieldBinding();
        public FieldBinding PeriodField { get; set; } = new FieldBinding();
        public FieldBinding EndPeriodField { get; set; } = new FieldBinding();
        public FieldBinding StartDateField { get; set; } = new FieldBinding();
        public FieldBinding EndDateField { get; set; } = new FieldBinding();
        public FieldBinding CurrencyField { get; set; } = new FieldBinding();
        public FieldBinding CurrencyTypeField { get; set; } = new FieldBinding();
        public FieldBinding ActualFlagField { get; set; } = new FieldBinding();
        public FieldBinding BudgetField { get; set; } = new FieldBinding();
        public FieldBinding EncumbranceField { get; set; } = new FieldBinding();
        public FieldBinding JournalSourceField { get; set; } = new FieldBinding();
        public FieldBinding JournalCategoryField { get; set; } = new FieldBinding();
        public FieldBinding AccountAssignmentField { get; set; } = new FieldBinding();

        public class ActivityModel : NotifyBase
        {
            public long CubeId { get; set; }
            public long LedgerId { get; set; }

            private string? _activityType;
            public string? ActivityType
            {
                get => _activityType;
                set
                {
                    if (_activityType != value)
                    {
                        _activityType = value;
                        OnPropertyChanged(nameof(ActivityType));
                        OnPropertyChanged(nameof(DisplayName));
                        OnPropertyChanged(nameof(ShortName));
                    }
                }
            }

            public string DisplayName
            {
                get
                {
                    if (ActivityType == null || string.IsNullOrWhiteSpace(ActivityType)) return "";
                    var parts = ActivityType.Split(':');
                    return parts.Length > 1 ? parts[1] : ActivityType;
                }
            }

            public string ShortName
            {
                get
                {
                    if (ActivityType == null || string.IsNullOrWhiteSpace(ActivityType)) return "";
                    var parts = ActivityType.Split(':');
                    return parts[0];
                }
            }

            public override string ToString()
            {
                return DisplayName;
            }
        }

        public class ActualFlagsModel : NotifyBase
        {
            private string _name = string.Empty;
            public string Name
            {
                get => _name;
                set
                {
                    if (_name != value)
                    {
                        _name = value;
                        OnPropertyChanged(nameof(Name));
                    }
                }
            }

            private string _shortName = string.Empty;
            public string ShortName
            {
                get => _shortName;
                set
                {
                    if (_shortName != value)
                    {
                        _shortName = value;
                        OnPropertyChanged(nameof(ShortName));
                    }
                }
            }

            public override string? ToString()
            {
                return Name;
            }
        }

        public class CurrencyTypeModel : NotifyBase
        {
            private string _name = string.Empty;
            public string Name
            {
                get => _name;
                set
                {
                    if (_name != value)
                    {
                        _name = value;
                        OnPropertyChanged(nameof(Name));
                    }
                }
            }

            private string _shortName = string.Empty;
            public string ShortName
            {
                get => _shortName;
                set
                {
                    if (_shortName != value)
                    {
                        _shortName = value;
                        OnPropertyChanged(nameof(ShortName));
                    }
                }
            }

            public override string? ToString()
            {
                return Name;
            }
        }

        private string _balanceType = string.Empty;
        public string BalanceType
        {
            get => _balanceType;
            set
            {
                if (_balanceType != value)
                {
                    _balanceType = value;
                    OnPropertyChanged(nameof(BalanceType));
                    OnPropertyChanged(nameof(IsEndPeriodsEnabled));
                }
            }
        }

        private PeriodModel _selectedPeriod = null!;
        public PeriodModel SelectedPeriod
        {
            get => _selectedPeriod;
            set
            {
                _selectedPeriod = value;
                OnPropertyChanged(nameof(SelectedPeriod));
                UpdateEndPeriods();
            }
        }
        private void UpdateEndPeriods()
        {
            LogUtility.LogDebug($"GLConfiguratorViewModel.UpdateEndPeriods: SelectedPeriod={SelectedPeriod?.PeriodName}, IsBalanceTypeCtd={IsBalanceTypeCtd()}");
            EndPeriods.Clear();
            if (!IsBalanceTypeCtd() || SelectedPeriod == null || Periods == null)
                return;

            var nextPeriods = Periods.SkipWhile(p => p.PeriodName != SelectedPeriod.PeriodName).Skip(1);
            foreach (var p in nextPeriods)
            {
                EndPeriods.Add(p);
            }
            LogUtility.LogDebug($"GLConfiguratorViewModel.UpdateEndPeriods: populated {EndPeriods.Count} end period(s).");
        }

        private bool IsBalanceTypeCtd()
        {
            var balanceType = GetBalanceTypeText();
            return string.Equals(balanceType, AppConstants.BalanceTypeCTD, StringComparison.OrdinalIgnoreCase);
        }

        private string GetBalanceTypeText()
        {
            var field = BalanceTypeField;
            if (field == null)
                return string.Empty;

            var comboValue = field.ComboValue?.ToString() ?? string.Empty;
            if (!string.IsNullOrWhiteSpace(comboValue))
                return comboValue.Trim();

            var refValue = field.RefValue ?? string.Empty;
            if (string.IsNullOrWhiteSpace(refValue))
                return string.Empty;

            if (ExcelRangeHelper.IsRealRange(refValue))
            {
                var resolvedValue = GetRangeValueSafe(refValue);
                if (!string.IsNullOrWhiteSpace(resolvedValue))
                    return resolvedValue!.Trim();
            }

            return refValue.Trim();
        }

        private bool _IsEndPeriodsEnabled = false;
        public bool IsEndPeriodsEnabled
        {
            get => _IsEndPeriodsEnabled;
            set
            {
                if (_IsEndPeriodsEnabled != value)
                {
                    LogUtility.LogDebug($"GLConfiguratorViewModel.IsEndPeriodsEnabled: changing from {_IsEndPeriodsEnabled} to {value}");
                    _IsEndPeriodsEnabled = value;
                    OnPropertyChanged(nameof(IsEndPeriodsEnabled));
                    EndPeriodField.RefreshEnableState();
                }
            }
        }

        private bool _isBudgetEnabled = false;
        public bool IsBudgetEnabled
        {
            get => _isBudgetEnabled;
            set
            {
                if (_isBudgetEnabled != value)
                {
                    LogUtility.LogDebug($"GLConfiguratorViewModel.IsBudgetEnabled: changing from {_isBudgetEnabled} to {value}");
                    _isBudgetEnabled = value;
                    OnPropertyChanged(nameof(IsBudgetEnabled));
                    BudgetField.RefreshEnableState();
                }
            }
        }

        private bool _isEncumbranceEnabled = false;
        public bool IsEncumbranceEnabled
        {
            get => _isEncumbranceEnabled;
            set
            {
                if (_isEncumbranceEnabled != value)
                {
                    LogUtility.LogDebug($"GLConfiguratorViewModel.IsEncumbranceEnabled: changing from {_isEncumbranceEnabled} to {value}");
                    _isEncumbranceEnabled = value;
                    OnPropertyChanged(nameof(IsEncumbranceEnabled));
                    EncumbranceField.RefreshEnableState();
                }
            }
        }

        private bool _isLedgerEnabled = false;
        public bool IsLedgerEnabled
        {
            get => _isLedgerEnabled;
            set
            {
                if (_isLedgerEnabled != value)
                {
                    LogUtility.LogDebug($"GLConfiguratorViewModel.IsLedgerEnabled: changing from {_isLedgerEnabled} to {value}");
                    _isLedgerEnabled = value;
                    OnPropertyChanged(nameof(IsLedgerEnabled));
                    LedgerField.RefreshEnableState();
                }
            }
        }
        // Visibility properties to collapse/show rows based on balance type and field enable state
        public bool IsPeriodVisible => !IsBalanceTypeJEDVariant();
        public bool IsStartEndVisible => IsBalanceTypeJEDVariant();
        public bool IsEndPeriodVisible => IsBalanceTypeCtd();

        public bool IsBudgetVisible => (BudgetField != null && (BudgetField.IsComboEnabled || BudgetField.IsRefEnabled));
        public bool IsEncumbranceVisible => (EncumbranceField != null && (EncumbranceField.IsComboEnabled || EncumbranceField.IsRefEnabled));
        public bool IsJournalSourceVisible => (JournalSourceField != null && (JournalSourceField.IsComboEnabled || JournalSourceField.IsRefEnabled));
        public bool IsJournalCategoryVisible => (JournalCategoryField != null && (JournalCategoryField.IsComboEnabled || JournalCategoryField.IsRefEnabled));

        // Evaluate the same logic as JournalValidationConverter so ViewModel can
        // decide whether journal comboboxes should be enabled/visible.
        // Not logged: this is invoked from FieldBinding's computed enable-state
        // properties, which WPF's binding engine re-evaluates on essentially every
        // dependent-property notification - thousands of times in a normal session.
        // An "entry" log with no real data here drowned out everything else in the
        // debug log (6000+ of ~10000 lines in one session were just this).
        public bool IsJournalValidationSatisfied()
        {
            try
            {
                var activity = GetFieldValue(ActivityField);
                var balanceType = GetFieldValue(BalanceTypeField);
                var currencyType = GetFieldValue(CurrencyTypeField);

                // Base enable is whether the field is allowed by its own enable state
                // (e.g., ActualFlag / BalanceType). Here we assume callers check
                // the FieldBinding.IsComboEnabled as well; this method focuses on
                // the content validation rules (activity/balanceType/currencyType).

                var validActivities = new[] { "Debit", AppConstants.ActivityDR, "Credit", AppConstants.ActivityCR, "Net" };
                bool isValidActivity = !string.IsNullOrEmpty(activity) && validActivities.Any(v => v.Equals(activity, StringComparison.OrdinalIgnoreCase));

                var validBalanceTypes = new[] {AppConstants.BalanceTypePTD, AppConstants.BalanceTypeYTD, AppConstants.BalanceTypeCTD, AppConstants.BalanceTypeJED, AppConstants.BalanceTypeJEDP, AppConstants.BalanceTypeJEDU };
                bool isValidBalanceType = !string.IsNullOrEmpty(balanceType) && validBalanceTypes.Any(v => v.Equals(balanceType, StringComparison.OrdinalIgnoreCase));

                var validCurrencyTypes = new[] { "E", AppConstants.CurrencyTypeEntered.ToUpperInvariant(), AppConstants.CurrencyTypeTotal.ToUpperInvariant() };
                bool isValidCurrencyType = !string.IsNullOrEmpty(currencyType) && validCurrencyTypes.Any(v => v.Equals(currencyType, StringComparison.OrdinalIgnoreCase));

                return isValidActivity && isValidBalanceType && isValidCurrencyType;
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex, "GLConfiguratorViewModel.IsJournalValidationSatisfied");
                return false;
            }
        }

        private bool IsBalanceTypeJEDVariant()
        {
            var bt = GetBalanceTypeText();
            return string.Equals(bt, AppConstants.BalanceTypeJED, StringComparison.OrdinalIgnoreCase)
                || string.Equals(bt, AppConstants.BalanceTypeJEDP, StringComparison.OrdinalIgnoreCase)
                || string.Equals(bt, AppConstants.BalanceTypeJEDU, StringComparison.OrdinalIgnoreCase);
        }
        private readonly Dispatcher _dispatcher;
        // Field-change coalescing helpers
        private readonly object _fieldChangeLock = new object();
        private bool _fieldChangeScheduled = false;
        private readonly HashSet<string> _pendingFieldNames = new HashSet<string>(StringComparer.Ordinal);
        // Guard to prevent re-entrant UpdateParameterSummary calls
        private bool _isUpdatingParameterSummary = false;

        // Actions for window overlay controls
        public Action<string>? ShowWarningAction { get; set; }
        public Action<string>? ShowInfoAction { get; set; }
        public Func<string, Func<Task>, Task>? ShowBusyAction { get; set; }
        public Func<Task>? HideBusyAsyncAction { get; set; }
        public GLConfiguratorViewModel(Dispatcher dispatcher)
        {
            LogUtility.LogDebug("GLConfiguratorViewModel.ctor: entry");
            _dispatcher = dispatcher;

            // Initialize collections to avoid null refs
            ConfiguratorSegments = new ObservableCollection<SegmentModel>();
            Ledgers = new ObservableCollection<GenericLedgerModel>();
            Activities = new ObservableCollection<ActivityModel>();
            BalanceTypes = new ObservableCollection<BalanceTypeModel>();
            CurrencyTypes = new ObservableCollection<CurrencyTypeModel>();
            Periods = new ObservableCollection<PeriodModel>();
            EndPeriods = new ObservableCollection<PeriodModel>();
            Currencies = new ObservableCollection<CurrencyModel>();
            Budgets = new ObservableCollection<BudgetModel>();
            Encumbrances = new ObservableCollection<EncumbranceModel>();
            JournalSources = new ObservableCollection<JournalSourceModel>();
            JournalCategories = new ObservableCollection<JournalCategoryModel>();
            ActualFlags = new ObservableCollection<ActualFlagsModel>();

            LedgerField = new FieldBinding { OwnerViewModel = this, Type = FieldBinding.FieldType.Ledger };
            ActivityField = new FieldBinding { OwnerViewModel = this, Type = FieldBinding.FieldType.Activity };
            BalanceTypeField = new FieldBinding { OwnerViewModel = this, Type = FieldBinding.FieldType.BalanceType };
            PeriodField = new FieldBinding { OwnerViewModel = this, Type = FieldBinding.FieldType.Period };
            EndPeriodField = new FieldBinding { OwnerViewModel = this, Type = FieldBinding.FieldType.EndPeriod };
            StartDateField = new FieldBinding { OwnerViewModel = this, Type = FieldBinding.FieldType.StartDate };
            EndDateField = new FieldBinding { OwnerViewModel = this, Type = FieldBinding.FieldType.EndDate };
            CurrencyField = new FieldBinding { OwnerViewModel = this, Type = FieldBinding.FieldType.Currency };
            CurrencyTypeField = new FieldBinding { OwnerViewModel = this, Type = FieldBinding.FieldType.CurrencyType };
            ActualFlagField = new FieldBinding { OwnerViewModel = this, Type = FieldBinding.FieldType.ActualFlag };
            BudgetField = new FieldBinding { OwnerViewModel = this, Type = FieldBinding.FieldType.Budgets };
            EncumbranceField = new FieldBinding { OwnerViewModel = this, Type = FieldBinding.FieldType.Encumbrances };
            JournalSourceField = new FieldBinding { OwnerViewModel = this, Type = FieldBinding.FieldType.JournalSources };
            JournalCategoryField = new FieldBinding { OwnerViewModel = this, Type = FieldBinding.FieldType.JournalCategories };
            AccountAssignmentField = new FieldBinding { OwnerViewModel = this, Type = FieldBinding.FieldType.AccountAssignments };

            LedgerField.PropertyChanged += OnFieldChanged;
            ActivityField.PropertyChanged += OnFieldChanged;
            BalanceTypeField.PropertyChanged += OnFieldChanged;
            PeriodField.PropertyChanged += OnFieldChanged;
            EndPeriodField.PropertyChanged += OnFieldChanged;
            StartDateField.PropertyChanged += OnFieldChanged;
            EndDateField.PropertyChanged += OnFieldChanged;
            CurrencyField.PropertyChanged += OnFieldChanged;
            CurrencyTypeField.PropertyChanged += OnFieldChanged;
            ActualFlagField.PropertyChanged += OnFieldChanged;
            BudgetField.PropertyChanged += OnFieldChanged;
            EncumbranceField.PropertyChanged += OnFieldChanged;
            JournalSourceField.PropertyChanged += OnFieldChanged;
            JournalCategoryField.PropertyChanged += OnFieldChanged;
            AccountAssignmentField.PropertyChanged += OnFieldChanged;
        }

        private void OnFieldChanged(string propertyName)
        {
            OnFieldChanged(this, new PropertyChangedEventArgs(propertyName));
        }

        private void OnFieldChanged(object sender, PropertyChangedEventArgs e)
        {
            // Coalesce rapid property-change events into a single update to avoid
            // repeated UpdateParameterSummary calls which can cause re-entrancy
            // and StackOverflow. We collect property names and schedule a single
            // dispatcher callback at Background priority.
            try
            {
                var propName = e?.PropertyName ?? string.Empty;
                lock (_fieldChangeLock)
                {
                    if (!string.IsNullOrEmpty(propName))
                        _pendingFieldNames.Add(propName);

                    if (_fieldChangeScheduled)
                        return;

                    _fieldChangeScheduled = true;
                    _dispatcher.BeginInvoke(new Action(ProcessPendingFieldChanges), DispatcherPriority.Background);
                }
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex, "GLConfiguratorViewModel.OnFieldChanged: scheduling failed (non-fatal)");
            }
        }

        private void ProcessPendingFieldChanges()
        {
            // Not logged here (matches AIPowered's port of this same method): this
            // runs once per coalesced batch of field changes, but a bare "entry" with
            // no data was still frequent enough (1000+ times in one session) to be
            // noise. UpdateParameterSummary below logs the actual field data instead.
            try
            {
                // Clear the pending set and mark scheduling flag reset.
                // We previously captured the pending names into a local
                // variable for possible per-field handling, but currently
                // we only need to clear the set and proceed with a single
                // summary update — so avoid allocating an unused list.
                lock (_fieldChangeLock)
                {
                    _pendingFieldNames.Clear();
                    _fieldChangeScheduled = false;
                }

                // Perform a single update for all pending changes
                UpdateParameterSummary();

                // Notify visibility properties so bound UI elements can collapse/show rows
                OnPropertyChanged(nameof(IsPeriodVisible));
                OnPropertyChanged(nameof(IsStartEndVisible));
                OnPropertyChanged(nameof(IsEndPeriodVisible));
                OnPropertyChanged(nameof(IsBudgetVisible));
                OnPropertyChanged(nameof(IsEncumbranceVisible));
                OnPropertyChanged(nameof(IsJournalSourceVisible));
                OnPropertyChanged(nameof(IsJournalCategoryVisible));

                // Refresh enable state for fields that may depend on other fields
                try
                {
                    PeriodField?.RefreshEnableState();
                    StartDateField?.RefreshEnableState();
                    EndDateField?.RefreshEnableState();
                    EndPeriodField?.RefreshEnableState();
                    BudgetField?.RefreshEnableState();
                    EncumbranceField?.RefreshEnableState();
                    JournalSourceField?.RefreshEnableState();
                    JournalCategoryField?.RefreshEnableState();
                }
                catch (Exception ex)
                {
                    LogUtility.LogException(ex, "GLConfiguratorViewModel.ProcessPendingFieldChanges: failed to refresh field enable states (non-fatal)");
                }
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex, "GLConfiguratorViewModel.ProcessPendingFieldChanges");
            }
        }
        public async Task LoadConfiguratorAsync(bool ZeroesChecked, LedgerRecord ledger, List<string>? FuncArgs = null, List<string>? FuncValues = null)
        {
            LogUtility.LogDebug($"GLConfiguratorViewModel.LoadConfiguratorAsync: entry. ZeroesChecked={ZeroesChecked}, LedgerId={ledger?.LedgerId}, FuncArgs.Count={FuncArgs?.Count ?? 0}, FuncValues.Count={FuncValues?.Count ?? 0}");
            if (ledger == null)
            {
                LogUtility.LogWarn("GLConfiguratorViewModel.LoadConfiguratorAsync: ledger is null, aborting configurator load.");
                return;
            }
            await ResetWindowAsync();
            await LoadDataAsync(ledger);
            await UpdateUIAsync();

            if (FuncArgs == null && FuncValues == null)
            {
                LogUtility.LogDebug("GLConfiguratorViewModel.LoadConfiguratorAsync: no formula params supplied, applying default selections.");
                ApplyDefaultSelections();
            }
            else if (FuncArgs != null && FuncValues != null)
            {
                LogUtility.LogDebug("GLConfiguratorViewModel.LoadConfiguratorAsync: formula params supplied, applying formula params.");
                await ApplyFormulaParamsAsync(ZeroesChecked, FuncArgs, FuncValues);
            }

            RefreshAllFields();
            LogUtility.LogDebug("GLConfiguratorViewModel.LoadConfiguratorAsync: exit");
        }

        private async Task UpdateUIAsync()
        {
            LogUtility.LogDebug("GLConfiguratorViewModel.UpdateUIAsync: entry");
            await _dispatcher.InvokeAsync(() =>
            {
                ResetUIState();
                InitializeStaticCollections();

                OnPropertyChanged(nameof(Activities));
                OnPropertyChanged(nameof(Periods));
                OnPropertyChanged(nameof(Currencies));
                OnPropertyChanged(nameof(Budgets));
                OnPropertyChanged(nameof(JournalSources));
                OnPropertyChanged(nameof(JournalCategories));
            });
        }

        private void ResetUIState()
        {
            IsSignChecked = false;
            IsZeroesChecked = true;
            FactorText = "1";
        }

        private async Task LoadDataAsync(LedgerRecord ledger)
        {
            LogUtility.LogDebug($"GLConfiguratorViewModel.LoadDataAsync: entry. LedgerId={ledger?.LedgerId}, Coaid={ledger?.Coaid}");
            if (ledger == null)
            {
                LogUtility.LogWarn("GLConfiguratorViewModel.LoadDataAsync: ledger is null, aborting data load.");
                return;
            }
            var repo = new DataRepository();
            var appState = AppState.Instance;

            if (appState.SelectedCube == null)
            {
                LogUtility.LogWarn("GLConfiguratorViewModel.LoadDataAsync: AppState.Instance.SelectedCube is null, aborting data load.");
                return;
            }

            // Capture the ledger/cube identifiers into value-type locals before spinning up the
            // Task.Run lambdas below: these are simple long values, so the lambdas no longer need
            // to dereference "ledger" or "appState.SelectedCube" (reference types) from within the
            // closures, which keeps nullable flow analysis (and any concurrent access) unambiguous.
            long cubeId = appState.SelectedCube.CubeId;
            long ledgerId = ledger.LedgerId;
            long coaid = ledger.Coaid;

            // Force a fresh pull of ledger setup data (Periods, Activity, Currencies, etc.)
            // from the source system every time the configurator loads for a ledger, instead
            // of relying on whatever was cached the first time this ledger was ever opened.
            // Some ledgers use custom fiscal calendars that get extended with new future
            // periods over time (e.g. a "GOV Calendar" period set growing past its original
            // last period); without this refresh, PERIODS (and the other tables below) stayed
            // frozen at whatever was cached on first load, so the Start/End Date pickers in
            // GLBalanceConfigurator silently capped out at stale data even though the source
            // system had newer periods available.
            // LedgerDataRepository.InsertLedgerDataAsync (called at the end of this pipeline)
            // already does a proper DELETE-then-INSERT per cubeId/ledgerId/table (see its
            // ClearExistingData step), so it's safe to invoke on every load, not just the first.
            // Failures here (e.g. offline, API error) are logged and swallowed so the
            // configurator still falls back to whatever is already cached instead of blocking.
            //
            // TEMP DIAGNOSTIC (LogWarn, so it shows up even without ribbon Debug mode enabled) -
            // remove once the periods-refresh symptom is confirmed fixed:
            int periodsCountBeforeDiag = DataRepository.GetTableItemsCount(cubeId, ledgerId, "PERIODS");
            LogUtility.LogWarn($"[PeriodsRefreshDiag] LoadDataAsync: CubeId={cubeId}, LedgerId={ledgerId}. Cached PERIODS row count BEFORE refresh attempt: {periodsCountBeforeDiag}.");
            try
            {
                using var refreshCts = new CancellationHelper();
                await CommonFunctions.FillResponsibilitiesAsync(ledgerId, cubeId, refreshCts.GetToken());
                int periodsCountAfterDiag = DataRepository.GetTableItemsCount(cubeId, ledgerId, "PERIODS");
                LogUtility.LogWarn($"[PeriodsRefreshDiag] LoadDataAsync: CubeId={cubeId}, LedgerId={ledgerId}. Refresh call completed without throwing. Cached PERIODS row count AFTER refresh: {periodsCountAfterDiag}.");
            }
            catch (Exception ex)
            {
                LogUtility.LogWarn($"[PeriodsRefreshDiag] LoadDataAsync: CubeId={cubeId}, LedgerId={ledgerId}. Refresh call THREW: {ex.GetType().Name}: {ex.Message}");
                LogUtility.LogException(ex, $"GLConfiguratorViewModel.LoadDataAsync: failed to refresh ledger setup data from source for CubeId={cubeId}, LedgerId={ledgerId}; falling back to cached data.");
            }

            var tasks = new Task<object>[]
            {
                Task.Run(() => (object)repo.GetSegments(cubeId, ledgerId)),
                Task.Run(() => (object)repo.GetConfiguratorLedgers(cubeId, coaid, false)),
                Task.Run(() => (object)repo.GetActivities(cubeId, ledgerId)),
                Task.Run(() => (object)repo.GetPeriods(cubeId, ledgerId)),
                Task.Run(() => (object)repo.GetCurrencies(cubeId, ledgerId)),
                Task.Run(() => (object)repo.GetBudgets(cubeId, ledgerId)),
                Task.Run(() => (object)repo.GetEncumbrances(cubeId, ledgerId)),
                Task.Run(() => (object)repo.GetJournalSources(cubeId, ledgerId)),
                Task.Run(() => (object)repo.GetJournalCategories(cubeId, ledgerId))
            };

            var results = await Task.WhenAll(tasks);
            LogUtility.LogDebug("GLConfiguratorViewModel.LoadDataAsync: all data repository tasks completed.");

            ConfiguratorSegments = results[0] as ObservableCollection<SegmentModel> ?? new ObservableCollection<SegmentModel>();
            var ledgersData = results[1] as ObservableCollection<GenericLedgerModel> ?? new ObservableCollection<GenericLedgerModel>();
            PopulateDynamicCollections(Ledgers, ledgersData, l => l.PropertyChanged += Ledger_PropertyChanged);
            Activities = results[2] as ObservableCollection<ActivityModel> ?? new ObservableCollection<ActivityModel>();
            Periods = results[3] as ObservableCollection<PeriodModel> ?? new ObservableCollection<PeriodModel>();
            Currencies = results[4] as ObservableCollection<CurrencyModel> ?? new ObservableCollection<CurrencyModel>();
            Budgets = results[5] as ObservableCollection<BudgetModel> ?? new ObservableCollection<BudgetModel>();

            var encumbrancesData = results[6] as ObservableCollection<EncumbranceModel> ?? new ObservableCollection<EncumbranceModel>();
            PopulateDynamicCollections(Encumbrances, encumbrancesData, e => e.PropertyChanged += Encumbrance_PropertyChanged);
            // Enable ledger combo/ref if ledgers exist
            try
            {
                IsLedgerEnabled = Ledgers != null && Ledgers.Any();
                LedgerField.RefreshEnableState();
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex, "GLConfiguratorViewModel.LoadDataAsync: failed to refresh ledger field enable state (non-fatal)");
            }
            JournalSources = results[7] as ObservableCollection<JournalSourceModel> ?? new ObservableCollection<JournalSourceModel>();
            JournalCategories = results[8] as ObservableCollection<JournalCategoryModel> ?? new ObservableCollection<JournalCategoryModel>();
            LogUtility.LogDebug($"GLConfiguratorViewModel.LoadDataAsync: exit. Segments={ConfiguratorSegments?.Count ?? 0}, Ledgers={Ledgers?.Count ?? 0}, Activities={Activities?.Count ?? 0}, Periods={Periods?.Count ?? 0}, Currencies={Currencies?.Count ?? 0}, Budgets={Budgets?.Count ?? 0}, Encumbrances={Encumbrances?.Count ?? 0}, JournalSources={JournalSources?.Count ?? 0}, JournalCategories={JournalCategories?.Count ?? 0}");
        }

        private static void PopulateDynamicCollections<T>(ObservableCollection<T> collection, IEnumerable<T> data, Action<T> onAdd)
        {
            collection ??= new ObservableCollection<T>();
            collection.Clear();

            foreach (var item in data)
            {
                collection.Add(item);
                onAdd(item);
            }
            LogUtility.LogDebug($"GLConfiguratorViewModel.PopulateDynamicCollections: populated {collection.Count} item(s) of type {typeof(T).Name}.");
        }

        private void ApplyDefaultSelections()
        {
            LogUtility.LogDebug("GLConfiguratorViewModel.ApplyDefaultSelections: entry");
            try
            {
                ApplyDefaultLedgerSelection();
                ActivityField.ComboValue = Activities.FirstOrDefault(a => a.DisplayName == "NET");
                BalanceTypeField.ComboValue = BalanceTypes.FirstOrDefault(b => b.Name == AppConstants.BalanceTypePTD);
                CurrencyField.ComboValue = Currencies.FirstOrDefault(c => c.CurrencyCode == AppState.Instance.SelectedLedger.CurrencyCode);
                CurrencyTypeField.ComboValue = CurrencyTypes.FirstOrDefault(c => c.Name == AppConstants.CurrencyTypeTotal);
                ActualFlagField.ComboValue = ActualFlags.FirstOrDefault(a => a.Name == AppConstants.ActivityFlagActual);
                IsSignChecked = false;
                IsZeroesChecked = true;
                FactorText = "1";
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex, "GLConfiguratorViewModel.ApplyDefaultSelections");
            }
        }

        private void RefreshAllFields()
        {
            LogUtility.LogDebug("GLConfiguratorViewModel.RefreshAllFields: entry");
            var fields = new[]
            {
                LedgerField, ActivityField, BalanceTypeField, EndPeriodField,
                CurrencyField, CurrencyTypeField, ActualFlagField, BudgetField, EncumbranceField,
                StartDateField, EndDateField
            };

            foreach (var field in fields)
            {
                try { field.RefreshEnableState(); }
                catch (Exception ex) { LogUtility.LogException(ex, "GLConfiguratorViewModel.RefreshAllFields"); }
            }
        }

        private void InitializeStaticCollections()
        {
            LogUtility.LogDebug("GLConfiguratorViewModel.InitializeStaticCollections: entry");
            BalanceTypes = [
                new() { Name = AppConstants.BalanceTypePTD },
                new() { Name = AppConstants.BalanceTypeYTD },
                new() { Name = "QTD" },
                new() { Name = "PJTD" },
                new() { Name = AppConstants.BalanceTypeCTD },
                new() { Name = AppConstants.BalanceTypeJED },
                new() { Name = AppConstants.BalanceTypeJEDP },
                new() { Name = AppConstants.BalanceTypeJEDU }
            ];
            OnPropertyChanged(nameof(BalanceTypes));

            CurrencyTypes = [
                new() { Name = AppConstants.CurrencyTypeTotal, ShortName = AppConstants.CurrencyTypeTotal },
                new() { Name = AppConstants.CurrencyTypeEntered, ShortName = "E" },
                new() { Name = "Translated", ShortName = "T" },
                new() { Name = "Converted", ShortName = "C" }
            ];
            OnPropertyChanged(nameof(CurrencyTypes));

            ActualFlags = [
                new() { Name = AppConstants.ActivityFlagActual, ShortName = "A" },
                new() { Name = Budget, ShortName = "B" },
                new() { Name = Encumbrance, ShortName = "E" },
                new() { Name = AE, ShortName = AppConstants.ActualEncumbranceShort }
            ];
            OnPropertyChanged(nameof(ActualFlags));
        }

        private async Task ResetWindowAsync()
        {
            LogUtility.LogDebug("GLConfiguratorViewModel.ResetWindowAsync: entry");
            await _dispatcher.InvokeAsync(() =>
            {
                try
                {
                    IsSignChecked = false;
                    FactorText = "1";
                    SelectedPeriod = null!;

                    LedgerField.ComboText = string.Empty;

                    if (Ledgers != null && Ledgers.Any())
                    {
                        foreach (var led in Ledgers)
                        {
                            led.IsSelected = false;
                        }
                    }
                    LedgerField.ComboValue = null;
                    LedgerField.RefValue = null;

                    ActivityField.ComboText = string.Empty;
                    ActivityField.ComboValue = null;
                    ActivityField.RefValue = null;

                    PeriodField.ComboText = string.Empty;
                    PeriodField.ComboValue = null;
                    PeriodField.RefValue = null;

                    EndPeriodField.ComboText = string.Empty;
                    EndPeriodField.ComboValue = null;
                    EndPeriodField.RefValue = null;

                    StartDateField.ComboText = string.Empty;
                    StartDateField.ComboValue = null;
                    StartDateField.RefValue = null;
                    StartDateSelected = null;

                    EndDateField.ComboText = string.Empty;
                    EndDateField.ComboValue = null;
                    EndDateField.RefValue = null;
                    EndDateSelected = null;

                    BalanceTypeField.ComboText = string.Empty;
                    BalanceTypeField.ComboValue = null;
                    BalanceTypeField.RefValue = null;

                    CurrencyField.ComboText = string.Empty;
                    CurrencyField.ComboValue = null;
                    CurrencyField.RefValue = null;

                    CurrencyTypeField.ComboText = string.Empty;
                    CurrencyTypeField.ComboValue = null;
                    CurrencyTypeField.RefValue = null;

                    ActualFlagField.ComboText = string.Empty;
                    ActualFlagField.ComboValue = null;
                    ActualFlagField.RefValue = null;

                    BudgetField.ComboText = string.Empty;
                    BudgetField.ComboValue = null;
                    BudgetField.RefValue = null;

                    if (Encumbrances != null && Encumbrances.Any())
                    {
                        foreach (var encum in Encumbrances)
                        {
                            encum.IsSelected = false;
                        }
                    }

                    EncumbranceField.ComboText = string.Empty;
                    EncumbranceField.ComboValue = null;
                    EncumbranceField.RefValue = null;

                    JournalSourceField.ComboText = string.Empty;
                    JournalSourceField.ComboValue = null;
                    JournalSourceField.RefValue = null;

                    JournalCategoryField.ComboText = string.Empty;
                    JournalCategoryField.ComboValue = null;
                    JournalCategoryField.RefValue = null;

                    AccountAssignmentField.ComboText = string.Empty;
                    AccountAssignmentField.ComboValue = null;
                    AccountAssignmentField.RefValue = null;

                }
                catch (Exception ex)
                {
                    LogUtility.LogException(ex, "GLConfiguratorViewModel.ResetWindowAsync");
                }
            });
        }

        private async Task ApplyFormulaParamsAsync(bool zeroesChecked, List<string> FuncArgs, List<string> FuncValues)
        {
            LogUtility.LogDebug($"GLConfiguratorViewModel.ApplyFormulaParamsAsync: entry. zeroesChecked={zeroesChecked}, FuncArgs.Count={FuncArgs?.Count ?? 0}, FuncValues.Count={FuncValues?.Count ?? 0}");
            if (FuncArgs == null || FuncValues == null)
            {
                LogUtility.LogWarn("GLConfiguratorViewModel.ApplyFormulaParamsAsync: FuncArgs or FuncValues is null, aborting formula param application.");
                return;
            }
            await _dispatcher.InvokeAsync(() =>
            {
                IsZeroesChecked = zeroesChecked;
            });

            await ProcessSignAndFactor(FuncValues[0]);
            await ProcessLedgerFieldAsync(1, FuncArgs, FuncValues);
            await ProcessFieldAsync_Activity(2, ActivityField, Activities, x => x.DisplayName ?? x.ShortName, FuncArgs, FuncValues);
            await ProcessBalanceTypeAndPeriod(FuncArgs, FuncValues);
            await ProcessFieldAsync(5, CurrencyField, Currencies, x => x.CurrencyCode, FuncArgs, FuncValues);
            await ProcessFieldAsync_CurrencyType(6, CurrencyTypeField, CurrencyTypes, x => x.Name ?? x.ShortName, FuncArgs, FuncValues);
            await ProcessActualFlagAndBudgetEncumbrance(FuncArgs, FuncValues);
            await ProcessJls(FuncArgs, FuncValues); //Process Journal Sources and Categories
            await ProcessAccountAssignments(FuncArgs, FuncValues);

            UpdateParameterSummary();
            LogUtility.LogDebug("GLConfiguratorViewModel.ApplyFormulaParamsAsync: exit");
        }

        // Sign and Factor Processing from formula
        private async Task ProcessSignAndFactor(string rawValue)
        {
            LogUtility.LogDebug($"GLConfiguratorViewModel.ProcessSignAndFactor: rawValue={rawValue}");
            string cleanedArg = rawValue.Replace("\"", "");
            string operatorStr = "+";
            string valueStr = cleanedArg;

            if (!string.IsNullOrEmpty(cleanedArg))
            {
                char firstChar = cleanedArg[0];
                if (firstChar == '+' || firstChar == '-')
                {
                    operatorStr = firstChar.ToString();
                    valueStr = cleanedArg.Substring(1);
                }
            }

            await _dispatcher.InvokeAsync(() =>
            {
                IsSignChecked = operatorStr == "-";
                FactorText = valueStr;
            });
        }

        // End of Sign and Factor Processing from formula
        // Ledger Field Processing from formula
        private async Task ProcessLedgerFieldAsync(int index, List<string> FuncArgs, List<string> FuncValues)
        {
            string arg = FuncArgs[index].Replace("\"", "");
            string value = FuncValues[index].Replace("\"", "").Trim();
            LogUtility.LogDebug($"GLConfiguratorViewModel.ProcessLedgerFieldAsync: arg={arg}, value={value}");

            if (ExcelRangeHelper.IsRealRange(arg))
            {
                await SetLedgerField(arg, value);
            }
            else
            {
                await HandleLedgerValue(value);
            }
        }

        private async Task SetLedgerField(string? refValue, string? comboValue)
        {
            LogUtility.LogDebug($"GLConfiguratorViewModel.SetLedgerField: refValue={refValue}, comboValue={comboValue}");
            await _dispatcher.InvokeAsync(() =>
            {
                // Set RefValue - if null, set as null
                LedgerField.RefValue = refValue;

                // Handle comboValue - if null or empty, set both ComboValue and ComboText to null
                if (string.IsNullOrWhiteSpace(comboValue))
                {
                    LedgerField.ComboValue = null;
                    LedgerField.ComboText = null;
                    return;
                }

                // Parse comboValue (e.g. "Ledger1;Ledger2")
                var parts = (comboValue ?? string.Empty).Split(';').Select(p => p?.Trim() ?? string.Empty)
                    .Where(p => !string.IsNullOrEmpty(p)).ToList();

                // Track which parts actually exist in ledgers collection
                var existingParts = new List<string>();

                // Reset all IsSelected to false first
                if (Ledgers != null && Ledgers.Any())
                {
                    foreach (var ledger in Ledgers)
                    {
                        var isSelectedProp = ledger.GetType().GetProperty(AppConstants.PropIsSelected);
                        if (isSelectedProp != null)
                        {
                            isSelectedProp.SetValue(ledger, false);
                        }
                    }

                    // Update IsSelected for each ledger based on existing parts
                    foreach (var ledger in Ledgers)
                    {
                        // Get ledger name property (adjust based on your Ledger class)
                        var ledgerName = ledger.GetType().GetProperty(AppConstants.PropLedgerName)?.GetValue(ledger)?.ToString();

                        if (!string.IsNullOrWhiteSpace(ledgerName))
                        {
                        var matchedPart = parts.FirstOrDefault(part =>
                                string.Equals(ledgerName, part, StringComparison.OrdinalIgnoreCase));

                            bool isMatched = !string.IsNullOrEmpty(matchedPart);

                            if (isMatched)
                            {
                                var isSelectedProp = ledger.GetType().GetProperty(AppConstants.PropIsSelected);
                                if (isSelectedProp != null)
                                {
                                    isSelectedProp.SetValue(ledger, true);
                                }
                                existingParts.Add(matchedPart!);
                            }
                        }
                    }
                }

                // Set ComboValue and ComboText based on existing parts
                var finalComboValue = existingParts.Any()
                    ? string.Join(";", existingParts)
                    : null;

                LedgerField.ComboValue = finalComboValue;
                LedgerField.ComboText = finalComboValue;
            });
        }

        private async Task HandleLedgerValue(string value)
        {
            LogUtility.LogDebug($"GLConfiguratorViewModel.HandleLedgerValue: value={value}");
            // Don't clear here - let the processing methods handle it
            if (string.IsNullOrWhiteSpace(value))
            {
                await ProcessEmptyValue();
                return;
            }

            if (value.Contains(";"))
            {
                await ProcessMultipleLedgers(value);
            }
            else
            {
                await ProcessSingleLedger(value);
            }
        }
        private async Task ProcessEmptyValue()
        {
            // Reset all IsSelected to false
            if (Ledgers != null)
            {
                foreach (var ledger in Ledgers)
                {
                    var isSelectedProp = ledger.GetType().GetProperty(AppConstants.PropIsSelected);
                    if (isSelectedProp != null)
                    {
                        isSelectedProp.SetValue(ledger, false);
                    }
                }
            }

            // Set all fields to null
            await SetLedgerField(null, null);
        }

        private async Task ProcessMultipleLedgers(string value)
        {
            LogUtility.LogDebug($"GLConfiguratorViewModel.ProcessMultipleLedgers: value={value}");
            // Parse the input values
            var names = value.Split(';')
                .Select(n => n.Trim())
                .Where(n => !string.IsNullOrWhiteSpace(n))
                .ToList();

            if (Ledgers == null || !Ledgers.Any())
            {
                await SetLedgerField(null, null);
                return;
            }

            // Reset all IsSelected to false
            foreach (var ledger in Ledgers)
            {
                var isSelectedProp = ledger.GetType().GetProperty(AppConstants.PropIsSelected);
                if (isSelectedProp != null)
                {
                    isSelectedProp.SetValue(ledger, false);
                }
            }

            // Find existing ledgers and set IsSelected true
            var existingNames = new List<string>();

            foreach (string name in names)
            {
                var ledgerMatch = Ledgers.FirstOrDefault(x =>
                {
                    var ledgerName = x.GetType().GetProperty(AppConstants.PropLedgerName)?.GetValue(x)?.ToString();
                    return ledgerName == name;
                });

                if (ledgerMatch != null)
                {
                    var isSelectedProp = ledgerMatch.GetType().GetProperty(AppConstants.PropIsSelected);
                    if (isSelectedProp != null)
                    {
                        isSelectedProp.SetValue(ledgerMatch, true);
                    }
                    existingNames.Add(name);
                }
            }

            // Set ComboValue only with existing names
            var comboValueToSet = existingNames.Any()
                ? string.Join(";", existingNames)
                : null;

            await SetLedgerField(null, comboValueToSet);
        }

        private async Task ProcessSingleLedger(string value)
        {
            LogUtility.LogDebug($"GLConfiguratorViewModel.ProcessSingleLedger: value={value}");
            if (string.IsNullOrWhiteSpace(value))
            {
                await ProcessEmptyValue();
                return;
            }

            if (Ledgers == null || !Ledgers.Any())
            {
                await SetLedgerField(null, null);
                return;
            }

            // Reset all IsSelected to false first
            foreach (var ledger in Ledgers)
            {
                var isSelectedProp = ledger.GetType().GetProperty(AppConstants.PropIsSelected);
                if (isSelectedProp != null)
                {
                    isSelectedProp.SetValue(ledger, false);
                }
            }

            var trimmedValue = value.Trim();
            var match = Ledgers.FirstOrDefault(x =>
            {
                var ledgerName = x.GetType().GetProperty(AppConstants.PropLedgerName)?.GetValue(x)?.ToString();
                return ledgerName == trimmedValue;
            });

            if (match != null)
            {
                var isSelectedProp = match.GetType().GetProperty(AppConstants.PropIsSelected);
                if (isSelectedProp != null)
                {
                    isSelectedProp.SetValue(match, true);
                }
                await SetLedgerField(null, match.GetType().GetProperty(AppConstants.PropLedgerName)?.GetValue(match)?.ToString());
            }
            else
            {
                await SetLedgerField(null, null);
            }
        }

        // End of Ledger Field Processing from formula
        // Processing Period and BalanceType from formula
        private async Task ProcessBalanceTypeAndPeriod(List<string> FuncArgs, List<string> FuncValues)
        {
            string btText = FuncValues[4].Replace("\"", "").Trim();
            LogUtility.LogDebug($"GLConfiguratorViewModel.ProcessBalanceTypeAndPeriod: btText={btText}");
            await ProcessFieldAsync(4, BalanceTypeField, BalanceTypes, x => x.Name, FuncArgs, FuncValues);

            // Handle CTD (end period) logic
            if (string.Equals(btText, AppConstants.BalanceTypeCTD, StringComparison.OrdinalIgnoreCase) || IsBalanceTypeCtd())
            {
                await _dispatcher.InvokeAsync(() => IsEndPeriodsEnabled = true);
                string periodArg = FuncArgs[3].Replace("\"", "");
                string periodValues = FuncValues[3].Replace("\"", "");
                if (!string.IsNullOrEmpty(periodArg) && periodArg.Contains("~"))
                {
                    var parts = periodArg.Split('~');
                    var valueParts = periodValues.Split('~');
                    await ProcessPeriodPart(parts[0].Replace("&", ""), valueParts[0].Replace("&", ""), PeriodField, Periods);
                    if (parts.Length > 1)
                    {
                        await ProcessPeriodPart(parts[1].Replace("&", ""), valueParts[1].Replace("&", ""), EndPeriodField, Periods);
                    }
                }
            }
            else if (!string.IsNullOrWhiteSpace(btText) && (btText.Equals(AppConstants.BalanceTypeJED, StringComparison.OrdinalIgnoreCase) || btText.Equals(AppConstants.BalanceTypeJEDP, StringComparison.OrdinalIgnoreCase) || btText.Equals(AppConstants.BalanceTypeJEDU, StringComparison.OrdinalIgnoreCase)))
            {
                // JED variants: Period is not used. Parse StartDate~EndDate from arg/values (index 3)
                await _dispatcher.InvokeAsync(() =>
                {
                    // disable end periods UI
                    IsEndPeriodsEnabled = false;
                    EndPeriodField.ComboValue = null;
                    EndPeriodField.RefValue = null;

                    // clear period field
                    PeriodField.ComboValue = null;
                    PeriodField.RefValue = null;
                });

                string periodArg = FuncArgs[3].Replace("\"", "");
                string periodValues = FuncValues[3].Replace("\"", "");
                if (!string.IsNullOrEmpty(periodArg) && periodArg.Contains("~"))
                {
                    var parts = periodArg.Split('~');
                    var valueParts = periodValues.Split('~');

                    // Left part -> StartDate
                    var startArg = parts[0].Replace("&", "");
                    var startVal = valueParts.Length > 0 ? valueParts[0].Replace("&", "") : string.Empty;
                    if (ExcelRangeHelper.IsRealRange(startArg))
                    {
                        StartDateField.ComboValue = null;
                        StartDateField.RefValue = startArg;
                    }
                    else
                    {
                        StartDateField.RefValue = null;
                        StartDateField.ComboValue = startVal;
                        // try parse iso date into SelectedDate
                        if (DateTime.TryParse(startVal, CultureInfo.InvariantCulture, DateTimeStyles.None, out DateTime sd))
                            StartDateSelected = sd;
                    }

                    // Right part -> EndDate
                    if (parts.Length > 1)
                    {
                        var endArg = parts[1].Replace("&", "");
                        var endVal = valueParts.Length > 1 ? valueParts[1].Replace("&", "") : string.Empty;
                        if (ExcelRangeHelper.IsRealRange(endArg))
                        {
                            EndDateField.ComboValue = null;
                            EndDateField.RefValue = endArg;
                        }
                        else
                        {
                            EndDateField.RefValue = null;
                            EndDateField.ComboValue = endVal;
                            if (DateTime.TryParse(endVal, CultureInfo.InvariantCulture, DateTimeStyles.None, out DateTime ed))
                                EndDateSelected = ed;
                        }
                    }
                }
            }
            else
            {
                await _dispatcher.InvokeAsync(() =>
                {
                    IsEndPeriodsEnabled = false;
                    EndPeriodField.ComboValue = null;
                    EndPeriodField.RefValue = null;
                });
                await ProcessFieldAsync(3, PeriodField, Periods, x => x.PeriodName, FuncArgs, FuncValues);
            }
        }
        private async Task ProcessJls(List<string> FuncArgs, List<string> FuncValues)
        {
            LogUtility.LogDebug("GLConfiguratorViewModel.ProcessJls: entry");
            string btText = FuncValues[4].Replace("\"", "").Trim();
            var validBtTypes = new[] { AppConstants.BalanceTypePTD, AppConstants.BalanceTypeYTD, AppConstants.BalanceTypeCTD, AppConstants.BalanceTypeJED, AppConstants.BalanceTypeJEDP, AppConstants.BalanceTypeJEDU };
            bool validBt = validBtTypes.Any(valid => valid.Equals(btText, StringComparison.OrdinalIgnoreCase));

            string activityText = FuncValues[2].Replace("\"", "").Trim();
            var validActivities = new[] { "Debit", AppConstants.ActivityDR, "Credit", AppConstants.ActivityCR, "Net" };
            bool validActivity = validActivities.Any(valid => valid.Equals(activityText, StringComparison.OrdinalIgnoreCase));

            string currencyTypeText = FuncValues[6].Replace("\"", "").Trim();
            var validCurrencyTypes = new[] { "E", AppConstants.CurrencyTypeEntered.ToUpperInvariant(), AppConstants.CurrencyTypeTotal.ToUpperInvariant() }; // Example currency types
            bool validCurrencyType = validCurrencyTypes.Any(valid => valid.Equals(currencyTypeText, StringComparison.OrdinalIgnoreCase));

            if (!validBt || !validActivity || !validCurrencyType)
            {
                await _dispatcher.InvokeAsync(() =>
                {
                    JournalSourceField.ComboValue = null;
                    JournalSourceField.RefValue = null;

                    JournalCategoryField.ComboValue = null;
                    JournalCategoryField.RefValue = null;
                });
            }
            else
            {
                await ProcessFieldAsync(9, JournalSourceField, JournalSources, x => x.SourceName, FuncArgs, FuncValues);
                await ProcessFieldAsync(10, JournalCategoryField, JournalCategories, x => x.CategoryName, FuncArgs, FuncValues);
            }
        }
        private async Task ProcessPeriodPart(string partArg, string partValue, dynamic field, IEnumerable<dynamic> periods)
        {
            LogUtility.LogDebug($"GLConfiguratorViewModel.ProcessPeriodPart: partArg={partArg}, partValue={partValue}");
            string cleanArg = partArg.Replace("\"", "");
            string cleanValue = partValue.Replace("\"", "");
            if (ExcelRangeHelper.IsRealRange(cleanArg))
            {
                await _dispatcher.InvokeAsync(() =>
                {
                    field.ComboValue = null;
                    field.RefValue = cleanArg;
                });
            }
            else
            {
                await _dispatcher.InvokeAsync(() => field.RefValue = null);
            }
            var match = periods.FirstOrDefault(x => x.PeriodName == cleanValue);
            if (match != null)
            {
                await _dispatcher.InvokeAsync(() =>
                {
                    field.ComboValue = match;
                    if (field == PeriodField)
                    {
                        SelectedPeriod = match;
                        if (field == PeriodField)
                        {
                            UpdateEndPeriods();
                        }
                    }
                });
            }
        }

        // End of Processing Period and BalanceType from formula

        private async Task ProcessFieldAsync<T>(int index, dynamic field, IEnumerable<T> items, Func<T, string> selector, List<string> FuncArgs, List<string> FuncValues) where T : class
        {
            string arg = FuncArgs[index].Replace("\"", "");
            string value = FuncValues[index].Replace("\"", "").Trim();
            LogUtility.LogDebug($"GLConfiguratorViewModel.ProcessFieldAsync<{typeof(T).Name}>: index={index}, arg={arg}, value={value}");

            if (ExcelRangeHelper.IsRealRange(arg))
            {
                await _dispatcher.InvokeAsync(() =>
                {
                    field.ComboValue = null;
                    field.RefValue = arg;
                });
            }
            else
            {
                await _dispatcher.InvokeAsync(() =>
                {
                    field.RefValue = null;
                    field.ComboValue = null;
                });
            }
            var match = items.FirstOrDefault(x => selector(x) == value);

            if (match != null)
            {
                await _dispatcher.InvokeAsync(() =>
                {
                    field.ComboValue = match;
                });
            }
        }
        private async Task ProcessFieldAsync_Activity<T>(int index, dynamic field, IEnumerable<T> items, Func<T, string> selector, List<string> FuncArgs, List<string> FuncValues) where T : class
        {
            string arg = FuncArgs[index].Replace("\"", "");
            string value = FuncValues[index].Replace("\"", "").Trim();
            LogUtility.LogDebug($"GLConfiguratorViewModel.ProcessFieldAsync_Activity: index={index}, arg={arg}, value={value}");

            if (ExcelRangeHelper.IsRealRange(arg))
            {
                await _dispatcher.InvokeAsync(() =>
                {
                    field.ComboValue = null;
                    field.RefValue = arg;
                });
            }
            else
            {
                await _dispatcher.InvokeAsync(() =>
                {
                    field.RefValue = null;
                    field.ComboValue = null;
                });

                var match = items.FirstOrDefault(x =>
                {
                    var selectedValue = selector(x);
                    if (!string.IsNullOrWhiteSpace(selectedValue) && selectedValue.Equals(value, StringComparison.OrdinalIgnoreCase))
                        return true;

                    var nameProp = x.GetType().GetProperty("DisplayName");
                    var shortNameProp = x.GetType().GetProperty(AppConstants.PropShortName);
                    var nameVal = nameProp?.GetValue(x)?.ToString();
                    var shortNameVal = shortNameProp?.GetValue(x)?.ToString();

                    return (!string.IsNullOrWhiteSpace(nameVal) && string.Equals(nameVal, value, StringComparison.OrdinalIgnoreCase)) ||
                           (!string.IsNullOrWhiteSpace(shortNameVal) && string.Equals(shortNameVal, value, StringComparison.OrdinalIgnoreCase));
                });

                if (match != null)
                {
                    await _dispatcher.InvokeAsync(() =>
                    {
                        field.ComboValue = match;
                    });
                }
            }
        }
        private async Task ProcessFieldAsync_CurrencyType<T>(int index, dynamic field, IEnumerable<T> items, Func<T, string> selector, List<string> FuncArgs, List<string> FuncValues) where T : class
        {
            string arg = FuncArgs[index].Replace("\"", "");
            string value = FuncValues[index].Replace("\"", "").Trim();
            LogUtility.LogDebug($"GLConfiguratorViewModel.ProcessFieldAsync_CurrencyType: index={index}, arg={arg}, value={value}");

            if (ExcelRangeHelper.IsRealRange(arg))
            {
                await _dispatcher.InvokeAsync(() =>
                {
                    field.ComboValue = null;
                    field.RefValue = arg;
                });
            }
            else
            {
                await _dispatcher.InvokeAsync(() =>
                {
                    field.RefValue = null;
                    field.ComboValue = null;
                });

                var match = items.FirstOrDefault(x =>
                {
                    var selectedValue = selector(x);
                    if (!string.IsNullOrWhiteSpace(selectedValue) && selectedValue.Equals(value, StringComparison.OrdinalIgnoreCase))
                        return true;

                    var nameProp = x.GetType().GetProperty("Name");
                    var shortNameProp = x.GetType().GetProperty(AppConstants.PropShortName);
                    var nameVal = nameProp?.GetValue(x)?.ToString();
                    var shortNameVal = shortNameProp?.GetValue(x)?.ToString();

                    return (!string.IsNullOrWhiteSpace(nameVal) && string.Equals(nameVal, value, StringComparison.OrdinalIgnoreCase)) ||
                           (!string.IsNullOrWhiteSpace(shortNameVal) && string.Equals(shortNameVal, value, StringComparison.OrdinalIgnoreCase));
                });

                if (match != null)
                {
                    await _dispatcher.InvokeAsync(() =>
                    {
                        field.ComboValue = match;
                    });
                }
            }
        }
        private async Task ProcessActualFlagFieldAsync(int index, dynamic field, List<string> FuncArgs, List<string> FuncValues)
        {
            string arg = FuncArgs[index].Replace("\"", "");
            string value = FuncValues[index].Replace("\"", "").Trim();
            LogUtility.LogDebug($"GLConfiguratorViewModel.ProcessActualFlagFieldAsync: index={index}, arg={arg}, value={value}");

            if (ExcelRangeHelper.IsRealRange(arg))
            {
                await _dispatcher.InvokeAsync(() =>
                {
                    field.ComboValue = null;
                    field.RefValue = arg;
                });
            }
            else
            {
                await _dispatcher.InvokeAsync(() =>
                {
                    field.RefValue = null;
                    field.ComboValue = null;
                });

                var match = ActualFlags.FirstOrDefault(x => x.Name == value || x.ShortName == value);

                if (match != null)
                {
                    await _dispatcher.InvokeAsync(() =>
                    {
                        field.ComboValue = match;
                    });
                }
            }
        }
        private static bool ValidateBudgetForCurrencyType(List<string> FuncValues)
        {
            var currencyTypeValue = FuncValues[6].Replace("\"", "").Trim();
            if (!string.IsNullOrEmpty(currencyTypeValue) && currencyTypeValue.Equals(AppConstants.CurrencyTypeTotal, StringComparison.OrdinalIgnoreCase))
            {
                return true;
            }
                return false;
        }
        private async Task ProcessActualFlagAndBudgetEncumbrance(List<string> FuncArgs, List<string> FuncValues)
        {
            LogUtility.LogDebug("GLConfiguratorViewModel.ProcessActualFlagAndBudgetEncumbrance: entry");
            string afText = FuncValues[7].Replace("\"", "").Trim();
            var matchAF = ActualFlags.FirstOrDefault(x => (x.Name ?? "") == afText || (x.ShortName ?? "") == afText);

            if (matchAF != null)
            {
                string afType = matchAF.ShortName;

                if (afType == "B" && !ValidateBudgetForCurrencyType(FuncValues))
                {
                    await _dispatcher.InvokeAsync(() =>
                    {
                        ActualFlagField.ComboValue = null;
                        ActualFlagField.RefValue = null;

                        IsBudgetEnabled = false;
                        BudgetField.ComboValue = null;
                        BudgetField.RefValue = null;

                        ClearEncumbranceSelections();
                    });
                    return;
                }

                await ProcessActualFlagFieldAsync(7, ActualFlagField, FuncArgs, FuncValues);

                if (afType == "B")
                {
                    await _dispatcher.InvokeAsync(() =>
                    {
                        IsBudgetEnabled = true;
                        IsEncumbranceEnabled = false;
                        ClearEncumbranceSelections();
                    });
                    await ProcessSimpleField(8, BudgetField, FuncArgs, FuncValues);
                }
                else if (afType == "E" || afType == AppConstants.ActualEncumbranceShort)
                {
                    await _dispatcher.InvokeAsync(() =>
                    {
                        IsBudgetEnabled = false;
                        IsEncumbranceEnabled = true;
                        BudgetField.ComboValue = null;
                        BudgetField.RefValue = null;
                    });
                    await ProcessEncumbranceField(8, FuncArgs, FuncValues);
                }
                else
                {
                    await _dispatcher.InvokeAsync(() =>
                    {
                        IsBudgetEnabled = false;
                        IsEncumbranceEnabled = false;
                        BudgetField.ComboValue = null;
                        BudgetField.RefValue = null;
                        ClearEncumbranceSelections();
                    });
                }
            }
        }

        private async Task ProcessSimpleField(int index, dynamic field, List<string> FuncArgs, List<string> FuncValues)
        {
            string arg = FuncArgs[index].Replace("\"", "");
            string value = FuncValues[index].Replace("\"", "").Trim();
            LogUtility.LogDebug($"GLConfiguratorViewModel.ProcessSimpleField: index={index}, arg={arg}, value={value}");

            if (ExcelRangeHelper.IsRealRange(arg))
            {
                await _dispatcher.InvokeAsync(() =>
                {
                    field.ComboValue = null;
                    field.RefValue = arg;
                });
            }
            else
            {
                await _dispatcher.InvokeAsync(() =>
                {
                    field.RefValue = null;
                    field.ComboValue = value;
                });
            }
        }

        private async Task ProcessEncumbranceField(int index, List<string> FuncArgs, List<string> FuncValues)
        {
            string arg = FuncArgs[index].Replace("\"", "");
            string value = FuncValues[index].Replace("\"", "");
            LogUtility.LogDebug($"GLConfiguratorViewModel.ProcessEncumbranceField: index={index}, arg={arg}, value={value}");

            if (ExcelRangeHelper.IsRealRange(arg))
            {
                await _dispatcher.InvokeAsync(() =>
                {
                    EncumbranceField.ComboValue = null;
                    EncumbranceField.RefValue = arg;
                });
                return;
            }

            await _dispatcher.InvokeAsync(() =>
            {
                EncumbranceField.RefValue = null;

                if (Encumbrances != null)
                {
                    foreach (var encumbrance in Encumbrances)
                    {
                        encumbrance.IsSelected = false;
                    }
                }

                EncumbranceField.ComboValue = null;
                EncumbranceField.ComboText = string.Empty;
            });

            var selectedEncumbrances = new List<EncumbranceModel>();

            if (value.Contains(";"))
            {
                var names = value.Split(';').Select(n => n.Trim()).Where(n => !string.IsNullOrWhiteSpace(n));
                foreach (string name in names)
                {
                    var match = Encumbrances.FirstOrDefault(x => x.EncumbranceType == name);
                    if (match != null)
                    {
                        match.IsSelected = true;
                        selectedEncumbrances.Add(match);
                    }
                }

                var selectedText = string.Join(";", selectedEncumbrances.Select(e => e.EncumbranceType));
                await _dispatcher.InvokeAsync(() =>
                {
                    EncumbranceField.ComboText = selectedText;
                    EncumbranceField.ComboValue = selectedEncumbrances.FirstOrDefault();
                });
            }
            else
            {
                var match = Encumbrances.FirstOrDefault(x => x.EncumbranceType == value.Trim());
                if (match != null)
                {
                    match.IsSelected = true;
                    await _dispatcher.InvokeAsync(() =>
                    {
                        EncumbranceField.ComboText = match.EncumbranceType;
                        EncumbranceField.ComboValue = match;
                    });
                }
            }
        }

        private void ProcessAccountAssignments(FieldBinding field, string refText, string? rngValue)
        {
            AccountAssignmentField.RefValue = refText;
            AccountAssignmentField.ComboValue = rngValue;
            AccountAssignmentField.RefreshEnableState();
        }

        private async Task ProcessAccountAssignments(List<string> FuncArgs, List<string> FuncValues)
        {
            string arg11 = GetFormulaTextAt(FuncArgs, 11);
            string val11 = GetFormulaTextAt(FuncValues, 11);
            LogUtility.LogDebug($"GLConfiguratorViewModel.ProcessAccountAssignments(FuncArgs,FuncValues): arg11={arg11}, val11={val11}, ConfiguratorSegments.Count={ConfiguratorSegments?.Count ?? 0}");

            if (ExcelRangeHelper.IsRealRange(arg11) && val11.Contains(";"))
            {
                await _dispatcher.InvokeAsync(() =>
                {
                    AccountAssignmentField.RefValue = arg11;
                    AccountAssignmentField.ComboValue = val11;
                });
            }
            else
            {
                var segList = new List<string>();
                int segmentCount = ConfiguratorSegments?.Count ?? 0;
                for (int i = 0; i < segmentCount; i++)
                {
                    segList.Add(GetFormulaTextAt(FuncArgs, 11 + i));
                }

                string finalResult = string.Join(";", segList);
                await _dispatcher.InvokeAsync(() =>
                {
                    AccountAssignmentField.RefValue = null;
                    AccountAssignmentField.ComboValue = finalResult;
                });
            }
        }

        private static string GetFormulaTextAt(IReadOnlyList<string> values, int index)
        {
            if (values == null || index < 0 || index >= values.Count)
                return string.Empty;

            return values[index].Replace("\"", "").Trim();
        }

        private void ApplyDefaultLedgerSelection()
        {
            LogUtility.LogDebug("GLConfiguratorViewModel.ApplyDefaultLedgerSelection: entry");
            if (AppState.Instance.SelectedLedger == null)
            {
                LogUtility.LogWarn("GLConfiguratorViewModel.ApplyDefaultLedgerSelection: AppState.Instance.SelectedLedger is null, aborting default ledger selection.");
                return;
            }

            var match = Ledgers.FirstOrDefault(x => x.LedgerId == AppState.Instance.SelectedLedger.LedgerId);

            if (match != null)
            {
                match.IsSelected = true;
                LedgerField.ComboValue = match.LedgerName;
                LedgerField.ComboText = match.LedgerName;
                OnPropertyChanged(nameof(IsLedgerEnabled));
            }
            else
            {
                LogUtility.LogWarn($"GLConfiguratorViewModel.ApplyDefaultLedgerSelection: no ledger match found for LedgerId={AppState.Instance.SelectedLedger.LedgerId}.");
            }
        }

        private void Ledger_PropertyChanged(object sender, PropertyChangedEventArgs e)
        {
            LogUtility.LogDebug($"GLConfiguratorViewModel.Ledger_PropertyChanged: PropertyName={e?.PropertyName}");
            if (e == null)
            {
                LogUtility.LogWarn("GLConfiguratorViewModel.Ledger_PropertyChanged: PropertyChangedEventArgs is null, ignoring event.");
                return;
            }
            if (e.PropertyName == nameof(GenericLedgerModel.IsSelected))
            {
                OnPropertyChanged(nameof(SelectedLedgers));
                UpdateParameterSummary();
                // Ensure the LedgerField enable state is refreshed so refedit becomes enabled
                try { LedgerField?.RefreshEnableState(); } catch (Exception ex) { LogUtility.LogException(ex, "GLConfiguratorViewModel.Ledger_PropertyChanged: failed to refresh ledger field enable state (non-fatal)"); }
            }
        }

        private void Encumbrance_PropertyChanged(object sender, PropertyChangedEventArgs e)
        {
            LogUtility.LogDebug($"GLConfiguratorViewModel.Encumbrance_PropertyChanged: PropertyName={e?.PropertyName}");
            if (e == null)
            {
                LogUtility.LogWarn("GLConfiguratorViewModel.Encumbrance_PropertyChanged: PropertyChangedEventArgs is null, ignoring event.");
                return;
            }
            if (e.PropertyName == nameof(EncumbranceModel.IsSelected))
            {
                OnPropertyChanged(nameof(SelectedEncumbrances));
                UpdateParameterSummary();
                // Ensure the EncumbranceField enable state is refreshed so refedit becomes enabled
                try { EncumbranceField?.RefreshEnableState(); } catch (Exception ex) { LogUtility.LogException(ex, "GLConfiguratorViewModel.Encumbrance_PropertyChanged: failed to refresh encumbrance field enable state (non-fatal)"); }
            }
        }

        public virtual void OnFieldDependencyChanged(FieldBinding changedField)
        {
            LogUtility.LogDebug($"GLConfiguratorViewModel.OnFieldDependencyChanged: FieldType={changedField?.Type}, ComboValue={changedField?.ComboValue}, RefValue={changedField?.RefValue}");
            if (changedField == null) return;

            switch (changedField.Type)
            {
                case FieldBinding.FieldType.AccountAssignments:
                    if ((changedField.ComboValue == null || string.IsNullOrWhiteSpace(changedField.ComboValue.ToString())) &&
                        (changedField.RefValue == null || string.IsNullOrWhiteSpace(changedField.RefValue.ToString())))
                    {
                        ResetField(changedField);
                    }
                    break;
                case FieldBinding.FieldType.Activity:
                case FieldBinding.FieldType.CurrencyType:
                    ValidateJournalFields();
                    break;
                case FieldBinding.FieldType.ActualFlag:
                    var val = changedField.ComboValue?.ToString();

                    // Handle dependency clearing only — no enable toggles
                    switch (val)
                    {
                        case Budget:
                        case "B":
                            IsBudgetEnabled = true;
                            IsEncumbranceEnabled = false;
                            ClearEncumbranceSelections();
                            break;
                        case Encumbrance:
                        case "E":
                        case AE:
                        case AppConstants.ActualEncumbranceShort:
                            IsBudgetEnabled = false;
                            IsEncumbranceEnabled = true;
                            ResetField(BudgetField);
                            break;
                        case AppConstants.ActivityFlagActual:
                        case "A":
                        case null:
                        case "":
                        default:
                            IsBudgetEnabled = false;
                            IsEncumbranceEnabled = false;
                            ResetField(BudgetField);
                            ClearEncumbranceSelections();
                            break;
                    }
                    break;

                case FieldBinding.FieldType.BalanceType:
                    // Handle BalanceType → EndPeriod dependency and also journal enabling
                    var balanceVal = GetBalanceTypeText();
                    var isCtd = string.Equals(balanceVal, AppConstants.BalanceTypeCTD, StringComparison.OrdinalIgnoreCase);
                    var isJED = !string.IsNullOrWhiteSpace(balanceVal) && (
                        balanceVal.Equals(AppConstants.BalanceTypeJED, StringComparison.OrdinalIgnoreCase) ||
                        balanceVal.Equals(AppConstants.BalanceTypeJEDP, StringComparison.OrdinalIgnoreCase) ||
                        balanceVal.Equals(AppConstants.BalanceTypeJEDU, StringComparison.OrdinalIgnoreCase)
                    );

                    // Clear/Reset fields depending on balance type
                    if (isJED)
                    {
                        // JED uses dates - clear any period selections
                        ResetField(PeriodField);
                        ResetField(EndPeriodField);
                        IsEndPeriodsEnabled = false;
                        EndPeriods.Clear();

                        // Ensure date fields refresh their enable state
                        StartDateField?.RefreshEnableState();
                        EndDateField?.RefreshEnableState();
                    }
                    else
                    {
                        // Non-JED use periods - clear date selections
                        ResetField(StartDateField);
                        ResetField(EndDateField);

                        // EndPeriod enabled only for CTD
                        ResetField(EndPeriodField);
                        IsEndPeriodsEnabled = isCtd;
                        if (!isCtd)
                            EndPeriods.Clear();
                        else
                            // This case fires on every interactive Balance Type combo pick
                            // (FieldBinding.ComboValue's setter calls OnFieldDependencyChanged).
                            // Previously only ResetField/IsEndPeriodsEnabled ran here, so
                            // switching to CTD enabled the End Period row but never actually
                            // populated its EndPeriods collection - UpdateEndPeriods() was only
                            // wired up for the Period-changed case, ApplyDefaultSelections, and
                            // formula-param loading, not for this direct BalanceType change.
                            UpdateEndPeriods();
                    }

                    OnPropertyChanged(nameof(IsEndPeriodsEnabled));
                    // Refresh period/date enable states
                    PeriodField?.RefreshEnableState();
                    EndPeriodField?.RefreshEnableState();
                    StartDateField?.RefreshEnableState();
                    EndDateField?.RefreshEnableState();

                    ValidateJournalFields();
                    // Update available currency types depending on whether this is a JED variant.
                    try
                    {
                        UpdateCurrencyTypesForBalanceType(isJED);
                    }
                    catch (Exception ex)
                    {
                        LogUtility.LogException(ex, "GLConfiguratorViewModel.OnFieldDependencyChanged: UpdateCurrencyTypesForBalanceType failed (non-fatal)");
                    }

                    break;
                case FieldBinding.FieldType.JournalSources:
                case FieldBinding.FieldType.JournalCategories:
                    if ((changedField.ComboValue != null && !string.IsNullOrWhiteSpace(changedField.ComboValue.ToString())) ||
                       (changedField.RefValue != null && !string.IsNullOrWhiteSpace(changedField.RefValue.ToString())))
                    {
                        try
                        {
                            UpdateCurrencyTypesForBalanceType(true);
                        }
                        catch (Exception ex)
                        {
                            LogUtility.LogException(ex, "GLConfiguratorViewModel.OnFieldDependencyChanged: UpdateCurrencyTypesForBalanceType when changing journals failed (non-fatal)");
                        }
                    }
                    break;
                case FieldBinding.FieldType.Period:
                    // Handle BalanceType → EndPeriod dependency
                    ResetField(EndPeriodField);
                    SelectedPeriod = (PeriodModel)changedField.OwnerViewModel.PeriodField.ComboValue;
                    UpdateEndPeriods();
                    break;
            }
        }
        private void ValidateJournalFields()
        {
            LogUtility.LogDebug("GLConfiguratorViewModel.ValidateJournalFields: entry");
            var activityValue = GetFieldValue(ActivityField);
            var balanceTypeValue = GetFieldValue(BalanceTypeField);
            var currencyTypeValue = GetFieldValue(CurrencyTypeField);

            bool IsJournalActivityValid;
            bool IsJournalBalanceTypeValid;
            bool IsJournalCurrencyTypeValid;

            // Validate Activity for Journals
            if (!string.IsNullOrEmpty(activityValue))
            {
                var validActivities = new[] { "DEBIT", AppConstants.ActivityDR, "CREDIT", AppConstants.ActivityCR, "NET" };
                IsJournalActivityValid = validActivities.Any(valid => valid.Equals(activityValue, StringComparison.OrdinalIgnoreCase));
            }
            else
            {
                IsJournalActivityValid = true; // No activity selected, don't block
            }

            // Validate Balance Type for Journals
            if (!string.IsNullOrEmpty(balanceTypeValue))
            {
                var validBalanceTypes = new[] { AppConstants.BalanceTypePTD, AppConstants.BalanceTypeYTD, AppConstants.BalanceTypeCTD };
                IsJournalBalanceTypeValid = validBalanceTypes.Any(valid => valid.Equals(balanceTypeValue, StringComparison.OrdinalIgnoreCase));
            }
            else
            {
                IsJournalBalanceTypeValid = true; // No balance type selected, don't block
            }

            if (!string.IsNullOrEmpty(currencyTypeValue))
            {
                var validCurrencyTypes = new[] { "E", AppConstants.CurrencyTypeEntered.ToUpperInvariant(), AppConstants.CurrencyTypeTotal.ToUpperInvariant()};
                IsJournalCurrencyTypeValid = validCurrencyTypes.Any(valid => valid.Equals(currencyTypeValue, StringComparison.OrdinalIgnoreCase));
            }
            else
            {
                IsJournalCurrencyTypeValid = true; // No currency type selected, don't block
            }

            if (!IsJournalActivityValid || !IsJournalBalanceTypeValid || !IsJournalCurrencyTypeValid)
            {
                ResetField(JournalSourceField);
                ResetField(JournalCategoryField);
            }
            else
            {
                // If all are valid, we can allow the user to select Journal Source and Category
                JournalSourceField.RefreshEnableState();
                JournalCategoryField.RefreshEnableState();
            }
        }
        public void OnRefEditTextChanged(FieldBinding field, string refText)
        {
            LogUtility.LogDebug($"GLConfiguratorViewModel.OnRefEditTextChanged: FieldType={field?.Type}, refText={refText}");
            if (field == null)
            {
                LogUtility.LogWarn("GLConfiguratorViewModel.OnRefEditTextChanged: field is null, ignoring ref-edit text change.");
                return;
            }
            if (string.IsNullOrWhiteSpace(refText))
            {
                // Reset if cleared
                ResetField(field);
                return;
            }

            string? rngValue = GetRangeValueSafe(refText);
            if (string.IsNullOrWhiteSpace(rngValue))
            {
                LogUtility.LogWarn($"GLConfiguratorViewModel.OnRefEditTextChanged: referenced cell \"{refText}\" resolved to an empty value, resetting field {field?.Type}.");
                // field was already confirmed non-null by the guard at the top of this
                // method; null-forgiving here just satisfies flow analysis across the
                // intervening GetRangeValueSafe call.
                ResetField(field!);
                return;
            }

            ProcessFieldReference(field, refText, rngValue);
        }

        private void ProcessFieldReference(FieldBinding field, string refText, string? rngValue)
        {
            LogUtility.LogDebug($"GLConfiguratorViewModel.ProcessFieldReference: FieldType={field?.Type}, refText={refText}, rngValue={rngValue}");
            if (field == null)
            {
                LogUtility.LogWarn("GLConfiguratorViewModel.ProcessFieldReference: field is null, cannot resolve a field processor.");
                return;
            }
            var processor = GetFieldProcessor(field.Type);
            processor?.Invoke(field, refText, rngValue);
        }

        private Action<FieldBinding, string, string?>? GetFieldProcessor(FieldBinding.FieldType type)
        {
            return type switch
            {
                FieldBinding.FieldType.Ledger => ProcessLedger,
                FieldBinding.FieldType.Activity => ProcessActivity,
                FieldBinding.FieldType.BalanceType => ProcessBalanceType,
                FieldBinding.FieldType.Period => ProcessPeriod,
                FieldBinding.FieldType.EndPeriod => ProcessEndPeriod,
                FieldBinding.FieldType.StartDate => (f, r, v) => ProcessStartDate(f, r, v),
                FieldBinding.FieldType.EndDate => (f, r, v) => ProcessEndDate(f, r, v),
                FieldBinding.FieldType.Currency => (f, r, v) => ApplyRefHelper(f, Currencies, r, v, (c, v) => c.CurrencyCode.Equals(v, StringComparison.OrdinalIgnoreCase), false),
                FieldBinding.FieldType.CurrencyType => ProcessCurrencyType,
                FieldBinding.FieldType.ActualFlag => (f, r, v) => ApplyRefHelper(f, ActualFlags, r, v, (a, v) => a.Name.Equals(v, StringComparison.OrdinalIgnoreCase) || a.ShortName.Equals(v, StringComparison.OrdinalIgnoreCase), false),
                FieldBinding.FieldType.Budgets => (f, r, v) => ApplyRefHelper(f, Budgets, r, v, (b, v) => b.BudgetName.Equals(v, StringComparison.OrdinalIgnoreCase), false),
                FieldBinding.FieldType.Encumbrances => ProcessEncumbrances,
                FieldBinding.FieldType.JournalSources => (f, r, v) => ApplyRefHelper(f, JournalSources, r, v, (j, v) => j.JeSourceName.Equals(v, StringComparison.OrdinalIgnoreCase) || j.SourceName.Equals(v, StringComparison.OrdinalIgnoreCase), false),
                FieldBinding.FieldType.JournalCategories => (f, r, v) => ApplyRefHelper(f, JournalCategories, r, v, (j, v) => j.JeCategoryName.Equals(v, StringComparison.OrdinalIgnoreCase) || j.CategoryName.Equals(v, StringComparison.OrdinalIgnoreCase), false),
                FieldBinding.FieldType.AccountAssignments => ProcessAccountAssignments,
                _ => null
            };
        }

        private void ProcessLedger(FieldBinding field, string refText, string? rngValue)
        {
            LogUtility.LogDebug($"GLConfiguratorViewModel.ProcessLedger: refText={refText}, rngValue={rngValue}");
            ApplyRefHelper(field, Ledgers, refText, rngValue,
                (l, t) => l.LedgerName.Equals(t, StringComparison.OrdinalIgnoreCase), true);
        }

        private void ProcessActivity(FieldBinding field, string refText, string? rngValue)
        {
            LogUtility.LogDebug($"GLConfiguratorViewModel.ProcessActivity: refText={refText}, rngValue={rngValue}");
            ApplyRefHelper(field, Activities, refText, rngValue,
                (a, t) =>
                {
                    if (a == null || a.ActivityType == null || string.IsNullOrEmpty(a.ActivityType)) return false;
                    var parts = a.ActivityType.Split(':');
                    var left = parts[0];
                    var right = parts.Length > 1 ? parts[1] : "";
                    return left.Equals(t, StringComparison.OrdinalIgnoreCase) || right.Equals(t, StringComparison.OrdinalIgnoreCase);
                }, false);

            ValidateJournalFields();
        }

        private void ProcessBalanceType(FieldBinding field, string refText, string? rngValue)
        {
            LogUtility.LogDebug($"GLConfiguratorViewModel.ProcessBalanceType: refText={refText}, rngValue={rngValue}");
            ApplyRefHelper(field, BalanceTypes, refText, rngValue,
                (b, t) => b.Name.Equals(t, StringComparison.OrdinalIgnoreCase), false);

            if (!string.Equals(field.ComboValue?.ToString(), AppConstants.BalanceTypeCTD, StringComparison.OrdinalIgnoreCase))
            {
                ResetField(EndPeriodField);
                OnPropertyChanged(nameof(IsEndPeriodsEnabled));
            }

            ValidateJournalFields();
            // Ensure currency types are updated when balance type is set via refedit
            try
            {
                var bt = GetBalanceTypeText();
                var isJED = !string.IsNullOrWhiteSpace(bt) && (
                    bt.Equals(AppConstants.BalanceTypeJED, StringComparison.OrdinalIgnoreCase) ||
                    bt.Equals(AppConstants.BalanceTypeJEDP, StringComparison.OrdinalIgnoreCase) ||
                    bt.Equals(AppConstants.BalanceTypeJEDU, StringComparison.OrdinalIgnoreCase)
                );

                UpdateCurrencyTypesForBalanceType(isJED);
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex, "GLConfiguratorViewModel.ProcessBalanceType: UpdateCurrencyTypesForBalanceType failed (non-fatal)");
            }
        }
        private void ProcessCurrencyType(FieldBinding field, string refText, string? rngValue)
        {
            LogUtility.LogDebug($"GLConfiguratorViewModel.ProcessCurrencyType: refText={refText}, rngValue={rngValue}");
            ApplyRefHelper(field, CurrencyTypes, refText, rngValue,
                (c, v) => c.Name.Equals(v, StringComparison.OrdinalIgnoreCase) || c.ShortName.Equals(v, StringComparison.OrdinalIgnoreCase), false);

            ValidateJournalFields();
        }

        private void ProcessPeriod(FieldBinding field, string refText, string? rngValue)
        {
            LogUtility.LogDebug($"GLConfiguratorViewModel.ProcessPeriod: refText={refText}, rngValue={rngValue}");
            ApplyRefHelper(field, Periods, refText, rngValue,
                (p, t) => p.PeriodName.Equals(t, StringComparison.OrdinalIgnoreCase), false);
            UpdateEndPeriods();
        }

        private void ProcessEndPeriod(FieldBinding field, string refText, string? rngValue)
        {
            LogUtility.LogDebug($"GLConfiguratorViewModel.ProcessEndPeriod: refText={refText}, rngValue={rngValue}");
            if (!IsValidEndPeriodSequence(GetFieldValue(PeriodField), rngValue))
            {
                LogUtility.LogWarn($"GLConfiguratorViewModel.ProcessEndPeriod: invalid end period sequence for rngValue={rngValue}, resetting field.");
                ShowWarningAction?.Invoke("Please select or reference a Period before selecting End Period.");
                ResetField(field);
                return;
            }

            ApplyRefHelper(field, Periods, refText, rngValue,
                (p, t) => p.PeriodName.Equals(t, StringComparison.OrdinalIgnoreCase), false);
        }

        private bool IsValidEndPeriodSequence(string? periodVal, string? endPeriodValue)
        {
            LogUtility.LogDebug($"GLConfiguratorViewModel.IsValidEndPeriodSequence: periodVal={periodVal}, endPeriodValue={endPeriodValue}");
            if (Periods == null)
                return true;

            if (string.IsNullOrWhiteSpace(periodVal) || string.IsNullOrWhiteSpace(endPeriodValue))
                return false;

            string? newPeriodValue = null;

            // Resolve reference if needed
            if (!string.IsNullOrEmpty(periodVal) && ExcelRangeHelper.IsRealRange(periodVal))
            {
                try
                {
                    if (_excelApp != null)
                    {
                        newPeriodValue = (string?)_excelApp.Range[periodVal]?.Value;
                    }
                    else
                    {
                        return false;
                    }
                }
                catch (Exception ex)
                {
                    LogUtility.LogException(ex, $"GLConfiguratorViewModel.IsValidEndPeriodSequence: failed to read Excel range value for periodVal={periodVal}");
                    return false;
                }
            }

            if (string.IsNullOrWhiteSpace(newPeriodValue))
                return false;

            var startPeriod = Periods?.FirstOrDefault(p => p.PeriodName.Equals(periodVal, StringComparison.OrdinalIgnoreCase));
            var endPeriod = Periods?.FirstOrDefault(p => p.PeriodName.Equals(endPeriodValue, StringComparison.OrdinalIgnoreCase));

            if (startPeriod == null || endPeriod == null)
                return true; // Can't validate without matches, allow

            if (Periods == null)
                return true;

            int startIdx = Periods.IndexOf(startPeriod);
            int endIdx = Periods.IndexOf(endPeriod);

            return startIdx < 0 || endIdx < 0 || endIdx >= startIdx;
        }

        private void ProcessEncumbrances(FieldBinding field, string refText, string? rngValue)
        {
            ApplyRefHelper(EncumbranceField, Encumbrances, refText, rngValue,
                (e, t) => e.EncumbranceType.Equals(t, StringComparison.OrdinalIgnoreCase) ||
                          (int.TryParse(t, out int id) && e.EncumbranceTypeId == id), true);
        }

        private void ClearEncumbranceSelections()
        {
            EncumbranceField.ComboValue = null;
            EncumbranceField.ComboText = null;
            EncumbranceField.RefValue = null;

            if (Encumbrances != null)
            {
                foreach (var encumbrance in Encumbrances)
                {
                    encumbrance.IsSelected = false;
                }
            }

            EncumbranceField.ComboText = string.Empty;
            OnPropertyChanged(nameof(SelectedEncumbrances));
            EncumbranceField.RefreshEnableState();
        }

        private string? GetRangeValueSafe(string refText)
        {
            LogUtility.LogDebug($"GLConfiguratorViewModel.GetRangeValueSafe: refText={refText}");
            try
            {
                var val = _excelApp?.Range[refText]?.Value2;
                if (val == null)
                    return null;

                // If Excel returns a string, return it directly
                if (val is string s)
                    return s;

                // If Excel returns a DateTime, format to ISO yyyy-MM-dd
                if (val is DateTime dt)
                    return dt.ToString(AppConstants.DateFormatIso);

                // Excel often returns dates as doubles (OADate)
                if (val is double d)
                {
                    try
                    {
                        var date = DateTime.FromOADate(d);
                        return date.ToString(AppConstants.DateFormatIso);
                    }
                    catch (Exception ex)
                    {
                        LogUtility.LogException(ex, $"GLConfiguratorViewModel.GetRangeValueSafe: failed to convert OADate to DateTime for refText={refText} (falling back)");
                    }
                }

                // Fallback: attempt to convert to string
                return val.ToString();
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex, $"GLConfiguratorViewModel.GetRangeValueSafe: failed to read Excel range value for refText={refText}");
                return null;
            }
        }

        private void ResetField(FieldBinding field)
        {
            LogUtility.LogDebug($"GLConfiguratorViewModel.ResetField: FieldType={field?.Type}");
            if (field == null) return;

            if (field.Type == FieldBinding.FieldType.Ledger && field.ComboValue is GenericLedgerModel && Ledgers != null)
            {
                foreach (var ledgerItem in Ledgers)
                {
                    ledgerItem.IsSelected = false;
                }
            }

            if (field.Type == FieldBinding.FieldType.Encumbrances && field.ComboValue is IEnumerable<EncumbranceModel> && Encumbrances != null)
            {
                foreach (var encumbrance in Encumbrances)
                {
                    encumbrance.IsSelected = false;
                }
            }

            // Clear core field tokens
            field.RefValue = null;
            field.ComboText = null;
            field.ComboValue = null;

            // Clear any related selected properties so UI controls (DatePickers, Period selection) reflect the reset
            try
            {
                switch (field.Type)
                {
                    case FieldBinding.FieldType.StartDate:
                        StartDateSelected = null;
                        // Also clear combo tokens already done above
                        break;
                    case FieldBinding.FieldType.EndDate:
                        EndDateSelected = null;
                        break;
                    case FieldBinding.FieldType.Period:
                        // Clear selected period and update end periods
                        SelectedPeriod = null!;
                        UpdateEndPeriods();
                        // Reset EndPeriod as well
                        if (EndPeriodField != null)
                        {
                            EndPeriodField.ComboValue = null;
                            EndPeriodField.ComboText = null;
                            EndPeriodField.RefValue = null;
                        }
                        IsEndPeriodsEnabled = false;
                        OnPropertyChanged(nameof(IsEndPeriodsEnabled));
                        break;
                    case FieldBinding.FieldType.EndPeriod:
                        // Nothing extra beyond clearing combo/ref
                        break;
                    case FieldBinding.FieldType.Ledger:
                        // Ensure all ledger selections cleared
                        try
                        {
                            if (Ledgers != null)
                            {
                                foreach (var led in Ledgers)
                                    led.IsSelected = false;
                            }
                        }
                        catch (Exception ex)
                        {
                            LogUtility.LogException(ex, "GLConfiguratorViewModel.ResetField: failed to clear ledger selections (non-fatal)");
                        }
                        break;
                    case FieldBinding.FieldType.Encumbrances:
                        // Ensure encumbrance selections cleared
                        try
                        {
                            if (Encumbrances != null)
                            {
                                foreach (var enc in Encumbrances)
                                    enc.IsSelected = false;
                            }
                        }
                        catch (Exception ex)
                        {
                            LogUtility.LogException(ex, "GLConfiguratorViewModel.ResetField: failed to clear encumbrance selections (non-fatal)");
                        }
                        break;
                }
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex, $"GLConfiguratorViewModel.ResetField: failed to clear dependent state for {field.Type}");
            }

            // Refresh enable state for this field and any related fields so UI enables are updated
            field.RefreshEnableState();
            try { StartDateField?.RefreshEnableState(); }
            catch (Exception ex)
            {
                LogUtility.LogException(ex, "GLConfiguratorViewModel.ResetField: failed to refresh StartDateField enable state (non-fatal)");
            }
            try { EndDateField?.RefreshEnableState(); }
            catch (Exception ex)
            {
                LogUtility.LogException(ex, "GLConfiguratorViewModel.ResetField: failed to refresh EndDateField enable state (non-fatal)");
            }
            try { PeriodField?.RefreshEnableState(); }
            catch (Exception ex)
            {
                LogUtility.LogException(ex, "GLConfiguratorViewModel.ResetField: failed to refresh PeriodField enable state (non-fatal)");
            }
            try { EndPeriodField?.RefreshEnableState(); }
            catch (Exception ex)
            {
                LogUtility.LogException(ex, "GLConfiguratorViewModel.ResetField: failed to refresh EndPeriodField enable state (non-fatal)");
            }
        }

        // Update CurrencyTypes collection based on whether BalanceType is a JED variant.
        // For JED/JEDP/JEDU only allow AppConstants.CurrencyTypeTotal and AppConstants.CurrencyTypeEntered. Clear the field binding
        // if the currently selected currency type is not allowed.
        private void UpdateCurrencyTypesForBalanceType(bool isJED)
        {
            LogUtility.LogDebug($"GLConfiguratorViewModel.UpdateCurrencyTypesForBalanceType: isJED={isJED}");
            try
            {
                // Ensure operation runs on UI thread
                _dispatcher.InvokeAsync(() =>
                {
                    if (CurrencyTypes == null)
                        CurrencyTypes = new System.Collections.ObjectModel.ObservableCollection<CurrencyTypeModel>();

                    CurrencyTypes.Clear();

                    if (isJED)
                    {
                        CurrencyTypes.Add(new CurrencyTypeModel { Name = AppConstants.CurrencyTypeTotal, ShortName = AppConstants.CurrencyTypeTotal });
                        CurrencyTypes.Add(new CurrencyTypeModel { Name = AppConstants.CurrencyTypeEntered, ShortName = "E" });

                        // If current currency type is not allowed for JED, clear it
                        var current = GetFieldValue(CurrencyTypeField);
                        if (!string.IsNullOrWhiteSpace(current))
                        {
                            var allowed = new[] { AppConstants.CurrencyTypeTotal, AppConstants.CurrencyTypeEntered, "E" };
                            if (!allowed.Any(a => a.Equals(current, StringComparison.OrdinalIgnoreCase)))
                            {
                                ResetField(CurrencyTypeField);
                            }
                        }
                    }
                    else
                    {
                        CurrencyTypes.Add(new CurrencyTypeModel { Name = AppConstants.CurrencyTypeTotal, ShortName = AppConstants.CurrencyTypeTotal });
                        CurrencyTypes.Add(new CurrencyTypeModel { Name = AppConstants.CurrencyTypeEntered, ShortName = "E" });
                        CurrencyTypes.Add(new CurrencyTypeModel { Name = "Translated", ShortName = "T" });
                        CurrencyTypes.Add(new CurrencyTypeModel { Name = "Converted", ShortName = "C" });
                    }

                    OnPropertyChanged(nameof(CurrencyTypes));
                    CurrencyTypeField?.RefreshEnableState();
                });
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex, $"GLConfiguratorViewModel.UpdateCurrencyTypesForBalanceType: failed (non-fatal). isJED={isJED}");
            }
        }
        private static void ApplyRefHelper<T>(
                FieldBinding field,
                IEnumerable<T> collection,
                string refText, string? rngValue,
                RefMatch<T> matchFunc,
                bool isMultiSelect = false) where T : class
        {
            LogUtility.LogDebug($"GLConfiguratorViewModel.ApplyRefHelper<{typeof(T).Name}>: refText={refText}, rngValue={rngValue}, isMultiSelect={isMultiSelect}");
            if (collection == null) return;

            if (isMultiSelect)
            {
                // Split input (e.g. Ledger1;Ledger2)
                var parts = rngValue?.Split(';').Select(p => p.Trim())
                    .Where(p => !string.IsNullOrEmpty(p)).ToList() ?? new List<string>();

                // Track which parts exist in collection
                var existingParts = new List<string>();

                foreach (var item in collection)
                {
                    var prop = item.GetType().GetProperty(AppConstants.PropIsSelected);
                    if (prop != null)
                    {
                        // Check if this item matches any part
                        var matchedPart = parts.FirstOrDefault(part => matchFunc(item, part));
                        bool isMatched = matchedPart != null;
                        prop.SetValue(item, isMatched);

                        // If matched, add the original part value to existingParts
                        if (isMatched && matchedPart != null)
                        {
                            existingParts.Add(matchedPart);
                        }
                    }
                }

                field.RefValue = refText;

                // Set ComboValue based on existing parts
                field.ComboValue = existingParts.Any()
                    ? string.Join(";", existingParts)
                    : null;
                field.ComboText = existingParts.Any()
                    ? string.Join(";", existingParts)
                    : null;
            }
            else
            {
                var match = collection?.FirstOrDefault(m => matchFunc(m, rngValue ?? string.Empty));
                field.RefValue = refText;
                if (match != null)
                {
                    field.ComboValue = match;
                }
                else
                {
                    field.ComboValue = default;
                }
            }
        }

        public void UpdateParameterSummary()
        {
            // Prevent re-entrant updates that can cause recursive PropertyChanged
            // calls and lead to StackOverflowException. If an update is already
            // in progress, skip this invocation.
            if (_isUpdatingParameterSummary)
            {
                LogUtility.LogDebug("GLConfiguratorViewModel.UpdateParameterSummary: re-entrant call detected, skipping.");
                return;
            }

            try
            {
                _isUpdatingParameterSummary = true;

                var fieldValues = CollectAllFieldValues();
                // Not logged: this is the data the user already sees rendered live in
                // the parameter summary text on the window itself, so logging it here
                // too was pure duplication of what's already visible on screen.

                // Use existing FlowDocument instance to avoid re-assigning the
                // document property (which can trigger bindings and layout).
                var doc = ParameterDisplayText ?? CreateFormattedDocument();

                // Ensure basic formatting is applied
                doc.PagePadding = new Thickness(4, 4, 4, 4);
                doc.TextAlignment = TextAlignment.Left;
                doc.FontSize = 11;
                doc.FontFamily = new FontFamily("Segoe UI");

                var blueBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString(AppConstants.GLAccentHex));
                var blackBrush = Brushes.Black;

                // Create the main paragraph
                var paragraph = new Paragraph
                {
                    Margin = new Thickness(0),
                    LineHeight = 18, // Consistent line height
                    TextAlignment = TextAlignment.Left
                };

                // Format each parameter with proper spacing
                for (int i = 0; i < fieldValues.Count; i++)
                {
                    var kvp = fieldValues.ElementAt(i);

                    // Add comma and space for all except first item
                    if (i > 0)
                    {
                        paragraph.Inlines.Add(new Run(", ") { Foreground = blackBrush, FontSize = 11 });
                    }

                    // Add parameter name in black
                    paragraph.Inlines.Add(new Run(kvp.Key) { Foreground = blackBrush, FontSize = 11, FontWeight = FontWeights.Normal });

                    // Add equals sign and space
                    paragraph.Inlines.Add(new Run(" = ") { Foreground = blackBrush, FontSize = 11 });

                    // Add quoted value in blue
                    paragraph.Inlines.Add(new Run($"\"{kvp.Value}\"") { Foreground = blueBrush, FontSize = 11 });
                }

                // Add final period
                paragraph.Inlines.Add(new Run(".") { Foreground = blackBrush, FontSize = 11 });

                // Replace blocks on existing document instead of assigning new instance
                try
                {
                    doc.Blocks.Clear();
                    doc.Blocks.Add(paragraph);
                }
                catch (Exception ex)
                {
                    LogUtility.LogException(ex, "GLConfiguratorViewModel.UpdateParameterSummary: failed to update FlowDocument blocks");
                }

                // Ensure ParameterDisplayText points to the current doc instance.
                if (!ReferenceEquals(ParameterDisplayText, doc))
                {
                    ParameterDisplayText = doc;
                }
            }
            finally
            {
                _isUpdatingParameterSummary = false;
            }
        }

        private static FlowDocument CreateFormattedDocument()
        {
            return new FlowDocument
            {
                PagePadding = new Thickness(4, 4, 4, 4),
                ColumnWidth = double.PositiveInfinity,
                TextAlignment = TextAlignment.Left,
                FontSize = 11,
                FontFamily = new FontFamily("Segoe UI"),
                IsOptimalParagraphEnabled = true,
                IsHyphenationEnabled = false
            };
        }

        private Dictionary<string, string> CollectAllFieldValues()
        {
            var values = new Dictionary<string, string>();

            values["Change Sign"] = IsSignChecked ? "True" : "False";
            values["Ledger"] = GetFieldValue(LedgerField);
            values["Activity"] = GetFieldValue(ActivityField);
            values["Balance Type"] = GetFieldValue(BalanceTypeField);
            // Period handling: always expose a single key "Periods". Prefer references, then values.
            var btType = BalanceTypeField?.ComboValue?.ToString() ?? string.Empty;
            string periodsValue;
            bool isJED = !string.IsNullOrWhiteSpace(btType) && (
                btType.Equals(AppConstants.BalanceTypeJED, StringComparison.OrdinalIgnoreCase) ||
                btType.Equals(AppConstants.BalanceTypeJEDP, StringComparison.OrdinalIgnoreCase) ||
                btType.Equals(AppConstants.BalanceTypeJEDU, StringComparison.OrdinalIgnoreCase));

            // NOTE: this "Periods" token is the one summary value with its own explicit
            // "prefer references, then values" contract (FormatSummaryToken's own comment,
            // just below) - it needs the RAW reference (GetFormulaFieldValue), not
            // GetFieldValue's resolved value, or FormatSummaryToken's own IsRealRange check
            // would never see a reference to prefer in the first place. Every other field in
            // this summary intentionally shows GetFieldValue's resolved value (this panel is
            // a read-only preview of current values, not the formula itself).
            if (isJED)
            {
                var startRaw = GetFormulaFieldValue(StartDateField);
                var endRaw = GetFormulaFieldValue(EndDateField);
                periodsValue = FormatSummaryToken(startRaw) + "~" + FormatSummaryToken(endRaw);
            }
            else if (!string.IsNullOrWhiteSpace(btType) && btType.Equals(AppConstants.BalanceTypeCTD, StringComparison.OrdinalIgnoreCase))
            {
                var periodRaw = GetFormulaFieldValue(PeriodField);
                var endPeriodRaw = GetFormulaFieldValue(EndPeriodField);
                periodsValue = FormatSummaryToken(periodRaw) + "~" + FormatSummaryToken(endPeriodRaw);
            }
            else
            {
                periodsValue = FormatSummaryToken(GetFormulaFieldValue(PeriodField));
            }

            values["Periods"] = periodsValue;
            values["Currency"] = GetFieldValue(CurrencyField);
            values["Currency Type"] = GetFieldValue(CurrencyTypeField);
            values["Actual Flag"] = GetFieldValue(ActualFlagField);
            AddConditionalBudgetEncumbrance(values);
            values["Journal Source"] = GetFieldValue(JournalSourceField);
            values["Journal Category"] = GetFieldValue(JournalCategoryField);

            // Conditional fields
            AddAccountSegments(values);

            return values;
        }

        private void AddConditionalBudgetEncumbrance(Dictionary<string, string> values)
        {
            var afType = ActualFlagField?.ComboValue?.ToString();
            if (string.IsNullOrWhiteSpace(afType)) return;

            switch (afType)
            {
                case Budget:
                case "B":
                    values[Budget] = GetFieldValue(BudgetField);
                    break;

                case Encumbrance:
                case "E":
                case AE:
                case AppConstants.ActualEncumbranceShort:
                    var encVal = GetFieldValue(EncumbranceField) ?? "";
                    // Only collect value for summary — do NOT mutate field state here.
                    // Mutating field properties while building the summary can cause
                    // recursive PropertyChanged notifications and lead to stack
                    // overflow. Keep UpdateParameterSummary read-only.
                    values[Encumbrance] = encVal;
                    break;
            }
        }
        private void AddAccountSegments(Dictionary<string, string> values)
        {
            string accountVal = GetResolvedAccountAssignmentValue();

            string[] segments;

            if (string.IsNullOrWhiteSpace(accountVal))
            {
                segments = new string[ConfiguratorSegments.Count];
            }
            else
            {
                segments = accountVal.Split(new[] { ';' }, StringSplitOptions.None);
            }

            for (int i = 0; i < ConfiguratorSegments.Count; i++)
            {
                string segmentValue = i < segments.Length && !string.IsNullOrWhiteSpace(segments[i])
                    ? segments[i].Replace("\"", "")
                    : "";
                values[ConfiguratorSegments[i].SegmentName] = segmentValue;
            }
        }

        public void WriteFormulaToCell(Microsoft.Office.Interop.Excel.Range rng)
        {
            LogUtility.LogDebug("GLConfiguratorViewModel.WriteFormulaToCell: entry");
            try
            {
                var formulaArgs = BuildFormulaArguments();

                if (formulaArgs.Count == 0)
                {
                    // Error already shown in BuildFormulaArguments
                    LogUtility.LogWarn("GLConfiguratorViewModel.WriteFormulaToCell: no formula arguments produced (validation failed), aborting write.");
                    return;
                }

                //--- Step 9: Construct and write formula
                var finalFormula = "=@" + "GLSense_GetBalance(" + string.Join(",", formulaArgs) + ",\"\")";
                try
                {
                    rng.Formula = finalFormula;
                    LogUtility.LogDebug($"GLConfiguratorViewModel.WriteFormulaToCell: wrote formula '{finalFormula}'");
                }
                catch (Exception ex)
                {
                    ShowWarningAction?.Invoke("Failed to write formula to the specified cell.");
                    LogUtility.LogException(ex, $"GLConfiguratorViewModel.WriteFormulaToCell: failed to set formula on Excel range. Formula : {finalFormula}");
                    return;
                }
                rng.EntireColumn.ColumnWidth = 20;
                rng.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignRight;
                rng.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignBottom;
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex, "GLConfiguratorViewModel.WriteFormulaToCell: failed to write formula to excel cell.");
            }
        }

        private List<string> BuildFormulaArguments()
        {
            LogUtility.LogDebug("GLConfiguratorViewModel.BuildFormulaArguments: entry");
            var signChecked = GetSignFactor();

            try
            {
                ValidateMandatoryFields();
            }
            catch (InvalidOperationException ex)
            {
                LogUtility.LogException(ex, "GLConfiguratorViewModel.BuildFormulaArguments: mandatory field validation failed");
                ShowWarningAction?.Invoke(ex.Message);
                return new List<string>();
            }

            var finalPeriodVal = GetFinalPeriodValue();
            var budEncum = GetBudgetEncumbranceValue();
            var accountSegments = GetAccountSegments();

            var formulaParts = new List<string>
            {
                FormatFormulaArg(signChecked),
                FormatFormulaArg(GetFormulaFieldValue(LedgerField)),
                FormatFormulaArg(GetFormulaFieldValue(ActivityField)),
                FormatFormulaArg(finalPeriodVal),
                FormatFormulaArg(GetFormulaFieldValue(BalanceTypeField)),
                FormatFormulaArg(GetFormulaFieldValue(CurrencyField)),
                FormatFormulaArg(GetFormulaFieldValue(CurrencyTypeField)),
                FormatFormulaArg(GetFormulaFieldValue(ActualFlagField)),
                FormatFormulaArg(budEncum),
                FormatFormulaArg(GetFormulaFieldValue(JournalSourceField)),
                FormatFormulaArg(GetFormulaFieldValue(JournalCategoryField))
            };

            formulaParts.AddRange(accountSegments);
            return formulaParts;
        }

        private string GetSignFactor()
        {
            var textFactor = string.IsNullOrWhiteSpace(FactorText) ? "1" : FactorText;
            return (IsSignChecked ? "-" : "+") + textFactor;
        }
        private bool IsBtJournalsType()
        {
            var bt = GetFieldValue(BalanceTypeField);
            if (!string.IsNullOrWhiteSpace(bt) &&
                (bt.Equals(AppConstants.BalanceTypeJED, StringComparison.OrdinalIgnoreCase) ||
                 bt.Equals(AppConstants.BalanceTypeJEDP, StringComparison.OrdinalIgnoreCase) ||
                 bt.Equals(AppConstants.BalanceTypeJEDU, StringComparison.OrdinalIgnoreCase)))
            {
                return true;
            }

            return false;
        }
        private void ValidateMandatoryFields()
        {
            LogUtility.LogDebug("GLConfiguratorViewModel.ValidateMandatoryFields: entry");
            ValidateRequiredFields(
                (LedgerField, "Ledger"),
                (ActivityField, "Activity"),
                (BalanceTypeField, "Balance Type"),
                (CurrencyField, "Currency"),
                (CurrencyTypeField, "Currency Type"),
                (ActualFlagField, "Actual Flag"));

            ValidateBalanceTypeSpecificFields();

            if (!HasValidAccountAssignment())
                throw new InvalidOperationException("Account Assignments not selected.");
        }

        private void ValidateBalanceTypeSpecificFields()
        {
            if (IsBtJournalsType())
            {
                ValidateRequired(StartDateField, "Start Date is required for selected Balance Type.");
                ValidateRequired(EndDateField, "End Date is required for selected Balance Type.");
                ValidateDateRange();
                return;
            }

            ValidateRequired(PeriodField, "Period");
        }

        private void ValidateDateRange()
        {
            if (StartDateSelected.HasValue &&
                EndDateSelected.HasValue &&
                StartDateSelected.Value > EndDateSelected.Value)
            {
                throw new InvalidOperationException("Start Date cannot be greater than End Date.");
            }
        }

        private void ValidateRequiredFields(params (FieldBinding Field, string NameOrMessage)[] fields)
        {
            foreach (var (field, nameOrMessage) in fields)
                ValidateRequired(field, nameOrMessage);
        }

        private void ValidateRequired(FieldBinding field, string nameOrMessage)
        {
            if (field == null)
                throw new InvalidOperationException("Field binding is missing.");

            if (!string.IsNullOrWhiteSpace(GetFieldValue(field)))
                return;

            var message = nameOrMessage.IndexOf("required", StringComparison.OrdinalIgnoreCase) >= 0
                ? nameOrMessage
                : $"{nameOrMessage} is required.";

            throw new InvalidOperationException(message);
        }
        private string GetFinalPeriodValue()
        {
            var btType = BalanceTypeField?.ComboValue?.ToString() ?? string.Empty;
            LogUtility.LogDebug($"GLConfiguratorViewModel.GetFinalPeriodValue: btType={btType}");

            // JED variants use StartDate~EndDate instead of Period
            if (!string.IsNullOrWhiteSpace(btType) &&
                (btType.Equals(AppConstants.BalanceTypeJED, StringComparison.OrdinalIgnoreCase) ||
                 btType.Equals(AppConstants.BalanceTypeJEDP, StringComparison.OrdinalIgnoreCase) ||
                 btType.Equals(AppConstants.BalanceTypeJEDU, StringComparison.OrdinalIgnoreCase)))
            {
                var s = GetFormulaFieldValue(StartDateField);
                var e = GetFormulaFieldValue(EndDateField);

                if (string.IsNullOrWhiteSpace(s) || string.IsNullOrWhiteSpace(e))
                    throw new InvalidOperationException("Start Date and End Date are required for selected Balance Type.");

                // Return formatted as StartDate~EndDate
                return CombinePeriod(s, e);
            }

            // CTD Balance Type uses Period~EndPeriod
            if (btType?.Equals(AppConstants.BalanceTypeCTD, StringComparison.OrdinalIgnoreCase) == true)
            {
                var endPeriodVal = GetFormulaFieldValue(EndPeriodField);
                if (string.IsNullOrWhiteSpace(endPeriodVal))
                    throw new InvalidOperationException("End Period is required for CTD Balance Type.");
                return CombinePeriod(GetFormulaFieldValue(PeriodField), endPeriodVal);
            }

            // Other balance types use single Period
            return GetFormulaFieldValue(PeriodField);
        }

        private string GetBudgetEncumbranceValue()
        {
            var afType = ActualFlagField?.ComboValue?.ToString();
            if (string.IsNullOrWhiteSpace(afType)) return "";

            return afType switch
            {
                Budget or "B" => GetFormulaFieldValue(BudgetField),
                Encumbrance or "E" or AE or AppConstants.ActualEncumbranceShort => GetFormulaFieldValue(EncumbranceField),
                _ => ""
            };
        }

        private List<string> GetAccountSegments()
        {
            var accountVal = GetFormulaFieldValue(AccountAssignmentField);
            if (string.IsNullOrWhiteSpace(accountVal)) return new List<string>();

            if (ExcelRangeHelper.IsRealRange(accountVal))
            {
                return new List<string> { accountVal.Trim() };
            }

            var segments = SplitAccountAssignmentSegments(accountVal);
            if (segments.Count == 0)
                return new List<string>();

            return segments
                .Select(p => FormatFormulaArg(p.Replace("\"", "").Trim()))
                .ToList();
        }

        private bool HasValidAccountAssignment()
        {
            var accountVal = GetFieldValue(AccountAssignmentField);
            return !string.IsNullOrWhiteSpace(accountVal);
        }

        private static string FormatFormulaArg(string value)
        {
            if (string.IsNullOrEmpty(value))
            {
                // Truly empty argument → ""
                return "\"\"";
            }

            // Cell reference (contains ! or $) → do not quote
            if (value.Contains("!") || value.Contains("$"))
            {
                return value;
            }

            // Already a concatenation (e.g., period & "~" & endPeriod)
            if (value.Contains("&"))
            {
                return value;
            }

            // Otherwise, quote everything
            return $"\"{value.Replace("\"", "")}\"";
        }

        private static string CombinePeriod(string periodVal, string endPeriodVal)
        {
            var periodStr = FormatFormulaArg(periodVal);
            var endPeriodStr = FormatFormulaArg(endPeriodVal);
            // combine respecting reference/value
            return periodStr + "&\"~\"&" + endPeriodStr;
        }

        // Format token for summary display: prefer Excel references (unquoted) otherwise quote literal values
        private static string FormatSummaryToken(string token)
        {
            if (string.IsNullOrWhiteSpace(token)) return string.Empty;
            var t = token.Trim();

            try
            {
                if (ExcelRangeHelper.IsRealRange(t))
                    return t;
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex, $"GLConfiguratorViewModel.FormatSummaryToken: ExcelRangeHelper.IsRealRange failed for token={t} (falling back to quoting)");
                // If helper fails, fall back to quoting logic below
            }

            // Strip outer quotes if present
            if (t.Length >= 2 && t.StartsWith("\"") && t.EndsWith("\""))
            {
                t = t.Substring(1, t.Length - 2).Trim();
            }

            return $"\"{t}\"";
        }

        // Utility helper: safely read field
        private string GetFieldValue(FieldBinding field)
        {
            if (field == null) return string.Empty;

            // Step 1: Highest priority — RefValue
            var refVal = field.RefValue;
            if (!string.IsNullOrWhiteSpace(refVal))
            {
                var trimmedRefVal = refVal.Trim();

                // Resolve the cell reference through Excel (mirrors
                // GetResolvedAccountAssignmentValue's pattern below) instead of
                // returning the raw, unresolved cell-address string. Returning the
                // raw address (e.g. "'Sheet1'!$B$2") here meant every RefValue-driven
                // field could never match the hardcoded validation tokens in
                // ValidateJournalFields()/IsJournalValidationSatisfied(), so setting
                // Activity/BalanceType/CurrencyType via Reference (instead of the
                // ComboBox) always failed validation and left Journal Source/Category
                // disabled.
                if (ExcelRangeHelper.IsRealRange(trimmedRefVal))
                {
                    var resolvedVal = GetRangeValueSafe(trimmedRefVal) ?? string.Empty;
                    if (!string.IsNullOrWhiteSpace(resolvedVal))
                    {
                        return resolvedVal.Trim();
                    }
                }

                return trimmedRefVal;
            }

            if (field.Type == FieldBinding.FieldType.Ledger)
            {
                // Special handling for multi-select Ledgers
                var selectedTypes = SelectedLedgers
                    .Select(l => l.LedgerName)
                    .ToList();
                return string.Join(";", selectedTypes);
            }

            if (field.Type == FieldBinding.FieldType.Encumbrances)
            {
                // Special handling for multi-select Ledgers
                var selectedTypes = SelectedEncumbrances
                    .Select(e => e.EncumbranceType)
                    .ToList();
                return string.Join(";", selectedTypes);
            }

            if (field.Type == FieldBinding.FieldType.AccountAssignments)
            {
                var input = field.ComboValue?.ToString() ?? string.Empty;

                if (string.IsNullOrWhiteSpace(input))
                    return string.Empty;

                // escapeInnerQuotes=true if you need to double embedded quotes for Excel CSV-like contexts
                var output = NormalizeSemicolonSegments(input, escapeInnerQuotes: false);

                return output;
            }

            // Step 2: Handle ComboValue (can be String or model)
            var comboVal = field.ComboValue;
            if (comboVal == null) return string.Empty;

            // If it's a simple string
            if (comboVal is string str)
            {
                return str.Trim();
            }

            // Step 3: If it's a model — find a suitable display name
            var props = comboVal.GetType().GetProperties();
            var preferredOrder = new[] { AppConstants.PropShortName, "Name", "Description", "Desc", "PeriodName", "Code", "CurrencyCode", "SourceName", "CategoryName", "BudgetName" };

            foreach (var key in preferredOrder)
            {
                var prop = props.FirstOrDefault(p => string.Equals(p.Name, key, StringComparison.OrdinalIgnoreCase));
                if (prop != null)
                {
                    var val = prop.GetValue(comboVal);
                    if (val != null)
                    {
                        return val.ToString().Trim();
                    }
                }
            }

            // Step 4: Fallback
            return comboVal.ToString();
        }

        // Formula-building counterpart to GetFieldValue(). GetFieldValue() deliberately
        // resolves a field's RefValue through Excel to its CURRENT literal value (see that
        // method's own comment) - that resolve-first behavior is required for business-
        // logic/validation call sites (e.g. IsJournalValidationSatisfied matching
        // "PTD"/"Actual"/etc. against a Reference-bound Activity/BalanceType/CurrencyType
        // field), but it must NOT be used when constructing the actual Excel formula
        // string: baking in today's resolved value there would silently turn a live cell
        // reference into a hardcoded snapshot, so the GLSense formula would stop
        // recalculating if the referenced cell's value ever changed - defeating the entire
        // point of letting a field be bound via Reference instead of the ComboBox/manual
        // entry. Regression reported directly against this behavior: "when building any
        // formula, the top priority is for the references" - a Reference must always win
        // and be written into the formula AS a reference (FormatFormulaArg leaves it
        // unquoted because it contains "!"/"$"), never resolved to its value first.
        //
        // Used ONLY by the formula-argument builders below (BuildFormulaArguments,
        // GetFinalPeriodValue, GetBudgetEncumbranceValue, GetAccountSegments) - every other
        // GetFieldValue call site (validation, enable-state checks, IsBtJournalsType, etc.)
        // is untouched and keeps resolving RefValue, which is still correct there.
        private string GetFormulaFieldValue(FieldBinding field)
        {
            if (field == null) return string.Empty;

            var refVal = field.RefValue;
            if (!string.IsNullOrWhiteSpace(refVal))
            {
                // Reference wins outright for formula-building - return the raw,
                // unresolved cell address so FormatFormulaArg recognizes it and leaves it
                // unquoted as a live reference.
                return refVal.Trim();
            }

            // No reference set - RefValue is already blank, so GetFieldValue's own "Step 1"
            // RefValue branch is a no-op here; this reaches straight into its ComboValue/
            // multi-select/model-property handling, unchanged.
            return GetFieldValue(field);
        }

        private string GetResolvedAccountAssignmentValue()
        {
            var refValue = AccountAssignmentField?.RefValue;

            if (!string.IsNullOrWhiteSpace(refValue))
            {
                var trimmedRefValue = refValue?.Trim();

                if (ExcelRangeHelper.IsRealRange(trimmedRefValue))
                {
                    var resolvedValue = GetRangeValueSafe(trimmedRefValue ?? string.Empty);
                    if (!string.IsNullOrWhiteSpace(resolvedValue))
                    {
                        var trimmedResolvedValue = resolvedValue?.Trim();
                        if (!string.IsNullOrWhiteSpace(trimmedResolvedValue))
                            return trimmedResolvedValue ?? string.Empty;
                    }
                }

                return trimmedRefValue ?? string.Empty;
            }

            var comboValue = AccountAssignmentField?.ComboValue;
            return comboValue?.ToString()?.Trim() ?? string.Empty;
        }

        private static List<string> SplitAccountAssignmentSegments(string accountVal)
        {
            if (string.IsNullOrWhiteSpace(accountVal))
                return new List<string>();

            return accountVal
                .Split(new[] { ';' }, StringSplitOptions.None)
                .Select(segment => segment.Trim())
                .ToList();
        }

        private static string NormalizeSemicolonSegments(string input, bool escapeInnerQuotes = false)
        {
            input ??= string.Empty;

            return string.Join(";",
                input.Split(';')
                     .Select(s => NormalizeSegment(s, escapeInnerQuotes))
            );
        }

        private static string NormalizeSegment(string s, bool escapeInnerQuotes)
        {
            if (string.IsNullOrWhiteSpace(s))
                return "\"\"";

            s = s.Trim();

            // Strip one pair of OUTER quotes if present
            if (s.Length >= 2 && s[0] == '"' && s[s.Length - 1] == '"')
                s = s.Substring(1, s.Length - 2).Trim();

            if (string.IsNullOrWhiteSpace(s))
                return "\"\"";

            // If it's an Excel reference (cell/range/sheet/book), DO NOT QUOTE
            if (ExcelRangeHelper.IsRealRange(s))
                return s;

            // Optionally escape embedded quotes if you need CSV/Excel-safe tokens
            if (escapeInnerQuotes)
                s = s.Replace("\"", "\"\"");

            // Quote literal value exactly once
            return $"\"{s}\"";
        }

        // ---------------- Excel interop for refedit controls ----------------
        private Microsoft.Office.Interop.Excel.Application? _excelApp;
        public Microsoft.Office.Interop.Excel.Application? ExcelApp
        {
            get => _excelApp;
            set
            {
                _excelApp = value;
                OnPropertyChanged(nameof(ExcelApp));
            }
        }

        // === INotifyPropertyChanged ===
        public event PropertyChangedEventHandler? PropertyChanged;

        protected void OnPropertyChanged([CallerMemberName] string name = "")
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
        }
    }
#nullable disable
}