// FieldBinding.cs in GLSense.Addin.Core
// Group H (Balance Configurator) - port of GLSense\Bindings\FieldBinding.cs
// (FinalWorkingCode). Backs each of GLConfiguratorViewModel's 15 configurator rows
// (Ledger/Activity/BalanceType/Period/EndPeriod/StartDate/EndDate/Currency/CurrencyType/
// ActualFlag/Budgets/Encumbrances/JournalSources/JournalCategories/AccountAssignments) -
// each row is one FieldBinding instance exposing ComboText/ComboValue/RefValue plus
// computed IsRefEnabled/IsComboEnabled that switch on FieldType and read back into the
// owning GLConfiguratorViewModel for cross-field enable/disable rules.
//
// Placed in a dedicated Bindings\ folder (matching the old monolith's convention) rather
// than alongside GLConfiguratorViewModel in ViewModels\, even though this project already
// has a Bindings\ folder (ComboFieldBindings.cs) for a different, incompatible
// FieldBinding-style class used by the Group C/D "Get Period"/segment-function windows
// (different FieldType enum values, different IFieldDependencyProvider-based design).
// This class is tightly coupled to GLConfiguratorViewModel specifically (its
// IsRefEnabled/IsComboEnabled switch statements read GLConfiguratorViewModel members like
// IsLedgerEnabled/IsBudgetEnabled/BalanceTypeField/Ledgers/Encumbrances directly) and is
// not a generic reusable binding - same relationship as in the old monolith, where
// FieldBinding.OwnerViewModel is typed as GLConfiguratorViewModel, not an interface.
//
// This project has no NotifyBase equivalent (see Models\PeriodModels.cs header for why -
// NotifyBase drags in the old project's static LogUtility for debug-mode
// property-change logging) - FieldBinding implements INotifyPropertyChanged directly here,
// matching the convention already used throughout this project's Models/ViewModels.
// No other logic changes vs. the original.
using GLSense.Addin.Core.ViewModels;
using System;
using System.ComponentModel;
using System.Linq;

namespace GLSense.Addin.Core.Bindings
{
    public class FieldBinding : INotifyPropertyChanged
    {
        public GLConfiguratorViewModel OwnerViewModel { get; set; }

        // ==============================
        // ComboText
        // ==============================
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

        // ==============================
        // ComboValue
        // ==============================
        private object _comboValue;
        public object ComboValue
        {
            get => _comboValue;
            set
            {
                if (!Equals(_comboValue, value))
                {
                    _comboValue = value;
                    OnPropertyChanged(nameof(ComboValue));
                    OnPropertyChanged(nameof(IsComboEnabled));
                    OnPropertyChanged(nameof(IsRefEnabled));

                    OwnerViewModel?.OnFieldDependencyChanged(this);
                }
            }
        }
        // ==============================
        // RefValue
        // ==============================
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
                    OnPropertyChanged(nameof(IsRefEnabled));
                    OnPropertyChanged(nameof(IsComboEnabled));

                    OwnerViewModel?.OnRefEditTextChanged(this, value);
                }
            }
        }

        // ==============================
        // Enable / Disable Logic
        // ==============================
        public bool IsRefEnabled
        {
            get
            {
                var parent = OwnerViewModel;
                if (parent == null)
                    return false;

                return Type switch
                {
                    FieldType.Ledger => IsLedgerRefEnabled(parent),
                    FieldType.Budgets => parent.IsBudgetEnabled && HasEmptyComboOrRefValue(),
                    FieldType.Encumbrances => IsEncumbranceRefEnabled(parent),
                    FieldType.EndPeriod => parent.IsEndPeriodsEnabled && HasEmptyComboOrRefValue(),
                    FieldType.StartDate => IsJedType(parent) && HasNullComboOrRefValue(),
                    FieldType.EndDate => IsJedType(parent) && HasNullComboOrRefValue(),
                    _ => IsDefaultRefEnabled()
                };
            }
        }

        private bool IsLedgerRefEnabled(GLConfiguratorViewModel parent) =>
            parent.IsLedgerEnabled &&
            (!HasSelectedLedgers(parent) || HasRefValue());

        private bool IsEncumbranceRefEnabled(GLConfiguratorViewModel parent) =>
            parent.IsEncumbranceEnabled &&
            (!HasSelectedEncumbrances(parent) || HasRefValue());

        private static bool HasSelectedLedgers(GLConfiguratorViewModel parent) =>
            parent.Ledgers?.Any(l => l.IsSelected) ?? false;

        private static bool HasSelectedEncumbrances(GLConfiguratorViewModel parent) =>
            parent.Encumbrances?.Any(e => e.IsSelected) ?? false;

        private bool HasRefValue() =>
            !string.IsNullOrEmpty(RefValue);

        private bool HasEmptyComboOrRefValue() =>
            ComboValue == null || string.IsNullOrEmpty(ComboValue.ToString()) || HasRefValue();

        private bool HasNullComboOrRefValue() =>
            ComboValue == null || HasRefValue();

        private bool IsDefaultRefEnabled() =>
            (IsComboEnabled && HasNullComboOrRefValue()) || HasRefValue();

        private static bool IsJedType(GLConfiguratorViewModel parent)
        {
            var balanceType = parent.BalanceTypeField?.ComboValue?.ToString();

            return !string.IsNullOrWhiteSpace(balanceType) &&
                   (balanceType.Equals("JED", StringComparison.OrdinalIgnoreCase) ||
                    balanceType.Equals("JEDP", StringComparison.OrdinalIgnoreCase) ||
                    balanceType.Equals("JEDU", StringComparison.OrdinalIgnoreCase));
        }

        public bool IsComboEnabled
        {
            get
            {
                GLConfiguratorViewModel parent = OwnerViewModel;
                if (parent == null) return false;

                var balanceTypeCombo = parent.BalanceTypeField?.ComboValue?.ToString();
                bool isJEDType = !string.IsNullOrWhiteSpace(balanceTypeCombo) && (
                    balanceTypeCombo.Equals("JED", System.StringComparison.OrdinalIgnoreCase) ||
                    balanceTypeCombo.Equals("JEDP", System.StringComparison.OrdinalIgnoreCase) ||
                    balanceTypeCombo.Equals("JEDU", System.StringComparison.OrdinalIgnoreCase)
                );

                return Type switch
                {
                    FieldType.Encumbrances => parent.IsEncumbranceEnabled && string.IsNullOrEmpty(RefValue),// Enable combo only if ActualFlag allows it and refedit empty
                    FieldType.Budgets => parent.IsBudgetEnabled && string.IsNullOrEmpty(RefValue),// Enable only when ActualFlag allows
                    FieldType.EndPeriod => parent.IsEndPeriodsEnabled && string.IsNullOrEmpty(RefValue),// Enable when BalanceType allows
                    // Period selection should be disabled for JED types (dates used instead)
                    FieldType.Period => !isJEDType && string.IsNullOrEmpty(RefValue),
                    FieldType.StartDate => isJEDType && string.IsNullOrEmpty(RefValue),
                    FieldType.EndDate => isJEDType && string.IsNullOrEmpty(RefValue),
                    FieldType.JournalSources => parent.IsJournalValidationSatisfied() && string.IsNullOrEmpty(RefValue),// Respect journal validation rules
                    FieldType.JournalCategories => parent.IsJournalValidationSatisfied() && string.IsNullOrEmpty(RefValue),// Respect journal validation rules
                    _ => string.IsNullOrEmpty(RefValue),
                };
            }
        }

        // ==============================
        // Utility
        // ==============================
        public void RefreshEnableState()
        {
            OnPropertyChanged(nameof(IsRefEnabled));
            OnPropertyChanged(nameof(IsComboEnabled));
        }

        // ==============================
        // Enum
        // ==============================
        public enum FieldType
        {
            Ledger,
            Activity,
            BalanceType,
            Period,
            EndPeriod,
            StartDate,
            EndDate,
            Currency,
            CurrencyType,
            ActualFlag,
            Budgets,
            Encumbrances,
            JournalSources,
            JournalCategories,
            AccountAssignments
        }

        public FieldType Type { get; set; }

        // === INotifyPropertyChanged ===
        public event PropertyChangedEventHandler PropertyChanged;

        protected void OnPropertyChanged(string propertyName)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
    }
}
