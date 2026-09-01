using System;
using System.Collections.Generic;

namespace GLSense.Models
{
#nullable enable
    /// <summary>
    /// One saved Balance Configurator parameter set, scoped to a single cube. Mirrors
    /// FieldBinding's Combo-vs-Reference duality per field (see GetFormulaFieldValue,
    /// GLConfiguratorViewModel.cs:3528-3545) - for each pair below, exactly one of
    /// XxxCombo/XxxRef is populated, never both, matching "Reference always wins".
    /// Persisted as JSON inside a workbook Custom XML Part by BalanceConfigXmlStore.
    /// </summary>
    public class SavedBalanceConfig
    {
        public string ConfigName { get; set; } = string.Empty;

        public string? LedgerCombo { get; set; }        // semicolon-joined ledger names, or null
        public string? LedgerRef { get; set; }
        public string? ActivityCombo { get; set; }
        public string? ActivityRef { get; set; }
        public string? BalanceTypeCombo { get; set; }
        public string? BalanceTypeRef { get; set; }
        public string? PeriodCombo { get; set; }
        public string? PeriodRef { get; set; }
        public string? EndPeriodCombo { get; set; }     // CTD only
        public string? EndPeriodRef { get; set; }
        public string? StartDateCombo { get; set; }     // JED/JEDP/JEDU only (ISO date string)
        public string? StartDateRef { get; set; }
        public string? EndDateCombo { get; set; }
        public string? EndDateRef { get; set; }
        public string? CurrencyCombo { get; set; }
        public string? CurrencyRef { get; set; }
        public string? CurrencyTypeCombo { get; set; }
        public string? CurrencyTypeRef { get; set; }
        public string? ActualFlagCombo { get; set; }
        public string? ActualFlagRef { get; set; }
        public string? BudgetCombo { get; set; }
        public string? BudgetRef { get; set; }
        public string? EncumbranceCombo { get; set; }   // semicolon-joined
        public string? EncumbranceRef { get; set; }
        public string? JournalSourceCombo { get; set; }
        public string? JournalSourceRef { get; set; }
        public string? JournalCategoryCombo { get; set; }
        public string? JournalCategoryRef { get; set; }
        public string? AccountAssignmentCombo { get; set; } // delimited per-segment literal string
        public string? AccountAssignmentRef { get; set; }   // single Excel range reference

        public bool IsSignChecked { get; set; }
        public string FactorText { get; set; } = "1";
        public bool IsZeroesChecked { get; set; } = true;
    }
}
