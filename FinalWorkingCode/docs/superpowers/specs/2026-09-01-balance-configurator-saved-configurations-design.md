# Balance Configurator — Saved Configurations — Design

**Date:** 2026-09-01
**Scope:** `GLSense/Views/GLBalanceConfigurator.xaml` / `.xaml.cs`, `GLSense/ViewModels/GLConfiguratorViewModel.cs`, new `GLSense/Common/BalanceConfigXmlStore.cs`, new `GLSense/Models/SavedBalanceConfig.cs`

## Problem

The Balance Configurator lets a user pick ~14 parameters (ledger, activity, balance type,
period, currency, currency type, actual flag, budget/encumbrance, journal source/category,
account-assignment segments, plus sign/factor/zeroes) and Insert writes a single
`=@GLSense_GetBalance(...)` formula into the selected cell
(`GLConfiguratorViewModel.WriteFormulaToCell`, line 3149). If the active cell already contains
a balance formula, the dialog re-populates every field from it
(`ApplyFormulaParamsAsync`); otherwise it falls back to default selections
(`ApplyDefaultSelections`).

There is no way to save a parameter combination under a name and reload it later. A user who
repeats the same combination across many cells/sheets/workbooks has to reselect everything
each time, or copy an existing formula cell and hand-edit it.

## Goals

1. Save the configurator's current field selections under a user-given name, scoped to the
   currently selected cube (`CubeId` + `CubeName`).
2. List, load, update (overwrite in place), and delete saved configurations for the current
   cube, from inside the Balance Configurator dialog itself.
3. Persist saved configurations inside the **workbook's Custom XML Parts** — they travel with
   the workbook, not the user account or the machine.
4. Loading a saved configuration must trigger the exact same cascading field-dependency
   behavior (Journal Source/Category enable/disable, CTD End Period population, currency
   defaulting, etc.) as reading an existing formula from a cell — no parallel/shortcut field
   -setting logic.
5. The new UI must not grow the dialog's width or existing minimum height when unused.
6. No selection state survives a dialog close/reopen — the saved-configuration picker always
   starts unselected.

## Non-goals

- No rename-only action (separately changing a saved config's name without touching its
  values) — out of scope for this iteration; "Update" always means "re-save current field
  values under the existing name."
- No cross-workbook or cross-cube sharing/export of saved configurations.
- No automated test project (this repository has none — see the `webview2-navigation-
  resilience` spec's Non-goals for the same, already-established rationale). Verified by
  building + exercising the real dialog + reading NLog output, matching this codebase's
  existing convention.
- No change to `WriteFormulaToCell`/`BuildFormulaArguments`/Insert itself — saving, loading,
  updating, and deleting a configuration are all independent of Insert. Saving never writes to
  the worksheet; Insert never implicitly saves.

## Architecture

### `GLSense/Models/SavedBalanceConfig.cs` (new)

One plain DTO per saved configuration. Mirrors `FieldBinding`'s own Combo-vs-Reference duality
per field — a "structured, self-documenting" shape (chosen over storing a raw formula-argument
list) so the raw XML payload is directly inspectable field-by-field:

```csharp
public class SavedBalanceConfig
{
    public string ConfigName { get; set; }

    public string LedgerCombo { get; set; }        // semicolon-joined ledger names, or null
    public string LedgerRef { get; set; }
    public string ActivityCombo { get; set; }
    public string ActivityRef { get; set; }
    public string BalanceTypeCombo { get; set; }
    public string BalanceTypeRef { get; set; }
    public string PeriodCombo { get; set; }
    public string PeriodRef { get; set; }
    public string EndPeriodCombo { get; set; }      // CTD only
    public string EndPeriodRef { get; set; }
    public DateTime? StartDateSelected { get; set; } // JED/JEDP/JEDU only
    public DateTime? EndDateSelected { get; set; }
    public string CurrencyCombo { get; set; }
    public string CurrencyRef { get; set; }
    public string CurrencyTypeCombo { get; set; }
    public string CurrencyTypeRef { get; set; }
    public string ActualFlagCombo { get; set; }
    public string ActualFlagRef { get; set; }
    public string BudgetCombo { get; set; }
    public string BudgetRef { get; set; }
    public string EncumbranceCombo { get; set; }    // semicolon-joined
    public string EncumbranceRef { get; set; }
    public string JournalSourceCombo { get; set; }
    public string JournalSourceRef { get; set; }
    public string JournalCategoryCombo { get; set; }
    public string JournalCategoryRef { get; set; }
    public string AccountAssignmentCombo { get; set; } // delimited per-segment literal string
    public string AccountAssignmentRef { get; set; }   // single Excel range reference

    public bool IsSignChecked { get; set; }
    public string FactorText { get; set; }
    public bool IsZeroesChecked { get; set; }
}
```

Per field: exactly one of `XxxCombo`/`XxxRef` is populated, never both — mirroring
`GetFormulaFieldValue`'s existing "Reference always wins" rule (`GLConfiguratorViewModel.cs:
3528-3545`). Capturing a field at Save time is a direct copy of `FieldBinding.RefValue` (if
non-blank) or the field's resolved combo text/model key — no new precedence logic, just reading
the same two properties `GetFormulaFieldValue` already reads.

### `GLSense/Common/BalanceConfigXmlStore.cs` (new)

CubeId-keyed `Workbook.CustomXMLParts` store, one part per cube holding the **entire list** of
that cube's saved configurations as JSON. Deliberately modeled on the existing
`Common/DrilldownMetadataXmlStore.cs` (delete-then-recreate on every write, 1-based backward
iteration, cheap substring check before `XDocument.Parse`) — same idiom, new root marker so
the two stores' lookups can never collide with each other or with the older
`DDDatatoWorksheet.cs` sheet-name-keyed mechanism:

```csharp
public static class BalanceConfigXmlStore
{
    private const string RootElementName = "BALANCECONFIGSAVE";
    private const string CubeIdElementName = "CUBEID";
    private const string CubeNameElementName = "CUBENAME";
    private const string PayloadElementName = "PAYLOAD"; // JSON array of SavedBalanceConfig

    public static void Save(Excel.Workbook wb, long cubeId, string cubeName,
                             List<SavedBalanceConfig> configs);
    public static bool TryRead(Excel.Workbook wb, long cubeId,
                                out List<SavedBalanceConfig> configs);
}
```

`Save` always replaces the whole part (same as `DrilldownMetadataXmlStore.Save`) — every
add/update/delete of a single saved configuration is: `TryRead` the current list → mutate it in
memory (add/replace/remove one entry by `ConfigName`) → `Save` the mutated list back. This is
safe at the realistic scale here (a handful to a few dozen manually-saved entries per cube, not
thousands) and keeps the store itself trivially simple — no per-entry XML nodes to manage.

### `GLConfiguratorViewModel.cs` additions

- `ObservableCollection<SavedBalanceConfig> SavedConfigurations` — populated from
  `BalanceConfigXmlStore.TryRead` for the current `CubeId` whenever the dialog loads (see
  Mechanics).
- `SavedBalanceConfig SelectedSavedConfig` — bound to the new combo box; starts `null` every
  time the dialog opens (see Goal 6).
- `SaveNewConfigurationAsync(string name)`, `UpdateSelectedConfigurationAsync()`,
  `DeleteSelectedConfigurationAsync()`, `LoadSavedConfigurationAsync(SavedBalanceConfig config)`.
- `ApplyFormulaParamsAsync` changes from `private` to `internal` (same precedent as making
  `GLSenseExcelFunctions.XLLContainer.ParsePeriodDate` `internal` for the Discover "Periods By
  Date" work) so `LoadSavedConfigurationAsync` can call it directly.

### `GLBalanceConfigurator.xaml` — UI placement

A new `Expander`, styled like the existing `BalanceParametersExpander` (`GLBalanceConfigurator.
xaml:786-792`, `BlueCircleExpanderStyle`), inserted as a new `Grid.Row` right below the header
row, **`IsExpanded="False"` by default** — costs zero height until the user opens it, and its
content (one combo box + three compact icon buttons, plus a toggle-visibility inline textbox for
naming) never widens past the dialog's existing 600px minimum.

## Mechanics

### Save flow

1. User expands "Saved Configurations", clicks **Save New**.
2. An inline textbox + Confirm/Cancel appears next to the buttons (no new window).
3. On Confirm: validate the name is non-blank and not already present in
   `SavedConfigurations` for this cube (case-insensitive) — if it collides, show a validation
   message and keep the textbox open; this codebase's existing `GLSenseMessage` helper covers
   the message box.
4. Build a `SavedBalanceConfig` by copying `RefValue`/resolved-combo-text from each
   `FieldBinding` (`LedgerField`, `ActivityField`, `BalanceTypeField`, `PeriodField`,
   `EndPeriodField`, `StartDateSelected`/`EndDateSelected`, `CurrencyField`,
   `CurrencyTypeField`, `ActualFlagField`, `BudgetField`/`EncumbranceField`,
   `JournalSourceField`, `JournalCategoryField`, `AccountAssignmentField` — the last one is a
   single `FieldBinding` whose `ComboValue` is itself a delimited multi-segment literal string,
   split/joined by the existing `SplitAccountAssignmentSegments`/`GetAccountSegments`, so it
   round-trips through `AccountAssignmentCombo`/`AccountAssignmentRef` exactly like every other
   field, no special-casing needed), plus `IsSignChecked`, `FactorText`, `IsZeroesChecked`.
5. Append to `SavedConfigurations`, call `BalanceConfigXmlStore.Save(workbook, cubeId,
   cubeName, SavedConfigurations.ToList())`.

### Update flow

Same capture step as Save (step 4 above), but replaces the entry matching
`SelectedSavedConfig.ConfigName` in place instead of appending, then persists the whole list.
Disabled while nothing is selected in the combo.

### Delete flow

Confirm prompt (`GLSenseMessage`, Yes/No — same pattern used elsewhere for destructive ribbon
actions), then removes the matching entry from `SavedConfigurations` and persists. Disabled
while nothing is selected.

### Load flow — the part that has to match "reading an existing formula" exactly

This is the piece Goal 4 is about, and it reuses the existing formula-reading path rather than
introducing a second way to populate the fields:

1. Build **synthetic `FuncArgs`/`FuncValues`** lists, at the exact same positional indices
   `BuildFormulaArguments()` already uses (0=sign+factor, 1=ledger, 2=activity,
   3=period[`~`endPeriod or startDate`~`endDate], 4=balanceType, 5=currency, 6=currencyType,
   7=actualFlag, 8=budget/encumbrance, 9=journalSource, 10=journalCategory, 11+=account
   segments):
   - For a field saved with `XxxRef` set: `FuncArgs[i] = XxxRef` (the stored reference text,
     unchanged); `FuncValues[i]` = that reference's **current** resolved value, read fresh from
     the active workbook (`ExcelApp.Range[refText].Value2`) — never a frozen value from save
     time, since the whole point of preserving the reference (per your call on the earlier
     question) is that it stays live.
   - For a field saved with `XxxCombo` set: `FuncArgs[i] = FuncValues[i] = XxxCombo` (a plain
     literal — `ExcelRangeHelper.IsRealRange` on it returns false, so the existing `Process*`
     methods naturally treat it as a ComboValue-mode field, exactly like today's "read an
     existing formula containing a literal argument" path).
   - Period argument is reassembled from `PeriodCombo`/`PeriodRef` (+`EndPeriodCombo`/`Ref` for
     CTD, or `StartDateSelected`/`EndDateSelected` for JED types) using the same `~`-joining
     `GetFinalPeriodValue()` already does for Insert.
2. Call the **existing, unmodified** `ApplyFormulaParamsAsync(config.IsZeroesChecked,
   funcArgs, funcValues)`. Every `Process*` helper it already calls (`ProcessLedgerFieldAsync`,
   `ProcessBalanceTypeAndPeriod`, `ProcessActualFlagAndBudgetEncumbrance`, `ProcessJls`,
   `ProcessAccountAssignments`, etc.) runs in the same order it always does, so every
   downstream `INotifyPropertyChanged`-driven cascade (Journal Source/Category enablement,
   CTD End Period population, ledger-change side effects) fires exactly as if the user had
   selected a cell already containing that formula. Zero new field-application logic — the
   only new code is step 1's list-building, which is the mirror image of what
   `CommonFunctions.FormulaParameters`/`FormulaValues` already produce when parsing a real
   formula string.

### Dialog open/close reset (Goal 6)

`SelectedSavedConfig` is set back to `null` (and `SavedConfigurations` re-populated fresh from
`BalanceConfigXmlStore.TryRead`) at the same point `ReLoadConfigurator()`/`LoadDataAsync`
already reset everything else on open (`GLBalanceConfigurator.xaml.cs:332`,
`GLConfiguratorViewModel.cs:814-906`) — no new lifecycle hook needed, this is strictly "one
more thing that gets reset where everything else already does."

## Error handling & logging

- `BalanceConfigXmlStore.Save`/`TryRead` wrap all Custom XML Part / JSON (de)serialization in
  `try/catch` + `LogUtility.LogException`, matching `DrilldownMetadataXmlStore` exactly — a
  corrupt or missing part degrades to "no saved configurations for this cube" rather than
  throwing into the UI.
- If a saved reference no longer resolves at Load time (deleted cell, wrong sheet, workbook
  reused elsewhere), only **that field** falls back to blank/default — logged via
  `LogUtility.LogDebug`, the rest of the load proceeds normally. This matches the existing
  graceful-degradation posture of `ApplyFormulaParamsAsync`'s own `Process*` methods, which
  already tolerate individual unresolvable arguments today.
- Duplicate-name validation on Save is a user-facing `GLSenseMessage`, not a silent failure.

## Verification

Build the project, then exercise directly in Excel (no automated test project in this repo,
per Non-goals):

1. Save a configuration with only literal (ComboValue) selections; close and reopen the
   dialog; confirm the combo starts unselected; select the saved config; confirm every field
   populates identically to selecting a cell that already has the equivalent formula.
2. Save a configuration where at least one field (e.g. Ledger) is Reference-bound; change that
   referenced cell's value; reload the saved configuration; confirm the field reflects the
   *current* cell value, not the value at save time.
3. Update an existing saved configuration with different field values; confirm the name is
   unchanged and the stored values are replaced.
4. Delete a saved configuration; confirm the confirmation prompt appears and the entry is
   gone from both the combo and the workbook's Custom XML Parts (inspect via the same
   technique used to verify `DrilldownMetadataXmlStore` previously — save, close/reopen the
   workbook, confirm persistence survives a full save/reopen cycle).
5. Attempt to save a second configuration under a name already used for the same cube; confirm
   it's blocked with a message, not silently overwritten.
6. Confirm a saved configuration for Cube A does not appear in the list when Cube B is
   selected.
7. Confirm the Expander's collapsed state doesn't change the dialog's width or its collapsed
   minimum height versus the current shipped version.

## Open risks / follow-ups

- `Workbook.CustomXMLParts` has an undocumented-but-real practical size ceiling per part and
  per workbook; not a concern at "a few dozen saved configs," worth revisiting only if usage
  patterns turn out to need hundreds of entries per cube.
- No rename-only action (see Non-goals) — if requested later, it's a small addition (update
  just `ConfigName` on the matching entry without recapturing field values).
- No expand-to-full-list-with-per-row-actions UI (the earlier-considered Option B) — if a
  cube's list grows large enough that a single combo box becomes unwieldy, revisit that layout;
  nothing in the storage design (`BalanceConfigXmlStore`, `SavedBalanceConfig`) would need to
  change to support a richer list UI later, only `GLBalanceConfigurator.xaml`.
