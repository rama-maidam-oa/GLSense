# Balance Configurator Saved Configurations Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Let a user save the Balance Configurator's current field selections under a name (scoped to the current cube), then list/load/update/delete those saved configurations from inside the dialog, persisted in the workbook's Custom XML Parts.

**Architecture:** A new `SavedBalanceConfig` DTO mirrors each `FieldBinding`'s Combo-vs-Reference duality. A new `BalanceConfigXmlStore` (modeled directly on the existing `Common/DrilldownMetadataXmlStore.cs`) persists the whole per-cube list as one Custom XML Part, JSON payload, delete-then-recreate on every write. `GLConfiguratorViewModel` gains Save/Update/Delete methods that capture the live `FieldBinding` state into the DTO, and a Load method that converts a DTO back into the same `FuncArgs`/`FuncValues` positional-list shape the existing (unmodified) `ApplyFormulaParamsAsync` already consumes when reading a real cell formula — so every cascading field-dependency behavior fires identically regardless of where the values came from. UI is a new collapsed-by-default `Expander` in `GLBalanceConfigurator.xaml`, matching the dialog's existing `BalanceParametersExpander` pattern.

**Tech Stack:** C# / .NET Framework 4.8, WPF, `Microsoft.Office.Interop.Excel` (Custom XML Parts), `System.Text.Json` (via the existing `GLSense.Helpers.JsonGlobals.Options`).

**Spec:** `docs/superpowers/specs/2026-09-01-balance-configurator-saved-configurations-design.md`

## Global Constraints

- No automated test project exists in this repository. Every task's verification step is "build with 0 errors" plus, for the final task, manual exercise in Excel + reading NLog output — matching this repo's established convention (see the spec's Non-goals and the `2026-08-10-webview2-navigation-resilience-design.md` spec for the same precedent).
- The new UI must not grow the dialog's width past its existing 600px minimum, and must add zero height while the new Expander is collapsed (its default state).
- `ApplyFormulaParamsAsync` and every `Process*` helper it calls are **never modified** by this plan — the whole point of the Load design is to feed them unchanged.
- Every saved-configuration field mirrors `FieldBinding`'s "Reference always wins over Combo" rule (`GetFormulaFieldValue`, `GLConfiguratorViewModel.cs:3528-3545`) — capture whichever of `RefValue`/resolved-combo-text is populated, never both.
- Namespace/using conventions: `GLSense.Models` for the new DTO, `GLSense.Common` for the new XML store (matching `DrilldownMetadataXmlStore`'s existing namespace).

---

## File Structure

- **Create** `GLSense/Models/SavedBalanceConfig.cs` — the DTO (Task 1).
- **Create** `GLSense/Common/BalanceConfigXmlStore.cs` — Custom XML Part persistence (Task 2).
- **Modify** `GLSense/ViewModels/GLConfiguratorViewModel.cs` — new properties + Save/Update/Delete/Load methods (Tasks 3-6).
- **Modify** `GLSense/Views/GLBalanceConfigurator.xaml` — new Expander UI (Task 7).
- **Modify** `GLSense/Views/GLBalanceConfigurator.xaml.cs` — click handlers wiring the new UI to the ViewModel (Task 7).

---

### Task 1: `SavedBalanceConfig` DTO

**Files:**
- Create: `GLSense/Models/SavedBalanceConfig.cs`

**Interfaces:**
- Produces: `GLSense.Models.SavedBalanceConfig` — a plain, parameterless-constructible class with public get/set properties, used by every later task.

- [ ] **Step 1: Create the DTO file**

```csharp
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
```

- [ ] **Step 2: Build to verify it compiles**

Run: `& "C:\Program Files\Microsoft Visual Studio\2022\Community\MSBuild\Current\Bin\MSBuild.exe" "D:\SQLLite_Test\GLSense\FinalWorkingCode\GLSense\GLSense.csproj" /t:Build /p:Configuration=Debug /p:RegisterForComInterop=false /nologo /v:minimal /clp:ErrorsOnly`
Expected: exit code 0, no output (matches this repo's established build-verification pattern for every prior fix this session).

- [ ] **Step 3: Commit**

```bash
git add GLSense/Models/SavedBalanceConfig.cs
git commit -m "Add SavedBalanceConfig DTO for Balance Configurator saved configurations"
```

---

### Task 2: `BalanceConfigXmlStore` — Custom XML Part persistence

**Files:**
- Create: `GLSense/Common/BalanceConfigXmlStore.cs`
- Reference (read-only, do not modify): `GLSense/Common/DrilldownMetadataXmlStore.cs` (the pattern this mirrors)

**Interfaces:**
- Consumes: `GLSense.Models.SavedBalanceConfig` (Task 1).
- Produces: `GLSense.Common.BalanceConfigXmlStore.Save(Excel.Workbook wb, long cubeId, string cubeName, List<SavedBalanceConfig> configs)` and `BalanceConfigXmlStore.TryRead(Excel.Workbook wb, long cubeId, out List<SavedBalanceConfig> configs)`, used by Tasks 3-6.

- [ ] **Step 1: Create the store file**

```csharp
using GLSense.Helpers;
using GLSense.Models;
using GLSense.Utilities;
using System;
using System.Collections.Generic;
using System.Text.Json;
using System.Xml.Linq;
using Excel = Microsoft.Office.Interop.Excel;

namespace GLSense.Common
{
#nullable enable
    /// <summary>
    /// CubeId-keyed CustomXMLPart storage for the full list of a cube's saved Balance
    /// Configurator configurations. Modeled directly on DrilldownMetadataXmlStore.cs
    /// (delete-then-recreate on every write, 1-based backward iteration, cheap substring
    /// check before XDocument.Parse) - same idiom, a distinct root marker so this store's
    /// lookups can never collide with DrilldownMetadataXmlStore's DRILLDOWNMETADATA parts
    /// or the older DDDatatoWorksheet.cs DRILLDOWNSHEET parts.
    ///
    /// The whole per-cube list lives in ONE part (not one part per saved configuration) -
    /// every Save/Update/Delete of a single entry is: TryRead the current list, mutate it
    /// in memory, Save the mutated list back. Safe at the realistic scale here (a handful
    /// to a few dozen manually-saved entries per cube).
    /// </summary>
    public static class BalanceConfigXmlStore
    {
        private const string RootElementName = "BALANCECONFIGSAVE";
        private const string CubeIdElementName = "CUBEID";
        private const string CubeNameElementName = "CUBENAME";
        private const string PayloadElementName = "PAYLOAD";

        /// <summary>
        /// Deletes any existing BALANCECONFIGSAVE part for this cube, then stores the
        /// given list (serialized as JSON) as-is in a fresh part.
        /// </summary>
        public static void Save(Excel.Workbook wb, long cubeId, string cubeName, List<SavedBalanceConfig> configs)
        {
            if (wb == null)
            {
                LogUtility.LogWarn("BalanceConfigXmlStore.Save: no active workbook, cannot save configurations.");
                return;
            }

            RemoveExisting(wb, cubeId, out _);

            try
            {
                string rawJson = JsonSerializer.Serialize(configs ?? new List<SavedBalanceConfig>(), JsonGlobals.Options);

                var root = new XElement(
                    RootElementName,
                    new XElement(CubeIdElementName, cubeId),
                    new XElement(CubeNameElementName, cubeName ?? string.Empty),
                    new XElement(PayloadElementName, new XCData(rawJson)));

                wb.CustomXMLParts.Add(new XDocument(root).ToString());
                LogUtility.LogDebug($"BalanceConfigXmlStore.Save: stored {configs?.Count ?? 0} saved configuration(s) for cubeId={cubeId}");
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex, "BalanceConfigXmlStore.Save");
            }
        }

        /// <summary>
        /// Reads back the saved-configuration list for this cube, if any. Returns an
        /// empty list (not null) via the out parameter when nothing has been saved yet,
        /// or when the stored payload is corrupt.
        /// </summary>
        public static bool TryRead(Excel.Workbook wb, long cubeId, out List<SavedBalanceConfig> configs)
        {
            configs = new List<SavedBalanceConfig>();

            if (wb?.CustomXMLParts == null || wb.CustomXMLParts.Count == 0)
                return false;

            try
            {
                var cxps = wb.CustomXMLParts;

                // CustomXMLParts collections are 1-based
                for (int i = cxps.Count; i >= 1; i--)
                {
                    var xml = cxps[i]?.XML;
                    if (!ContainsConfigForCube(xml, cubeId))
                        continue;

                    var doc = XDocument.Parse(xml);
                    string? rawJson = doc.Root?.Element(PayloadElementName)?.Value;
                    if (string.IsNullOrEmpty(rawJson))
                        return false;

                    var deserialized = JsonSerializer.Deserialize<List<SavedBalanceConfig>>(rawJson, JsonGlobals.Options);
                    configs = deserialized ?? new List<SavedBalanceConfig>();
                    return true;
                }
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex, "BalanceConfigXmlStore.TryRead");
                configs = new List<SavedBalanceConfig>();
            }

            return false;
        }

        private static bool RemoveExisting(Excel.Workbook wb, long cubeId, out bool deletedAny)
        {
            deletedAny = false;

            if (wb?.CustomXMLParts == null || wb.CustomXMLParts.Count == 0)
                return true;

            try
            {
                var cxps = wb.CustomXMLParts;

                for (int i = cxps.Count; i >= 1; i--)
                {
                    var part = cxps[i];
                    if (ContainsConfigForCube(part?.XML, cubeId))
                    {
                        part.Delete();
                        deletedAny = true;
                    }
                }

                return true;
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex, "BalanceConfigXmlStore.RemoveExisting");
                return false;
            }
        }

        private static bool ContainsConfigForCube(string? xml, long cubeId)
        {
            if (string.IsNullOrEmpty(xml) || xml.IndexOf(RootElementName, StringComparison.OrdinalIgnoreCase) < 0)
                return false;

            try
            {
                var doc = XDocument.Parse(xml);
                var cubeIdValue = doc.Root?.Element(CubeIdElementName)?.Value;
                return !string.IsNullOrEmpty(cubeIdValue)
                    && long.TryParse(cubeIdValue, out long parsedCubeId)
                    && parsedCubeId == cubeId;
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex, "BalanceConfigXmlStore.ContainsConfigForCube");
                return false;
            }
        }
    }
}
```

- [ ] **Step 2: Build to verify it compiles**

Run: `& "C:\Program Files\Microsoft Visual Studio\2022\Community\MSBuild\Current\Bin\MSBuild.exe" "D:\SQLLite_Test\GLSense\FinalWorkingCode\GLSense\GLSense.csproj" /t:Build /p:Configuration=Debug /p:RegisterForComInterop=false /nologo /v:minimal /clp:ErrorsOnly`
Expected: exit code 0, no output.

- [ ] **Step 3: Commit**

```bash
git add GLSense/Common/BalanceConfigXmlStore.cs
git commit -m "Add BalanceConfigXmlStore for per-cube saved-configuration persistence"
```

---

### Task 3: ViewModel — saved-config list, selection, and open/reset wiring

**Files:**
- Modify: `GLSense/ViewModels/GLConfiguratorViewModel.cs`
  - Add properties near the other `ObservableCollection`/`FieldBinding` declarations (around line 251-282).
  - Extend `LoadDataAsync` (`GLConfiguratorViewModel.cs:814-906`) to populate the list.
  - Extend `ResetUIState` (`GLConfiguratorViewModel.cs:807-812`) to reset the selection.

**Interfaces:**
- Consumes: `BalanceConfigXmlStore.TryRead` (Task 2), `ExcelApp` (existing property, `GLConfiguratorViewModel.cs:3622`), `AppState.Instance.SelectedCube` (existing, `CubeId`/`CubeName`).
- Produces: `ObservableCollection<SavedBalanceConfig> SavedConfigurations`, `SavedBalanceConfig? SelectedSavedConfig` — consumed by Tasks 4-7.

- [ ] **Step 1: Add the two new properties**

Add near the other collection/field declarations (after `AccountAssignmentField` at line 282):

```csharp
public ObservableCollection<SavedBalanceConfig> SavedConfigurations { get; set; } = new();

private SavedBalanceConfig? _selectedSavedConfig;
public SavedBalanceConfig? SelectedSavedConfig
{
    get => _selectedSavedConfig;
    set
    {
        _selectedSavedConfig = value;
        OnPropertyChanged(nameof(SelectedSavedConfig));
    }
}
```

Add `using GLSense.Models;` and `using GLSense.Common;` to the file's using list if not already present (check the top of `GLConfiguratorViewModel.cs` first — `PeriodModel`/other `GLSense.Models` types are already used in this file, so `GLSense.Models` is very likely already imported; add only `using GLSense.Common;` if it's missing).

- [ ] **Step 2: Populate the list on load, inside `LoadDataAsync`**

Insert right after the `cubeId`/`ledgerId`/`coaid` locals are captured (`GLConfiguratorViewModel.cs:835-837`), before the `FillResponsibilitiesAsync` refresh call:

```csharp
try
{
    if (BalanceConfigXmlStore.TryRead(ExcelApp?.ActiveWorkbook, cubeId, out var savedConfigs))
    {
        await _dispatcher.InvokeAsync(() =>
        {
            SavedConfigurations.Clear();
            foreach (var saved in savedConfigs)
                SavedConfigurations.Add(saved);
        });
    }
    else
    {
        await _dispatcher.InvokeAsync(() => SavedConfigurations.Clear());
    }
}
catch (Exception ex)
{
    LogUtility.LogException(ex, "GLConfiguratorViewModel.LoadDataAsync: failed to load saved configurations (non-fatal)");
}
```

- [ ] **Step 3: Reset the selection every time the dialog opens, inside `ResetUIState`**

```csharp
private void ResetUIState()
{
    IsSignChecked = false;
    IsZeroesChecked = true;
    FactorText = "1";
    SelectedSavedConfig = null;
}
```

- [ ] **Step 4: Build to verify it compiles**

Run: `& "C:\Program Files\Microsoft Visual Studio\2022\Community\MSBuild\Current\Bin\MSBuild.exe" "D:\SQLLite_Test\GLSense\FinalWorkingCode\GLSense\GLSense.csproj" /t:Build /p:Configuration=Debug /p:RegisterForComInterop=false /nologo /v:minimal /clp:ErrorsOnly`
Expected: exit code 0, no output.

- [ ] **Step 5: Commit**

```bash
git add GLSense/ViewModels/GLConfiguratorViewModel.cs
git commit -m "Load saved configurations per cube and reset selection on dialog open"
```

---

### Task 4: ViewModel — capture current fields, Save New, Update Selected

**Files:**
- Modify: `GLSense/ViewModels/GLConfiguratorViewModel.cs` — add new private/public methods near `BuildFormulaArguments` (line 3168) since they read the same field set.

**Interfaces:**
- Consumes: `GetFormulaFieldValue(FieldBinding)` (existing, `GLConfiguratorViewModel.cs:3528`), `ShowWarningAction` (existing `Action<string>?`, line 618), `SavedConfigurations`/`SelectedSavedConfig` (Task 3), `BalanceConfigXmlStore.Save` (Task 2).
- Produces: `Task<bool> SaveNewConfigurationAsync(string name)`, `Task UpdateSelectedConfigurationAsync()` — consumed by Task 7's click handlers.

- [ ] **Step 1: Add a per-field capture helper and the full-configuration builder**

```csharp
// Captures one field's current state the same way GetFormulaFieldValue already
// resolves it for formula-building: Reference wins outright when present.
private (string? Combo, string? Ref) CaptureField(FieldBinding field)
{
    if (field == null) return (null, null);

    if (!string.IsNullOrWhiteSpace(field.RefValue))
        return (null, field.RefValue);

    var comboText = GetFieldValue(field);
    return (string.IsNullOrWhiteSpace(comboText) ? null : comboText, null);
}

private SavedBalanceConfig CaptureCurrentAsSavedConfig(string configName)
{
    var (ledgerCombo, ledgerRef) = CaptureField(LedgerField);
    var (activityCombo, activityRef) = CaptureField(ActivityField);
    var (btCombo, btRef) = CaptureField(BalanceTypeField);
    var (periodCombo, periodRef) = CaptureField(PeriodField);
    var (endPeriodCombo, endPeriodRef) = CaptureField(EndPeriodField);
    var (startDateCombo, startDateRef) = CaptureField(StartDateField);
    var (endDateCombo, endDateRef) = CaptureField(EndDateField);
    var (currencyCombo, currencyRef) = CaptureField(CurrencyField);
    var (currencyTypeCombo, currencyTypeRef) = CaptureField(CurrencyTypeField);
    var (actualFlagCombo, actualFlagRef) = CaptureField(ActualFlagField);
    var (budgetCombo, budgetRef) = CaptureField(BudgetField);
    var (encumbranceCombo, encumbranceRef) = CaptureField(EncumbranceField);
    var (journalSourceCombo, journalSourceRef) = CaptureField(JournalSourceField);
    var (journalCategoryCombo, journalCategoryRef) = CaptureField(JournalCategoryField);
    var (accountCombo, accountRef) = CaptureField(AccountAssignmentField);

    return new SavedBalanceConfig
    {
        ConfigName = configName,
        LedgerCombo = ledgerCombo,
        LedgerRef = ledgerRef,
        ActivityCombo = activityCombo,
        ActivityRef = activityRef,
        BalanceTypeCombo = btCombo,
        BalanceTypeRef = btRef,
        PeriodCombo = periodCombo,
        PeriodRef = periodRef,
        EndPeriodCombo = endPeriodCombo,
        EndPeriodRef = endPeriodRef,
        StartDateCombo = startDateCombo,
        StartDateRef = startDateRef,
        EndDateCombo = endDateCombo,
        EndDateRef = endDateRef,
        CurrencyCombo = currencyCombo,
        CurrencyRef = currencyRef,
        CurrencyTypeCombo = currencyTypeCombo,
        CurrencyTypeRef = currencyTypeRef,
        ActualFlagCombo = actualFlagCombo,
        ActualFlagRef = actualFlagRef,
        BudgetCombo = budgetCombo,
        BudgetRef = budgetRef,
        EncumbranceCombo = encumbranceCombo,
        EncumbranceRef = encumbranceRef,
        JournalSourceCombo = journalSourceCombo,
        JournalSourceRef = journalSourceRef,
        JournalCategoryCombo = journalCategoryCombo,
        JournalCategoryRef = journalCategoryRef,
        AccountAssignmentCombo = accountCombo,
        AccountAssignmentRef = accountRef,
        IsSignChecked = IsSignChecked,
        FactorText = string.IsNullOrWhiteSpace(FactorText) ? "1" : FactorText,
        IsZeroesChecked = IsZeroesChecked
    };
}

private void PersistSavedConfigurations()
{
    var cube = AppState.Instance.SelectedCube;
    if (cube == null)
    {
        LogUtility.LogWarn("GLConfiguratorViewModel.PersistSavedConfigurations: no cube selected, cannot persist.");
        return;
    }

    BalanceConfigXmlStore.Save(ExcelApp?.ActiveWorkbook, cube.CubeId, cube.CubeName, new List<SavedBalanceConfig>(SavedConfigurations));
}

/// <summary>
/// Saves the current field selections as a new named configuration for the current
/// cube. Returns false (and raises ShowWarningAction) if the name is blank or already
/// used by another saved configuration for this cube.
/// </summary>
public Task<bool> SaveNewConfigurationAsync(string name)
{
    var trimmedName = (name ?? string.Empty).Trim();
    if (string.IsNullOrWhiteSpace(trimmedName))
    {
        ShowWarningAction?.Invoke("Please enter a name for the saved configuration.");
        return Task.FromResult(false);
    }

    bool nameExists = SavedConfigurations.Any(c =>
        string.Equals(c.ConfigName, trimmedName, StringComparison.OrdinalIgnoreCase));

    if (nameExists)
    {
        ShowWarningAction?.Invoke($"A saved configuration named \"{trimmedName}\" already exists for this cube. Choose a different name.");
        return Task.FromResult(false);
    }

    var newConfig = CaptureCurrentAsSavedConfig(trimmedName);
    SavedConfigurations.Add(newConfig);
    PersistSavedConfigurations();
    SelectedSavedConfig = newConfig;

    LogUtility.LogDebug($"GLConfiguratorViewModel.SaveNewConfigurationAsync: saved '{trimmedName}'.");
    return Task.FromResult(true);
}

/// <summary>
/// Overwrites SelectedSavedConfig's stored values with the current field selections.
/// Name is unchanged. No-op (with a warning) if nothing is selected.
/// </summary>
public Task UpdateSelectedConfigurationAsync()
{
    if (SelectedSavedConfig == null)
    {
        ShowWarningAction?.Invoke("Select a saved configuration to update first.");
        return Task.CompletedTask;
    }

    var index = SavedConfigurations.IndexOf(SelectedSavedConfig);
    if (index < 0)
    {
        LogUtility.LogWarn("GLConfiguratorViewModel.UpdateSelectedConfigurationAsync: selected configuration no longer present in the list.");
        return Task.CompletedTask;
    }

    var updated = CaptureCurrentAsSavedConfig(SelectedSavedConfig.ConfigName);
    SavedConfigurations[index] = updated;
    PersistSavedConfigurations();
    SelectedSavedConfig = updated;

    LogUtility.LogDebug($"GLConfiguratorViewModel.UpdateSelectedConfigurationAsync: updated '{updated.ConfigName}'.");
    return Task.CompletedTask;
}
```

- [ ] **Step 2: Build to verify it compiles**

Run: `& "C:\Program Files\Microsoft Visual Studio\2022\Community\MSBuild\Current\Bin\MSBuild.exe" "D:\SQLLite_Test\GLSense\FinalWorkingCode\GLSense\GLSense.csproj" /t:Build /p:Configuration=Debug /p:RegisterForComInterop=false /nologo /v:minimal /clp:ErrorsOnly`
Expected: exit code 0, no output.

- [ ] **Step 3: Commit**

```bash
git add GLSense/ViewModels/GLConfiguratorViewModel.cs
git commit -m "Add Save New and Update Selected for Balance Configurator saved configurations"
```

---

### Task 5: ViewModel — Delete Selected

**Files:**
- Modify: `GLSense/ViewModels/GLConfiguratorViewModel.cs` — one more method alongside Task 4's.

**Interfaces:**
- Consumes: `SavedConfigurations`/`SelectedSavedConfig` (Task 3), `PersistSavedConfigurations()` (Task 4).
- Produces: `Task DeleteSelectedConfigurationAsync()`. The confirmation prompt itself lives in the View (Task 7), not here — this method performs the delete unconditionally once called.

- [ ] **Step 1: Add the method**

```csharp
/// <summary>
/// Removes SelectedSavedConfig from the list and persists the change. Confirmation
/// (are you sure?) is the View's responsibility (GLBalanceConfigurator.xaml.cs) - this
/// method performs the delete unconditionally once invoked. No-op if nothing selected.
/// </summary>
public Task DeleteSelectedConfigurationAsync()
{
    if (SelectedSavedConfig == null)
    {
        ShowWarningAction?.Invoke("Select a saved configuration to delete first.");
        return Task.CompletedTask;
    }

    var removedName = SelectedSavedConfig.ConfigName;
    SavedConfigurations.Remove(SelectedSavedConfig);
    SelectedSavedConfig = null;
    PersistSavedConfigurations();

    LogUtility.LogDebug($"GLConfiguratorViewModel.DeleteSelectedConfigurationAsync: deleted '{removedName}'.");
    return Task.CompletedTask;
}
```

- [ ] **Step 2: Build to verify it compiles**

Run: `& "C:\Program Files\Microsoft Visual Studio\2022\Community\MSBuild\Current\Bin\MSBuild.exe" "D:\SQLLite_Test\GLSense\FinalWorkingCode\GLSense\GLSense.csproj" /t:Build /p:Configuration=Debug /p:RegisterForComInterop=false /nologo /v:minimal /clp:ErrorsOnly`
Expected: exit code 0, no output.

- [ ] **Step 3: Commit**

```bash
git add GLSense/ViewModels/GLConfiguratorViewModel.cs
git commit -m "Add Delete Selected for Balance Configurator saved configurations"
```

---

### Task 6: ViewModel — Load a saved configuration through the existing formula-reading path

**Files:**
- Modify: `GLSense/ViewModels/GLConfiguratorViewModel.cs` — the last new method, calling the existing `ApplyFormulaParamsAsync` unchanged.

**Interfaces:**
- Consumes: `ApplyFormulaParamsAsync(bool, List<string>, List<string>)` (existing, unmodified, `GLConfiguratorViewModel.cs:1088`), `GetRangeValueSafe(string)` (existing, `GLConfiguratorViewModel.cs:2433`), `Budget`/`Encumbrance`/`AE` constants (existing, `GLConfiguratorViewModel.cs:30-32`), `AppConstants.BalanceTypeJED`/`JEDP`/`JEDU`/`CTD`/`ActualEncumbranceShort`.
- Produces: `Task LoadSavedConfigurationAsync(SavedBalanceConfig config)` — consumed by Task 7's combo-box selection handler.

- [ ] **Step 1: Add the field-resolution helper and the FuncArgs/FuncValues builder**

```csharp
// Mirrors ProcessLedgerFieldAsync/ProcessFieldAsync's own arg-vs-value split: when a
// reference was saved, the FORMULA-ARGUMENT slot is the raw reference text (so
// downstream Process* methods recognize it as live via ExcelRangeHelper.IsRealRange),
// and the VALUE slot is that reference's CURRENT resolved value - re-read from Excel
// now, never a frozen value from save time, since the whole point of preserving a
// reference is that it stays live. When a literal was saved, both slots are the same
// plain text (IsRealRange on it is false, so Process* naturally treats it as
// ComboValue-mode - exactly like today's "read an existing formula with a literal
// argument" path).
private (string Arg, string Val) ResolveSavedField(string? combo, string? refValue)
{
    if (!string.IsNullOrWhiteSpace(refValue))
    {
        var resolved = GetRangeValueSafe(refValue) ?? string.Empty;
        return (refValue, resolved);
    }

    var literal = combo ?? string.Empty;
    return (literal, literal);
}

/// <summary>
/// Builds the same positional FuncArgs/FuncValues shape CommonFunctions.FormulaParameters/
/// FormulaValues would produce from a real =@GLSense_GetBalance(...) formula, from a
/// saved configuration instead of formula text. Indices match BuildFormulaArguments()
/// exactly: 0=sign+factor, 1=ledger, 2=activity, 3=period (or JED start~end, or CTD
/// period~endPeriod), 4=balanceType, 5=currency, 6=currencyType, 7=actualFlag,
/// 8=budget/encumbrance, 9=journalSource, 10=journalCategory, 11+=account segments.
/// </summary>
private (List<string> FuncArgs, List<string> FuncValues) BuildFuncArgsFromSavedConfig(SavedBalanceConfig config)
{
    var signVal = (config.IsSignChecked ? "-" : "+") + (string.IsNullOrWhiteSpace(config.FactorText) ? "1" : config.FactorText);
    var (ledgerArg, ledgerVal) = ResolveSavedField(config.LedgerCombo, config.LedgerRef);
    var (activityArg, activityVal) = ResolveSavedField(config.ActivityCombo, config.ActivityRef);
    var (btArg, btVal) = ResolveSavedField(config.BalanceTypeCombo, config.BalanceTypeRef);
    var (currencyArg, currencyVal) = ResolveSavedField(config.CurrencyCombo, config.CurrencyRef);
    var (currencyTypeArg, currencyTypeVal) = ResolveSavedField(config.CurrencyTypeCombo, config.CurrencyTypeRef);
    var (actualFlagArg, actualFlagVal) = ResolveSavedField(config.ActualFlagCombo, config.ActualFlagRef);
    var (journalSourceArg, journalSourceVal) = ResolveSavedField(config.JournalSourceCombo, config.JournalSourceRef);
    var (journalCategoryArg, journalCategoryVal) = ResolveSavedField(config.JournalCategoryCombo, config.JournalCategoryRef);

    // Period (index 3) - branches on the (possibly reference-resolved) balance type text,
    // same three-way split GetFinalPeriodValue() already uses for Insert.
    string btText = (btVal ?? string.Empty).Trim();
    string periodArg, periodVal;

    if (btText.Equals(AppConstants.BalanceTypeJED, StringComparison.OrdinalIgnoreCase) ||
        btText.Equals(AppConstants.BalanceTypeJEDP, StringComparison.OrdinalIgnoreCase) ||
        btText.Equals(AppConstants.BalanceTypeJEDU, StringComparison.OrdinalIgnoreCase))
    {
        var (startArg, startVal) = ResolveSavedField(config.StartDateCombo, config.StartDateRef);
        var (endArg, endVal) = ResolveSavedField(config.EndDateCombo, config.EndDateRef);
        periodArg = $"{startArg}~{endArg}";
        periodVal = $"{startVal}~{endVal}";
    }
    else if (btText.Equals(AppConstants.BalanceTypeCTD, StringComparison.OrdinalIgnoreCase))
    {
        var (pArg, pVal) = ResolveSavedField(config.PeriodCombo, config.PeriodRef);
        var (epArg, epVal) = ResolveSavedField(config.EndPeriodCombo, config.EndPeriodRef);
        periodArg = $"{pArg}~{epArg}";
        periodVal = $"{pVal}~{epVal}";
    }
    else
    {
        (periodArg, periodVal) = ResolveSavedField(config.PeriodCombo, config.PeriodRef);
    }

    // Budget/Encumbrance (index 8) - branches on the resolved actual-flag text, same
    // switch GetBudgetEncumbranceValue() already uses for Insert.
    string afText = (actualFlagVal ?? string.Empty).Trim();
    string budEncumArg, budEncumVal;

    if (afText.Equals(Budget, StringComparison.OrdinalIgnoreCase) || afText == "B")
    {
        (budEncumArg, budEncumVal) = ResolveSavedField(config.BudgetCombo, config.BudgetRef);
    }
    else if (afText.Equals(Encumbrance, StringComparison.OrdinalIgnoreCase) || afText == "E" ||
             afText.Equals(AE, StringComparison.OrdinalIgnoreCase) ||
             afText.Equals(AppConstants.ActualEncumbranceShort, StringComparison.OrdinalIgnoreCase))
    {
        (budEncumArg, budEncumVal) = ResolveSavedField(config.EncumbranceCombo, config.EncumbranceRef);
    }
    else
    {
        budEncumArg = string.Empty;
        budEncumVal = string.Empty;
    }

    var funcArgs = new List<string>
    {
        signVal, ledgerArg, activityArg, periodArg, btArg,
        currencyArg, currencyTypeArg, actualFlagArg, budEncumArg,
        journalSourceArg, journalCategoryArg
    };
    var funcValues = new List<string>
    {
        signVal, ledgerVal, activityVal, periodVal, btVal,
        currencyVal, currencyTypeVal, actualFlagVal, budEncumVal,
        journalSourceVal, journalCategoryVal
    };

    // Account assignment (index 11+) - mirrors ProcessAccountAssignments' own two shapes:
    // a single range (index 11 only) when a reference was saved, or one plain literal
    // per COA segment starting at index 11 when a delimited combo string was saved.
    if (!string.IsNullOrWhiteSpace(config.AccountAssignmentRef))
    {
        var resolved = GetRangeValueSafe(config.AccountAssignmentRef) ?? string.Empty;
        funcArgs.Add(config.AccountAssignmentRef);
        funcValues.Add(resolved);
    }
    else
    {
        var segments = (config.AccountAssignmentCombo ?? string.Empty)
            .Split(new[] { ';' }, StringSplitOptions.None);

        foreach (var segment in segments)
        {
            var trimmed = segment.Trim();
            funcArgs.Add(trimmed);
            funcValues.Add(trimmed);
        }
    }

    return (funcArgs, funcValues);
}

/// <summary>
/// Loads a saved configuration into every field, via the exact same
/// ApplyFormulaParamsAsync path used when the active cell already contains a balance
/// formula - every cascading Process* call (Journal Source/Category enablement, CTD
/// End Period population, ledger-change side effects) fires identically, since this
/// is the same method, unchanged, just fed from a saved configuration instead of a
/// parsed formula string.
/// </summary>
public async Task LoadSavedConfigurationAsync(SavedBalanceConfig config)
{
    if (config == null)
    {
        LogUtility.LogWarn("GLConfiguratorViewModel.LoadSavedConfigurationAsync: config is null, aborting.");
        return;
    }

    var (funcArgs, funcValues) = BuildFuncArgsFromSavedConfig(config);
    LogUtility.LogDebug($"GLConfiguratorViewModel.LoadSavedConfigurationAsync: loading '{config.ConfigName}'.");
    await ApplyFormulaParamsAsync(config.IsZeroesChecked, funcArgs, funcValues);
}
```

- [ ] **Step 2: Build to verify it compiles**

Run: `& "C:\Program Files\Microsoft Visual Studio\2022\Community\MSBuild\Current\Bin\MSBuild.exe" "D:\SQLLite_Test\GLSense\FinalWorkingCode\GLSense\GLSense.csproj" /t:Build /p:Configuration=Debug /p:RegisterForComInterop=false /nologo /v:minimal /clp:ErrorsOnly`
Expected: exit code 0, no output.

- [ ] **Step 3: Commit**

```bash
git add GLSense/ViewModels/GLConfiguratorViewModel.cs
git commit -m "Load saved Balance Configurator configurations through the existing formula-reading path"
```

---

### Task 7: UI — collapsible Expander, wired to the ViewModel

**Files:**
- Modify: `GLSense/Views/GLBalanceConfigurator.xaml` — new `Grid.Row`, `Expander`, inserted between the header (`Grid.Row="0"`, ends line 138) and the Main Controls section (`Grid.Row="1"`, starts line 141). This shifts every existing `Grid.Row="N"` down by one and needs one more `RowDefinition`.
- Modify: `GLSense/Views/GLBalanceConfigurator.xaml.cs` — click handlers.

**Interfaces:**
- Consumes: `SaveNewConfigurationAsync`, `UpdateSelectedConfigurationAsync`, `DeleteSelectedConfigurationAsync`, `LoadSavedConfigurationAsync`, `SavedConfigurations`, `SelectedSavedConfig` (Tasks 3-6).

- [ ] **Step 1: Add a sixth `RowDefinition` and renumber existing rows**

In `GLBalanceConfigurator.xaml`, change (line 116-122):

```xml
<Grid.RowDefinitions>
    <RowDefinition Height="Auto"/>
    <RowDefinition Height="Auto"/>
    <RowDefinition Height="Auto"/>
    <RowDefinition Height="Auto"/>
    <RowDefinition Height="*"/>
</Grid.RowDefinitions>
```

to:

```xml
<Grid.RowDefinitions>
    <RowDefinition Height="Auto"/>
    <RowDefinition Height="Auto"/>
    <RowDefinition Height="Auto"/>
    <RowDefinition Height="Auto"/>
    <RowDefinition Height="Auto"/>
    <RowDefinition Height="*"/>
</Grid.RowDefinitions>
```

Then renumber every existing `Grid.Row="N"` in this same outer `Grid` (the Header at `Grid.Row="0"` stays 0; Main Controls moves from `Grid.Row="1"` to `Grid.Row="2"`; Options Section moves from `Grid.Row="2"` to `Grid.Row="3"`; Action Buttons moves from `Grid.Row="3"` to `Grid.Row="4"`; the Balance Parameters Expander moves from `Grid.Row="4"` to `Grid.Row="5"`).

- [ ] **Step 2: Insert the new Expander at `Grid.Row="1"`**

Insert immediately after the Header `Border` closes (after line 138, before the `<!-- ================= Main Controls Section ================= -->` comment):

```xml
<!-- ================= Saved Configurations (collapsed by default) ================= -->
<Expander Grid.Row="1"
          x:Name="SavedConfigurationsExpander"
          Style="{StaticResource BlueCircleExpanderStyle}"
          Header="Saved Configurations"
          IsExpanded="False"
          Background="White"
          BorderThickness="0"
          HorizontalAlignment="Stretch"
          Margin="0,0,0,8">
    <Grid Margin="12,8,12,8">
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>
        <Grid.ColumnDefinitions>
            <ColumnDefinition Width="*"/>
            <ColumnDefinition Width="Auto"/>
            <ColumnDefinition Width="Auto"/>
            <ColumnDefinition Width="Auto"/>
        </Grid.ColumnDefinitions>

        <ComboBox x:Name="CmbSavedConfigurations"
                  Grid.Column="0"
                  Style="{StaticResource CompactCombo}"
                  Margin="0,0,8,0"
                  ItemsSource="{Binding SavedConfigurations}"
                  DisplayMemberPath="ConfigName"
                  SelectedItem="{Binding SelectedSavedConfig, Mode=TwoWay}"
                  SelectionChanged="CmbSavedConfigurations_SelectionChanged"/>

        <Button x:Name="btnSaveNewConfig"
                Grid.Column="1"
                Content="Save New"
                Margin="0,0,4,0"
                Click="BtnSaveNewConfig_Click"/>

        <Button x:Name="btnUpdateConfig"
                Grid.Column="2"
                Content="Update"
                Margin="0,0,4,0"
                IsEnabled="False"
                Click="BtnUpdateConfig_Click"/>

        <Button x:Name="btnDeleteConfig"
                Grid.Column="3"
                Content="Delete"
                IsEnabled="False"
                Click="BtnDeleteConfig_Click"/>

        <StackPanel x:Name="SaveNamePanel"
                    Grid.Row="1"
                    Grid.Column="0"
                    Grid.ColumnSpan="4"
                    Orientation="Horizontal"
                    Margin="0,8,0,0"
                    Visibility="Collapsed">
            <TextBox x:Name="TxtNewConfigName"
                     Width="220"
                     Style="{StaticResource CompactTextBox}"
                     Margin="0,0,8,0"/>
            <Button Content="Confirm" Margin="0,0,4,0" Click="BtnConfirmSaveNewConfig_Click"/>
            <Button Content="Cancel" Click="BtnCancelSaveNewConfig_Click"/>
        </StackPanel>
    </Grid>
</Expander>
```

`BlueCircleExpanderStyle` (used identically by the existing `BalanceParametersExpander`, `GLBalanceConfigurator.xaml:788`), `CompactCombo` (used by every other combo in this file, e.g. line 157), and `CompactTextBox` (used by `TxtFactor`, line 742) are all existing resources already merged into this `UserControl`'s `Resources` — no new resource dictionary entries needed.

- [ ] **Step 3: Add the code-behind handlers**

In `GLBalanceConfigurator.xaml.cs`, add near `BtnCancelBottom_Click` (after line 631). This file already has `private readonly GLConfiguratorViewModel vm;` (line 45) and already wires `vm.ShowWarningAction` to `AppOverlayControl.ShowWarning` in the constructor (lines 78-79) — so `SaveNewConfigurationAsync`/`UpdateSelectedConfigurationAsync`/`DeleteSelectedConfigurationAsync`'s validation messages surface automatically through the existing overlay, no extra wiring needed here. `AppOverlayControl.ShowConfirmAsync(string) -> Task<bool?>` (an existing method on this same view's overlay control) is used for the destructive Delete confirmation, matching this view's own established confirm pattern (rather than reaching for `CommonFunctions.GLSenseMessage`, which other, unrelated windows use):

```csharp
private void CmbSavedConfigurations_SelectionChanged(object sender, SelectionChangedEventArgs e)
{
    bool hasSelection = vm.SelectedSavedConfig != null;
    btnUpdateConfig.IsEnabled = hasSelection;
    btnDeleteConfig.IsEnabled = hasSelection;

    if (!hasSelection)
        return;

    LogUtility.LogDebug($"GLBalanceConfigurator.CmbSavedConfigurations_SelectionChanged: loading '{vm.SelectedSavedConfig.ConfigName}'");
    _ = LoadSelectedSavedConfigurationAsync(vm.SelectedSavedConfig);
}

private async Task LoadSelectedSavedConfigurationAsync(SavedBalanceConfig config)
{
    try
    {
        await vm.LoadSavedConfigurationAsync(config);
    }
    catch (Exception ex)
    {
        LogUtility.LogException(ex, "GLBalanceConfigurator.LoadSelectedSavedConfigurationAsync");
    }
}

private void BtnSaveNewConfig_Click(object sender, RoutedEventArgs e)
{
    LogUtility.LogDebug("GLBalanceConfigurator.BtnSaveNewConfig_Click invoked");
    TxtNewConfigName.Text = string.Empty;
    SaveNamePanel.Visibility = Visibility.Visible;
    TxtNewConfigName.Focus();
}

private void BtnCancelSaveNewConfig_Click(object sender, RoutedEventArgs e)
{
    LogUtility.LogDebug("GLBalanceConfigurator.BtnCancelSaveNewConfig_Click invoked");
    SaveNamePanel.Visibility = Visibility.Collapsed;
}

private async void BtnConfirmSaveNewConfig_Click(object sender, RoutedEventArgs e)
{
    LogUtility.LogDebug($"GLBalanceConfigurator.BtnConfirmSaveNewConfig_Click invoked - name={TxtNewConfigName.Text}");
    try
    {
        bool saved = await vm.SaveNewConfigurationAsync(TxtNewConfigName.Text);
        if (saved)
            SaveNamePanel.Visibility = Visibility.Collapsed;
    }
    catch (Exception ex)
    {
        LogUtility.LogException(ex, "GLBalanceConfigurator.BtnConfirmSaveNewConfig_Click");
    }
}

private async void BtnUpdateConfig_Click(object sender, RoutedEventArgs e)
{
    LogUtility.LogDebug("GLBalanceConfigurator.BtnUpdateConfig_Click invoked");
    try
    {
        await vm.UpdateSelectedConfigurationAsync();
    }
    catch (Exception ex)
    {
        LogUtility.LogException(ex, "GLBalanceConfigurator.BtnUpdateConfig_Click");
    }
}

private async void BtnDeleteConfig_Click(object sender, RoutedEventArgs e)
{
    var configName = vm.SelectedSavedConfig?.ConfigName;
    LogUtility.LogDebug($"GLBalanceConfigurator.BtnDeleteConfig_Click invoked - config={configName}");
    if (configName == null)
        return;

    try
    {
        bool? confirmed = await AppOverlayControl.ShowConfirmAsync($"Delete saved configuration \"{configName}\"?");
        if (confirmed == true)
        {
            await vm.DeleteSelectedConfigurationAsync();
            btnUpdateConfig.IsEnabled = false;
            btnDeleteConfig.IsEnabled = false;
        }
    }
    catch (Exception ex)
    {
        LogUtility.LogException(ex, "GLBalanceConfigurator.BtnDeleteConfig_Click");
    }
}
```

Add `using GLSense.Models;` to `GLBalanceConfigurator.xaml.cs`'s using list if not already present.

- [ ] **Step 4: Build to verify it compiles**

Run: `& "C:\Program Files\Microsoft Visual Studio\2022\Community\MSBuild\Current\Bin\MSBuild.exe" "D:\SQLLite_Test\GLSense\FinalWorkingCode\GLSense\GLSense.csproj" /t:Build /p:Configuration=Debug /p:RegisterForComInterop=false /nologo /v:minimal /clp:ErrorsOnly`
Expected: exit code 0, no output.

- [ ] **Step 5: Launch the add-in in Excel and confirm the collapsed Expander doesn't change the dialog's size**

Open the Balance Configurator on a cell (with or without an existing balance formula). Confirm:
- The dialog's width and the collapsed-state height look identical to before this change (compare against a screenshot or your memory of the current shipped layout).
- Clicking the "Saved Configurations" header expands/collapses it, matching the existing "Balance Parameters" expander's behavior at the bottom of the same dialog.

- [ ] **Step 6: Commit**

```bash
git add GLSense/Views/GLBalanceConfigurator.xaml GLSense/Views/GLBalanceConfigurator.xaml.cs
git commit -m "Add Saved Configurations UI to Balance Configurator"
```

---

### Task 8: Full manual verification pass

**Files:** none (verification only — no automated test project exists in this repository).

**Interfaces:** none — this task exercises everything built in Tasks 1-7 together, in the real Excel add-in.

- [ ] **Step 1: Save a purely literal configuration and confirm round-trip fidelity**

In Excel, with the add-in loaded and logged in: open the Balance Configurator, select a ledger, activity, balance type (e.g. PTD), period, currency, currency type, actual flag, and an account-assignment segment — all via the ComboBoxes (no RefEdit references). Expand "Saved Configurations", click Save New, type a name (e.g. `Test-Literal`), click Confirm.

Close the dialog. Reopen it (on any cell). Confirm the "Saved Configurations" combo starts empty/unselected (Goal 6). Select `Test-Literal` from the combo. Confirm every field populates identically to what you selected before saving.

- [ ] **Step 2: Save a configuration with a live cell reference, then confirm it stays live**

In a fresh Balance Configurator session, set the Ledger field via its RefEdit control pointing at a cell containing a valid ledger name (instead of the ComboBox). Fill the rest of the fields via ComboBoxes as in Step 1. Save New as `Test-Reference`.

Change the referenced cell's value to a *different* valid ledger name. Reload `Test-Reference` from the combo. Confirm the Ledger field now reflects the cell's *current* value, not the value at save time — this is the "save the live reference itself" behavior confirmed earlier in the design discussion.

- [ ] **Step 3: Update an existing saved configuration**

With `Test-Literal` selected, change the Balance Type to CTD (which reveals the End Period field) and pick an End Period. Click Update. Confirm the button was enabled only because a configuration was selected. Close and reopen the dialog, reload `Test-Literal`, and confirm it now reflects CTD + the new End Period, still named `Test-Literal`.

- [ ] **Step 4: Delete a saved configuration and confirm persistence across a workbook save/reopen**

Select `Test-Reference`, click Delete. Confirm the confirm prompt appears (`AppOverlayControl.ShowConfirmAsync`) and, on confirming, the entry disappears from the combo. Save the workbook, close it, and reopen it. Open the Balance Configurator again — confirm `Test-Reference` is still gone and `Test-Literal` is still present (proves the delete was actually persisted to the workbook's Custom XML Parts, not just removed from the in-memory list).

- [ ] **Step 5: Duplicate-name validation**

With any configuration already saved (e.g. `Test-Literal`), try Save New with the exact same name again. Confirm a warning appears (via the overlay) and no second entry is created.

- [ ] **Step 6: Per-cube scoping**

Switch to a different cube (if more than one is available in your environment). Open the Balance Configurator and confirm the "Saved Configurations" combo is empty for this cube (does not show `Test-Literal` from the other cube). Switch back to the original cube and confirm `Test-Literal` reappears.

- [ ] **Step 7: Check the NLog output for this session**

Open the current day's log (`%LOCALAPPDATA%\ORBIT\Excel_Logs\GLSense_Logs\Logs\GLSense_Logs_<date>.log`, or `GLSense_Logs_New` depending on which logging path this environment is using — see this session's own earlier investigation of the log folder location). Confirm no `LogException`/`LogError` entries appeared during the above steps; only `LogDebug`/`LogWarn` lines from the new methods (`BalanceConfigXmlStore.Save/TryRead`, `SaveNewConfigurationAsync`, `LoadSavedConfigurationAsync`, etc.).

- [ ] **Step 8: Final commit (if Step 7 surfaced any fix)**

If verification surfaced a bug, fix it, rebuild, re-run the affected verification step(s), then:

```bash
git add -A
git commit -m "Fix issue found during Balance Configurator saved-configurations verification"
```

If nothing needed fixing, this task has no commit of its own — Tasks 1-7's commits already cover the complete, verified feature.

