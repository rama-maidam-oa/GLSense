# GLSense (FinalWorkingCode) - fix log

This codebase is the older monolith counterpart to `AIPowered\GLSense`. Most fix work
happens in AIPowered (see that project's `CLAUDE.md` for the full log and the reasons
behind each fix); the two items below were explicitly reported as bugs in **both**
codebases and mirrored here identically.

## `ViewModels\GLConfiguratorViewModel.cs`

- **Journal Source/Category always disabled**: `GetFieldValue()`'s `RefValue` branch
  returned the raw, unresolved cell-address string instead of resolving it through
  Excel, so setting Activity/BalanceType/CurrencyType via Reference (instead of the
  ComboBox) always failed `ValidateJournalFields()`'s token matching and left Journal
  Source/Category disabled. Fixed by resolving the ref through `GetRangeValueSafe()`
  first, mirroring the already-correct `GetResolvedAccountAssignmentValue()` pattern in
  the same file. (This file's `Converters\Converters.cs::JournalValidationConverter`
  already had the full token list - no change needed there, unlike AIPowered's copy.)

- **End Period not populating for CTD**: `OnFieldDependencyChanged`'s `BalanceType`
  case correctly enabled the End Period row for CTD but never called
  `UpdateEndPeriods()` to actually populate the `EndPeriods` collection. Added the
  missing call to the CTD branch.

See `AIPowered\GLSense\CLAUDE.md` sections 2.2/2.3 for the full write-up (exact
line-level reasoning, why the bug happened, what the "correct reference pattern"
looked like) - it applies here verbatim.

## `Views\GLBalanceConfigurator.xaml.cs`

- **DatePicker min/max wrong for non-standard fiscal calendars**: `DatePicker_CalendarOpenedEx`
  computed the selectable start/end range as `Periods[0].StartDate` and
  `Periods[Periods.Count - 1].EndDate` - i.e. it trusted list position instead of the actual
  dates. For a standard calendar this happens to work, but some ledgers use custom period
  sets (e.g. a "GOV Calendar" fiscal year running JUL-DEC-then-JAN-JUN, or other calendars
  shifted to start in a different quarter/month) where the first/last element of `Periods`
  is not guaranteed to be the true earliest/latest date. Reported symptom: a ledger whose
  periods actually extend to DEC-28 only showed JUN-28 as the last selectable date in the
  Start/End Date pickers, silently cutting off real periods from selection.
  Fixed by computing the range as `vm.Periods.Min(p => p.StartDate)` /
  `vm.Periods.Max(p => p.EndDate)` instead, so the DisplayDateStart/DisplayDateEnd bounds
  and blackout ranges are correct regardless of what order the repository returns periods in.
  **Status: fixed in FinalWorkingCode only so far** - port to
  `AIPowered\GLSense.Addin.Core\Views\GLBalanceConfigurator.xaml.cs`'s identical method once
  this is confirmed working here.
