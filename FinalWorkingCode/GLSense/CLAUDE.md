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
