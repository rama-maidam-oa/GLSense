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

## `Views\GLBalanceConfigurator.xaml.cs` / `ViewModels\GLConfiguratorViewModel.cs`

- **DatePicker min/max wrong for non-standard fiscal calendars**: `DatePicker_CalendarOpenedEx`
  computed the selectable start/end range as `Periods[0].StartDate` and
  `Periods[Periods.Count - 1].EndDate` - i.e. it trusted list position instead of the actual
  dates. For a standard calendar this happens to work, but some ledgers use custom period
  sets (e.g. a "GOV Calendar" fiscal year running JUL-DEC-then-JAN-JUN, or other calendars
  shifted to start in a different quarter/month) where the first/last element of `Periods`
  is not guaranteed to be the true earliest/latest date.
  Fixed by computing the range as `vm.Periods.Min(p => p.StartDate)` /
  `vm.Periods.Max(p => p.EndDate)` instead, so the DisplayDateStart/DisplayDateEnd bounds
  and blackout ranges are correct regardless of what order the repository returns periods in.
  This fix is correct but on its own did **not** resolve the reported symptom (see next item)
  - it's still kept since it's a valid defensive fix regardless of calendar shape.

- **Root cause of the still-broken calendar: PERIODS cache never refreshes after first load**.
  Reported symptom: a ledger ("Progress US Primary Ledger", GOV Calendar period set) whose
  fiscal calendar had since been extended by the source system to DEC-28 still only showed
  JUN-28 as the last selectable date, even after the Min/Max fix above and a clean
  rebuild/fresh Excel session. Traced via direct query of the live local SQLite cache
  (`GLSense.sqlite`, `PERIODS` table) - confirmed the cache genuinely had no rows past
  2028-06-30 for that cube/ledger. `GLConfiguratorViewModel.LoadDataAsync` reads periods via
  `DataRepository.GetPeriods(cubeId, ledgerId)`, a plain `SELECT` against this local cache
  with no refresh trigger anywhere in that code path - once a ledger's periods were cached
  once, ever, nothing ever re-synced them from the source system again, even as the source
  system's fiscal calendar grew.
  Fixed by having `LoadDataAsync` call `CommonFunctions.FillResponsibilitiesAsync(ledgerId,
  cubeId, token)` - the existing remote ledger-setup-data fetch - every time the Balance
  Configurator loads for a ledger, before reading from the cache. This is safe to call
  repeatedly: `LedgerDataRepository.InsertLedgerDataAsync` already does a proper
  DELETE-then-INSERT per cubeId/ledgerId for every affected table (PERIODS, ACTIVITY,
  CURRENCIES, BUDGETS, SEGMENTS, etc. - see its `ClearExistingData` step) inside a single
  transaction, so re-running it just replaces stale rows with fresh ones rather than
  duplicating or erroring. The refresh is awaited but fully async (network call + a
  `Task.Run`-wrapped SQLite write), so it doesn't block the UI thread; failures (e.g.
  offline) are caught and logged, and the configurator falls back to whatever was already
  cached rather than blocking the user.
  **Status: fixed in FinalWorkingCode only so far** - port both this fix and the Min/Max
  fix above to AIPowered's identical `GLSense.Addin.Core\Views\GLBalanceConfigurator.xaml.cs`
  / `GLConfiguratorViewModel.cs` once confirmed working here.
