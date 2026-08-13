# GLSense (FinalWorkingCode) - fix log

This codebase is the older monolith counterpart to `AIPowered\GLSense`. Most fix work
happens in AIPowered (see that project's `CLAUDE.md` for the full log and the reasons
behind each fix); the two items below were explicitly reported as bugs in **both**
codebases and mirrored here identically.

## `Utilities\DpiAwareWindow.cs`

- **Windows not centered on screen in the shipped MSI**: `WindowStartupLocation="CenterOwner"`
  (set per-window in XAML, or via `WindowHelper.SetExcelAsOwner`) only centers a window
  once, at the moment WPF applies it. Two methods in this base class can resize the window
  afterward without ever recalculating position: `FitToAvailableWorkArea()` (runs once from
  `OnLoaded`, can shrink Width/Height to fit the screen's work area based on measured content
  size) and `EnsureFitsWorkArea()` (runs on every `OnRenderSizeChanged` - e.g. a DataGrid
  populating with data after an async load, or a DPI change - clamps Width/Height against
  Min/Max bounds). Both only ever changed Width/Height, never Left/Top, so a resize always
  grew/shrank anchored at the window's current top-left corner - the window's true center
  silently drifted away from wherever `CenterOwner` originally centered it.
  Fixed by adding `RecenterAfterSizeChange(previousLeft, previousTop, previousWidth,
  previousHeight)`: both methods now capture Left/Top/Width/Height before making their
  change, and if they actually changed the size, recenter around the same center point
  afterward (clamped so it can't be pushed off the visible work area). This is scoped
  narrowly - it only fires when these two methods themselves changed the size, so a plain
  user drag-resize (`ResizeMode="CanResize"`, used by nearly every window here) is
  completely unaffected, since `EnsureFitsWorkArea` only reassigns Width/Height when a drag
  actually violates Min/MaxWidth/Height (in which case recentering after the forced clamp
  is correct anyway).
  **Status: confirmed working, ported to AIPowered.** See AIPowered's `CLAUDE.md` for the
  equivalent write-up in `BaseWindow.cs` (`CenterWindowInExcel()`, `FitToAvailableWorkArea()`,
  `ForceSizeToContentResettle()`), which had the identical resize-without-recenter shape.

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
  Verified via temporary `LogWarn` diagnostics (logged the cached PERIODS row count before
  and after the refresh call, regardless of ribbon Debug mode) - confirmed the refresh runs
  and completes without throwing, and the row count/date range is unchanged (176 rows,
  through 2028-06-30) before and after. That, plus a direct read of the live source dates,
  showed there was no remaining caching bug: the "-28" suffix in this ledger's period names
  is a fiscal-year label (FY28 = Jul 2027-Jun 2028), not a calendar year, so e.g. "DEC-28" is
  real calendar December 2027 (already well within the selectable range), and the genuinely
  missing months are real Jul-Dec 2028, which belong to the next fiscal year (FY29) and don't
  exist in the source system yet - a data/calendar-setup gap upstream, not an app bug. The
  diagnostic `LogWarn` calls have been removed now that this is confirmed.
  **Status: fixed in FinalWorkingCode only so far** - port both this fix and the Min/Max
  fix above to AIPowered's identical `GLSense.Addin.Core\Views\GLBalanceConfigurator.xaml.cs`
  / `GLConfiguratorViewModel.cs` once requested.

## `Utilities\SegmentDiscoverer.cs`

- **"COM object separated from its underlying RCW" on Explode All / Explode 1 Level**:
  `SegmentAction` caches the source worksheet in a class-level field (`HrWorksheet =
  CellActive.Worksheet`) before doing any async work. Many awaits later,
  `CreateSingleSheetAsync` reads that cached field (`HrWorksheet?.Index`, then
  `HrWorksheet?.Copy(...)`) - but just before that, it calls `SheetExists`/
  `GetWorksheetByName` to check whether a sheet with the target name already exists,
  and both of those helper methods called `ExcelComHelper.SafeRelease(ws, "Worksheet")`
  (`Marshal.FinalReleaseComObject`) on *every* worksheet enumerated that didn't match the
  target name - including the source sheet, since the new sheet's sanitized name is
  essentially never equal to the source sheet's own name. Since .NET's classic COM
  interop layer caches one RCW per underlying COM object (within the same execution
  context), the `ws` yielded for that sheet during enumeration is the *same* wrapper as
  the cached `HrWorksheet` field - so releasing it there detached `HrWorksheet`'s RCW
  too, and the very next lines in the same call (`HrWorksheet?.Index`,
  `HrWorksheet?.Copy(...)`) threw "COM object that has been separated from its
  underlying RCW cannot be used." This was fully deterministic on the very first child
  sheet of the very first click, not timing-dependent.
  Fixed by removing the `SafeRelease` calls from `SheetExists`/`GetWorksheetByName`
  entirely - both are lightweight name-lookup helpers with no long-lived object graph to
  worry about (unlike Application/Workbook/Worksheets, which is what actually keeps
  Excel.exe alive after close, per the section below on that topic), so they don't need
  to force-release anything, and doing so here was unsafe given `HrWorksheet`'s lifetime.
  Verified these two methods are called from nowhere else in the file, so this doesn't
  affect any other code path.
  **Status: confirmed working, ported to AIPowered.** See AIPowered's `CLAUDE.md` section 34.

## `ViewModels\SegmentSelectorViewModel.cs` (GLSegmentValues)

- **Hierarchy field not cleared when segment changes**: `SelectedSegment`'s setter kicks
  off `LoadSegmentValuesAsync()` for the newly selected segment, which repopulates
  `HierarchyItems` (the Hierarchy combo's `ItemsSource`) for the new segment - but nothing
  cleared `SelectedHierarchy` itself (the combo's bound `SelectedItem`), so after picking
  a hierarchy value and then switching segments, the Hierarchy combo kept showing the
  previous segment's stale selection even though it no longer applied.
  Fixed by clearing the `_selectedHierarchy` backing field directly (not through the
  `SelectedHierarchy` property setter) whenever `SelectedSegment` actually changes.
  Bypassing the property setter is deliberate: setting it normally would also fire
  `LoadHierarchySegmentValuesAsync()`, which would run concurrently with (and race
  against) the `LoadSegmentValuesAsync()` call already firing for the newly selected
  segment - clearing the field directly and raising `OnPropertyChanged` avoids that.
  **Status: confirmed working, ported to AIPowered.** See AIPowered's `CLAUDE.md` section 35.

## `AddinModule.cs` / `Utilities\CommonMethods.cs`

- **Add-in crash on drilldown hyperlink click** (found while triaging a colleague's
  `GLSense_Logs_13-Aug-2026.log`): `CommonMethods.EnableExcelSettings()`/
  `DisableExcelSettings()` deliberately log-and-`throw;` on failure (e.g. a transient
  `COMException 0x800A03EC` toggling `DisplayAlerts` while Excel is busy/closing). That's
  fine for callers inside a `try` with their own `catch`, but 5 call sites in
  `AddinModule.cs` invoked them completely unguarded - either before the enclosing `try`
  even starts, or bare inside a `finally` block - in methods reachable from `async void`
  Excel-event/ribbon-click handlers (`adxExcelAppEvents1_SheetFollowHyperlink`,
  `RibHighlight_OnClick`, `RibRefreshRange_OnClick`, `ResetBalances` fire-and-forgotten
  from `RibClear(Sheet)_OnClick`, `RowProcessor.ExecuteAsync` awaited from
  `RibHideRows_OnClick`/`RibUnHideRows_OnClick`). An exception escaping an `async void`
  method (or a `finally` block) has no catch to land in - it reaches
  `AppDomain.UnhandledException` and takes the whole add-in down. Confirmed in the log:
  the very last thing before a ~52s gap and an add-in restart was exactly this - a
  drilldown's own `finally` caught its own restore failure locally (see
  `Drilldowns\DDDatatoWorksheet.cs`'s `DD_DatetoWorksheet`, which already guards this
  correctly), but the outer `SheetFollowHyperlink` handler's own `finally` calling
  `EnableExcelSettings()` was unguarded and crashed.
  Fixed by adding `CommonMethods.TryDisableExcelSettings(context)` /
  `TryEnableExcelSettings(context)` - non-throwing wrappers that log and swallow instead -
  and switching all 5 vulnerable call sites in `AddinModule.cs` to use them (pre-`try`
  `Disable` calls now `return` early if it fails, instead of proceeding as if Excel were
  actually in the disabled state). `CommonMethods.EnableExcelSettings()`/
  `DisableExcelSettings()` themselves are unchanged, since their other callers (e.g.
  `Drilldowns\BalanceRefresh.cs`, `DD_BL.cs`/`DD_JL.cs`/`DD_SL.cs`,
  `Utilities\SegmentDiscoverer.cs`/`PeriodsDiscoverer.cs`) already run inside a `try` with
  a real `catch`, so the throw-on-failure contract is still exactly what they need.
  **Status: fixed in FinalWorkingCode only so far** - the same 5-call-site shape likely
  exists in AIPowered's `AddinModule.cs` too; port once confirmed there.

- **Separately (not yet root-caused): continuous `GetRangeValueSafe` COM exceptions**.
  The same log had 5941 occurrences of `GLConfiguratorViewModel.GetRangeValueSafe`
  failing to resolve `Excel.Application.Range["AABCJ!..."]` (4 distinct cell refs on a
  sheet named `AABCJ`) - literally every time it was attempted, for the entire ~5 hour
  session. `LogException`'s 5-second dedupe window (see `Utilities\LogUtility.cs`) is why
  the log shows a clean ~5s cadence per message instead of the true call frequency - the
  underlying re-validation (`IsJournalValidationSatisfied`/`GetFieldValue`, triggered via
  `FieldBinding.IsComboEnabled`/`IsRefEnabled` getters) runs far more often than that.
  Not a crash - `GetRangeValueSafe` already catches and returns `null` - but it means
  some Balance Configurator field (Activity/BalanceType/CurrencyType/AccountAssignment)
  had its `RefValue` pointing at a sheet reference that never resolves in that user's
  workbook (renamed/deleted sheet, or wrong active workbook at read time are the two
  likely causes). Needs the reporting colleague's workbook/repro steps to pin down which
  field and why `AABCJ` doesn't resolve - not fixed yet.