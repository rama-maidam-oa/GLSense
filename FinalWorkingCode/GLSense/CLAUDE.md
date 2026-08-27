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
  actually in the disabled state).
  **Correction (same pass, caught on a second look prompted by "is that really all the
  call sites?")**: the note above originally claimed every *other* `Disable`/
  `EnableExcelSettings()` caller in this codebase was already safely inside a `try`/
  `catch`. That was only checked for the `Disable` half - the `Enable` half (almost always
  bare in a `finally`) was not actually re-checked file-by-file at the time, and turned out
  to have the exact same unguarded-in-`finally` shape in **8 more files**:
  `Drilldowns\BalanceRefresh.cs` (`InitializeAsync`'s pre-try `Disable` and
  `CleanupAsync`'s `Enable`, both inside a `RunExcelAsync` marshaling lambda),
  `DD_BL.cs`/`DD_JL.cs`/`DD_ExcelPrecedents.cs`/`DrillCellHighlighter.cs`/
  `Utilities\PeriodsDiscoverer.cs`/`SegmentDiscoverer.cs`/`Views\GLSegmentDiscovery.xaml.cs`
  (`Enable` bare in each one's outer `finally`), and `DD_SL.cs` specifically also had its
  `Disable` call sitting at the very top of the method with **no enclosing `try` at all**
  (matching `AddinModule.cs`'s worst case exactly). All 9 fixed the same way. Found by
  auditing AIPowered's equivalents first (see that repo's `CLAUDE.md` section 36) and
  finding the identical shape there, which is what prompted re-checking these
  FinalWorkingCode originals rather than assuming the first pass had been exhaustive.
  **Status: fixed in FinalWorkingCode (build-verified) and ported to AIPowered.**

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

## `Views\GLSegmentDiscovery.xaml.cs`

- **Excel crash/freeze double-clicking Insert (Hierarchy Discoverer)**: ported from
  `11.1.0_NewUI`'s already-fixed version of this same file (that branch had already hit
  and fixed this independently). Root cause: `BtnSubmit_Click` was a plain synchronous
  method with no re-entry guard and no busy overlay - a second "Insert" click queued on
  the message loop while the first write was still running (`WriteValuesToExcel` is a
  long, fully synchronous run of Excel COM calls) reached the handler a second time and
  started a second overlapping write into a cell range that already had formulas from the
  first write. With Calculation left on Automatic (the code only ever toggled
  ScreenUpdating/DisplayAlerts/EnableEvents via `DisableExcelSettings()`), each formula
  written references the previous cell in the chain, so every single `cell.Value`
  assignment in the write loop dirtied and immediately recalculated its own entire
  downstream suffix - an O(n^2) recalculation storm (525 cells measured as ~275,000 UDF
  calls elsewhere in this codebase for the identical class of bug) indistinguishable from
  a permanent freeze, with `DisplayAlerts=false` hiding any dialog that might have hinted
  Excel was still (uselessly) working. Separately, the busy overlay wasn't reliably
  showing before the freeze either, since nothing forced WPF to paint it before the
  synchronous write loop blocked the UI thread.
  Fixed by porting `11.1.0_NewUI`'s version of `BtnSubmit_Click` (now `async void`):
  bails out immediately if `btnSubmit.IsEnabled` is already `false` (a write in
  progress); disables `btnSubmit` and shows the busy overlay before the write, calling a
  new `PumpDispatcherFrame()` helper (WPF's "DoEvents" equivalent - pushes a nested
  dispatcher frame processed at Background priority) to force that overlay to actually
  paint first; switches `Excel.Calculation` to Manual for the duration of the write and
  does exactly one `Calculate()` pass afterward (O(n) instead of O(n^2)); restores
  `Calculation`/`btnSubmit.IsEnabled`/the busy overlay in `finally` regardless of outcome.
  Also added a `PumpEveryNWrites` helper (pumps the dispatcher every 20 writes) hooked
  into all four write loops (`WriteVerticalForward`/`WriteHorizontalForward`/
  `WriteVerticalBackward`/`WriteHorizontalBackward`) so the busy overlay's spinner/"Time
  Elapsed" counter keeps visibly animating during a large write instead of freezing on
  whatever frame was current when the loop started.
  `PumpDispatcherFrame()` didn't exist anywhere in this codebase (it lives on
  `11.1.0_NewUI`'s `BaseWindow.cs`, a shared base class this branch's `GLSegmentDiscovery`
  doesn't have - it derives from `DpiAwareWindow` instead), so it was added as a private
  method scoped to this one file rather than added to `DpiAwareWindow.cs` itself, to avoid
  changing behavior for the ~15 other windows that derive from it.
  **Status: build-verified on `11.1.0`.** Also needs the equivalent port into AIPowered's
  `GLSense.Addin.Core\Views\GLSegmentDiscovery.xaml.cs`, which has a different (already
  reliable, async-yielding) busy-overlay mechanism but is missing both the re-entry guard
  and the Calculation-manual/single-`Calculate()`-pass fix - see AIPowered's `CLAUDE.md`.

## Blank window on open (release-blocking, reported via video + screenshots showing all 3 stages of one sample window, `GLLOVs`)

Two distinct root causes, both fixed together per the user's request:

- **Universal, every window - WPF cold-start blank first frame.** The very first time WPF
  ever shows a `Window`/`DataGrid`/custom control of a given type in this process, it pays
  a one-time cost to parse XAML, apply styles/`ControlTemplate`s, and JIT-compile the
  generated code behind them - and the native HWND becomes visible (`Show()`/
  `ShowDialog()` returns control to Windows) before that first frame is actually
  composited, so the user sees a completely blank/white rectangle (confirmed in the first
  of the three shared screenshots - no title text, no static "Ledger:" label, nothing at
  all, not even elements that don't depend on any data binding) until WPF catches up. This
  is a well-known, generic WPF effect, not specific to `GLLOVs` or any one window - every
  `DpiAwareWindow`-derived window shares the same `DataGrid`/`ExcelRefEditControl`/
  `AppOverlay` controls that pay this cost.
  Fixed with a new `Utilities\WpfWarmup.cs`: `WpfWarmup.WarmUpInBackground()`, called once
  from `AddinModule.cs`'s `AddinModule_OnRibbonLoaded` right after
  `MahAppsBootstrapper.PreloadResources()` (the only prior "warm-up" in this codebase, and
  it only forces XAML *resource dictionaries* to parse - never instantiates an actual
  `Window`/`DataGrid`, so it never touched the JIT/template-application/first-composite
  costs that actually cause this). Dispatches at `DispatcherPriority.ApplicationIdle` (so
  it never blocks ribbon load or Excel's responsiveness) to construct a throwaway `Window`
  containing a `DataGrid` (with sample rows so its row/cell templates actually get
  exercised), an `ExcelRefEditControl`, and an `AppOverlay` - the controls shared by nearly
  every real window - positioned far off-screen with `Opacity=0`, `ShowActivated=false`,
  `ShowInTaskbar=false` so the user never perceives it, then `Show()`s and immediately
  `Close()`s it once `ContentRendered` fires. This pays the entire one-time JIT/style cost
  silently in the background before the user ever opens a real window, instead of it being
  visible on whichever window they happen to open first. `ExcelRefEditControl`'s own
  `Loaded` handler looks for an `IWarningHost` ancestor and finds none in the throwaway
  window, so `ExcelRefManager.SetupControl` is never called - no side effects from the
  warm-up itself. New file needed an explicit `<Compile Include>` entry in `GLSense.csproj`
  (old-style project format, no implicit globbing).

- **Per-window - initial data load runs with no loading indicator.** Separately from the
  above, `GLLOVs` (the second/third screenshots - static chrome rendered, but the LOVs
  `DataGrid` sitting empty for a further, unbounded stretch) and 10 other windows do a
  non-trivial data load in `Window_Loaded` where the busy overlay either never shows at
  all, or only covers a *later* phase of the load, leaving a real gap with zero loading
  feedback:
  - `GLLovViewModel.LoadLovRowsAsync`: busy overlay was gated on `!ledgerDataExist` (only
    shown on a cache-miss remote fetch) - in the common case (data already cached), the 8+
    local SQLite queries this method runs (SEGMENTS/ACTIVITY/BUDGETS/CURRENCIES/etc.) had
    no indicator at all. Also never hid the overlay in an exception path (no `finally`).
  - `GLGetPeriodModel`/`GLPeriodByDateModel`/`GLGetPeriodByYearModel`/`GLPeriodDetails`
    (shared by `GLGetPeriodDetails`/`GLGetPeriodStartEnd`) - all four `LoadDataAsync`
    methods fetch the initial ledger list (`GetConfiguratorLedgers`) with no overlay; the
    overlay only starts once `LoadPeriodsForLedger` runs afterward.
  - `SegmentSelectorViewModel.LoadSegmentsAsync` (`GLSegmentValues`/`GLSegmentRef`) - the
    initial `GetSegments` call had no overlay at all, unlike the hierarchy-loading path
    elsewhere in the same class which already does this correctly.
  - `SimpleSegmentViewModel.LoadSegmentsAsync` (`GLRollerGroups`) - this ViewModel had no
    `ShowBusyAction`/`HideBusyAsyncAction` properties at all; added them and wired them up
    in `GLRollerGroups.xaml.cs`'s constructor to match every other window's pattern.
  - `GLDailyRatesViewModel.LoadDataAsync` (`GLDailyRates`) - same as above, no busy-overlay
    mechanism existed on this ViewModel at all; added and wired up.
  - `GLCubeDetails.xaml.cs`: `LoadUserPreferencesForCube` (a real network API call) ran
    before `LoadCubeData`'s own `ShowBusyOverlayAsync` - overlay shown before it now (left
    up rather than shown-then-hidden-then-reshown, so `LoadCubeData`'s own call just
    updates the message with no flicker); added a safety-net `HideBusyAsync()` call in the
    outer `finally` at both call sites (`Window_Loaded` and `CmbCubes_SelectionCommitted`)
    in case `LoadUserPreferencesForCube` throws (cancellation) before `LoadCubeData` ever
    gets a chance to run its own hide.
  Fixed each by showing the busy overlay unconditionally for the actual full duration of
  the load (not gated on a cache-existence check), hiding in a `finally` so it can't get
  stuck showing on an exception path either. `GLUserConfig`/`GLJobsMonitor`/
  `GLSegmentFunctions`/`GLBalanceConfigurator`/`GLLogin`/`GLDrilldownCustomization`/`GLAbout`
  already did this correctly and needed no changes.
  **Status: build-verified (full solution).**

## `Views\GLDailyRates.xaml`

- **Cell Reference field's Select/Clear buttons unclickable**: reported as "unable to
  select or clear the reference" in `GLDailyRates`, while the same `ExcelRefEditControl`
  worked fine in `GLGetPeriod`. `ExcelRefEditControl.xaml`'s own layout puts its
  "Select Excel Cell" (`btnEdit`) and "Clear Reference" (`btnClear`) buttons at the
  control's right edge (`Grid.Column="1"`/`"2"`, `Auto`-width, after a `*`-width `TextBox`
  in column 0). In `GLDailyRates.xaml`'s Cell Reference row, the control spans
  `Grid.Column="1" Grid.ColumnSpan="2"` of the outer row Grid - but a leftover
  `<Border Grid.Column="2" ... Background="Transparent"/>` spacer was declared
  immediately after it in the same Grid. WPF hit-tests a `Background="Transparent"`
  element same as any opaque one (unlike a `null`/unset Background, which lets clicks
  pass through), and later-declared siblings paint on top - so this spacer sat directly
  over the right edge of the control, exactly where `btnEdit`/`btnClear` are, silently
  swallowing every click meant for them while leaving the left portion (the read-only
  text box) unaffected - matching the reported symptom precisely. `GLGetPeriod.xaml`'s
  equivalent Reference row has no such trailing Border, which is why it worked there. The
  same spacer pattern elsewhere in `GLDailyRates.xaml` (e.g. the Conversion Type row) is
  harmless, since those rows don't have a real interactive control spanning under it.
  Fixed by deleting the redundant spacer Border from the Cell Reference row, matching
  `GLGetPeriod.xaml`'s pattern exactly.
  **Status: build-verified in both codebases** (`GLSense.csproj` here; AIPowered's
  identical `GLSense.Addin.Core\Views\GLDailyRates.xaml` got the same fix, verified with
  `/p:SignAssembly=false`, this project's existing verification-only override for
  `GLSense.Contracts.pfx`).