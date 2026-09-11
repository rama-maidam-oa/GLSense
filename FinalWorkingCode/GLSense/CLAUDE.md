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
  **Status: same bug found on the `11.1.0-window-flash-redo`/`main`/`11.1.0` branches too
  (this one, `wpfui-removal-phase1`, has its own copy of `GLDailyRates.xaml` under a
  different base class - `views:BaseWindow` in AIPowered here vs. `utils:DpiAwareWindow`
  elsewhere - but the same Cell Reference row/spacer shape); fixed independently on each,
  in both FinalWorkingCode and AIPowered.

## `Drilldowns\DD_BL.cs` / `DD_JL.cs` / `DD_SL.cs`

- **GLWaitWindow processing title missing or wrong for some drilldowns**: reported as
  the processing/wait window not showing the drilldown's full name, or showing the wrong
  one, for some drilldown types. `GLWaitWindow.xaml`'s `txtTitle` defaults to
  "Refreshing Data" until `SetProcessTitle(...)` is called.
  - `DrilldownBl.ProcessBLDrilldown` (handles ddType `BL`, `BL_JL`, `BL_SL`, and `UF` -
    see `AddinModule.RunBalanceDrilldownAsync`) never called `SetProcessTitle` at all, so
    the window was stuck on the XAML default "Refreshing Data" for every one of those
    four drilldown types, regardless of which was actually running.
  - `DrilldownJl.ProcessJLDrilldown` (handles ddType `JL`, `BLDD_SL`, and `BLDD_UF` - see
    `AddinModule.RibJournalDD_OnClick`/`RibBalancesDDToSubLedger_OnClick`/
    `RibBalancesDDToUnified_OnClick`) stored `_ddType` in a field but hardcoded the title
    to the literal string `"Journals Drilldown"` regardless of its value - so the two
    Balances-Drilldown-to-X types launched through this class showed "Journals Drilldown"
    instead of their real names.
  - `DrilldownSl` hardcoded `"Subledgers Drilldown"` (lowercase "l"), which only differs
    from `DrilldownType.SL`'s canonical `[Description("SubLedgers Drilldown")]` by
    casing, but was still inconsistent with the single source of truth for these display
    strings.
  `Common\DrilldownMetadata.GetDisplay(DrilldownType)` (backed by
  `Common\DrilldownType.cs`'s `[Description(...)]` attributes) already existed as that
  source of truth and was already used correctly elsewhere (e.g.
  `DDDatatoWorksheet.cs`'s toast messages), just not wired into these three progress-window
  title call sites.
  Fixed by having `DrilldownBl`/`DrilldownJl` parse their own `_DDType`/`_ddType` field
  via `Enum.TryParse<DrilldownType>` and pass `DrilldownMetadata.GetDisplay(ddEnum)` as
  the title (falling back to the raw string if parsing fails), and switching `DrilldownSl`
  to call `DrilldownMetadata.GetDisplay(DrilldownType.SL)` instead of its hardcoded
  literal. Build-verified.
  **Status: fixed in both FinalWorkingCode and AIPowered.** AIPowered's
  `GLSense.Addin.Core\Drilldowns\DD_BL.cs`/`DD_JL.cs`/`DD_SL.cs` had the exact same three
  gaps and got the identical fix - see AIPowered's `CLAUDE.md` section 42.
  in both FinalWorkingCode and AIPowered.

## `Helpers\LogHelper.cs`

- **20MB log-file rollover used NLog's "Legacy/unstable" archive handler, and archived
  filenames carried no date**: `FileName` is a dynamic layout (`${date:format=dd-MMM-yyyy}`,
  see the header comment above about the log file being per-day), and the archive config
  additionally set `ArchiveFileName = "...\GLSense_Logs_{#}.log"`. NLog's own wiki warns
  this combination ("Dynamic FileName Archive Logic" + an explicit `ArchiveFileName`) causes
  "unexpected archive behavior" - confirmed by tracing NLog 6.1.4's own source
  (`FileTarget.cs`'s `CreateFileArchiveHandler`): merely setting `ArchiveFileName` at all,
  regardless of its content, forces the `LegacyArchiveFileNameHandler` path, which the
  source itself comments as `"Legacy / unstable because file-move can fail because of
  file-locks from other applications"` - a real risk here since this target also sets
  `KeepFileOpen = true` (an exclusive lock on the active file). Separately, `{#}` is
  deprecated syntax in NLog v6 (superseded by `ArchiveSuffixFormat`); its legacy
  compatibility shim only strips `{#}` when preceded by `.`/`_`/`-`, so archived files
  were named like `GLSense_Logs_00.log`, `GLSense_Logs_01.log` - the date was lost, and a
  size-rollover on one day's file shared the same flat sequence-number pool as any other
  day's rollovers, since nothing in the archive name distinguished dates.
  Fixed by removing `ArchiveFileName` entirely and setting `ArchiveSuffixFormat = "({0})"`
  instead. This routes size-based rollover through NLog 6's `RollingArchiveFileHandler`
  ("Updated dynamic sequence handling without file-move-logic" per its own source comment)
  - it opens a new, already-numbered file instead of renaming the full one, so there's no
  lock contention with `KeepFileOpen`. Per `FileTarget.cs`'s `BuildFullFilePath`, the suffix
  is only appended once `sequenceNumber > 0`, so the day's first/active chunk stays plain
  (`GLSense_Logs_{date}.log`), and each subsequent 20MB rollover produces
  `GLSense_Logs_{date}(1).log`, `(2).log`, etc. Numbering is naturally scoped per day too,
  since the wildcard NLog uses internally to find the next sequence number is derived from
  the already-dated active filename.
  Verified two ways: (1) a standalone NLog 6.1.4 console harness reproducing this exact
  `FileTarget` config against a 10KB threshold produced `GLSense_Logs_<date>.log` through
  `<date>(6).log` cleanly, no overwrites; (2) a real Excel run with
  `AppConstants.LogMaxFileSizeBytes` temporarily set to 10KB produced
  `GLSense_Logs_31-Aug-2026.log` through `(4).log` in the real logs folder, each ~10-11KB,
  content chronologically intact across the split (cross-checked via embedded HTTP response
  `Date:` headers in the logged API traffic). `LogMaxFileSizeBytes` was reverted to 20MB
  (`20 * 1024 * 1024`) after the test.
  Note: `FileTarget.Header` (the "Environment Snapshot" block) is written by NLog whenever
  it creates a new physical file - this was already true before this fix (any rollover, old
  handler or new, triggers it), so a 20MB size-rollover in production will still re-emit the
  header into the new file, not just once per day as the header's own placement comment
  assumes. Rare in practice at 20MB and not a regression from this fix, just worth knowing.
  **Status: fixed in both FinalWorkingCode and AIPowered.** AIPowered's
  `GLSense.Shared\Logger.cs` had the exact same `ArchiveFileName = "...\GLSense_Logs_{#}.log"`
  shape and got the identical fix - see AIPowered's `CLAUDE.md` section 43.

## `Views\AppOverlay.xaml.cs` and multiple `Views\*.xaml.cs` (GLJobsMonitor et al.)

- **GLJobsMonitor "Download Logs" window blurs but no success/error toast shown,
  intermittently on repeated clicks**: reported as the window going blurred without a
  message after clicking Download Logs, worse the more the user clicked - correctly
  suspected as a race condition. Root cause traced to two compounding gaps:
  1. None of `GLJobsMonitor.xaml.cs`'s footer buttons (Refresh/Download Logs/Download
     Outputs/Delete/Delete All) had a re-entrancy guard, so a second click before the
     first click's async operation finished started a second, concurrent call into
     `GLSubmittedJobsViewModel`, both driving the single shared `AppOverlayControl`.
  2. `AppOverlay.HideBusyAsync()` tracked its storyboard-completion callback in one
     instance field, `_hideBusyHandler`. When two `HideBusyAsync()` calls raced (from
     two overlapping operations, or from the busy overlay's own Cancel button firing
     `HideBusyAsync()` while the operation's own completion code called it again a
     moment later - `cancelAction` here doesn't actually cancel the underlying async
     work, it only hides the overlay early), the second call unsubscribed and
     discarded the first call's completion handler before it ever fired, and
     restarted the fade-out storyboard from scratch. The first call's
     `TaskCompletionSource` was then never completed, so its `await HideBusyAsync()` -
     called right before the success/error toast in every caller (e.g.
     `GLSubmittedJobsViewModel.DownloadLogsAsync`) - hung forever, and that toast never
     showed, while the overlay was left in whatever visual state the second call's
     animation produced.
  Fixed in two places:
  - `AppOverlay.HideBusyAsync()`: added a `_pendingHideBusyTcs` list. If a hide
    animation is already in flight (`_hideBusyHandler != null`) when a new
    `HideBusyAsync()` call arrives, it no longer steals/restarts the storyboard - it
    just adds its `TaskCompletionSource` to the pending list and lets the in-flight
    animation's completion handler (`CompletePendingHideBusy()`) resolve every pending
    caller at once. This is shared infrastructure used by every window that hosts an
    `AppOverlay`, so this half of the fix protects all of them, not just
    GLJobsMonitor.
  - Per-window re-entrancy guards, added to every window found (via a full audit of
    `Views\*.xaml.cs`) to have an `async void` button-click handler that touches the
    shared overlay with no existing guard: `GLJobsMonitor.xaml.cs` (all five footer
    buttons, plus the initial `Window_Loaded` load - one shared `_actionInProgress`
    flag + `SetActionButtonsEnabled(bool)`, since only one of these operations should
    ever run at a time regardless of which button started it),
    `GLDrilldownCustomization.xaml.cs` (`BtnSaveLocally_Click`), `GLLOVs.xaml.cs`
    (`CmdSubmit_Click`), `GLRollerGroups.xaml.cs` (`BtnOK_Click`),
    `GLSegmentValues.xaml.cs` (`BtnOK_Click`) - each via `if (_actionInProgress)
    return;` + disabling its own button for the duration, mirroring the pattern
    `GLSegmentDiscovery.xaml.cs`'s `BtnSubmit_Click` already used for the unrelated
    Explode-All double-click freeze bug (see that section above) - `if
    (!btnSubmit.IsEnabled) return;`.
    `GLCubeDetails.xaml.cs` (`BtnValidateCube_Click`/`BtnOK_Click`) and
    `GLUserConfig.xaml.cs` (`CmdSave_Click`/`CmdReset_Click`) got a narrower,
    same-button-only guard (`if (!btn.IsEnabled) return;`, toggled per-button) rather
    than a guard shared across both buttons in the pair, since those two already use a
    shared `_activeCancellation` field so that clicking one deliberately cancels an
    in-flight operation from the other (e.g. OK cancelling an in-flight Validate) -
    existing, intended behavior this fix does not change. `GLUserConfig.xaml.cs`'s
    Save/Reset buttons have no `x:Name` in XAML, so the guard toggles `IsEnabled` via
    `sender` instead of a named field.
  Audited every `Views\*.xaml.cs` file for this shape; windows with only synchronous
  click handlers (no `async void` touching the overlay) were left unchanged since they
  can't race this way.
  Build-verified (full solution). Needs the identical port to AIPowered's
  `GLSense.Addin.Core\Views\AppOverlay.xaml.cs`/`GLJobsMonitor.xaml.cs` (confirmed to
  have the exact same `_hideBusyHandler` shape) and its other affected windows.
  **Status: fixed in FinalWorkingCode; AIPowered port pending.**

- **Correction - the fix above didn't fully resolve it: toast still appeared while the
  busy overlay was still visible, and closed without waiting its full duration**
  (reported via `videoframe_26987.png`, showing the "Downloading logs..." spinner and
  the "Logs downloaded to..." toast on screen at once). Root cause is a second, deeper
  bug the `_hideBusyHandler`/re-entrancy fix above never touched: `GLJobsMonitor.xaml.cs`
  wired `ShowInfoAsyncAction`/`ShowWarningAsyncAction`/`ShowStatusAsyncAction`/
  `ShowBusyAction`/`HideBusyAsyncAction` as `async () => await
  Dispatcher.InvokeAsync(async () => await AppOverlayControl.XyzAsync(...))`. This is a
  classic WPF `Dispatcher.InvokeAsync` gotcha: since the inner delegate is itself
  `async`, C# infers `Dispatcher.InvokeAsync<TResult>(Func<TResult> callback)` with
  `TResult = Task` - the `DispatcherOperation<Task>` is considered *complete* as soon as
  the delegate returns control at its own first `await` (handing back a still-pending
  `Task` as its "result"), not when that inner `Task` actually finishes. The outer
  `await Dispatcher.InvokeAsync(...)` awaits the operation, receives that pending `Task`,
  and never awaits it further - so `HideBusyAsyncAction()`/`ShowInfoAsyncAction()` were
  returning to `GLSubmittedJobsViewModel.DownloadLogsAsync` (and every other caller)
  almost immediately: before `AppOverlay.HideBusyAsync()` had actually collapsed the busy
  overlay, and before `ShowInfoAsync()`'s toast had run for its real duration - explaining
  both halves of the reported symptom (overlay+toast overlapping, and the toast's
  lifetime no longer being honored by anything awaiting it).
  This exact shape (and its fix) already existed once in this codebase -
  `GLWaitWindow.ShowConfirmToastAsync` uses `Dispatcher.InvokeAsync(() =>
  AppOverlayControl.ShowConfirmAsync(message)).Task.Unwrap()` specifically to avoid it -
  but `GLJobsMonitor.xaml.cs`'s constructor wiring (added independently) never used that
  pattern.
  Fixed by switching all five `GLJobsMonitor.xaml.cs` delegates to the same non-async
  delegate + `.Task.Unwrap()` shape: e.g. `HideBusyAsyncAction = () =>
  Dispatcher.InvokeAsync(() => AppOverlayControl.HideBusyAsync()).Task.Unwrap()`. Passing
  a plain (non-`async`) lambda makes `Dispatcher.InvokeAsync<Task>` hand back the real
  inner `Task` from `.Task` (a `Task<Task>`), and `.Unwrap()` turns that into a single
  `Task` that only completes when the real async work does. Needed a new `using
  System.Threading.Tasks;` in `GLJobsMonitor.xaml.cs` for `Unwrap()` to resolve.
  This same `Dispatcher.InvokeAsync(async () => await ...)` shape also existed in several
  other windows' constructors and helper methods, and was fixed there too in a follow-up
  pass (same `.Task.Unwrap()` treatment throughout):
  - `GLBalanceConfigurator.xaml.cs`, `GLDailyRates.xaml.cs`, `GLGetPeriod.xaml.cs`,
    `GLGetPeriodByDate.xaml.cs`, `GLGetPeriodByYear.xaml.cs`, `GLGetPeriodDetails.xaml.cs`,
    `GLGetPeriodStartEnd.xaml.cs`, `GLRollerGroups.xaml.cs`, `GLSegmentFunctions.xaml.cs`,
    `GLSegmentRef.xaml.cs`, `GLSegmentValues.xaml.cs`, `GLLOVs.xaml.cs` - ctor wiring of
    `ShowBusyAction`/`ShowWarningAsyncAction`/`HideBusyAsyncAction` (and, for the five
    `GLGetPeriod*` windows, a second standalone `HideBusyAsync()` call later in the
    file). `GLLOVs.xaml.cs`'s `HideBusyAsyncAction` had a slightly different but equally
    broken variant - `async () => await Dispatcher.InvokeAsync(() => ...)` (inner
    delegate not async, but the outer `await` on the `DispatcherOperation<Task>` still
    only unwraps to the inner `Task` once and never awaits *that* - same missing
    `.Task.Unwrap()` fix applies).
  - `GLConfiguratorPane.RelaunchPane()` and `GLCubeDetails.UpdateGridAsync()` - same
    shape, not overlay-related (a WPF control reload and a DataGrid population + a
    trailing async `DgGridUpdate` call respectively); the latter could let its caller's
    `finally { HideBusyAsync() }` hide the busy overlay before the grid actually finished
    populating. Both split into a named async local function passed (undecorated) to
    `Dispatcher.InvokeAsync(...).Task.Unwrap()`, since their bodies do real synchronous
    work before/around the inner `await`, not just a single call to forward.
  - `GLLogin.xaml.cs` (two identical `finally` blocks) - `await Dispatcher.InvokeAsync(async
    () => { await AppOverlayControl.HideBusyAsync(); webView.Visibility = Visibility.Visible;
    });` had the same issue: `webView.Visibility` could be set (or the whole await return)
    before `HideBusyAsync()` genuinely finished. Same named-local-function fix.
  - `GLUserConfig.xaml.cs` - six near-identical `Hide-busy-and-show-<Error/Success/Warn/
    Info>Async` helpers all had this exact shape (`Dispatcher.InvokeAsync(async () => {
    await AppOverlayControl.HideBusyAsync(); await AppOverlayControl.ShowXAsync(...); })`)
    - the same class of bug as the originally-reported GLJobsMonitor symptom, just not
    reported for this window yet.
  Each file needing it got `using System.Threading.Tasks;` added for `.Unwrap()` to
  resolve (`GLJobsMonitor.xaml.cs`, `GLDailyRates.xaml.cs`) - most already had it.
  Build-verified (`GLSense.sln`, Debug config, full solution).
  **Status: fixed in FinalWorkingCode only so far** - not yet ported to AIPowered.

## `AddinModule.cs`

- **No confirmation before deleting a saved drilldown customization**: `RibDDDeleteConfiguration_OnClick`
  deleted the saved customization for the selected cube (`DrilldownMetadataXmlStore.Delete`)
  immediately on click, with no chance to back out of an accidental click.
  Fixed by prompting with the existing `GLMessageWindow` (via
  `CommonFunctions.GLSenseMessage(..., MessageBoxIcon.Question, MessageBoxButtons.YesNo)`,
  the same pattern already used elsewhere, e.g. the chart-of-account-change prompt in
  `RunBalanceDrilldownAsync`) before deleting, with wording calling out that the deletion
  cannot be undone. Anything other than `Yes` (`No`, or closing the window) returns
  without touching the store.
  **Status: needs the identical port to AIPowered's `AddinModule.cs`.**

## `ViewModels\GLConfiguratorViewModel.cs`

- **Budget accidentally hidden from Actual Flag when Balance Type is CTD**: `IsBalanceTypeSupportingBudget()`
  only allowed PTD/YTD/QTD/PJTD, omitting CTD - so `UpdateActualFlagsForConditions()`
  (which calls it to decide `hideBudget`) hid `Budget` from the Actual Flag dropdown
  whenever Balance Type was CTD, even though Budget is a valid Actual Flag for CTD.
  This contradicted the code's own intent elsewhere in the same file:
  `UpdateBalanceTypesForConditions()`'s Issue-3 comment already documents "ActualFlag=Budget
  restricts Balance Type to PTD/YTD/QTD/CTD/PJTD" and always keeps CTD in the rebuilt
  `BalanceTypes` list regardless of Actual Flag, i.e. CTD+Budget was always meant to be a
  valid combination in that direction - `IsBalanceTypeSupportingBudget()` just never
  matched it in the reverse direction (Balance Type → Actual Flag options).
  Fixed by adding `AppConstants.BalanceTypeCTD` to `IsBalanceTypeSupportingBudget()`'s
  allowed list.
  **Status: needs the identical port to AIPowered's identical
  `GLSense.Addin.Core\ViewModels\GLConfiguratorViewModel.cs`.**

## `Utilities\CommonMethods.cs` / new `Utilities\ComMessageFilter.cs`

- **Excel hangs on Hide Rows with Zeros / Unhide Rows if the user clicks into Excel while
  the process popup is showing** (reported via `GLSense_Logs_10-Sep-2026.log`, GLSense
  11.1.0 - confirmed present, byte-identical row-hide/unhide code and `CommonMethods.cs`
  between `11.1.0` and `11.1.1`). `RowProcessor.ExecuteAsync` (`AddinModule.cs`) disables
  `ScreenUpdating`/`DisplayAlerts`/`EnableEvents`, shows `GLWaitWindow` **non-modally**
  (`DpiAwareWindow.ShowWithOwner()` just calls `this.Show()`, not `ShowDialog()` - Excel
  behind the popup stays fully interactive), then loops setting `RowHeight` on Excel COM
  ranges with `await Task.Yield()`/`Dispatcher.InvokeAsync` between batches. Because the
  popup is non-modal and the loop repeatedly yields to the message pump, a user click into
  Excel during that window races with GLSense's own in-flight COM calls - Excel's
  automation layer rejects the incoming call as busy (`COMException 0x800AC472`,
  `VBA_E_IGNORE`). The codebase had no `IOleMessageFilter`/`CoRegisterMessageFilter`
  registered anywhere, which is the standard mechanism that would otherwise retry a
  busy-rejected automation call transparently instead of throwing. That exception then
  cascaded into cleanup: `CommonMethods.EnableExcelSettings()` set `ScreenUpdating`/
  `DisplayAlerts`/`EnableEvents` back to `true` **sequentially with no per-property
  handling** - if the first line (`ScreenUpdating = true`) hit the same busy rejection,
  the method threw immediately and never reached the other two properties, and
  `TryEnableExcelSettings` (the `finally`-block wrapper) just logged and swallowed it with
  no retry. Result: `ScreenUpdating` stuck at `false` with no recovery path - Excel stops
  redrawing and looks completely hung. The log's `16:17:33` and `17:31:22` entries show
  this exact sequence: `GetBalanceTotalRange`/`RowHeight` COM error → `Failed to enable
  Excel settings` → `Failed to restore Excel settings after RowProcessor.ExecuteAsync`.
  Fixed two ways:
  1. New `Utilities\ComMessageFilter.cs`: registers the classic OLE busy-retry
     `IOleMessageFilter` (`CoRegisterMessageFilter`) for the whole process, in
     `AddinModule_AddinInitialize`, revoked in `AddinModule_AddinBeginShutdown`. This lets
     a transient "Excel is busy" rejection retry automatically instead of throwing,
     addressing the underlying race for every COM call made from this add-in, not just
     the row hide/unhide path.
  2. `CommonMethods.cs`: `DisableExcelSettings()`/`EnableExcelSettings()` now set each of
     the three properties independently through a new `TrySetComProperty()` helper that
     retries up to 3 times (150ms apart) on a `COMException` before giving up, and no
     longer abandon the remaining properties when one throws. `DisableExcelSettings()`
     additionally rolls back whatever it did manage to disable if it can't fully succeed,
     so a partial failure never leaves e.g. `ScreenUpdating` stuck `false` with nothing
     queued to restore it (previously, `RowProcessor.ExecuteAsync` called
     `TryDisableExcelSettings` *before* its own `try`/`finally`, so a partial-disable
     failure there returned early with no `EnableExcelSettings` call ever reached).
  Since every `GLWaitWindow` non-modal-popup call site (`BalanceRefresh.cs`, `DD_BL.cs`,
  `DD_JL.cs`, `DD_SL.cs`, `DD_ExcelPrecedents.cs`, `DrillCellHighlighter.cs`,
  `PeriodsDiscoverer.cs`, `SegmentDiscoverer.cs` - found via a full `Show()`/
  `ShowDialog()` audit, see below) already routes through this same
  `CommonMethods.Disable/EnableExcelSettings`, this fix covers the identical race in all
  of them, not only Hide/Unhide Rows.
  Build-verified (`GLSense.csproj`, Debug config).
  **Status: fixed in `11.1.0` only so far - port to `11.1.1`/`11.1.2` and to AIPowered's
  identical `GLSense.Addin.Core\Utilities\CommonMethods.cs` once requested.**

  **`Show()` vs `ShowDialog()` audit (requested alongside this fix)**: repo-wide search of
  every `Window.Show()`/`ShowDialog()`/`ShowWithOwner()`/`ShowDialogWithOwner()` call in
  `GLSense\**\*.cs`.
  - **Non-modal on purpose, same exposure as the bug above (all `GLWaitWindow`
    progress/busy popups, all go through `CommonMethods.Disable/EnableExcelSettings`,
    now covered by the fix above)**: `AddinModule.cs` (`RowProcessor` - the reported bug),
    `Drilldowns\BalanceRefresh.cs`, `Drilldowns\DD_BL.cs`, `Drilldowns\DD_JL.cs`,
    `Drilldowns\DD_SL.cs`, `Drilldowns\DD_ExcelPrecedents.cs` (calls `SetExcelOwner()` +
    `Show()` directly instead of the `ShowWithOwner()` helper, same non-modal shape),
    `Drilldowns\DrillCellHighlighter.cs`, `Utilities\PeriodsDiscoverer.cs`,
    `Utilities\SegmentDiscoverer.cs`.
  - **Non-modal on purpose, unrelated to this bug (no Excel COM loop running while
    shown)**: `AddinModule.cs`'s `blpane.Show()` (a docked `ADXTaskPane` UserControl, not
    a `Window`), `Utilities\WebView2NavigationResilience.cs`'s `popup.Show()`,
    `Utilities\WindowLoadingPlaceholder.cs` (the shared cross-window loading placeholder),
    `Utilities\WpfWarmup.cs`'s off-screen invisible warm-up window (see the "Blank window
    on open" section above - deliberately never shown to the user).
  - **Everything else already modal** via `ShowDialogWithOwner()`/`ShowDialog()` - the
    large majority of real data/config windows, plus `Helpers\SnapshotDialogHelper.cs`,
    `Views\GLAccountsRef.xaml.cs`, and `GLMessageWindow` (`Utilities\CommonFunctions.cs`).
  No changes made from this audit alone - the `GLWaitWindow` popups are non-modal by
  design (so their Cancel button/live progress text stay usable while a long operation
  runs), and the message-filter + retry fix above addresses the actual race without
  changing that UX. Flagging in case there's an appetite to also make `GLWaitWindow`
  application-modal as a second layer of defense - that would prevent this specific race
  outright (no click can reach Excel while it's up) but is a bigger behavior change than
  what was asked for here.