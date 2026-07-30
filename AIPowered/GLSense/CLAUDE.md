# GLSense (AIPowered) - fix log / session overview

This file catalogs bugs found and fixed in this codebase across recent work sessions,
organized by area. Read this before touching a file listed below - it explains *why*
the code looks the way it does, so a fix doesn't accidentally get reverted or
re-derived from scratch. For the underlying architecture (AppDomain/hot-reload
boundary, thread-affinity rules, porting-from-FinalWorkingCode conventions), see
`PORTING_GUIDE.md` in this same folder.

Scope note: unless stated otherwise, every fix below is **AIPowered-only**
(`GLSense.Addin.Core` / `GLSense` host project). Two fixes were explicitly mirrored
into `FinalWorkingCode\GLSense` as well (called out below) because the user reported
them as bugs in both codebases.

---

## 1. Window "blank gap" saga (SizeToContent) - READ THIS FIRST

**STATUS: RESOLVED, user-confirmed.** The fix in **1.4e** (`Window.ContentRendered`
hook in `BaseWindow.cs`) was tested by the user after a clean rebuild + fresh Excel
session and confirmed to have "perfectly solved the issue" across `GLCubeDetails`,
`GLServerConfiguration`, and `GLMessageWindow` - the last 3 windows that were still
showing the gap. Combined with 1.4 (`SizeToContent="Width"` for static-content
windows) and 1.4b/1.4c (per-window post-async-load resettle hooks), this closes the
saga. If this exact symptom (blank gap on open, fixed by a manual resize or a click on
the resize border) ever reappears on a window not listed below, re-read 1.4e first -
don't re-derive 1.1/1.2/1.3 from scratch, they were real but insufficient steps along
the way, not the actual fix.

Many `BaseWindow`-derived dialogs (GLMessageWindow, GLLoginDetails, GLDailyRates,
GLGetPeriod family, GLSegmentDiscovery, GLSegmentFunctions, etc.) showed a blank gap
near the window edge/footer on open, which only corrected itself once the user
manually resized the window. This went through several iterations before landing on
the actual fix - the earlier iterations are documented here too because their
supporting code (`BaseWindow.ForceSizeToContentResettle`/`PumpDispatcherFrame`, the
DPI-change resettle) is still in place as defense-in-depth, even though it's no longer
the primary fix for the windows listed in 1.4.

### 1.1 First theory (superseded): `"*"` row collapse

Original theory: a bare `Grid.RowDefinition Height="*"` for a window's content row
collapses toward 0 during `SizeToContent`'s infinite-availSize measurement pass
(opposite of how DataGridColumn `"*"` behaves). Fix applied: convert the
content-driving row from `"*"` to `"Auto"`, add a genuine empty `"*"` spacer row
between content and the footer, bump the footer row's `Grid.Row` index. Applied to
~10 windows. This helped in some cases but did **not** fully explain the bug - the gap
still recurred (see 1.2-1.4).

### 1.2 `BaseWindow.OnLoaded` resettle timing fix

`BaseWindow.cs` has `ForceSizeToContentResettle()` (toggles `SizeToContent` off/on with
a sub-pixel Width/Height nudge to force a genuine native HWND resize, since
`UpdateLayout()` alone only flushes WPF's logical tree, not the Win32 window). This was
originally deferred via `Dispatcher.BeginInvoke(..., DispatcherPriority.ContextIdle)`,
which raced against the window's first paint: on a cold first open (JIT/resource
loading busy the dispatcher long enough), the deferred fix usually won; on a warm
reopen of the same window type, the stale layout would get painted first. Fixed by
calling `ForceSizeToContentResettle()` synchronously in `OnLoaded`, then pumping a
nested dispatcher frame (`PumpDispatcherFrame()`, WPF's "DoEvents" equivalent) to flush
all pending layout/native-resize work before `OnLoaded` returns - removing the race
entirely. Both methods live in `Views/BaseWindow.cs`.

**Correction (see 1.4d below): the nudge amount used here was itself broken from the
start** - it was only `+0.1` logical units, which rounds to under one physical device
pixel at every DPI scale WPF actually renders at, so the native HWND very likely never
changed size at all. 1.4d fixes this - read it before assuming this section's
`ForceSizeToContentResettle()` genuinely forces a resize.

### 1.3 DPI-change resettle

`BaseWindow.AdjustForDpiChange` (fired on `WM_DPICHANGED`) applies a `LayoutTransform`
to rescale window content and calls `InvalidateMeasure()` - a purely logical
invalidation, same category of problem as 1.2. Since this add-in only applies
Per-Monitor-V2 DPI awareness via a scoped thread context (not a real app manifest,
because it's hosted inside Excel's own process - see `DpiAwarenessHelper`), Windows
commonly fires `WM_DPICHANGED` moments after a window opens even without moving
monitors, especially at >100% display scaling (125%/150%, the norm on business
laptops). This could silently re-collapse a window right after `OnLoaded`'s resettle
already fixed it. Fixed by also calling `ForceSizeToContentResettle()` +
`PumpDispatcherFrame()` from inside `AdjustForDpiChange` for `SizeToContent` windows.

### 1.4 Actual root fix (current, canonical): drop the buggy axis entirely

After a clean rebuild + fresh Excel session (version folder deleted and recreated) the
gap **still occurred on every window**, disproving 1.1 as a full explanation (by then
`GLLoginDetails.xaml` had zero `"*"` rows at all and still showed it). The real,
user-confirmed fix: for any dialog whose content height is static (fixed number of
rows, nothing conditionally shown/hidden via `Visibility="{Binding ...}"`), don't use
`SizeToContent="WidthAndHeight"` at all - switch to `SizeToContent="Width"` and give the
window an explicit `MinHeight`/`MaxHeight` range instead. This removes the height axis
from WPF's (buggy, in this hosting environment) dual-axis auto-measurement entirely,
rather than trying to patch around it.

Applied to (all confirmed via grep to have zero conditional rows):
`GLLoginDetails.xaml` (user's own fix, MinHeight="100" MaxHeight="230"),
`GLDailyRates.xaml` (350/450), `GLGetPeriod.xaml` (345/450),
`GLGetPeriodByDate.xaml` (345/480), `GLGetPeriodByYear.xaml` (345/480),
`GLGetPeriodDetails.xaml` (300/450), `GLGetPeriodStartEnd.xaml` (300/450),
`GLSegmentDiscovery.xaml` (345/520), `GLSegmentFunctions.xaml` (440/520).

**`GLMessageWindow.xaml` was deliberately left on `SizeToContent="WidthAndHeight"`** -
its message text genuinely varies in length, so a fixed height range would clip long
messages or leave dead space for short ones. It still benefits from 1.2/1.3's
defensive resettle. If it ever shows the same gap, don't reach for 1.1/1.2/1.3 first -
work out the correct mirrored fix (`SizeToContent="Height"` with a narrower/fixed width
range) instead, following the same reasoning as 1.4.

**If a new window shows this same gap symptom:** check whether its content rows are
ever conditionally hidden. If no: apply 1.4 (this is the fix that actually works).
If yes (content height genuinely varies): leave `SizeToContent="WidthAndHeight"`,
rely on 1.2/1.3's resettle safety net, and consider whether an inner `ScrollViewer`
with a sane `MaxHeight` is a better fit than fighting `SizeToContent` on that axis.

### 1.4b GLCubeDetails: content that loads asynchronously *after* Loaded

`GLCubeDetails` correctly keeps `SizeToContent="WidthAndHeight"` (it has a real
`DataGrid` in a `"*"` row - see 1.5's legitimate-`*`-usage list) and still showed the
gap on initial open, clearing only once the user picked a cube. Root cause: its own
`Window_Loaded` handler is a *second* `Loaded` subscriber (wired via XAML,
`Views/GLCubeDetails.xaml.cs`) that runs after `BaseWindow.OnLoaded` in the same
routed-event dispatch. `BaseWindow.OnLoaded`'s resettle (1.2) is synchronous and
finishes before `Window_Loaded`'s async chain (`Window_Loaded` -> ...
-> `LoadCubeData` -> `UpdateGridAsync`) has set `dgCubes.ItemsSource` - several
`await`s deep - so the resettle always measures an empty grid. Selecting a cube later
runs the same `UpdateGridAsync` while the window is already visible, which is just an
ordinary live-content relayout WPF handles correctly on its own - that's why "picking a
cube fixes it" looked like the trigger, when really it was just the first time real
row data existed at all.

Fix: made `ForceSizeToContentResettle()`/`PumpDispatcherFrame()` `protected` (were
`private`) in `BaseWindow.cs` so a derived window can call them again once its own
async content has actually loaded. `GLCubeDetails.xaml.cs`'s `UpdateGridAsync` now
calls both right after `dgCubes.ItemsSource` is set and `DgGridUpdate` completes. If
another window shows this "gap until data loads" variant (as opposed to 1.1's
already-fixed variant or 1.4's), the fix is this pattern, not another row-structure
change: call `ForceSizeToContentResettle()` + `PumpDispatcherFrame()` again from
wherever that window's real async content finally lands, not just once in `OnLoaded`.

### 1.4c The 1.4b pattern applied to every other window with the same shape

Once 1.4b's pattern (resettle again after real async content lands, not just once in
`OnLoaded`) was identified, every other `BaseWindow`-derived window still on
`SizeToContent="WidthAndHeight"` with a `DataGrid`/data-driven content was audited for
the same bug shape. Two call-graph shapes turned up, needing two different fixes:

**Shape A - content-loading chain is fully `await`ed by `Window_Loaded` itself.** Just
call `ForceSizeToContentResettle()` + `PumpDispatcherFrame()` right after that `await`
returns, in the View's code-behind:
- `GLAbout.xaml.cs` - `AboutWindow_Loaded`, right after `await CheckInstanceCompatibility()`.
- `GLJobsMonitor.xaml.cs` - `Window_Loaded`, right after `await vm.LoadJobsAsync()`.
- `GLRollerGroups.xaml.cs` - `Window_Loaded`, right after the `Dispatcher.InvokeAsync`
  block that follows `await vm.LoadSegmentsAsync(...)` (verified `LoadSegmentsAsync`'s
  internal `SelectedSegment` setter calls the synchronous, non-fire-and-forget
  `LoadSegmentValues()` directly, so by the time the `await` returns the grid's real
  data is already in place).

**Shape B - the ViewModel populates the grid via a fire-and-forget call, detached from
`Window_Loaded`'s own `await` chain** (e.g. a property setter does
`_ = SomeAsyncLoad();`/`Task.Run(...)` rather than being awaited by the caller) - there
is no View-side `await` point to hook after. Fixed with a callback property on the
ViewModel (matching the existing `ShowWarningAction`/`HideBusyAsyncAction` convention
already used in these ViewModels), invoked from inside the ViewModel right after the
grid's backing collection/property is actually populated on the UI thread, with the
View wiring it in its constructor next to those other callbacks:
- `GLLovViewModel.cs` - added `Action DataLoadedAction`, invoked inside
  `LoadLovRowsAsync`'s `_dispatcher.InvokeAsync` block, right after the `LOVRows.Add(r)`
  loop (population is triggered by `LOV_SelectedLedger`'s setter calling
  `LoadLovRows()`, which does `Task.Run(async () => await LoadLovRowsAsync())` -
  fire-and-forget). Wired in `GLLOVs.xaml.cs`'s constructor.
- `SegmentSelectorViewModel.cs` - added `Action DataLoadedAction`, invoked at the end of
  `UpdatePagingAndGrid()` (population is triggered by `SelectedSegment`'s setter doing
  `_ = LoadSegmentValuesAsync();` - fire-and-forget). `UpdatePagingAndGrid()` is a
  shared choke point for every paging/filter/search update too, not just the initial
  load - deliberately left un-guarded (no "first load only" check) since resettling an
  already-correctly-sized window is a cheap no-op. This ViewModel is shared by three
  windows, all wired the same way in their constructors: `GLSegmentRef.xaml.cs`,
  `GLSegmentManager.xaml.cs`, `GLSegmentValues.xaml.cs`.

**Ruled out** (checked, don't need this fix): `GLUserConfig` (grid fully populated
synchronously in the constructor before the window shows), `GLServerConfiguration`
(synchronous `XDocument.Load`, no `await`), `GLLogin` (server list loaded
synchronously; WebView2 content isn't row-driven), `GLDrilldownCustomization`
(WebView2-only, not row-driven), `GLMessageWindow` (fully synchronous constructor, no
`Loaded` handler - see 1.4's own note on why it stays `WidthAndHeight`), `GLWaitWindow`
(no data-driven content). `AttachmentsDialog` has the same `Window_Loaded`-runs-after-
`OnLoaded` shape but populates synchronously and already has a defensive
`ScrollViewer.MinHeight="220"` floor (see 1.1) - flagged as low-priority/likely already
masked rather than fixed outright.

**If yet another window turns up with this symptom:** first figure out which shape it
is (does `Window_Loaded` fully `await` the load, or does a property setter fire it
off detached?) before choosing which of the two fixes above to copy.

### 1.4d The actual reason the gap survived every fix above: the nudge never worked

After 1.4b/1.4c shipped, the user re-tested and the gap was **still there on both
`GLCubeDetails` and `GLServerConfiguration`** - the second of which was in 1.4c's own
"ruled out" list (its grid loads fully synchronously before `Loaded` even fires, so
`OnLoaded`'s single resettle should have measured it correctly the first time). That a
window with zero async-loading quirks still showed the bug meant the resettle
mechanism itself - not its timing - was broken. The user also reported the specific,
crucial detail that a mere **click** on the resize border (not even a drag) fixes it.

Root cause: `ForceSizeToContentResettle()`'s Width/Height nudge was only `+0.1` logical
units. At every DPI scale WPF actually renders at (100%/125%/150%...), `0.1` logical
units rounds to under one physical device pixel - so the native HWND's actual on-screen
size never changed at all. The whole "toggle SizeToContent off/on + nudge" trick had
likely been a near-total no-op at the Win32 level since it was introduced in 1.2 - it
just happened to go unnoticed on windows where 1.1's row-structure fix or 1.4's
`SizeToContent="Width"` change independently masked the symptom. A user clicking
(not even dragging) the resize border works because Windows' modal sizing loop forces
an actual non-client-frame recalculation and lets DWM recompose the window - a Win32-
level action, not a WPF logical-layout one.

Fixed in `ForceSizeToContentResettle()` (`Views/BaseWindow.cs`) two ways:
1. The nudge amount changed from `+0.1` to `+1.0` - a full logical pixel, guaranteed to
   produce an actual device-pixel size delta regardless of DPI scale.
2. Added an explicit `SetWindowPos(hwnd, ..., SWP_NOMOVE | SWP_NOSIZE | SWP_NOZORDER |
   SWP_NOACTIVATE | SWP_FRAMECHANGED)` call (new P/Invoke declaration, next to the
   other Win32 imports near the bottom of the file) right after the `SizeToContent`
   toggle - this is the direct, well-known Win32 idiom for "force this window's
   non-client frame to recompute and let DWM redraw it" without actually moving or
   resizing anything, and is functionally what a resize-border click triggers
   internally. This does not depend on the Width/Height nudge working at all, so it's
   the more load-bearing half of this fix.

This should be the actual, final fix for every window still depending on
`ForceSizeToContentResettle()` (i.e. everything NOT already moved to
`SizeToContent="Width"` per 1.4) - `GLCubeDetails`, `GLServerConfiguration`,
`GLMessageWindow`, `GLAbout`, `GLJobsMonitor`, `GLRollerGroups`, `GLLOVs`,
`GLSegmentRef`/`GLSegmentManager`/`GLSegmentValues`, `GLUserConfig`, `GLLOVs`,
`GLDrilldownCustomization`, `GLLogin`, `GLWaitWindow`. If the gap survives THIS fix on
any of them, the resettle mechanism itself is no longer the suspect - look elsewhere
(e.g. genuinely re-verify the deployment/rebuild note at the bottom of this file first,
then reconsider whether that specific window's content is measured correctly at all).

**This diagnosis turned out to be incomplete - see 1.4e, the actual final fix.**

### 1.4e The REAL final fix: `SetWindowPos` only works on an already-painted window

1.4d's fix was real and necessary, but not sufficient. The user re-tested and precisely
scoped the remaining bug: the gap survived on **exactly** `GLCubeDetails`,
`GLServerConfiguration`, and `GLMessageWindow` - and on **no other window**, including
several with the identical DataGrid-in-`"*"`-row/`SizeToContent="WidthAndHeight"` shape
(`GLAbout`, `GLJobsMonitor`, `GLRollerGroups`, `GLLOVs`,
`GLSegmentRef`/`GLSegmentManager`/`GLSegmentValues`, `GLUserConfig`). That precise a
split meant there had to be a real, mechanical difference between these 3 and
everything else - not just "some windows have async content and some don't" (several
working windows have async content too).

The actual difference, found by tracing every one of these windows' invocation
(`ShowDialog()`, all of them, always `new`, no cached instances - ruled out),
`XAML` attributes/Min-Max ranges (no pattern), and exactly *when* each window's own
code calls `ForceSizeToContentResettle()` a second time:

- Every "working" window happens to get a **second** resettle call from somewhere that
  is only ever reached *after* `ShowDialog()`'s nested message loop has already painted
  the window once - either an `await`-chained continuation in its own `Window_Loaded`
  (1.4c Shape A: `GLAbout`, `GLJobsMonitor`, `GLRollerGroups`) or a `DataLoadedAction`
  ViewModel callback (1.4c Shape B: `GLLOVs`, the `SegmentSelectorViewModel` trio).
  `GLUserConfig` similarly has an unconditional `Window_Loaded` that always runs a
  busy-overlay/network round trip after first show. This was never a deliberate part of
  1.4b/1.4c's design - it was an accidental side effect of those windows needing
  *some* post-load hook for other reasons, which happened to also run after paint.
- `GLServerConfiguration` and `GLMessageWindow` have **no `Loaded` handler of their own
  at all** - all their content loads/builds fully synchronously in the constructor,
  before the window is ever shown. The only resettle call either of them ever receives
  is `BaseWindow.OnLoaded`'s, which runs *inside* the `Loaded` routed event - i.e.
  before `ShowDialog()`'s message loop has painted the window for the first time.
- `GLCubeDetails` has a second resettle call inside `UpdateGridAsync` (1.4b), but on a
  fresh open with no cube already selected in `AppState`, the `Window_Loaded` branch
  that would reach it is gated behind a "was a cube already selected" check and never
  runs at all - collapsing this window to the same single-pre-paint-resettle situation
  as the other two on that (common) code path.

The mechanism: `SetWindowPos(..., SWP_FRAMECHANGED)` (1.4d) tells Windows to recompute
a window's non-client frame and let DWM recompose it - but that only has something to
act on once the window actually has a first frame on screen. Calling it *before* the
window has ever been painted is as much of a no-op as a user trying to "click the
resize border" of a window that isn't visible yet. Every fix through 1.4d kept
targeting *how* the resettle worked; this is the first one that identifies *when* it
was actually running.

**Fix**: `BaseWindow`'s constructor now also subscribes
`this.ContentRendered += OnContentRendered;` - `Window.ContentRendered` is the one WPF
event guaranteed to fire only after the window's content has actually been rendered/
painted on screen for the first time. `OnContentRendered` calls the exact same
`ForceSizeToContentResettle()` + `PumpDispatcherFrame()` pair as `OnLoaded` does. This
gives every `BaseWindow`-derived window a guaranteed post-paint resettle pass
regardless of whether that specific window happens to have its own async
`Window_Loaded` continuation or `DataLoadedAction` callback - fixing `GLCubeDetails`,
`GLServerConfiguration`, and `GLMessageWindow` without any window-specific changes, and
without needing to remove the (now largely redundant, but harmless) per-window hooks
added in 1.4b/1.4c.

**If the gap survives this fix on anything**, `ForceSizeToContentResettle`/
`PumpDispatcherFrame`/`OnContentRendered` are no longer credible suspects - by this
point every plausible timing/mechanism gap in the resettle approach has been closed.
Re-verify the deployment/rebuild note at the bottom of this file first, and if that's
genuinely not it, treat it as a new, unrelated bug rather than another variant of this
one.

**CONFIRMED**: user rebuilt clean, relaunched Excel, and reported this "perfectly
solved the issue" on `GLCubeDetails`, `GLServerConfiguration`, and `GLMessageWindow` -
the whole saga (1.1 through 1.4e) is closed. No further action needed on this bug
unless a genuinely new trigger surfaces.

### 1.4f 1.4's own fix reintroduced a different, purely cosmetic gap - reverted for 8 windows

User supplied screenshots of `GLGetPeriod`, `GLGetPeriodByDate`, `GLGetPeriodByYear`,
`GLGetPeriodDetails`, `GLDailyRates`, and `GLSegmentFunctions`: every one showed a large
dead grey gap between the field content and the footer buttons - i.e. exactly the
visual symptom 1.4's fix was supposed to have eliminated.

Root cause: 1.4's fix (`SizeToContent="Width"` + fixed `MinHeight`/`MaxHeight` + a
genuine `"*"` spacer row between content and buttons) does stop the *collapse-until-
manual-resize* bug, but it does so by removing height from `SizeToContent` entirely -
the window opens at a **fixed height within its Min/MaxHeight range**, and whenever the
actual field content is shorter than that fixed height, the leftover space has nowhere
to go except that spacer row, appearing as a big dead gap. FinalWorkingCode's equivalent
windows don't have this problem because `DpiAwareWindow` (`Utilities/DpiAwareWindow.cs`)
implements its own custom height-fitting logic in `AdjustSizeAndScale`/similar - it
measures `root.DesiredSize.Height` and explicitly sets `Height` to fit the actual
content (clamped to available screen space), even though the XAML itself says
`SizeToContent="Manual"`. AIPowered's `BaseWindow` has no equivalent mechanism, so
1.4's fixed-height-range approach left a real, user-visible gap that 1.4 itself never
actually eliminated for these windows - it only masked the original collapse bug.

Since then, **1.4e** (`Window.ContentRendered`-based resettle in `BaseWindow`'s
constructor) was confirmed to fix the collapse-until-resize symptom **generically**,
for every `BaseWindow`-derived window, regardless of `SizeToContent` mode. That means
1.4's height-axis-removal workaround is no longer needed to avoid the collapse bug -
so windows can go back to `SizeToContent="WidthAndHeight"` (matching FinalWorkingCode's
actual shrink-to-fit visual result) without reintroducing the original problem.

**Fix**: reverted 8 windows from `SizeToContent="Width"` + spacer-row back to
`SizeToContent="WidthAndHeight"`, removed the dead `"*"` spacer row (3 total rows now:
title/content/buttons, matching FinalWorkingCode's row count exactly), and shifted the
buttons `Border`'s `Grid.Row` from `3` down to `2` accordingly: `GLGetPeriod.xaml`,
`GLGetPeriodByDate.xaml`, `GLGetPeriodByYear.xaml`, `GLGetPeriodDetails.xaml`,
`GLGetPeriodStartEnd.xaml`, `GLDailyRates.xaml`, `GLSegmentFunctions.xaml`,
`GLSegmentDiscovery.xaml`. `MinWidth`/`MaxWidth`/`MinHeight`/`MaxHeight` were left
unchanged on all of them (still useful as clamps for `WidthAndHeight`'s own sizing).

Worth calling out separately: `GLSegmentFunctions.xaml` should never have been put in
the "Width"-only bucket in the first place - its
`ParentChildSection`/`AttributesSection`/`ResultSection`/`IncludeValuesSection` borders
are conditionally `Visibility.Collapsed` in code-behind depending on the selected
function type, so its content height genuinely varies. 1.4's own documented decision
tree says exactly this case ("content height genuinely varies") should stay on
`WidthAndHeight` rather than being forced into a fixed range - this window was
mis-categorized in the original 1.4 pass.

**`GLLoginDetails.xaml` was deliberately left untouched** - it's called out in 1.4 as
"the user's own fix" (their own `MinHeight="100" MaxHeight="230"` choice), so it wasn't
touched here even though it uses the same pattern; if it turns out to have the same
dead-gap symptom, that's a decision for the user to make, not an automatic revert.

**Status**: implemented, not yet rebuilt/tested by the user.

### 1.5 Other window-chrome/layout facts established while debugging this

- `Wpf.Ui.Controls.FluentWindow` (the `BaseWindow` base class) does **not** add hidden
  chrome/margin for a normal, non-maximized window on Win11 - traced via the actual
  package source (`Wpf.Ui 4.3.0`). Don't re-suspect `WindowChrome`/glass-frame margins
  without new evidence.
- `BaseWindow.FitToAvailableWorkArea()` only shrinks (`Math.Min` against the screen
  work area) - it never grows `Height`/`Width`/`Max*` beyond a window's own
  content-driven size, so it can't be the cause of a gap getting bigger.
- Legitimate uses of a `"*"` row (NOT a bug, leave alone): `DataGrid` (GLCubeDetails,
  GLUserConfig, GLServerConfiguration, GLLOVs, GLRollerGroups, GLAbout's
  Instance-Compatibility grid), `TabControl` (GLDrilldownCustomization), `WebView2`
  (GLLogin).

---

## 2. GLBalanceConfigurator (`Views/GLBalanceConfigurator.xaml` +
`ViewModels/GLConfiguratorViewModel.cs`)

This view is a `UserControl` (not a `BaseWindow`), hosted via `ConfiguratorPaneHost.cs`
in a plain `Window` with `SizeToContent="Manual"` - none of section 1's fixes apply to
its outer host, only to its own internal layout.

### 2.1 Row spacing / label width

- Field rows (Ledger/Activity/Balance Type/.../Account Assignment(s)) were packed too
  tightly with a large unused blank area below (before the collapsed "Get Balance
  Function Parameters" accordion). Fixed by widening the shared `ConfigRowStyle`
  margin from `0,3` to `0,8`.
- That extra height, across up to ~12 simultaneously-visible rows, could exceed the
  inner field-list `ScrollViewer`'s `MaxHeight`, clipping the last row (reported as
  "Account Assignment(s) label not fully visible"). Fixed by bumping that `MaxHeight`
  from `500` to `620`.
- Separately, `ConfigLabelStyle`'s `Width` (and each row's label-column
  `ColumnDefinition`) was `140`, too narrow for longer labels like
  "Account Assignment(s):", causing visible text truncation. Widened to `190` in both
  places (15 rows + the shared style), taking the space from the combo/refedit `"*"`
  columns.

### 2.2 Journal Source/Category always disabled (fixed in **both** AIPowered and
FinalWorkingCode)

`GLConfiguratorViewModel.GetFieldValue()`'s `RefValue` branch returned the raw,
unresolved cell-address string (e.g. `'Sheet1'!$B$2`) instead of resolving it through
Excel - so whenever Activity/BalanceType/CurrencyType was set via Reference (instead of
the ComboBox), `ValidateJournalFields()`/`IsJournalValidationSatisfied()` could never
match the raw address against the hardcoded tokens (Debit/DR/Credit/CR/Net,
PTD/YTD/CTD/JED variants, E/Entered/Total), permanently disabling Journal
Source/Category. Fixed by resolving the ref through `GetRangeValueSafe()` first,
mirroring the already-correct `GetResolvedAccountAssignmentValue()` pattern in the same
file, falling back to the raw ref text only if resolution fails.

Also fixed a related token-list mismatch: AIPowered's `Converters.JournalValidationConverter`
(bound directly in the XAML `MultiBinding`s) only allowed `{PTD,YTD,CTD}` and
`{E,ENTERED}`, while `IsJournalValidationSatisfied()` (used by `FieldBinding`'s
enable-state logic) already allowed the fuller `{PTD,YTD,CTD,JED,JEDP,JEDU}` and
`{E,ENTERED,TOTAL}`. Brought the converter's lists in line with the ViewModel's. (No
change needed in FinalWorkingCode - its converter already had the full list.)

### 2.3 End Period not populating for CTD (fixed in **both** AIPowered and
FinalWorkingCode)

`OnFieldDependencyChanged`'s `BalanceType` case (which fires on every interactive
Balance Type combo pick, via `FieldBinding.ComboValue`'s setter) correctly set
`IsEndPeriodsEnabled = true` for CTD but never called `UpdateEndPeriods()` to actually
populate the `EndPeriods` collection - that method was only wired up for the
Period-changed handler, `ApplyDefaultSelections`, and formula-param loading. Added the
missing `UpdateEndPeriods()` call to the CTD branch.

### 2.3b Label-to-combo gap + combo/refedit width balance

- **Fixed 190px label column left a dead gap after short labels**: `ConfigLabelStyle`
  hardcoded `Width="190"`, so a short label like "Ledger:" left a big blank strip
  before its combo started, while the longest label ("Account Assignment(s):") used
  nearly all of it. Fixed properly (not by guessing a smaller fixed pixel value) using
  WPF's shared-size-group mechanism: each row's label `ColumnDefinition` changed from
  `Width="190"` to `Width="Auto" SharedSizeGroup="ConfigLabelCol"`, and the parent
  `StackPanel` wrapping all ~15 rows got `Grid.IsSharedSizeScope="True"`. This makes the
  label column auto-size to the single widest label across every row (so all rows'
  combos still line up at the same x), while removing the dead space after shorter
  labels down to a 1px margin (`ConfigLabelStyle`'s `Margin` changed from `0` to
  `0,0,1,0`). `ConfigLabelStyle`'s own `Width` setter was removed entirely since the
  column now controls sizing.
- **Combo vs RefEdit width balance**: previously both were `Width="*"` (50/50 split).
  Per request ("increase combo size, reduce refedit width"), changed each row's third
  column (RefEdit) from `Width="*"` to a fixed `Width="150"`, so the combo (still `*`)
  now claims all the space the fixed-width label + fixed-width refedit don't use -
  effectively larger than before, while RefEdit is narrower and consistent across rows.
- **Status**: implemented, not yet rebuilt/tested by the user.

### 2.4 Cross-thread `NotSupportedException` on reopen

Closing and reopening the Balance Configurator threw
`System.NotSupportedException: This type of CollectionView does not support changes to
its SourceCollection from a thread different from the Dispatcher thread` out of
`PopulateDynamicCollections`. Root cause: `GLConfiguratorViewModel.LoadDataAsync`'s
continuation after `await Task.WhenAll(tasks)` could resume on a ThreadPool thread
instead of the UI thread, because this add-in's HWND-reparented WPF host (unlike
FinalWorkingCode's `ElementHost`) doesn't reliably install a
`DispatcherSynchronizationContext`. Fixed by wrapping the whole post-fetch
collection-population block in `_dispatcher.InvokeAsync(...)`, matching the pattern
already used everywhere else in that class.

Also fixed (same investigation): `GLBalanceConfigurator.xaml.cs`'s
`ShowBusyOverlayAsync` now calls `ShowOverlay()` directly via
`Dispatcher.CheckAccess()` instead of always hopping through
`Dispatcher.InvokeAsync(..., Background)`, and `ReLoadConfigurator`'s lambda has a
defensive `BalanceParametersExpander.Dispatcher.CheckAccess()/Invoke` guard around
`IsExpanded = false`.

### 2.5 Reopen: values silently fail to populate (same cross-thread root cause, different call site)

User reported: first open populates every field correctly; close and reopen the
configurator and no fields populate - no visible error (no toast), but an exception was
in the logs. Confirmed from the live log:

```
System.InvalidOperationException: The calling thread cannot access this object
because a different thread owns it.
   at System.Windows.Threading.Dispatcher.VerifyAccess()
   at System.Windows.DependencyObject.SetValue(DependencyProperty dp, Object value)
   at GLConfiguratorViewModel.UpdateParameterSummary()
   at GLConfiguratorViewModel.Ledger_PropertyChanged(...)
   at GenericLedgerModel.OnPropertyChanged(String propertyName)
   at GenericLedgerModel.set_IsSelected(Boolean value)
   at GLConfiguratorViewModel.ApplyDefaultLedgerSelection()
   at GLConfiguratorViewModel.ApplyDefaultSelections()
```

Same root cause family as 2.4 (this add-in's HWND-reparented WPF host doesn't reliably
install a `DispatcherSynchronizationContext`, so `await` continuations can resume on a
ThreadPool thread) but a different, previously-unguarded call site:
`LoadConfiguratorAsync` calls `ApplyDefaultSelections()` directly (not wrapped in
`_dispatcher.InvokeAsync`) right after `await UpdateUIAsync()`. On reopen, that
resumption can land off the UI thread; `ApplyDefaultSelections -> 
ApplyDefaultLedgerSelection` sets `GenericLedgerModel.IsSelected = true`, which
synchronously fires `PropertyChanged -> Ledger_PropertyChanged -> 
UpdateParameterSummary()`, which does a `DependencyObject.SetValue` - throwing off the
UI thread.

Compounding it: `ApplyDefaultSelections`'s own `catch` only called
`ServiceLocator.Logger?.LogException(ex)` - logged, never surfaced to the user - so the
symptom looked like silent failure rather than a visible error. Because the exception
aborted the method partway through (right after the ledger selection), every field
after it in that method (Activity/BalanceType/Currency/CurrencyType/ActualFlag/Sign/
Zeroes/Factor) was left unset - this is why "nothing was populated," not just the
ledger.

**Fix**:
- Wrapped the call in `LoadConfiguratorAsync`:
  `await _dispatcher.InvokeAsync(() => ApplyDefaultSelections());` instead of a bare
  `ApplyDefaultSelections();` call - guarantees the whole synchronous chain (including
  `UpdateParameterSummary`'s `SetValue`) runs on the UI thread, matching the pattern
  used everywhere else in this class.
- Hardened the catch block: now also calls `ShowWarningAction?.Invoke(...)` (the same
  toast delegate used elsewhere in this ViewModel) so any future failure here is
  visible to the user instead of log-only.
- **Status**: implemented, not yet rebuilt/tested by the user.

---

## 3. Ribbon / login-state

- `RibbonController.SetControlLabel` was calling `GetProperty("Label")` via reflection,
  but AddinExpress `ADXRibbonButton`'s visible-text property is actually `"Caption"` -
  the reflection was silently no-op'ing every ribbon-caption update (e.g. selected cube
  name never appearing in `RibGetCube`) with no exception/log. Fixed to
  `GetProperty("Caption")`.
- `AddinEntry.Logout()` only cleared in-memory `AppState` (`SelectedCube`/
  `SelectedLedger`), not the ribbon UI. Added explicit `SetControlLabel("RibGetCube",
  "Selected Cube : Select cube")`, `ClearComboItems("Ribledger")`,
  `SetComboText("Ribledger", string.Empty)`, `ClearComboItems("RibSegS")` right after
  `AppState.Instance.Reset()`.
- Ribbon caption format standardized to `"Selected Cube : {name}"` /
  `"Selected Cube : Select cube"` in both `AddinEntry.SyncRibbonSelectionWithAppState`
  and `GLCubeDetails.xaml.cs`'s `UpdateRibbonForCube`.
- **Ribledger dropdown only showing one value**: `GLCubeDetails.xaml.cs`'s
  `ProcessCubeSelectionNew`/`ProcessCubeSelectionReload` called
  `await LoadCubeLedgers(cube, ...)` (which populates the full `Ribledger` item list via
  `SetComboItems`) *before* committing `AppState.Instance.SelectedCube`/`SelectedLedger`
  - even though the identical-looking ordering fix for `UpdateRibbonForCube`'s caption
  update (the "commit to AppState before touching the ribbon" comment already in this
  method) was already in place. `LoadCubeLedgers` marshals its `SetComboItems` call onto
  the dispatcher at `DispatcherPriority.Background` - low/preemptible - so if Excel's
  `WorkbookActivate` handler (`AddinEntry.SyncRibbonSelectionWithAppState`) fired
  mid-flight (e.g. from a focus change caused by the busy overlay), it would read the
  *old* `AppState.Instance.SelectedCube` and repopulate `Ribledger` from the old cube's
  (possibly single-ledger) list right after the new cube's full list had just been set.
  Fixed by moving the `AppState.Instance.SelectedCube = cube; AppState.Instance.SelectedLedger
  = ledger;` commit to before `LoadCubeLedgers` in both methods (it was already correctly
  placed before `UpdateRibbonForCube`, just not before `LoadCubeLedgers`).
  - **Re-verified** after a follow-up report worrying this might be a "we're only
    setting the caption, not actually filling the items" bug: it isn't.
    `RibbonController.SetComboItems` (`GLSense\RibbonController.cs`) genuinely rebuilds
    the AddinExpress control's real `Items` collection via reflection (`Clear()`, then
    constructs an `ADXRibbonItem` per string and `Add()`s it) - it's a completely
    separate method from `SetComboText` (which only sets the display string), and this
    reflection logic isn't Ribledger-specific (shared verbatim with `RibSegS`). Also
    re-confirmed `UpdateRibbonForCube` never touches `Ribledger` at all (only
    `RibGetCube`/`RibSegS`), and `AddinEntry.SyncRibbonSelectionWithAppState` (the
    `WorkbookActivate` handler that originally raced with this) reads
    `AppState.Instance.SelectedCube.Ledgers` - the same object the ordering fix above
    now guarantees is already the NEW cube by the time this could fire, so even a
    concurrent race would populate the correct, full ledger list. If this symptom is
    still observed after re-testing, it's almost certainly the deployment/rebuild trap
    at the bottom of this file (stale build), not a code issue - re-verify a clean
    rebuild + fresh Excel relaunch happened before assuming a new bug here.
  - **Cross-checked against FinalWorkingCode's reference implementation**
    (`BtnOK_Click`/`ProcessCubeSelectionNew`/`ProcessCubeSelectionReload`/
    `LoadCubeLedgers` in `FinalWorkingCode\GLSense\Views\GLCubeDetails.xaml.cs`), per a
    direct request to compare against it. Findings:
    - `LoadCubeLedgers` is structurally identical in both codebases, including the exact
      same `DispatcherPriority.Background` on the `Dispatcher.InvokeAsync` that populates
      the ledger list - so this isn't a priority-difference bug.
    - FinalWorkingCode commits `AppState.Instance.SelectedCube`/`SelectedLedger` only
      once, in `BtnOK_Click` itself, *after* `ProcessCubeSelectionNew`/`Reload` return
      (i.e. after `LoadCubeLedgers` already ran) - AIPowered has that exact same outer
      commit in `BtnOK_Click`'s `if (result.IsSuccess)` block. The fix above adds a
      second, *earlier* commit inside `ProcessCubeSelectionNew`/`Reload`, before
      `LoadCubeLedgers` runs - it's additive, not a deviation from the reference.
    - FinalWorkingCode has the identical `WorkbookActivate` -> `SyncRibbonSelectionWithAppState`
      mechanism (`AddinModule.cs`), also rebuilding `Ribledger.Items` from
      `AppState.Instance.SelectedCube.Ledgers` - so the same theoretical race exists
      there too. It likely hasn't manifested in FinalWorkingCode because its
      `LoadCubeLedgers` writes directly to the real `ADXRibbonComboBox` (`ribbon.Items.Add(...)`),
      while AIPowered's `SetComboItems` call has to cross an AppDomain via .NET Remoting
      first - that marshaling overhead widens the async race window, which is almost
      certainly why this symptom only ever showed up in AIPowered. Conclusion: the
      earlier-commit fix is correct and necessary for AIPowered specifically; no further
      change needed, and no change made to FinalWorkingCode (out of scope per this
      session's AIPowered-only instruction).
  - **Correction - user confirmed with screenshots this was still reproducing**: cube
    with 4 ledgers selected in AIPowered showed only the selected ledger's name as text,
    but clicking the dropdown arrow showed an EMPTY list (0 items) - not a race-condition
    symptom, a genuinely reachable bad state. Root cause found in
    `AddinEntry.SyncRibbonSelectionWithAppState()` (the `WorkbookActivate` handler, fires
    on every workbook/window activation, including - very plausibly - right after
    `GLCubeDetails.Close()` hands focus back to Excel): it called
    `ClearComboItems("Ribledger")` **unconditionally**, then only refilled via
    `SetComboItems` **conditionally** (`if (cube.Ledgers != null)`). Clear and refill
    were NOT atomic - any time the refill condition didn't hold (or `Ledgers` was an
    empty-but-non-null list, which the old `!= null` check didn't catch), Ribledger was
    left cleared with only `SetComboText` still setting the display text afterward -
    text shows a value, dropdown is empty. Same non-atomic pattern also existed for
    RibSegS right below it in the same method.
    - Fixed by making both atomic: check `cube.Ledgers?.Count ?? 0 > 0` up front: if
      there's data, do `SetComboItems` + `SetComboText` together; if not, log a warning
      and leave Ribledger/RibSegS exactly as they were instead of blanking them. Also
      added `LogDebug` of the ledger count in both `SyncRibbonSelectionWithAppState` and
      `GLCubeDetails.LoadCubeLedgers` so if this recurs, the logs pinpoint exactly which
      code path saw zero ledgers instead of requiring more guesswork.
    - Files: `GLSense.Addin.Core\AddinEntry.cs` (`SyncRibbonSelectionWithAppState`),
      `GLSense.Addin.Core\Views\GLCubeDetails.xaml.cs` (`LoadCubeLedgers` - log only, no
      behavior change there, it was already safe: throws before reaching the ribbon call
      if ledgers are missing).
    - **If this still reproduces after a clean rebuild + full Excel restart**, check the
      new debug logs first - `ledgerCount=0` logged from `SyncRibbonSelectionWithAppState`
      would mean `AppState.Instance.SelectedCube.Ledgers` is genuinely empty at that
      moment (a `CubeCache`/data-loading issue upstream, not a ribbon-population issue),
      which would need further investigation into how/when `CubeCache.AllCubes` gets its
      `Ledgers` populated per cube.
  - **Log-verified after a real rebuild + repro** (`GLSense_Logs_20-Jul-2026.log`,
    19:33:19-19:33:28): user selected cube "101 standard view imp" (4 ledgers) and the
    log confirms the fix above IS working end to end - `LoadCubeLedgers:
    cubeId=38141286, ledgerNames.Count=4` immediately followed by `RibbonController.
    SetComboItems: 'Ribledger' <- 4 item(s)`, no errors, and nothing subsequently
    cleared it before the dialog closed. `SyncRibbonSelectionWithAppState` never even
    fired in this session (no `WorkbookActivate` was raised) - so the earlier
    AppState-ordering/race theory wasn't actually exercised here, and the item
    population itself is confirmed correct at the .NET object level.
  - **New, separate symptom - user confirmed FinalWorkingCode does NOT have this, so
    it's an AIPowered-only regression, not a shared AddinExpress/Ribbon limitation**:
    whatever Ribledger shows gets cleared when switching Excel ribbon TABS (Home, Data,
    etc.) and back. Switching ribbon tabs raises no Excel COM event at all (confirmed no
    log line correlates with it), so nothing in our C# code is "reacting" to the tab
    switch - this has to be Windows Ribbon Framework re-querying the control's cached
    display state when its tab regains visibility. Since the log above proves the
    underlying `Ribledger.Items`/`Text` values ARE correct at the moment they're set,
    the most likely gap is that the blanket `_ribbon.Invalidate()` (called from
    `RibbonController.SetState`'s `InvalidateAll()`, fired right after population) is
    reliably refreshing `getEnabled`/`getText`-style state but not necessarily forcing a
    fresh `getItemCount`/`getItemLabel` pull for the dropdown's cached item list - some
    RibbonX hosts only re-pull a control's item list when THAT SPECIFIC control is
    invalidated, not on every whole-ribbon invalidate.
    - `RibbonController.Invalidate(string controlId)` (calls the AddinExpress control's
      own `.Invalidate()` via reflection) already existed but was **never called
      anywhere** in the codebase - only the blanket `InvalidateAll()` was ever used.
    - Fixed by calling `ServiceLocator.RibbonController?.Invalidate("Ribledger")`
      immediately after `SetComboItems`/`SetComboText("Ribledger", ...)` in both
      `GLCubeDetails.LoadCubeLedgers` and `AddinEntry.SyncRibbonSelectionWithAppState`,
      and the equivalent `Invalidate("RibSegS")` after `RibSegS`'s `SetComboItems` in
      `SyncRibbonSelectionWithAppState` and `GLCubeDetails.UpdateRibbonForCube` - forcing
      a per-control refresh right where the data was actually set, instead of relying
      solely on the later blanket invalidate.
    - This is a plausible, low-risk, additive fix, not a confirmed root cause - the
      tab-switch trigger itself couldn't be captured in a log (it raises no Excel event),
      so if it still reproduces after this + a clean rebuild, the next step is a fresh
      repro focused ONLY on switching tabs (nothing else) with a log capture, to see
      whether `Invalidate("Ribledger")` is even being reached by that point, or whether
      the real cause is something else entirely (e.g. AddinExpress-specific tab-
      visibility behavior needing a different hook).
    - **Verified with folder access + a real repro, and corrected**: gained direct
      Read/Grep access to the live deployment folder
      (`%LOCALAPPDATA%\ORBIT\Excel_Logs\GLSense_Logs_New\`) - confirmed both
      `GLSense.dll` (host, loaded straight from
      `AIPowered\GLSense\GLSense\bin\Debug\`, no separate deploy step - COM-registered
      directly via `RegisterForComInterop`) and `GLSense.Addin.Core.dll` (shadow-copied
      into `Versions\v11.1.0\`) were freshly rebuilt (same timestamp, same byte size as
      the project's own bin output) before this repro - ruling out the deployment trap
      for this particular test. Log again showed `LoadCubeLedgers:
      ledgerNames.Count=4` -> `SetComboItems: 'Ribledger' <- 4 item(s)` with zero
      errors - the .NET-side data is confirmed correct for the SECOND time across two
      separate sessions. User still saw literally no dropdown popup open for either
      Ribledger or RibSegS. Root cause of why the `Invalidate("Ribledger")` fix above
      had no effect: `RibbonController.Invalidate(string controlId)` was calling
      `ctrl.GetType().GetMethod("Invalidate")` - i.e. reflecting for an `Invalidate`
      method on the CONTROL object itself. That's the wrong target and unreliable
      (whatever `GetMethod` happens to resolve first isn't guaranteed to be the one
      that actually tells Excel's Ribbon engine to re-pull this control's cached
      item list) - no exception was thrown/logged, so it silently "succeeded" at
      calling *something* that did nothing useful.
    - **Real fix**: `RibbonController._ribbon` is `AddinExpress.MSO.IRibbonUI` - a
      faithful wrapper around Office's standard Ribbon extensibility `IRibbonUI` COM
      interface, which has a well-documented `InvalidateControl(string controlID)`
      member specifically for refreshing one control's cached Ribbon-side state
      (including its item list) without touching the rest of the ribbon - this is the
      correct, standard API for this, as opposed to guessing at a method on the
      control object. `RibbonController.Invalidate(string controlId)` now calls
      `_ribbon?.InvalidateControl(controlId)` instead of the old control-reflection
      approach. Existing call sites (`Invalidate("Ribledger")`/`Invalidate("RibSegS")`
      added in the previous round) are unchanged - only the implementation underneath
      them changed. Needs a fresh rebuild + repro to confirm this actually fixes the
      "no popup opens" symptom; not yet user-confirmed.
    - **`InvalidateControl` fix re-tested, still no visible change**: user rebuilt again
      (DLLs confirmed fresh, 20:08 timestamps), log again shows correct population
      (`ledgerNames.Count=4` -> `SetComboItems: 'Ribledger' <- 4 item(s)`, `RibSegS <- 5
      item(s)`, zero errors, `InvalidateControl` didn't throw) - but the dropdown
      popups still don't open for either Ribledger or RibSegS. Population is now
      confirmed correct across three separate sessions; the gap is entirely on the
      Excel-rendering side and neither invalidate approach tried so far has closed it.
    - **New diagnostic added per user's own suggestion**: RibbonX comboBox controls have
      no "dropdown opened" callback to hook (only `OnChange`, which fires on selection -
      never, if the popup won't open at all), so there's no event to attach a
      read-the-live-Items-back log to directly. Instead, wired a
      `DumpRibbonComboState()` method into the existing `RibDebug_OnClick` handler
      (`GLSense\AddinModule.cs`, host-side, direct field access - no reflection/
      remoting) that logs `Ribledger`/`RibSegS`'s live `Enabled`, `Visible`, `Text`,
      `Items.Count`, and every item's `Caption` at the exact moment Debug is
      clicked. Next repro: reproduce "dropdown won't open", then click the Debug
      ribbon button again (even if just toggling it back on/off) to capture a
      `[RibbonDiagnostic]` log line for both controls at that moment - this tells us
      definitively whether the live control still holds 4/5 items with
      Enabled/Visible=true at the moment the popup fails to open (a genuine Windows-
      Ribbon-Framework rendering mystery needing a different kind of fix entirely) or
      whether something has changed it by then (a code path we haven't found yet).
    - **First attempt produced nothing usable, and why**: user's first dump fired at the
      very start of the session (Debug turned ON before Login), so it just showed the
      expected pre-login empty state (Items.Count=0, Enabled=False - correct for that
      moment). The user then did the real repro (login, cube select, dialog close - 4/5
      items confirmed set again) and clicked Debug a second time to capture the
      post-repro dump - but that second click turned Debug OFF, and NOTHING was
      logged for it at all. Root cause: `RibDebug_OnClick` calls
      `OnRibbonAction("DebugLogsToggled", pressed)` FIRST, which flips `DebugMode` to
      false immediately; `Logger.LogDebug` (`GLSense.Shared\Logger.cs`) starts with
      `if (!DebugMode) return;`, so every `LogDebug` call made afterward in the same
      click - including the entire diagnostic dump - silently vanished. `LogWarn`/
      `LogError` have no such check. Fixed by switching `DumpRibbonComboState()`'s two
      log lines from `LogDebug` to `LogWarn`, so the dump always gets written
      regardless of which direction the Debug toggle is going.
    - **Next repro**: reproduce the bug, then click Debug once (either direction) -
      the `[RibbonDiagnostic]` lines will now show up in the log no matter what.
  - **ACTUAL ROOT CAUSE - found by the user directly stepping through
    `RibbonController.SetComboItems` in a debugger**, correcting a misdiagnosis that
    persisted across every round above: `captionProp` (`itemParamType.GetProperty
    ("Caption")`) was coming back **null**, so the method hit `if (captionProp == null)
    return;` and returned WITHOUT adding a single item to the real collection, every
    single time, in every session logged above.
    - Why the logs never caught this: `_logger?.LogDebug($"RibbonController.
      SetComboItems: '{controlName}' <- {items?.Count() ?? 0} item(s)")` runs at the
      very TOP of the method, before `itemsCollection`/`addMethod`/`itemParamType`/
      `captionProp` are even resolved - it only reports the size of the INPUT list
      (`ledgerNames`/segment names), never whether anything was actually written to the
      control. Every "SetComboItems: 'Ribledger' <- 4 item(s)" log line in every session
      this whole saga was read as "4 items were successfully set" - it actually only
      ever meant "4 items were passed in as the argument." This is why population kept
      looking "confirmed correct" from the logs while the live control was empty the
      entire time - a real gap in my own reasoning, not something the user missed.
    - Why `captionProp` was null: the method resolves which "Add" overload to use by
      filtering `itemsType.GetMethods()` for single-parameter methods named "Add", then
      picking whichever ISN'T `Add(object)` (to avoid the `IList.Add(object)` overload).
      That's insufficient disambiguation - the Items collection apparently also exposes
      an `Add(string)` convenience overload, which also isn't `Add(object)` and can tie
      with (or win over, depending on `GetMethods()`'s enumeration order) the real
      `Add(ADXRibbonItem)` overload. If `Add(string)` got picked, `itemParamType` became
      `System.String`, which has no `Caption` property at all - hence null, hence the
      silent early return, hence zero items ever actually reaching the control despite
      `SetComboItems` never throwing or logging an error.
    - **Fix**: rewrote the overload resolution to pick by capability instead of by
      exclusion - iterate every single-parameter `Add` overload and use the first one
      whose parameter type has a writable `Caption` property. That's unambiguous: only
      the real `ADXRibbonItem`-shaped overload qualifies, regardless of how many other
      `Add` overloads exist or what order reflection returns them in. Also added a
      second, genuine "actually populated" log line AFTER the real `Add` loop runs
      (reading the live collection's `Count` back), so a future log can distinguish
      "N items were passed in" from "N items are actually in the control" - closing the
      exact gap that caused the repeated misdiagnosis.
    - This is the real fix for the "shows 1 item / dropdown never opens / items log
      as 4 but nothing shows" symptom across this entire saga - not the AppState-commit
      ordering, not `InvalidateAll`/`InvalidateControl`, both of which were real,
      legitimate hardening but were never going to fix an Items collection that was
      never actually being written to in the first place. Needs a rebuild + repro to
      get final user confirmation, but this is a concrete, verified (by the user's own
      debugger session, not by log inference) bug with a direct fix.
    - **First attempt at this fix (capability-based: "does the candidate's parameter
      type have a writable Caption property") was ALSO wrong** - confirmed by the new
      error log actually firing: `SetComboItems: could not resolve an Add(x) overload
      with a writable Caption property for 'Ribledger' - no items were set.` Every
      candidate's `GetProperty("Caption")` came back null - meaning the real
      `Add(ADXRibbonItem)`-shaped overload's declared parameter type doesn't expose
      "Caption" via a direct `GetProperty` call the way expected (most likely the Add
      overload's parameter is typed as a base/interface type that doesn't itself
      declare `Caption` - only the concrete `ADXRibbonItem` does).
    - **Real fix**: stopped reflecting for the item type entirely. `RibbonController.cs`
      already has `using AddinExpress.MSO;` at the top (the "we can't reference
      AddinExpress types by name" premise in the original code's doc-comment was
      simply wrong for this file) - so `new ADXRibbonItem { Caption = text }` is
      constructed directly, strongly typed, zero ambiguity. Reflection is now used only
      to find which "Add" overload can ACCEPT an `ADXRibbonItem` - via
      `parameters[0].ParameterType.IsAssignableFrom(typeof(ADXRibbonItem))`, which
      correctly matches whether the parameter type is `ADXRibbonItem` itself or any
      base/interface type of it (unlike checking for a property by name, which only
      works if that exact type declares it directly). This is the third and (hopefully)
      final iteration of this specific bug - not yet rebuilt/retested.
    - **CONFIRMED FIXED - user-verified via the RibbonDiagnostic dump, both symptoms at
      once**: two dumps taken 12 seconds apart (before/after switching Excel ribbon
      tabs) both show `Ribledger: Text='IAS Reporting Vision Ops', Items.Count=4,
      Items=[IAS Reporting Vision Ops, Vision Belgium, Vision Operations (USA), Vision
      UK]` - identical in both. Dropdown popups now show their items, AND the
      tab-switch-clears-text symptom is gone.
    - **Why one fix closed both symptoms**: they were never two separate bugs. Windows
      Ribbon Framework re-validates a comboBox's displayed text against its Items list
      whenever it re-renders the control (e.g. on tab-visibility change) - with Items
      genuinely empty (the `Add`-overload resolution bug), there was nothing for the
      displayed text to match against, so any re-render blanked it. Once Items are
      actually populated, the match succeeds and the text persists correctly across
      tab switches with no other change needed.
    - **STATUS: RESOLVED.** The AppState-commit-ordering fix, `InvalidateAll`, and the
      later `InvalidateControl` calls earlier in this section were all real, reasonable
      hardening at the time, but none of them were the actual fix - this
      `SetComboItems` Add-overload bug (found by the user directly stepping through
      `RibbonController.cs` in a debugger) was the sole root cause of the entire
      "Ribledger/RibSegS dropdown empty / no popup / text clears on tab switch" saga.
      Left the earlier hardening in place (harmless, no evidence it caused any
      problem) rather than churning further without cause.

- **Ribledger ribbon dropdown font/rendering looked worse than FinalWorkingCode's**: the
  user confirmed (via screenshots) that FinalWorkingCode's item population/selection
  behavior for the Ledger ribbon combo is exactly what's expected (populates all N
  ledgers for the selected cube, pre-selects the chosen one) and that the AppState/
  `SetComboItems` fix above already reproduces that correctly - but the combo's dropdown
  list font/visual polish was noticeably worse in AIPowered. Compared
  `Ribledger`/`RibSegS`'s `ADXRibbonComboBox` designer declarations between
  `FinalWorkingCode\GLSense\AddinModule.Designer.cs` and
  `AIPowered\GLSense\GLSense\AddinModule.Designer.cs` - identical in both (same
  `SizeString`, no explicit `Font` set on either). The real difference:
  `FinalWorkingCode\GLSense\GLSense.csproj` wires
  `<ApplicationManifest>GLSense.app.manifest</ApplicationManifest>` (that manifest
  declares a `Microsoft.Windows.Common-Controls` v6.0.0.0 dependency + `dpiAware=true`);
  AIPowered's `GLSense.csproj` had no `ApplicationManifest` at all and no manifest file
  existed in the host project. Without the ComCtl6 dependency, native ribbon/owner-drawn
  combo popups fall back to unthemed classic rendering (smaller default font, flatter
  look) - fixed by adding `AIPowered\GLSense\GLSense\GLSense.app.manifest` (identical
  content to FinalWorkingCode's) and the matching `<ApplicationManifest>` property in
  `GLSense.csproj`. Requires a full rebuild + Excel restart to take effect (manifest is
  embedded as a native resource at build time).

- **Logging hardening pass, directly motivated by how long the Ribledger saga above
  took to pin down**: every reflection-based method in `GLSense\RibbonController.cs`
  used to silently no-op on failure - a missing control, a missing/non-writable
  property, or (the actual root cause of the whole saga) a reflection resolution that
  quietly picked the wrong thing - with zero trace in the log. All of that only showed
  up because the user had source access and a debugger; at a client site, with neither,
  this class of bug would be unfindable. Added `LogWarn` (not `LogDebug` - **always**
  written to the log file regardless of whether Debug mode is toggled on) at every
  point where one of these methods used to fail silently:
  - `GetRibbonControl` (the single choke point every other method below goes through):
    warns once when a control name can't be resolved at all.
  - `SetControlEnabled`/`SetControlVisible`/`SetControlPressed`/`SetControlLabel`: warn
    when the expected property (`Enabled`/`Visible`/`Pressed`+`Checked`/`Caption`)
    isn't found or isn't writable.
  - `SetComboItems`: warns if the `Items` property can't be read, AND (this is the
    important one) does a **self-verifying expected-vs-actual count check** after
    populating - reads the live collection's `Count` back and warns if it doesn't match
    the input list's size, instead of just logging "N items were passed in" and
    assuming that means they were set (exactly the misleading signal that dragged out
    the diagnosis this time).
  - `SetComboText`: warns if `Text` isn't writable, AND reads the value back
    immediately after setting it, warning on any mismatch - catches the control
    silently rejecting/altering a value right where it happens.
  - `ClearComboItems`: warns if `Items` can't be read.
  - Deliberately left `LogDebug` status lines (e.g. "actually populated with N item(s)")
    as debug-only on the success path, so normal operation at a client site doesn't get
    noisy - only genuine anomalies are always-on.
  - Goal: if anything in this family of bugs recurs at a client site with no source
    access, the log alone should point at the exact failing control and property
    instead of requiring a live debugging session to find, as happened here.

## 4. Visual/theme consistency

- **FontAwesome migration**: replaced WPF-UI's `ui:SymbolIcon`/`Symbol=` usage
  project-wide with `MahApps.Metro.IconPacks.FontAwesome`
  (`iconPacks:PackIconFontAwesome`/`Kind=`), since it was already a proven dependency
  (used by FinalWorkingCode) and avoided a class of icon-rendering issues. Every
  `Kind=`/`IconSymbol=` value used was verified to actually exist in the installed DLL.
  Touches `Themes/GlobalStyles.xaml`, `Themes/Generic.xaml`, `Views/BaseWindow.cs`
  (`IconSymbol` default), and ~22 window `.xaml` files.
- **CheckBox `Storyboard.Seek` warning**: WPF's default CheckBox template's
  VisualStateManager/Storyboard-driven "CheckStates" transitions caused a
  `Storyboard.Seek "never applied to this object"` warning on bound CheckBoxes. Fixed
  via a shared `SafeCheckBoxTemplate` (plain property-`Trigger`-based, no VSM/Storyboard)
  in `Themes/GlobalStyles.xaml`, applied through `ModernCheckBox`/`CompactCheckBox` and
  a new implicit (`x:Key`-less) `CheckBox` style so it covers every unstyled CheckBox
  too.
- **Button hover chrome**: WPF's default Button chrome layers its own theme
  hover/pressed overlay on top of a programmatically-toggled `Background`, washing out
  manually-computed hover colors (GLMessageWindow's OK button rendered blue at rest,
  near-transparent on hover - the opposite of the correctly-behaving Close button).
  Fixed via a shared `PlainButtonTemplate` (plain `Border` bound to
  `TemplateBinding Background`, no chrome), applied in `GLMessageWindow.xaml.cs`'s
  `AddDialogButton`.
- **Field-label color standardization**: every field-prompt label is now
  `PrimaryBrush` (blue) instead of a mix of blue (`HeaderTextBlock`/`LabelTextStyle`)
  and near-black (`TextPrimaryBrush`, used by plain `<Label>`/`ConfigLabelStyle`).
  Touched `GlobalStyles.xaml`'s `LabelTextStyle`, `GLBalanceConfigurator.xaml`'s local
  `ConfigLabelStyle`, and per-instance `Foreground=` overrides across ~10 window files.
- `BaseWindow.IconSymbol` default changed from `"Key24"` (WPF-UI naming, no longer
  valid) to `"KeySolid"` (FontAwesome naming).
- **`SuggestAppendComboBox` font size**: this custom control's own `Style`
  (`Themes\Generic.xaml`, `Style TargetType="{x:Type ctrls:SuggestAppendComboBox}"`)
  never set `FontSize` anywhere - not on the style itself, not on `PART_TextBox`, not
  on the dropdown's `ListBoxItem` container style - so every window's combo box AND
  its popup item list rendered at whatever `FontSize` happened to be ambiently
  inherited, inconsistent across windows and larger than intended. Added
  `<Setter Property="FontSize" Value="12"/>` to the control-level `Style` (12 matches
  this codebase's existing "compact control" convention, e.g. `ModernButton` in
  `GlobalStyles.xaml` - a size step down from the `13` used for primary input text
  like `ModernTextBox`/labels). Being inherited, this one change flows into both the
  editable text box and (via the `Popup`'s logical-tree inheritance) the dropdown
  `ListBox`/`ListBoxItem`s, fixing all ~16 windows that use this control without
  touching any of them individually.
  - **Correction**: that "flows into the dropdown via the Popup's logical-tree
    inheritance" claim was wrong. WPF `Popup` content is a separate visual tree root
    and does NOT participate in property-value inheritance from its logical parent -
    so the control-level `FontSize="12"` Setter was reaching `PART_TextBox` (inherits
    normally, same template) but NOT `PART_ListBox` inside `PART_Popup`, which kept
    rendering at the ambient/system default. User caught this directly: "font size is
    set only for the TextBox... the popup ListBox... should match." Fixed by adding
    `FontSize="{TemplateBinding FontSize}"` directly on `PART_ListBox` in
    `Themes\Generic.xaml` - `TemplateBinding` reads straight from the templated
    parent regardless of Popup boundaries, so this bypasses the broken inheritance
    chain. `ListBoxItem`s then inherit correctly from their own immediate parent
    (`PART_ListBox`), which is now set right, so no separate fix needed on the
    `ListBoxItem` style. Not yet rebuilt/retested.

## 5. Other BaseWindow fixes

- `FitToAvailableWorkArea()` used to unconditionally overwrite
  `this.MaxWidth`/`MaxHeight` with the screen work-area size, stomping a window's own
  tighter XAML-declared `MaxWidth`/`MaxHeight` (e.g. GLCubeDetails expanding way beyond
  the screen when a value was selected). Fixed to `Math.Min(this.MaxWidth,
  availableWidth)` / `Math.Min(this.MaxHeight, availableHeight)`.
- Escape-to-close wired for every `BaseWindow`-derived dialog (opt-out via
  `EnableEscapeToClose`, used by `GLWaitWindow` since Cancel should be the only way out)
  - bound on bubbling `KeyDown` (not `PreviewKeyDown`) and gated on `!e.Handled` so a
    DataGridCell/ComboBox that already consumed Escape for its own purpose isn't
    double-handled.

## 6. Structural additions

- `GLSegmentRef` ported from FinalWorkingCode to AIPowered.
- `GLSegmentManager` built as a master-detail redesign of `GLSegmentRef`.
- `GLServerConfiguration`'s status/toast UX reworked (AIPowered only).

## 6.1 Sizing/collapse parity pass across 15 other windows

User asked to bring the same "FinalWorkingCode always renders perfectly" sizing/scroll
behavior established on GLSegmentValues to a batch of other windows: AttachmentsDialog,
GLAbout, GLDailyRates, GLDrilldownCustomization, GLGetPeriod(+ByDate/ByYear/Details/
StartEnd), GLJobsMonitor, GLLOVs, GLRollerGroups, GLSegmentDiscovery,
GLSegmentFunctions, GLUserConfig. Per-file outcome:

- **GLDailyRates, GLGetPeriod, GLGetPeriodByDate, GLGetPeriodByYear,
  GLGetPeriodDetails, GLGetPeriodStartEnd, GLSegmentDiscovery, GLSegmentFunctions**: no
  changes needed - these are the "8 static windows" already fixed in an earlier pass
  (`SizeToContent="Width"` + fixed Min/MaxHeight + Auto-row-holding-a-ScrollViewer
  pattern), confirmed still in place.
- **GLAbout**: brought fully in line with FinalWorkingCode - added the footer's own
  Close button (previously only the title-bar Close existed) and aligned the instances
  DataGrid's `MinHeight` to FinalWorkingCode's value.
- **GLRollerGroups**: wrapped the Segment/Search fields section in a `ScrollViewer`
  (`VerticalScrollBarVisibility="Auto"`) so those fields stay reachable instead of
  collapsing when the window can't fully fit on screen.
- **GLUserConfig**: wrapped the whole tab-control content area in a `ScrollViewer`,
  matching FinalWorkingCode, for the same DPI/resolution reachability reason.
- **AttachmentsDialog**: no changes needed - already uses a deliberate
  `ScrollViewer MinHeight="220"` pattern from an earlier fix pass for its
  variable-length attachment list; reasoning matches the current pattern exactly.
- **GLDrilldownCustomization**: no changes needed - this window is just a full-bleed
  WebView2 host with fixed, narrow Min/Max bounds; no field rows or DataGrids that can
  collapse.
- **GLJobsMonitor**: added `MinHeight="300"` to the Jobs DataGrid's Border (`Grid.Row="2"`,
  a `"*"` row) - same collapse-during-SizeToContent-measurement risk as
  GLSegmentValues' old Dual DataGrids row, now guarded the same way.
- **GLLOVs**: added `MinHeight="250"` to the LOVs DataGrid's Border (`Grid.Row="1"`,
  a `"*"` row) for the identical reason.
- **Status**: implemented, not yet rebuilt/tested by the user.

### 6.2 GLLOVs: missing Comments watermark

User screenshot showed the Comments `TextBox` completely blank when empty - no hint
text at all, unlike FinalWorkingCode which shows italic placeholder text explaining
what the field does and what happens if it's left blank. AIPowered's `TxtComments` was
just plain `Style="{StaticResource ModernTextBox}"` with no watermark mechanism (plain
`TextBox` has no built-in placeholder support). Ported FinalWorkingCode's
`WatermarkTextBox` style (window-scoped resource in `GLLOVs.xaml`, `BasedOn`
`ModernTextBox`) - it overlays a `TextBlock` in the `ControlTemplate` next to
`PART_ContentHost`, toggled visible via triggers on `Text=""`/`Text={x:Null}`/
`IsFocused`. Applied it to `TxtComments` in place of `ModernTextBox`.
- **Status**: implemented, not yet rebuilt/tested by the user.

## 7. GLSegmentValues.xaml (Segment Values Configurator) layout fixes

User supplied a screenshot with color-coded annotations; cross-checked against
FinalWorkingCode's `Views/GLSegmentValues.xaml` for each fix.

- **Left DataGrid unwanted horizontal scrollbar**: `Description` column was
  `Width="Auto"`, which sizes to longest content and forces horizontal scroll once
  content is wider than the grid. Changed to `Width="*"` so Description absorbs all
  remaining space instead (matches FinalWorkingCode). Also widened `Is-Summary` column
  `90->100` and `Value` column `100->90` so both display fully without truncation.
  Added `ToolTip` Setters (bound to `DisplaySegmentValue` / `Description`) on the
  `ElementStyle` of both columns so hovering a cell shows that row's full value -
  FinalWorkingCode already does this per-cell hover tooltip.
- **Paging footer cramped / oversized**: reduced footer `TextBlock` font size from the
  inherited `HeaderTextBlock` 13px down to an explicit `FontSize="11"` (2px down, as
  requested). Added `iconPacks:PackIconFontAwesome` icons (`FileLinesSolid` before "Per
  Page:", `DatabaseSolid` before "Total:") matching FinalWorkingCode. Wrapped "Page X of
  Y" in a bordered/padded box (`BackgroundBrush`, `CornerRadius="4"`, `Padding="8,4"`).
  Increased margins between footer groups so elements aren't crowded on top of each
  other or spaced too far apart.
- **Cell Reference control (`excelRefEdit`) nudge**: added `Margin="2,0,0,0"` to shift
  it 2px right per the annotation.
- **SizeToContent / high-DPI scrollability**: FinalWorkingCode's version of this window
  uses `SizeToContent="Manual"` with fixed `Width`/`Height`/Min/Max - it never needed
  any of AIPowered's `BaseWindow` resettle machinery (those hooks are gated on
  `SizeToContent != Manual`, so they simply don't run for Manual windows). Switched
  AIPowered's `GLSegmentValues` from `SizeToContent="WidthAndHeight"` to `Manual` with
  an explicit starting `Width="740" Height="700"` (within the existing Min/Max bounds).
  Additionally wrapped the whole "Middle Content" `Grid.Row="1"` section in
  `<ScrollViewer VerticalScrollBarVisibility="Auto" HorizontalScrollBarVisibility="Disabled"
  CanContentScroll="False">`, and added `MaxHeight="450"` to the Dual DataGrids border -
  this is FinalWorkingCode's exact pattern, so that content otherwise clipped at 150%+
  DPI scaling or lower screen resolutions becomes reachable via scrolling instead of
  hidden off-window.
- **Status**: implemented, not yet rebuilt/tested by the user.

### 7.0b User-supplied screenshots exposed 3 real regressions - see 1.4f, 7.1b, and section 3.x below

The user copied rendered PNG screenshots of several windows into the logs folder for
direct visual comparison. These exposed genuine bugs that code review alone had missed:
the "8 static windows" dead-gap regression (now section **1.4f**), GLSegmentValues'
Is-Summary column vanishing (below, folded into 7.1b), and GLLOVs' missing Comments
watermark (section **3.6**). Screenshots-in-a-shared-folder turned out to be a very
effective way to catch layout bugs that are easy to reason past when just reading XAML.

### 7.1 Follow-up round: Is-Summary header clipping + unlabeled page-range

- **Is-Summary header text clipped**: root cause was that AIPowered's
  `DataGridColumnHeader` style had no custom `Template`, so it fell back to the default
  WPF header template, which reserves extra internal space for the sort-arrow glyph
  even when the column isn't sorted - clipping "Is-Summary" at `Width="100"`.
  FinalWorkingCode's header style defines a custom `ControlTemplate` that puts the sort
  arrow in its own `Auto` column so the `ContentPresenter` gets the full remaining
  width. Ported that exact template to both `dgLeft`'s and `dgRight`'s
  `ColumnHeaderStyle` in AIPowered.
- **"1 - 24" page-range text meaningless without context**: `PageRangeText` (bound at
  the bottom-left of the paging footer) was a bare `"{start} - {end}"` string with no
  label, so it wasn't obvious what it represented. Added a "Showing:" label before it
  and a tooltip ("Record range currently displayed in the left grid"). Nudged the
  "Total:" group's left margin `24->32` to keep both groups from crowding each other
  now that "Showing:" adds width.
- **Status**: implemented, not yet rebuilt/tested by the user.

### 7.1b Is-Summary column vanishing entirely (not just clipped) - `DataGridColumnFillHelper` never re-runs after data loads

A user screenshot after the 7.1 fix still showed the left grid with only "Value" and
"Description" visible - no "Is-Summary" header text, no checkboxes, just blank white
space where the column should be (not merely truncated text this time - the whole
column was effectively gone from the visible area).

Root cause: `GLSegmentValues.xaml.cs`'s constructor wires
`DataGridColumnFillHelper.EnableFillColumn(dgLeft, dgLeft.Columns[1])` (Description is
the designated "fill" column - see `Utilities/DataGridColumnFillHelper.cs`), which only
hooks the DataGrid's own `Loaded` and `SizeChanged` events to (re)compute Description's
width as `grid.ActualWidth - (every other visible column's ActualWidth) - scrollbar
allowance`. `Loaded` fires once, early - typically before the user has selected a
segment/hierarchy and before `PagedSegmentValues` has any real rows. Populating
`PagedSegmentValues` later (after the user picks a segment) changes row content but
does **not** raise the grid's own `SizeChanged` event, so the fill column's width, once
computed against that initial near-empty state, never gets recalculated for the real
data - it can end up computed too wide, pushing the fixed-width `Is-Summary` column
(third/rightmost column in `dgLeft`) out past the grid's visible/clipped viewport
entirely. The helper's own doc comment explicitly calls out this exact scenario
("Safe to call manually after code populates or replaces a grid's ItemsSource at
runtime") - that follow-up call was simply never added when `GLSegmentValues` was
built.

**Fix**: `SegmentSelectorViewModel`'s `DataLoadedAction` callback (already wired in the
constructor to fire `ForceSizeToContentResettle()`/`PumpDispatcherFrame()` once real
data has loaded - see CLAUDE.md 1.4b) now also calls
`DataGridColumnFillHelper.Refresh(dgLeft, dgLeft.Columns[1])` and
`DataGridColumnFillHelper.Refresh(dgRight, dgRight.Columns[2])` first, so the fill
column width is recomputed against the real, final data instead of the stale
near-empty-grid measurement. Also added `ScrollViewer.HorizontalScrollBarVisibility="Auto"`
explicitly to `dgLeft` as a defensive fallback, so if this class of bug recurs for any
other reason the overflow content becomes reachable via scrollbar instead of silently
vanishing.
- **Status**: implemented, not yet rebuilt/tested by the user.

## 8. Second round of user-tested screenshot feedback (post section 7)

### 8.1 GLSegmentValues: Is-Summary STILL broken after 7.1b - real fix was to stop using DataGridColumnFillHelper here at all

7.1b's "call Refresh again in DataLoadedAction" patch was not enough - user reported the
column now shows at ~40% width on open and vanishes/shifts right after scrolling. The
actual insight: `DataGridColumnFillHelper` exists specifically to work around
`DataGridColumn Width="*"` reporting a huge desired width when measured with an
**infinite available width** - which only happens for `SizeToContent="WidthAndHeight"`
windows. `GLSegmentValues` is `SizeToContent="Manual"` with a fixed, explicit `Width`
(see section 7's SizeToContent fix) - it is **never** measured with infinite available
width, so the bug this helper works around cannot occur here in the first place. Using
the helper anyway introduced a *different*, timing-fragile bug (one-shot
Loaded/SizeChanged-driven width calculation going stale relative to virtualized/
scrollbar-affected real layout).

**Fix**: removed `DataGridColumnFillHelper.EnableFillColumn`/`.Refresh` entirely for
`dgLeft`/`dgRight` in `GLSegmentValues.xaml.cs`; changed `dgRight`'s "Segment" column
from `Width="Auto"` to native `Width="*"` (Description on `dgLeft` was already `"*"`).
Native WPF star-width columns now handle this correctly and robustly across scroll/
resize/data-reload, since the window's own width is fixed from the very first measure
pass. Also added `ScrollViewer.HorizontalScrollBarVisibility="Auto"` to `dgLeft` as a
defensive fallback.
- **Status**: implemented, not yet rebuilt/tested by the user.

### 8.2 GLRollerGroups / GLLOVs: perfect when empty, grow too tall once populated

Both size correctly with no data but sprawl downward once their DataGrid has real rows
- risky on smaller/lower-resolution screens. Added a `MaxHeight` cap to each window's
DataGrid-hosting `Border` (matching `GLSegmentValues`' Dual DataGrids `MaxHeight="450"`
pattern already established in section 7): `GLRollerGroups.xaml` -> `MaxHeight="450"`,
`GLLOVs.xaml` -> `MaxHeight="400"` (kept its existing `MinHeight="250"`). The DataGrid's
own internal `ScrollViewer.VerticalScrollBarVisibility="Auto"` takes over once content
exceeds the cap, instead of the whole window growing unbounded.
- **Status**: implemented, not yet rebuilt/tested by the user.

### 8.3 GLSegmentFunctions: window sized for the Next/Previous layout regardless of which function is open

User: opening this window for any function OTHER than Next/Previous Segment leaves a
gap before the footer, as if it were sized for the Next/Previous case. Root cause:
`NEXTSEGMENT`/`PREVIOUSSEGMENT` are the only two function types that leave
`ParentChildSection` visible (every other function type - `ENABLEDFLAG`, `SUMMARYFLAG`,
`DESCRIPTION`, `DFF` - collapses it in the constructor's `switch`), so those two show
one more row-border than all the others. The window's `MinHeight="440"` had been
calibrated to comfortably fit that *tallest* (Next/Previous) case - which is exactly
backwards for a `MinHeight`: it become a **floor** that the shorter, far more common
"other options" content can't shrink below, leaving dead space before the footer.
FinalWorkingCode's equivalent window uses `MinHeight="375"` - tuned to the *shortest*
content case instead, letting `SizeToContent` (there, `"Height"`; here,
`"WidthAndHeight"`) grow up to `MaxHeight` for the taller case naturally in both
directions.

**Fix**: lowered `MinHeight` from `440` to `375` (matching FinalWorkingCode exactly);
`MaxHeight="520"` and width bounds left unchanged.
- **Status**: implemented, not yet rebuilt/tested by the user.

### 8.4 Period functions family (GLGetPeriod/ByDate/ByYear/Details/StartEnd) confirmed perfect

No action needed - explicitly confirmed correct by the user after the section 1.4f
revert. Noted here so this doesn't get re-touched without a new, specific report.

### 8.5 DatePicker calendar icon too small/barely visible (GLGetPeriodByDate, GLDailyRates, GLBalanceConfigurator)

Neither codebase has a shared `DatePicker` style - all 3 windows only ever overrode
`DatePickerTextBox`'s template (replacing its internal chrome with a plain `TextBox`),
leaving the separate calendar toggle **button** to WPF's stock Aero2/Fluent-theme
template, which draws a small vector glyph that's hard to see against this app's other,
consistently larger `iconPacks` glyphs used everywhere else in the UI.

**Fix**: added a new shared `ModernDatePicker` style/`ControlTemplate` to
`Themes/GlobalStyles.xaml` - a bordered `Grid` with an explicit `PART_TextBox`
(`DatePickerTextBox`) column and a dedicated 30px `PART_Button` column showing
`iconPacks:PackIconFontAwesome Kind="CalendarDaysSolid"` at a fixed, clearly visible
16x16 size (the same icon already used for this app's period/date windows elsewhere).
`PART_Popup` is left bare - `DatePicker` itself creates and assigns the `Calendar` into
it at runtime, no `Calendar` element needs to be declared. Applied
`Style="{StaticResource ModernDatePicker}"` to `dtpDate` in `GLGetPeriodByDate.xaml`/
`GLDailyRates.xaml` and to `dtpStartDate`/`dtpEndDate` in `GLBalanceConfigurator.xaml`,
removing the now-redundant local `BorderThickness`/`BorderBrush`/`Background` on each
(the shared style's `Border` supplies those). Each window's existing per-window
`DatePickerTextBox` override (which customizes the inner text/binding) is untouched and
still applies, since it targets the `DatePickerTextBox` *type*, which is still present
as `PART_TextBox` inside the new outer template.
- **Status**: implemented, not yet rebuilt/tested by the user.

### 8.6 GLAbout: missing company logo

`Views/GLAbout.xaml.cs`'s own header comment had documented "no Images folder exists in
GLSense.Addin.Core" as the reason the logo was replaced with a generic icon badge back
when this window was first ported - but the actual `orbit_logo.png` asset does exist in
FinalWorkingCode (`GLSense\Images\orbit_logo.png`), it just was never copied over.

**Fix**: copied the real `orbit_logo.png` into
`GLSense.Addin.Core\Images\orbit_logo.png`, added a `<Resource Include="Images\orbit_logo.png" />`
item to `GLSense.Addin.Core.csproj` so it's embedded, and changed `GLAbout.xaml`'s
product-badge `Border` from a `PackIconFontAwesome` placeholder to
`<Image Source="/GLSense.Addin.Core;component/Images/orbit_logo.png" Stretch="Uniform"/>`,
matching FinalWorkingCode's `Border Width="200" Height="100"` + `Image` structure.
- **Status**: implemented, not yet rebuilt/tested by the user.

### 8.7 GLBalanceConfigurator: Start/End Date rows not displaying properly

Beyond the shared calendar-icon fix (8.5, applied here too), this window's own
`DatePickerTextBox` override was missing `VerticalAlignment`/`VerticalContentAlignment`/
`Padding` on its inner `TextBox` (present on the GLGetPeriodByDate/GLDailyRates
equivalents but not here) - at this window's compact `Height="28"` (vs the 32px used
elsewhere), the date text likely sat top-aligned/cramped instead of vertically centered.
Added `VerticalAlignment="Stretch"`, `VerticalContentAlignment="Center"`, and
`Padding="4,0,0,0"` to both `dtpStartDate`'s and `dtpEndDate`'s inner `TextBox`, and
changed `BorderBrush="White"` to `BorderThickness="0" Background="Transparent"` since
the new shared `ModernDatePicker` style's own outer `Border` now supplies the visible
border/background - the old inner white border was redundant.
- **Status**: implemented, not yet rebuilt/tested by the user.

---

## 9. Third round of user-tested screenshot feedback (post section 8)

User confirmed almost all UI issues from section 8 were fixed, and reported 2 remaining items with fresh screenshots.

### 9.1 DatePicker inner textbox border still visible (GLDailyRates, GLGetPeriodByDate)

`Logs/GLDailyRates.png` showed a red-circled small bordered box around the "Currency
Date :" field's text area, distinct from `ModernDatePicker`'s own outer border (added in
8.5). Root cause: when 8.5's shared `ModernDatePicker` style was added, only
`GLBalanceConfigurator.xaml`'s per-window `DatePickerTextBox` override was updated to
drop its own border (see 8.7) - `GLDailyRates.xaml` and `GLGetPeriodByDate.xaml` still had
their original `BorderThickness="1" BorderBrush="{StaticResource SoftControlBorderBrush}"
Background="White"` on the inner `PART_TextBox`, which now rendered as a nested "double
border" once the new outer control border was added around the whole DatePicker.

**Fix**: changed both files' inner `PART_TextBox` to `BorderThickness="0"
Background="Transparent"`, matching the GLBalanceConfigurator pattern from 8.7, so only
`ModernDatePicker`'s outer border is visible.
- **Status**: implemented, not yet rebuilt/tested by the user.

### 9.2 GLSegmentValues: scrolling the grid also scrolled away the Segment/Search rows

User: "we are placing windows contents inside the scroll viewer hence when i scrolled the
segments and next row items are also scrolled up and not visible bring the top 2 rows out
the scroll viewer. So scroll viewer should be only for the DGV's row."

Section 7's original layout wrapped the entire Middle Content (Segment & Hierarchy row +
Search row + Dual DataGrids row) in one outer `ScrollViewer`, matching FinalWorkingCode's
literal structure. That meant scrolling the DataGrid content also scrolled the
Segment/Hierarchy pickers and Search box out of view, since they were siblings inside the
same scrollable area.

**Fix**: removed the outer `ScrollViewer` entirely from `GLSegmentValues.xaml`. The
Segment & Hierarchy Border and Search Border are now direct, non-scrolling children of the
outer `Grid Grid.Row="1"` (their own `Auto` rows), and the Dual DataGrids Border (the "DGV's
row") is the only vertically-growing (`*`) row. It doesn't need its own outer ScrollViewer
either - it already has a `MaxHeight="450"` cap plus each DataGrid's own internal
`ScrollViewer.VerticalScrollBarVisibility="Auto"`, so the grids scroll internally while the
selector/search rows above stay fixed and always visible.
- **Status**: implemented, not yet rebuilt/tested by the user.

---

## 10. GLLOVs: comments-clear resize glitch + Items Count tooltip

### 10.1 Window width jumped the instant comments text was cleared

User: "When the window is launched it is resized perfectly. When i try to enter comments
by clearing then it is immediately resizing its width based on comments text."

Root cause: `GLLOVs.xaml` is `SizeToContent="WidthAndHeight"`, and (per section 1's
"infinite-measure" root cause) WPF re-measures Width using `PositiveInfinity` every time
content invalidates layout - not just once at startup. The `WatermarkTextBox` style's
`WatermarkText` TextBlock (added earlier to show placeholder guidance in the Comments box)
toggles from `Collapsed` to `Visible` the moment the user clears the text. That TextBlock
has `TextWrapping="Wrap"` but no width constraint, so on the infinite-width measure pass it
reported its full unwrapped single-line desired width (the whole sentence on one line),
briefly ballooning the window width toward `MaxWidth` right as the user cleared the field.

**Fix**: added `MaxWidth="480"` to the `WatermarkText` TextBlock in the `WatermarkTextBox`
style's `ControlTemplate` (`GLLOVs.xaml`). A hard `MaxWidth` forces wrapping/a stable
desired size regardless of available measure space, so showing/hiding the watermark no
longer changes the window's width.
- **Status**: implemented, not yet rebuilt/tested by the user.

### 10.2 Items Count column had no tooltip

User: DataGrid has "Available LOVs" and "Items Count" columns; the first has a tooltip,
the second didn't - it should say how many items are available for that row's LOV.

**Fix**: added a `ToolTip` to the "Items Count" column's `ElementStyle`, using a
`MultiBinding` with `StringFormat="{}{0} item(s) available for '{1}'"` over `ItemsCount`
and `Name`, matching the existing tooltip pattern already used on the "Available LOVs"
column.
- **Status**: implemented, not yet rebuilt/tested by the user.

---

## 11. AddinModule.CountFormulaCells logging a normal condition as an exception

Log excerpt:
```
Context: AddinModule.CountFormulaCells
Type: System.Runtime.InteropServices.COMException
Message: No cells were found.
   at ... Microsoft.Office.Interop.Excel.Range.SpecialCells(XlCellType Type, Object Value)
   at GLSense.AddinModule.CountFormulaCells(Worksheet sheet)
```

Excel's `Range.SpecialCells(xlCellTypeFormulas)` throws `COMException` with HRESULT
`0x800A03EC` and message "No cells were found" whenever the target range has zero cells of
the requested type - e.g. a worksheet with no formulas anywhere. This is normal, expected
behavior from that Excel API, not an actual error condition (`CommonFunctions.cs` already
handles the identical case correctly in three other places -
`BalanceFormulaExists`/`GetFormulaCellsWithinArea`/`TryGetFormulaCells` - all specifically
catch `COMException` there and treat it as "no formulas found"). `AddinModule.cs`'s
`CountFormulaCells` was the one place still catching `Exception` generically and routing
every occurrence through `Logger.LogException`, so the exception log filled up with a
"failure" entry every time a user opened/switched to a sheet with no formulas at all.

**Fix**: added a dedicated `catch (COMException ex) when ((uint)ex.ErrorCode ==
0x800A03EC)` before the generic `catch (Exception ex)` in `AddinModule.cs`'s
`CountFormulaCells`, returning `0` silently for that specific, expected HRESULT while still
logging any other, genuinely unexpected exception.
- **Status**: implemented, not yet rebuilt/tested by the user.

---

## 12. New feature: GLSense_GetAccountType (completed in FinalWorkingCode, then ported to AIPowered)

User was mid-way through adding a brand-new Excel UDF, `GLSense_GetAccountType`, to
FinalWorkingCode: request/URL construction was done, parsing/return was reportedly not.
Instructed to finish FinalWorkingCode first, then port to AIPowered. Sample API response
provided: `{"msg":"default","records":"Assets","status":"success"}`.

### 12.1 FinalWorkingCode: investigation showed parsing already worked - the real bug was elsewhere

Read `GLSenseExcelFunctions.cs`'s `GLSense_GetAccountType` end to end: it already called
`ParseSegmentDFFResponse(apiResponse)` and `asyncCallObject.ReturnResult(output)` - parsing
was NOT missing. Traced the parsing pipeline (`ApiResponseHelper.Parse<JsonElement>` ->
`HandleWrappedResponse` -> `TryGetRecordsNode` -> `ExtractStringValue`) against the exact
sample response and confirmed it already correctly returns `"Assets"` (finds the `"records"`
property, which is a plain JSON string, and returns it directly).

Two real bugs were found and fixed instead:
1. A copy-paste leftover log line said `"GLSense_GetSegmentDFF invoked..."` instead of
   `"GLSense_GetAccountType invoked..."` - cosmetic, but fixed for clean diagnostics.
2. **A real, function-blocking bug** in `GetSegmentSequenceIndex`: it did
   `int.TryParse(match.Value, ...)` where `match.Value` is the *whole* regex match (e.g.
   `"SEGMENT3"`), not just the captured digits - `int.TryParse("SEGMENT3", ...)` always
   fails, so this method returned `-1` for every segment, which means
   `GLSense_GetAccountType` would always hit its "Segment sequence not found" check and
   return an error *before ever reaching the API call*, regardless of how correct the
   parsing logic was. Fixed to use `match.Groups[1].Value` (the `"(\d+)"` capture) instead.
- **Status**: both fixed in `FinalWorkingCode\GLSense\GLSenseExcelFunctions.cs`. Not yet
  rebuilt/tested by the user.

### 12.2 AIPowered: full port (function didn't exist there at all yet)

AIPowered's architecture splits every UDF across two projects: the host project
(`GLSense\GLSenseExcelFunctions.cs`, an `ADXXLLModule`) holds only ADX-only/host-only
plumbing ([ExcelParam] validation, async plumbing) and forwards to
`GLSense.Addin.Core.Udf.UdfDispatcher` (hot-reloadable, holds all real business logic) via
`CrossToExecuteUdf`/`AddinEntry.ExecuteUdf`. Ribbon buttons are wired through a third path:
`AddinModule.Designer.cs` (button definition) -> `AddinModule.cs` (`_OnClick` handler calling
`_ribbonController?.ExecuteAction("...")`) -> `AddinEntry.cs`'s `OnRibbonAction` switch ->
`ShowSegmentWindow(string funcName)` (opens the shared `GLSegmentFunctions` WPF window).

Ported piece by piece, mirroring the already-working `GLSense_GetSegmentDFF`/`RibSegmentDFF`
wiring at every step:
- **`GLSense.Addin.Core\Udf\UdfDispatcher.cs`**: added `case "GLSense_GetAccountType"` to
  the `Execute` switch, a new `HandleGetAccountType(args)` method (GET
  `.../account-type?cubeId=...&segmentValue=...&segmentNumber=...&ledgerId=...`, reusing the
  existing `ParseSegmentDFFResponse`/`TryGetRecordsNode`/`ExtractStringValue` string-parsing
  pipeline since the response shape is identical to `GLSense_GetSegmentDFF`'s), and a new
  `GetSegmentSequenceIndex` helper (ported from FinalWorkingCode with the regex bug from 12.1
  pre-fixed, i.e. `match.Groups[1].Value` from the start). Added the missing
  `System.Text.RegularExpressions` using.
- **`GLSense\GLSenseExcelFunctions.cs`** (host): added the async wrapper
  `GLSense_GetAccountType(SegmentValue, SegmentName, Ledger, asyncCallObject)`, identical in
  structure to `GLSense_GetSegmentDFF`'s wrapper (`ValidateInputs` -> `Task.Run` ->
  `CrossToExecuteUdf("GLSense_GetAccountType", parameters)`).
- **`GLSense\AddinModule.Designer.cs`**: added a `RibSegmentAccountType` ribbon button
  (field declaration, `RibFunctionsMenu.Controls.Add`, full control setup block copied from
  FinalWorkingCode's own Designer.cs including the same `Id` GUID/Caption/SuperTip text, and
  the public field declaration).
- **`GLSense\AddinModule.cs`**: added `RibSegmentAccountType_OnClick`, calling
  `_ribbonController?.ExecuteAction("ShowSegmentAccountType")` (same pattern as
  `RibSegmentDFF_OnClick`).
- **`GLSense.Addin.Core\AddinEntry.cs`**: added `case "ShowSegmentAccountType":
  ShowSegmentWindow("ACCOUNTTYPE"); break;` to the `OnRibbonAction` dispatch switch.
- **`GLSense.Addin.Core\ViewModels\GLSegmentFuncsViewModel.cs`**: added `"ACCOUNTTYPE" => 3`
  to `GetLedgerName()`'s expected-arg-count switch, and `case "ACCOUNTTYPE": return
  "GLSense_GetAccountType";` to `BuildFormulaName()`. `FormulaParameters()` needed no new
  case (ACCOUNTTYPE's 2 base args + ledger = 3, matching the default path, same as
  FinalWorkingCode).
- **`GLSense.Addin.Core\Views\GLSegmentFunctions.xaml.cs`**: added an `"ACCOUNTTYPE"` case to
  the constructor's UI-section-visibility switch (title "Get Segment Account Type";
  collapses `ParentChildSection`/`AttributesSection`/`IncludeValuesSection`/`ResultSection`,
  matching FinalWorkingCode exactly), added `"ACCOUNTTYPE" => 3` to this file's own duplicate
  `GetLedgerName()` switch, and added a `GLSense_GetAccountType` string check to
  `IsSegmentFormula()` so re-opening a cell already containing this formula is recognized.
- **`RibbonStateHelper.cs`**: no changes - confirmed neither FinalWorkingCode's nor
  AIPowered's version of this file references individual segment-function ribbon buttons by
  name (it's a generic enable/disable helper keyed by control name string), and
  FinalWorkingCode's `ApplyDrilldownState` control list doesn't include
  `RibSegmentAccountType` either, so AIPowered's equivalent list was left untouched for
  parity.
- **`GLSenseExcelFunctions.Designer.cs`** (Excel Insert-Function-dialog help text): checked
  - FinalWorkingCode hasn't added a descriptor for this function there either, so left
  AIPowered's untouched too, for parity. Purely cosmetic (Insert Function dialog tooltip),
  doesn't affect the formula working.
- **Status**: full port implemented, not yet rebuilt/tested by the user. Once rebuilt,
  `=GLSense_GetAccountType(SegmentValue, SegmentName, Ledger)` should behave identically in
  both codebases, and the "Account Type" ribbon button (Functions menu) should open the
  same `GLSegmentFunctions` dialog FinalWorkingCode does.

---

## 13. GLSense_GetAccountType: segment argument changed from name to 1-based dropdown index

Change request, applied to both codebases: the formula's 2nd argument used to be the
segment NAME (e.g. `=GLSense_GetAccountType("1000", "Account", "Vision Ops")`), resolved
server-side via `ResolveSegmentName` + the `ApplicationColumnName`-regex helper
`GetSegmentSequenceIndex`. It now carries the segment's 1-based position within the
Segment picker's dropdown directly, e.g. `=GLSense_GetAccountType("1000", 3, "Vision Ops")`
- if the dropdown holds `Company, Department, Account` (0-based indices 0/1/2), picking
"Account" (either directly or via a cell reference that resolves to "Account") now bakes in
`3` (index 2 + 1), not the text `"Account"`.

### 13.1 UDF side: index is now used directly, no more name resolution

**Both `FinalWorkingCode\GLSense\GLSenseExcelFunctions.cs`'s `GLSense_GetAccountType` and
AIPowered's `GLSense.Addin.Core\Udf\UdfDispatcher.cs`'s `HandleGetAccountType`**: removed
the `ResolveSegmentName`/`GetSegmentSequenceIndex` calls entirely (per the user's explicit
instruction to drop the added helper, since the index is now supplied directly). Added a
`TryParseSegmentIndex(object value, out int index)` helper (both codebases) that accepts
`double` (Excel passes an un-quoted numeric literal as a boxed double), `int`, or a
numeric string, and parses it defensively - the parsed value is used directly as the API
call's `segmentNumber` query parameter. `GetSegmentSequenceIndex` was deleted from both
files, along with the now-unused `System.Text.RegularExpressions` using in each.
- **Status**: implemented, not yet rebuilt/tested by the user.

### 13.2 ViewModel side: build the index when writing the formula, reverse it when reopening

**Both `FinalWorkingCode\GLSense\ViewModels\GLSegmentFuncsViewModel.cs` and AIPowered's
`GLSense.Addin.Core\ViewModels\GLSegmentFuncsViewModel.cs`** (identical changes in each):

- **`FormulaParameters()`**: added an `_formulaName == "ACCOUNTTYPE"` branch that emits a new
  `GetSelectedSegmentIndex()` helper's result as a bare (unquoted) number instead of
  `FormatFormulaArg(segmentVal)` (the segment name). Every other formula type is unaffected.
- **`GetSelectedSegmentIndex()`** (new): returns `Segments.IndexOf(SegmentField.ComboValue as
  SegmentModel) + 1`, or `-1` if nothing's selected. This works identically whether the
  segment was picked directly from the combo (WPF binds `SelectedItem` to `ComboValue`, so
  it's the same `SegmentModel` instance that's in `Segments`) or resolved from a cell
  reference (`SetSegmentFromLookup` already sets `ComboValue` to the matching `SegmentModel`
  from `Segments` too, confirmed by reading that method) - either way `Segments.IndexOf` finds
  it, so the formula always gets a literal index, never a live cell reference, regardless of
  how the user picked the segment. This matches the user's own description ("if i select a
  reference to 'Account' then the combo box will be populated with account and we can get
  that index").
- **`ProcessSegmentField`** (reopening the picker from an existing formula cell): added an
  `else if (_formulaName == "ACCOUNTTYPE")` branch calling a new `SetSegmentFromIndex(string
  cleanValue)` method instead of the default `SetSegmentFromList` (which matches by name).
  `SetSegmentFromIndex` parses the argument as a 1-based integer and sets `SegmentField.
  ComboValue = Segments[index - 1]` (with bounds/parse validation and a warning message on
  failure), so the combobox shows the right segment name again when the window reopens. The
  existing "raw cell reference" branch (`SetSegmentAsReference`, when the argument is a real
  Excel range) is untouched and still takes priority, same as before.
- **Status**: implemented, not yet rebuilt/tested by the user.

### 13.3 Live cell-reference lookup also needed the same index-based fix

User screenshot: referencing a cell containing `1` (a valid segment index) into the
Segment field while the window is open (not reopening an existing formula - live
interactive use) showed `Value "1" not found in available options.`

Root cause: that live-reference path goes through a DIFFERENT method than 13.2's reopen
path - `ProcessRefEditValue` -> `TrySetFieldFromLookup` -> `SetSegmentFromLookup(cellValue)`,
which unconditionally matched `Segments` by NAME regardless of formula type, so a
referenced cell holding `"1"` never matched any segment named "1".

**Fix (both codebases' `GLSegmentFuncsViewModel.cs`)**: `SetSegmentFromLookup` now branches
on `_formulaName == "ACCOUNTTYPE"`: for ACCOUNTTYPE it parses `cellValue` as a 1-based
integer and resolves `Segments[index - 1]` (matching the user's stated convention - numbers
in referenced cells are 1-based, converted to a 0-based list lookup), instead of the
default name-matching `FirstOrDefault`. Everything else (`IsValueFromRefEdit`,
`ComboValue`, `SelectedSegment`, failure handling via `HandleMatchFailure`) is unchanged.
- **Status**: implemented, not yet rebuilt/tested by the user.

### 13.4 Formula construction also had to respect reference-over-value priority

User clarified an important convention this window already follows for every other field:
the picker lets you choose a direct value OR a cell reference (that's what
`ComboFieldBindings`' `ComboValue`/`RefValue`/`IsComboEnabled`/`IsRefEnabled` exist for), and
whenever a live reference is active, it takes priority over a resolved value both when
*building* the formula and when *parsing* it back open - `GetFieldValue`'s Step 1 already
checks `RefValue` before `ComboValue` for every other argument.

13.2's original `GetSelectedSegmentIndex()`-based construction didn't follow this: it always
resolved and baked in a static index from `ComboValue`, even when the user had an active
cell reference in the Segment field, silently discarding the live reference instead of
embedding it.

**Fix (both codebases' `GLSegmentFuncsViewModel.cs`)**: added `GetAccountTypeSegmentArg()`,
called from `FormulaParameters()`'s ACCOUNTTYPE branch instead of `GetSelectedSegmentIndex()`
directly: if `SegmentField.RefValue` is set (reference mode active), it returns that
reference as-is (unquoted, embedded live in the formula, e.g. `$A$1`); otherwise it falls
back to `GetSelectedSegmentIndex()`'s resolved 1-based number. The parsing side already
handled this correctly by construction - `ProcessSegmentField` checks
`ExcelRangeHelper.IsRealRange` first regardless of formula type, and
`SetSegmentAsReference` setting `RefValue` triggers `OnRefEditTextChanged` ->
`SetSegmentFromLookup` (13.3's fix), so reopening a reference-driven ACCOUNTTYPE formula
already re-resolved the right combo item correctly - only the construction side needed
this fix.
- **Status**: implemented, not yet rebuilt/tested by the user.

### 13.5 What did NOT need to change

The `GLSegmentFunctions.xaml.cs` UI-visibility switch, `IsSegmentFormula()`, and the
duplicate `GetLedgerName()` (arg-count-based, unaffected by whether arg 2 is a name or a
number) all needed no changes - the Segment picker UI itself is untouched; only what gets
written into/read from the formula text changed. The host wrapper's `GLSense_GetAccountType`
parameter was renamed `SegmentName` -> `SegmentIndex` in both `GLSenseExcelFunctions.cs`
files for clarity (purely cosmetic - the host still just forwards the boxed value through).

---

## 14. Update-system folder/file nomenclature: "Version" -> "Manifest"

Preparatory rename for the upcoming production auto-update feature (see the
architecture discussion notes for the three-tier update model: pre-AppDomain-load
silent check/download in `GLSense.Loader.Core`, manual "Select Update Source" picker in
`GLSense.Addin.Core`). Before writing any new update-flow code, the user asked to first
fix the existing nomenclature: the update-tracking JSON file/folder was named
"Version"/`version.json`, which is easy to confuse with the pre-existing, unrelated
"Versions" (plural) folder that holds the actual hot-reloadable DLL payloads
(`VersionsPath` / `V{version}\` subfolders). Renamed the singular concept to
"Manifest"/`manifest.json` everywhere; the plural "Versions" DLL-payload folder was left
completely untouched.

**Changed:**
- `GLSense.Contracts\IPathProvider.cs` - interface members `VersionFile`/`VersionDirectory`
  renamed to `ManifestFile`/`ManifestDirectory`.
- `GLSense.Shared\PathProvider.cs` - property implementations renamed to match
  (`ManifestDirectory => Path.Combine(_root, "Manifest")`, `ManifestFile => Path.Combine
  (ManifestDirectory, "manifest.json")`), plus every internal call site in `Ensure()`,
  `InitializeVersion()`, and the renamed `CreateDefaultManifestFile()` (was
  `CreateDefaultVersionFile()`) updated to use `ManifestFile`.
- `GLSense\UpdateManager.cs` - `_context.Paths.VersionFile` -> `.ManifestFile` in
  `GetLocalVersionInfoAsync()` (this was a compile-breaking reference after the
  `IPathProvider` rename, now fixed). Also renamed every literal `"version.json"` string
  to `"manifest.json"`: the remote URL segment (`$"{domainUrl}/version.json"`, used in
  both `DownloadVersionJsonAsync` and `DownloadAndExtractVersionAsync`) and the local
  per-version-folder copy path (`Path.Combine(versionPath, "version.json")`).
- `GLSense.Shared\VersionParser.cs` - `ParseVersionFromDirectory`'s hardcoded
  `Path.Combine(directoryPath, "version.json")` -> `"manifest.json"`, plus doc-comments
  and debug/error log text updated from "version file"/"version.json" to "manifest
  file"/"manifest.json" for consistency. (Confirmed via grep this method is currently
  unused/dead code, but kept consistent anyway.)
- `GLSense.Addin.Core\post_build.cmd` and `GLSense.Loader.Core\post_build.cmd` - the
  commented-out (`REM`) manifest-seeding lines near the bottom of each file updated from
  `%DEPLOY_ROOT%\Version\version.json` to `%DEPLOY_ROOT%\Manifest\manifest.json`. Still
  commented out in both - this seeding step isn't active yet.

**Deliberately NOT renamed** (represent the version *concept*, not a file/folder path):
`GLSense.Contracts\VersionInfo.cs` (`{Version, ReleaseDate}` DTO),
`IVersionParser`/`VersionParser`/`VersionParseResult` and their method names
(`ParseVersionJson`/`ParseVersionFile`/`ParseVersionFromDirectory`/`GetLatestVersion`),
`IPathProvider.VersionsPath`/`LatestVersion`/`LatestReleaseDate`/`AllVersions`,
`IGLSenseContext.Version`/`ReleaseDate`/`AllVersions`. The plural `Versions` DLL-payload
folder (`VersionsPath`, `V{version}\` subfolders) is a completely separate concept from
this manifest file and was untouched throughout.

**Verified via grep** across all `.cs`/`.cmd` files that no other reference to the old
`VersionFile`/`VersionDirectory`/literal `version.json` remains anywhere in the AIPowered
project - the only surviving `VersionFile`-shaped hits are the intentionally-kept
`ParseVersionFile` method name.

**Not yet done, deliberately out of scope for this rename** (noted for the next
update-system work session, not yet raised with the user): `VersionParser
.GetLatestVersion()` and `UpdateManager.IsUpdateAvailableAsync()` both pick "latest"
via `System.Version` semantic comparison, but the user's stated design freezes the
version number for months and treats `releaseDate` as the sole comparison signal - worth
revisiting once the actual update-check logic is built, since two manifest entries could
share a version number with different release dates.

- **Status**: implemented, not yet rebuilt/tested by the user.

### 14.1 Post-build scripts weren't actually seeding the Manifest folder - they were commented out

User reported: after building, no `Manifest` folder appeared under
`%LOCALAPPDATA%\ORBIT\Excel_Logs\GLSense_Logs_New\`. Root cause: the manifest-seeding
lines in both `post_build.cmd` files (renamed from "Version" to "Manifest" in 14 above)
had been commented out (`REM`) the whole time - they were only a single-line placeholder
(`echo %VERSION% > ...`), never actually enabled. `PathProvider`'s own
`CreateDefaultManifestFile()` (in `GLSense.Shared`) does auto-create a default manifest,
but only lazily, the first time `PathProvider`'s constructor runs (i.e. when Excel
actually loads the add-in) - not at build time, which is why nothing appeared right
after a build with Excel not yet (re)launched.

**Fix**: uncommented and rewrote the manifest-seeding block in both
`GLSense.Addin.Core\post_build.cmd` and `GLSense.Loader.Core\post_build.cmd`
(identical in both). It now:
- Creates `%DEPLOY_ROOT%\Manifest\` if missing.
- Only seeds `manifest.json` if the file doesn't already exist (mirrors
  `CreateDefaultManifestFile()`'s own "don't clobber an existing manifest" behavior -
  a rebuild won't overwrite a manifest a developer is using to test the update flow,
  e.g. a manually-set `downloadUrl`).
- Writes a JSON **array** (`[ { ... } ]`), matching what `PathProvider.InitializeVersion`/
  `VersionParser` currently deserialize (`List<VersionInfo>`) - a flat single object
  (no array wrapper) would fail to parse with the existing code.
- Each object contains the full decided manifest schema: `version`, `releaseDate`,
  `downloadUrl`, `checksum`, `notes`, `mandatory`. `releaseDate` is computed via
  `powershell -Command "[DateTime]::UtcNow.ToString('yyyy-MM-ddTHH:mm:ssZ')"` (locale-
  independent, unlike parsing `%DATE%`/`%TIME%`; avoided `wmic` since it's deprecated/
  removed on newer Windows builds). `downloadUrl`/`checksum` are left empty (IT populates
  these at ship time per the agreed design); `notes` defaults to `"Local dev build"`;
  `mandatory` defaults to `false`.

**Naming note, flagged for the user rather than silently decided**: the field is named
`"version"` here, not `"versionNumber"` as originally proposed in the schema discussion -
`"version"` is what actually matches `VersionInfo.Version` (the C# property
`JsonSerializer` binds to, case-insensitively). If `versionNumber` is preferred as the
wire name going forward, `VersionInfo.cs` needs a `[JsonPropertyName("versionNumber")]`
attribute (or the property renamed) to match - not done here since it touches the shared
DTO used by all existing parsing code, and only the field NAME would change, not the
rename task's scope.

**Also worth noting**: `downloadUrl`/`checksum`/`notes`/`mandatory` are written into the
file now but are currently silently ignored by `VersionParser`/`PathProvider` - neither
`VersionInfo.cs` nor any parsing code has been extended to read them yet, since the
actual update-check/download logic hasn't been built. They're present in the seeded
JSON now so the schema is future-ready, but reading/using them is a separate, not-yet-
started piece of work.

- **Status**: implemented, not yet rebuilt/tested by the user.

### 14.2 Codebase adapted to actually model the new manifest fields (not just ignore them)

After 14.1 shipped, `downloadUrl`/`checksum`/`notes`/`mandatory` existed in the seeded
JSON but nothing in the codebase read them - `VersionInfo.cs` only declared
`Version`/`ReleaseDate`, so `System.Text.Json` (both call sites already use
`PropertyNameCaseInsensitive = true`, no `UnmappedMemberHandling.Disallow` set anywhere)
silently dropped the unknown JSON properties rather than throwing - technically "safe"
but the new fields were write-only, dead weight in the file. User asked for the codebase
to actually adapt to the change and confirmed parsing must not fail.

**Fix - extended the model all the way through the existing pipeline, no new pipeline
needed:**
- `GLSense.Contracts\VersionInfo.cs` - added `DownloadUrl`, `Checksum`, `Notes` (string)
  and `Mandatory` (bool). Binds automatically via case-insensitive matching against the
  JSON keys `downloadUrl`/`checksum`/`notes`/`mandatory` - no `JsonPropertyName`
  attributes needed.
- `GLSense.Shared\VersionParser.cs` - `VersionParseResult` gained matching
  `DownloadUrl`/`Checksum`/`Notes`/`Mandatory` properties, populated in
  `ParseVersionJson()` from the resolved `latest` entry right alongside
  `Version`/`ReleaseDate`.
- `GLSense.Contracts\IPathProvider.cs` / `GLSense.Shared\PathProvider.cs` - added
  `LatestDownloadUrl`/`LatestChecksum`/`LatestNotes`/`LatestMandatory`, backed by new
  static fields populated in `InitializeVersion()` alongside `_latestVersion`/
  `_latestReleaseDate` (null-coalesced to empty string for the string fields so a
  manifest entry that omits a field never surfaces a null down the line).
- `PathProvider.CreateDefaultManifestFile()`'s fallback default entry (used when
  `manifest.json` is missing entirely, e.g. a genuinely first-run install with no
  post-build seed and no IT-provided manifest yet) now also sets
  `DownloadUrl`/`Checksum`/`Notes`/`Mandatory` so both fallback paths (C#'s own default
  vs. the batch script's seed) produce the same full shape.
- Confirmed via grep that `PathProvider` is the only class implementing `IPathProvider`
  in the codebase, so adding interface members didn't require touching any other class.

**Why this can't fail parsing**: `System.Version` (used to pick "latest") only ever
reads `.Version` - untouched by the new fields. `ParseReleaseDate()`'s format list
doesn't include ISO-8601, but its final fallback, a bare `DateTime.TryParse(releaseDate,
...)` with no exact format, natively recognizes the `yyyy-MM-ddTHH:mm:ssZ` shape the
post-build script now writes - confirmed this is the actual code path exercised (the
`TryParseExact` formats all miss it, generic `TryParse` catches it), not a
theoretical fallback. Verified the exact JSON string produced by the post-build script's
`echo` block is syntactically valid JSON (parsed cleanly, correct field types
str/str/str/str/str/bool) as a structural check, since a live .NET Framework 4.8.1
runtime wasn't available in this environment to compile-and-run an actual round-trip
test - the binding-safety reasoning above (case-insensitive matching, no
`UnmappedMemberHandling.Disallow`, single `IPathProvider` implementer) is standard,
well-documented `System.Text.Json`/interface-compiler behavior, not a guess.

- **Status**: implemented, not yet rebuilt/tested by the user.

### 14.3 Removed redundant/duplicate steps across the 3 post-build scripts

First step of a broader post-build cleanup (see the deployment-pipeline redesign
discussion for the rest). Full inventory of what the 3 scripts do:
`GLSense.Addin.Core\sqlite_postbuild.cmd` (copies SQLite native interop DLLs from the
NuGet cache into Addin.Core's own `bin\{Config}\{x86,x64}\`, chained first via Addin.Core's
`<PostBuildEvent>`), `GLSense.Addin.Core\post_build.cmd` (deploys Addin.Core's build
output into `%LOCALAPPDATA%\...\GLSense_Logs_New\Versions\v11.1.0\`), and
`GLSense.Loader.Core\post_build.cmd` (deploys Loader.Core's build output into that exact
same shared `Versions\v11.1.0\` folder).

Found and removed:
- **Duplicate Manifest-seed block**: the entire "ensure Manifest folder + manifest.json
  exist" step (added in 14.1) was byte-for-byte identical in both `post_build.cmd`
  files - a genuine duplicate step, not just similarly-shaped code. Removed from
  `GLSense.Loader.Core\post_build.cmd`; kept only in `GLSense.Addin.Core\post_build.cmd`
  (Addin.Core is the project rebuilt on every local dev/hot-reload iteration, so keeping
  it there - rather than in the less-frequently-rebuilt Loader.Core - guarantees the
  manifest gets seeded on the machine's very next relevant build).
- **Duplicate folder-creation block**: both scripts independently created
  `Logs`/`Database`/`TempFiles`/`BrowserLogs`/`Versions\%VERSION%` under `DEPLOY_ROOT`.
  `GLSense.Loader.Core\post_build.cmd` doesn't write into `Logs`/`Database`/`TempFiles`/
  `BrowserLogs` at all - trimmed its folder-creation block down to just
  `DEPLOY_ROOT` + `Versions\%VERSION%` (the only two directories its own `xcopy` actually
  needs), leaving a comment pointing at Addin.Core's script as the owner of the rest.
- **Dead variable**: `GLSense.Loader.Core\post_build.cmd` computed
  `DATA_BIN_DIR=%SOLUTION_DIR%\GLSense.Addin.Core\bin\%CONFIG%` and echoed it, but never
  referenced it anywhere else in the script (no copy/xcopy used it) - misleading
  leftover, removed entirely.

**Deliberately NOT touched in this pass** (would require introducing a genuinely shared
script file, which is a bigger structural change than "remove duplicates" - see the
pipeline-redesign discussion instead): the ~20-line `TARGET_DIR`/`PROJECT_DIR`/
`SOLUTION_DIR`/`CONFIG` parameter-parsing boilerplate is still duplicated verbatim across
both `post_build.cmd` files, since MSBuild's `<PostBuildEvent>` mechanism gives each
project its own independent script with no built-in "include" - and the hardcoded
`VERSION=v11.1.0` literal is still duplicated in both files too (a live, tracked
follow-up, not resolved here).

- **Status**: implemented, not yet rebuilt/tested by the user.

### 14.4 Centralized the version number - one edit, not five (or two)

Second step of the post-build cleanup, ahead of the local-host update simulation (see
that design discussion for the full roadmap). User asked directly whether a value could
flow from "a class" into the build events, and specifically wanted bumping the version
in the future to require touching exactly one place. Confirmed a batch script can't read
a C# class's field (nothing left to read once compiled - a class isn't a build-event
concept), but two real mechanisms achieve the same effect. Checked first: every one of
the 5 projects' `Properties\AssemblyInfo.cs` had the generic Visual Studio default
`AssemblyVersion("1.0.0.0")` - completely disconnected from the "11.1.0" used everywhere
else (batch files, `VersionInfo` defaults) - so there was no existing source of truth to
reuse; this was built from scratch.

**Chosen approach**: a shared, linked assembly-info file - the standard pattern for
classic (pre-SDK-style) .NET Framework solutions, which don't support MSBuild
`<Version>` properties auto-flowing into compiled attributes the way SDK-style projects
do.

- Created `GLSenseSharedVersion.cs` at the **solution root**
  (`D:\SQLLite_Test\AIPowered\GLSense\GLSenseSharedVersion.cs`), containing only
  `[assembly: AssemblyVersion("11.1.0.0")]` / `[assembly: AssemblyFileVersion("11.1.0.0")]`
  plus a header comment explaining the pattern and warning not to re-add these
  attributes elsewhere.
- Removed the `AssemblyVersion`/`AssemblyFileVersion` attribute lines from all 5
  projects' own `Properties\AssemblyInfo.cs` (`GLSense`, `GLSense.Addin.Core`,
  `GLSense.Shared`, `GLSense.Contracts`, `GLSense.Loader.Core`) - left every other
  attribute (Title/Description/GUID/ComVisible/ThemeInfo/etc.) untouched. Note:
  `GLSense`'s (the host project) `AssemblyInfo.cs` only ever had `AssemblyVersion`, no
  `AssemblyFileVersion` at all - meant its compiled file-version metadata would have
  read `0.0.0.0` before this change; now it gets a real file version like every other
  project, for free.
- Linked `..\GLSenseSharedVersion.cs` into all 5 `.csproj` files (`<Compile
  Include="..\GLSenseSharedVersion.cs"><Link>Properties\GLSenseSharedVersion.cs</Link>
  </Compile>`, inserted next to each project's existing `AssemblyInfo.cs` `<Compile>`
  entry) - a *linked* file, not a copy, so all 5 assemblies compile the identical source;
  editing the one solution-root file and rebuilding updates every DLL's version
  simultaneously.
- Verified via grep: exactly one `[assembly: AssemblyVersion(...)]`/
  `[assembly: AssemblyFileVersion(...)]` pair exists in the whole solution (in the shared
  file), and all 5 `.csproj` files reference `GLSenseSharedVersion`.

**Post-build scripts updated to match** (`GLSense.Addin.Core\post_build.cmd` and
`GLSense.Loader.Core\post_build.cmd`): the hardcoded `set VERSION=v11.1.0` literal is
gone from both. Each script now runs
`powershell -NoProfile -Command "$v = (Get-Item '<its own DLL>').VersionInfo; '{0}.{1}.{2}' -f $v.FileMajorPart, $v.FileMinorPart, $v.FileBuildPart"`
against its own just-built DLL (`GLSense.Addin.Core.dll` / `GLSense.Loader.Core.dll`
respectively - each already exists in `$(TargetDir)` by the time a `PostBuildEvent`
fires, so no chicken-and-egg problem), captures the 3-part version via a temp file (same
redirect-to-file pattern as 14.1's `RELEASE_DATE`, avoiding `for /f` quote-nesting
issues), and falls back to `v0.0.0` with a `WARNING` echo if the read ever comes back
empty (DLL missing/PowerShell blocked) rather than silently producing a malformed
`Versions\v\` folder. `%VERSION:v=%` (used to populate `manifest.json`'s `version` field)
needed no changes - it operates on whatever `%VERSION%` holds, hardcoded or derived.

**Result**: to release a new version, edit the two literals in
`GLSenseSharedVersion.cs`, rebuild. Nothing else - not the 5 `AssemblyInfo.cs` files, not
either `post_build.cmd`, not `manifest.json`'s seeded content - needs a manual touch;
they all derive from the one file on the next build.

- **Status**: implemented, not yet rebuilt/tested by the user. Real .NET Framework
  compile/run verification wasn't possible in this environment (no MSBuild/dotnet/mono
  available) - correctness here rests on standard, well-documented MSBuild linked-file
  and Win32 file-version-resource behavior, not an executed test.

### 14.5 Removed all folder creation from both post-build scripts - PathProvider.cs owns it

Third step of the post-build cleanup, before the local-host update simulation. User's
direction: the build should not be responsible for ensuring any deployment folder
exists - that's `PathProvider.cs`'s job (`Ensure()` already creates
`Logs`/`Database`/`Temp`/`LoginBrowserPath`/`DrilldownBrowserPath`/`VersionsPath`/
`Resources`/`ManifestDirectory`, and `CreateDefaultManifestFile()` already creates the
`Manifest` folder specifically and seeds a default `manifest.json` - both already run
automatically the first time `PathProvider` is constructed, i.e. the next time Excel
loads the add-in).

**Removed from both `post_build.cmd` files**: every `if not exist ... mkdir ...` line
(`DEPLOY_ROOT` itself, `Logs`, `Database`, `TempFiles`, `BrowserLogs`,
`Versions\%VERSION%`), plus the entire manifest-seeding block (STEP 2, added in 14.1)
including its own `Manifest` folder `mkdir` - all replaced with comments pointing at
`PathProvider.cs` as the sole owner going forward. Also fixed a comment in
`GLSense.Loader.Core\post_build.cmd` left stale by this change (it previously said
manifest-seeding was "kept in Addin.Core's script only" - no longer true since it was
removed from there too).

**Why this doesn't break the DLL deployment step**: the `xcopy` commands that actually
place the built DLLs (and `x64`/`x86`/`de`/`runtimes` subfolders in Addin.Core's script)
were left untouched, and none of them need a preceding `mkdir` - every destination path
ends in a trailing backslash (`...\Versions\%VERSION%\`), and per `xcopy`'s documented
behavior, a trailing-backslash destination that doesn't exist while copying multiple
files is unambiguously treated as a directory and created automatically (this is the
same reasoning that already justified skipping `/I` on the main DLL `xcopy` in earlier
sections - no interactive "F = file, D = directory?" prompt to worry about). This holds
even on a completely fresh machine where `%LOCALAPPDATA%\...\GLSense_Logs_New\` doesn't
exist at all yet - `xcopy` creates the full missing chain, not just the last segment.

**Real, intentional behavior change worth flagging**: right after a build, on a machine
that has never run Excel with this add-in, the `Manifest` folder (and `manifest.json`)
will **not** exist yet - only appears once Excel actually loads and `PathProvider`'s
constructor runs. This reverses 14.1's original goal (make the manifest appear
immediately after building, without needing to run Excel first) in favor of a cleaner
ownership split ("build only places files; the running app owns its own folder
structure") - a deliberate trade-off the user chose, not an oversight. This also fits
naturally with the upcoming local-host/zip redesign, where manifest generation is
expected to move out of post-build entirely and into the new update-simulation flow.

**Not independently verified in this environment** (no Windows/`xcopy` available in this
sandbox): the claim that `xcopy` creates multi-level missing destination directory
chains from a bare trailing-backslash path with no pre-existing parent folders at all,
on the very first build ever on a brand-new machine. This is long-established, widely-
relied-upon `xcopy` behavior, not a guess, but flagging it as the one part of this change
that should get a real first-build-on-a-clean-machine test rather than just a normal
rebuild.

- **Status**: implemented, not yet rebuilt/tested by the user.

---

## 15. Automated update-bootstrap flow (local-host simulation)

**SUPERSEDED by section 17** - the `GLSense.LocalUpdateHost` project and all
remote/HTTP logic described below were removed after repeated real-world connection
failures during testing (see section 17). Kept here as a historical record of the
design and why it was tried, same as this file does for other superseded approaches
(e.g. section 1's SizeToContent saga) - don't re-derive this from scratch if "online
updates" comes back up; read 17 first for the current, simpler folder-only state, then
this section for what a future real remote-server version should probably still borrow
(the manifest schema, the `UpdateBootstrapper` decision-tree shape, the "silently fall
back on any failure" principle).

Real functional piece of the production update roadmap (not just build-script cleanup
like section 14). Goal: `AddinModule_OnRibbonLoaded` should automatically decide which
version of `GLSense.Addin.Core` to load - and fetch it if it isn't on disk yet - using
`manifest.json` as the readily-available, single source of truth, instead of the
previously hardcoded `GlobalsEx.Context.Version = "11.1.0"` (which had zero connection to
the manifest and meant nothing discussed in sections 9-14 actually influenced which DLLs
got loaded).

**Two decisions confirmed with the user before writing any code** (both affect Excel
startup reliability, too consequential to guess):
1. Nothing was running at `http://localhost/GLSense` yet - user asked for help setting
   one up, rather than pointing at an existing IIS/IIS Express site.
2. If the local host is unreachable during the startup check, the add-in should
   **silently fall back** to whatever's already installed - a network hiccup must never
   block Excel from opening with a working add-in.

### 15.1 UpdateBootstrapper (`GLSense.Loader.Core\UpdateBootstrapper.cs`) - new class

Lives in `GLSense.Loader.Core` (not `GLSense.Addin.Core`) per the architecture agreed
earlier in this engagement: the pre-AppDomain-load check has to live host-side, since
Addin.Core isn't loaded yet at this point and can't be responsible for replacing itself.

`ResolveVersionToLoad(IGLSenseContext context)` implements this decision tree:

1. **Manifest folder missing** -> return `null` (defensive only - `PathProvider.Ensure()`
   already creates this before `OnRibbonLoaded` ever runs, so this should never actually
   trigger in practice).
2. **`manifest.json` + a `.zip` both present** in the Manifest folder -> extract the zip
   into `Versions\V{version}\` (version read from the local `manifest.json`, already
   parsed into `PathProvider.LatestVersion`), delete the zip, adopt that version. Zero
   network calls - this is the "IT/dev dropped a zip here" case from the original
   three-tier design.
3. **Only `manifest.json`, no zip**:
   - If `Versions\V{version}\` already has `.dll` files on disk: try the remote host
     (`http://localhost:8080/GLSense/manifest.json` - see 15.3 for what serves this) for
     something newer, comparing `releaseDate`. **Any exception here (unreachable host,
     timeout, bad JSON, etc.) is caught and this falls back to the local version** - this
     is the "silently fall back" behavior confirmed above, and it's the path exercised
     on every normal Excel launch where nothing new has been published.
   - If that folder is missing/empty (nothing usable installed at all): the remote host
     **must** be reached to get something - there's no local fallback available. If that
     also fails, returns `null` so the caller skips loading the AppDomain instead of
     crashing Excel.

When a remote version is adopted (either branch), `DownloadZipAndAdopt` downloads the
zip to a temp file, extracts it into `Versions\V{version}\` (replacing any existing
folder of that name), calls the new `IPathProvider.WriteManifest(VersionInfo)` (see
15.2) to make the just-downloaded version the new local "installed" record, then deletes
the temp zip in a `finally` block.

HTTP calls use a `Task.Run(...).GetAwaiter().GetResult()` wrapper (`RunSync<T>`) rather
than a bare `.Result`/`.GetAwaiter().GetResult()` on the UI-thread-created task - Excel's
ADX thread has a WinForms message pump (`Application.EnableVisualStyles()` in
`AddinModule`'s constructor), so blocking directly on a task created on that thread risks
the classic "continuation wants to resume on the thread that's blocked waiting for it"
deadlock. Handing the async work to `Task.Run` first removes the captured context, so
blocking on *that* task from the UI thread is safe.

Remote manifest JSON is parsed via the existing `GLSense.Shared.VersionParser` (already
referenced by Loader.Core through the `GLSense.Shared` project reference) - no new JSON
library needed, and it already returns `DownloadUrl`/`Checksum`/`Notes`/`Mandatory`
thanks to the 14.2 model extension.

### 15.2 PathProvider/IPathProvider additions

`InitializeVersion()`'s parsing logic was already private and only ran once, in the
constructor - nothing existed to (a) re-read `manifest.json` after something outside
`PathProvider` changed it, or (b) write a new "currently installed" record after a
download. Added:

- `void Refresh()` - just re-calls the (still-private) `InitializeVersion()`. Public so
  callers outside `GLSense.Shared` can force a re-parse.
- `void WriteManifest(VersionInfo info)` - overwrites `manifest.json` with a single-entry
  array (mirrors `CreateDefaultManifestFile()`'s shape - `manifest.json` is a record of
  "what's currently installed," not a growing history) and immediately calls `Refresh()`
  so `LatestVersion`/`LatestReleaseDate`/etc. reflect the write without needing a new
  `PathProvider` instance.

Both added to `IPathProvider` too. `VersionInfo` (`GLSense.Contracts`) gained
`[Serializable]` - it now crosses the host<->Addin.Core AppDomain boundary not just as
list elements (`IPathProvider.AllVersions`) but as a direct method parameter
(`WriteManifest`), and .NET Remoting requires non-`MarshalByRefObject` types to be
`Serializable` to pass by value across domains - without this, any cross-domain call
touching it would throw `SerializationException` at runtime.

### 15.3 GLSense.LocalUpdateHost - new console project (the "local host")

Since nothing was already serving `http://localhost/GLSense`, built a minimal, dependency
-free static file server: a `HttpListener`-based console app, added to `GLSense.sln`.

- Listens on `http://localhost:8080/` - **not port 80**. Binding `HttpListener` to port
  80 needs either running as Administrator or a one-time
  `netsh http add urlacl url=http://localhost:80/ user=Everyone` reservation, since port
  80 is privileged on Windows; port 8080 needs neither. `UpdateBootstrapper`'s constant
  and the post-build script's `downloadUrl` both point at port 8080 - if a real port-80
  URL is wanted later, both need updating together (or this becomes configuration
  instead of a hardcoded constant, at the same time work moves toward real production).
- Serves `wwwroot\` relative to the **project source folder**, not the `bin\Debug\`
  build-output copy - computed as two directories above
  `AppDomain.CurrentDomain.BaseDirectory` (`Program.cs`'s `Main()`). This is deliberate:
  it means `GLSense.Addin.Core\post_build.cmd` can write straight into
  `GLSense.LocalUpdateHost\wwwroot\GLSense\` and this already-running server picks up
  the new files on the very next request - no rebuild of `GLSense.LocalUpdateHost`
  itself required, and no `CopyToOutputDirectory`/timing dependency to get wrong. Files
  are read fresh via `File.ReadAllBytes` on every request, never cached.
- No dependencies beyond the .NET Framework BCL (`System.Net.HttpListener`) - old-style
  `Exe` csproj, `TargetFrameworkVersion v4.8.1` to match the rest of the solution. Has
  its own `Properties\AssemblyInfo.cs` but deliberately does **not** link
  `GLSenseSharedVersion.cs` (section 14.4) - it isn't part of the shipped product, its
  own version number is irrelevant.
- Basic path-traversal guard (rejects any resolved path that escapes `wwwroot`) and
  content-type mapping for `.json`/`.zip`/`.txt`/`.html`.
- **To test the flow**: run `GLSense.LocalUpdateHost.exe` (or F5 it) *before* opening
  Excel. If it isn't running, `UpdateBootstrapper`'s HTTP calls simply fail and it falls
  back to whatever's already installed locally - by design (15's second confirmed
  decision), not a bug.

### 15.4 Addin.Core post_build.cmd STEP 2 - publish zip + remote manifest

Added after the existing DLL-deploy `xcopy` block (STEP 1). Using the already-resolved
`%FILE_VERSION%`/`%VERSION%` from 14.4's version-centralization work:

- Creates `GLSense.LocalUpdateHost\wwwroot\GLSense\` if missing (computed from
  `%SOLUTION_DIR%`, already resolved earlier in the script). This mkdir is **not** a
  violation of section 14.5's "build shouldn't create deployment folders" rule - that
  rule is specifically about the `%LOCALAPPDATA%\...\GLSense_Logs_New\` tree, which
  `PathProvider.cs` owns. This folder is this project's own pretend-remote-server
  storage, an entirely different concern with no runtime owner of its own.
- Zips `%DEPLOY_ROOT%\Versions\%VERSION%\` (the folder STEP 1 *just* populated) into
  `v{FILE_VERSION}.zip` via PowerShell's `Compress-Archive` - zipping the already-
  deployed folder (rather than re-selecting `*.dll`/`x86`/`x64`/`de`/`runtimes` from
  `CORE_BIN_DIR` with separate logic) guarantees the zip's contents are byte-for-byte
  identical to what a manually-dropped zip in the Manifest folder would extract.
- Computes a SHA256 checksum via `Get-FileHash`.
- Writes `manifest.json` (stable filename, **not** version-suffixed - `UpdateBootstrapper`
  needs one well-known URL to check before it knows what version is even published) with
  `version`/`releaseDate` (fresh UTC timestamp per build)/`downloadUrl`
  (`http://localhost:8080/GLSense/v{version}.zip`)/`checksum`/`notes`/`mandatory: false`.
  This is a **separate file** from the local `Manifest\manifest.json` under
  `GLSense_Logs_New` - that one records "what's installed on this machine," this one
  records "what's currently published," matching a real update server's manifest.
- Guarded with `if not exist "%HOST_ZIP%" goto :SkipHostPublish` so a `Compress-Archive`
  failure (e.g. PowerShell blocked by execution policy) degrades to a `WARNING` rather
  than breaking the rest of the build.

**Known, accepted side-effect worth understanding, not a bug**: because this step reruns
on every single Addin.Core build and always stamps a fresh UTC `releaseDate`, the local
host's manifest is essentially always "newer" than whatever's in the local
`Manifest\manifest.json` (which only updates when `UpdateBootstrapper` actually adopts a
download) - even though the version NUMBER hasn't changed. So the very next Excel launch
after any rebuild will exercise the full download-and-extract path, re-fetching and
re-extracting content that's actually byte-identical to what STEP 1 already xcopy'd
directly. This is redundant but harmless (idempotent end state), and is actually useful
for testing - it means every local rebuild is a fresh opportunity to verify the
download/extract/adopt code path end-to-end. It is **not** meant to replace the existing
`RibReload_OnClick` hot-reload button for day-to-day Addin.Core iteration - that remains
the fast local dev loop; this local-host path exists specifically to exercise the
production-mimicking update mechanism.

### 15.5 Not independently verified in this environment

No Windows machine, MSBuild, IIS, or .NET Framework runtime was available in this
sandbox (Linux-only, confirmed earlier in this engagement) - none of the following could
be compiled or executed here: `UpdateBootstrapper`'s HTTP/zip-extraction logic,
`GLSense.LocalUpdateHost`'s `HttpListener` server, the new `post_build.cmd` PowerShell
`Compress-Archive`/`Get-FileHash` calls, or the full `AddinModule_OnRibbonLoaded` startup
sequence end-to-end. Everything above was written and reasoned through carefully (grep-
verified references, traced call sites, matched existing patterns already proven to work
elsewhere in this codebase) but genuinely needs a real build + a real Excel launch with
`GLSense.LocalUpdateHost.exe` both running and stopped (to test both the "update found"
and "silently fall back" paths) before being trusted in daily use.

- **Status**: implemented, not yet rebuilt/tested by the user.

## 16. Post-build stops touching Versions\ entirely - UpdateBootstrapper is the only writer

User's explicit direction after 15 shipped: even the direct `xcopy` of freshly-built
DLLs into `%LOCALAPPDATA%\...\GLSense_Logs_New\Versions\vX\` (which both post-build
scripts still did, as an inherent side effect of literally copying files there - not a
forgotten `mkdir`, confirmed by grep) should stop. `Versions\` should now be driven
**exclusively** by `UpdateBootstrapper`.

**This has a real, load-bearing consequence, not just a build-script simplification**:
`GLSense.Loader.Core.dll` was only ever ending up in `Versions\vX\` because
`GLSense.Loader.Core\post_build.cmd` directly copied it there - `GLSense.Addin.Core`
never referenced `GLSense.Loader.Core` at all (confirmed via its `ProjectReference`
list), so Addin.Core's own bin output never contained it. But
`AddinDomainLoader.Load()`'s `CreateInstanceAndUnwrap(typeof(RemoteLoader)...)` call
needs `GLSense.Loader.Core.dll` (RemoteLoader's own assembly) resolvable from the new
AppDomain's `ApplicationBase`, which is set to wherever the zip gets extracted. Simply
deleting Loader.Core's xcopy without addressing this would have silently broken AppDomain
creation the next time `Versions\vX\` got populated from a zip missing that DLL.

**Fix**: added a `ProjectReference` from `GLSense.Addin.Core.csproj` to
`GLSense.Loader.Core.csproj` - not because Addin.Core's code actually uses anything from
Loader.Core, but purely so MSBuild (a) always builds Loader.Core first (real dependency
ordering instead of solution-file-order luck) and (b) automatically copies
`GLSense.Loader.Core.dll` into Addin.Core's own bin output. This means Addin.Core's own
`$(TargetDir)` now already contains everything the zip needs, with no staging/merging of
two separate projects' bin folders required.

**Changed:**
- `GLSense.Addin.Core\post_build.cmd` - removed the entire "STEP 1: Copy to deployment
  location" block (all four `xcopy`s into `Versions\%VERSION%\`) and the now-unused
  `DEPLOY_ROOT` variable. The zip step (formerly zipping
  `%DEPLOY_ROOT%\Versions\%VERSION%\`, which no longer gets populated by this script at
  all) now zips `%CORE_BIN_DIR%` directly - Addin.Core's own build output, which
  (thanks to the new `ProjectReference`) already includes `GLSense.Loader.Core.dll`.
  Renamed the old `VERSION`/`v11.1.0`-style variable usage down to just
  `FILE_VERSION` (e.g. `11.1.0`, no `v` prefix) since the `Versions\vX\` folder-naming
  concern this script used to care about doesn't exist here anymore - `v{FILE_VERSION}`
  is still used for the published zip's filename (`v11.1.0.zip`), matching the naming
  convention asked for earlier.
- `GLSense.Loader.Core\post_build.cmd` - reduced to a no-op stub with a comment header
  explaining why (points at Addin.Core's script as where deployment actually happens
  now). Left the `<PostBuildEvent>` wiring in the `.csproj` in place in case a
  Loader.Core-specific post-build step is needed again later, rather than removing the
  mechanism entirely.

### 16.1 The Reload button also had to change

`RibReload_OnClick` -> `ReloadAddinCore()` (`GLSense\AddinModule.cs`) previously just
called `loader.Load(GlobalsEx.Context)` again, reusing whatever `GlobalsEx.Context
.Version` was set to from the last successful load. Since post-build no longer deposits
fresh DLLs into `Versions\vX\` directly, leaving this unchanged would have made Reload
silently keep reloading stale code forever after a rebuild - the whole point of the
button (test your Addin.Core changes without restarting Excel) would have quietly
stopped working.

**Fix**: `ReloadAddinCore()` now calls `new UpdateBootstrapper().ResolveVersionToLoad
(GlobalsEx.Context)` again, between step 3 (unload the old AppDomain) and step 4 (load
the new one) - same call `AddinModule_OnRibbonLoaded` makes at Excel startup (15.1). If
it resolves `null` (no usable local install and the host unreachable), Reload now shows
an error and aborts rather than proceeding with a stale/missing version.

**Real operational requirement this creates, not a bug**: because `GLSense
.LocalUpdateHost`'s published manifest gets a fresh UTC `releaseDate` on every build
(14.4's "known accepted side effect", still true), `UpdateBootstrapper` will see a
"newer" remote release almost every time Reload or Excel-startup runs after a rebuild -
meaning **`GLSense.LocalUpdateHost.exe` must be running** for a rebuild to actually reach
either Excel startup or the Reload button now. If it isn't running, both paths silently
fall back to whatever's already extracted in `Versions\vX\` (by design - `Update
Bootstrapper`'s own "never block on a network hiccup" behavior from 15.1) - which will
look exactly like "my code changes aren't showing up" if the host is forgotten. This is
the single most important new habit for local dev iteration going forward: **start
`GLSense.LocalUpdateHost.exe` before opening Excel, and leave it running.**

**UPDATE - superseded by section 17**: the paragraph above describing
`GLSense.LocalUpdateHost.exe` as a required running process is no longer accurate -
that project and the whole remote-check step were removed (real connection failures
during testing, logged in 17). The rest of this subsection (`ReloadAddinCore()` now
re-running `UpdateBootstrapper` before reloading) is still exactly correct and
unaffected - `UpdateBootstrapper.ResolveVersionToLoad` just no longer has a network
branch to fail, so Reload now works purely off whatever's in the Manifest folder.

### 16.2 Zip excludes *.pdb - they're dead weight in the download

User asked why `*.pdb` (debug symbol) files were being zipped and published at all -
they aren't needed to run the add-in, just unnecessary size in the download. The
original one-line `Compress-Archive -Path '%CORE_BIN_DIR%\*'` zipped the whole bin
output indiscriminately.

**Fix**: `GLSense.Addin.Core\post_build.cmd` now mirrors `%CORE_BIN_DIR%` into a small
temp staging folder first (`xcopy /Y /E /I /EXCLUDE:<generated exclude-list file
containing ".pdb">`), preserving the `x86\`/`x64\`/`de\`/`runtimes\` subfolder structure,
then zips the staging folder instead of `%CORE_BIN_DIR%` directly - `Compress-Archive`
has no clean built-in way to exclude a pattern while zipping a whole folder without
either flattening the subfolder structure (piping filtered `FileInfo` objects into it
does this) or hand-rolling per-file compression, so a filtered copy first was the
reliable option. Both the exclude-list file and the staging folder are deleted after the
zip is created.

**Not a two-project merge** (that concern - `GLSense.Loader.Core.dll` needing to be in
the zip - was already solved in 16 via the `ProjectReference` addition, so `%CORE_BIN_DIR
%` alone already has everything needed). This staging step is single-source and exists
solely for the `.pdb` filter.

**Local debugging is unaffected**: this only changes what's inside the *published* zip.
The original `.pdb` files in `%CORE_BIN_DIR%` (i.e. `GLSense.Addin.Core\bin\%CONFIG%\`)
are untouched, so attaching a debugger / breakpoints in Visual Studio during local
development still works exactly as before.

### 16.3 Not independently verified in this environment

Same caveat as 15.5 - no Windows/MSBuild/.NET Framework runtime available here. The
`ProjectReference` addition, the rewritten `post_build.cmd` (particularly zipping
`%CORE_BIN_DIR%` directly and confirming `GLSense.Loader.Core.dll` actually lands there
via the new reference), and the `ReloadAddinCore()` change all need a real clean rebuild
+ a real Reload-button click to confirm before being trusted.

- **Status**: implemented, not yet rebuilt/tested by the user.

## 17. Local-host removed - back to a folder-only test flow

User hit `UpdateBootstrapper`'s "no local install and remote unreachable" error in a real
run (`WebException: Unable to connect to the remote server` - i.e. `GLSense
.LocalUpdateHost.exe` wasn't running). Rather than keep debugging the "remember to start
a separate console app" workflow, the user's call: **remove the local-host/HTTP path
entirely** and go back to the simplest possible mechanism - post-build drops a zip +
manifest.json directly into the real, local `Manifest` folder; `UpdateBootstrapper` sees
both and extracts. This is explicitly a **temporary simplification for testing** ("once
testing is completed we will come back to real-time development") - the online/remote
tier from section 15 is expected to return later, informed by whatever's learned here.

**Removed entirely:**
- `GLSense.LocalUpdateHost` project - deleted (`Program.cs`, its `.csproj`, `App.config`,
  `Properties\AssemblyInfo.cs`, `wwwroot\`) and removed from `GLSense.sln` (the `Project`
  entry and its 4 `ProjectConfigurationPlatforms` lines). Deleting the physical files
  required `mcp__cowork__allow_cowork_file_delete` (a plain `rm -rf` failed with
  "Operation not permitted" on this mounted folder) - granted once, then the delete
  succeeded.
- `UpdateBootstrapper.cs`'s entire remote branch: the `RemoteManifestUrl` constant,
  `CheckRemoteAndUpdateIfNewer`, `DownloadAndAdopt`, `DownloadZipAndAdopt`, the
  `RunSync<T>` deadlock-safe-HTTP helper, and the `System.Net.Http`/`System.Threading
  .Tasks`/`GLSense.Shared` (for `VersionParser`) usings that only existed to support them.

**`UpdateBootstrapper.ResolveVersionToLoad` is now a strictly 3-branch, folder-only
decision tree** (no network branch at all):
1. Manifest folder missing -> `null` (defensive, same as always).
2. Manifest folder has both `manifest.json` and a `.zip` -> extract into
   `Versions\V{version}\`, delete the zip, adopt. **This is now the only way a version
   ever gets installed or updated.**
3. Manifest folder has only `manifest.json` -> if `Versions\V{version}\` already has
   DLLs, use them; otherwise `null` - there's no remote fallback left to try.

**`GLSense.Addin.Core\post_build.cmd` rewritten again**: STEP 2 no longer publishes to
`GLSense.LocalUpdateHost\wwwroot\GLSense\` (that path no longer exists) - it now writes
directly into the REAL, local `%LOCALAPPDATA%\...\GLSense_Logs_New\Manifest\` folder:
`v{FILE_VERSION}.zip` (same staged, `.pdb`-excluded zip as before - the zip-building
mechanics from 16.2 are unchanged, only the destination moved) and an overwritten
`manifest.json` (`downloadUrl` is now always empty - nothing downloads this locally, the
zip sits right next to it). The `if not exist "%MANIFEST_DIR%" mkdir` here is a
**deliberate, explicitly-temporary override** of section 14.5's "build shouldn't create
`GLSense_Logs_New` folders, `PathProvider.cs` owns that" rule - scoped only to this
folder, only for this local testing setup, called out in the script's own header
comment so it doesn't get mistaken for a regression of 14.5 later.

**What did NOT need to change**: `AddinModule_OnRibbonLoaded` and `ReloadAddinCore`
(`GLSense\AddinModule.cs`) - both already just call `UpdateBootstrapper
().ResolveVersionToLoad(...)` and check for `null`, with no awareness of *how* that
method resolves a version. `PathProvider.Refresh()`/`WriteManifest()` (15.2) - left in
place as generically useful primitives even though nothing currently calls
`WriteManifest` (tier 2's already-correct local `manifest.json` never needs rewriting
after an extraction - it was already describing the version that just got installed).
`GLSense.Loader.Core\post_build.cmd` - already a no-op stub from section 16, still
accurate, no change needed.

**Net effect on the local dev loop**: rebuild `GLSense.Addin.Core` -> a fresh zip +
manifest.json land in the Manifest folder automatically -> restart Excel (or click
"Reload Add-in") -> `UpdateBootstrapper` extracts and loads it. No separate process to
remember to start, no network call, no port to worry about.

- **Status**: implemented, not yet rebuilt/tested by the user.

---

## 18. Journal drilldown (`DD_JL.cs`): "Multi Select" broken by parsing only the last row (fixed in **both** AIPowered and FinalWorkingCode)

User-reported bug, found in both codebases: `DrilldownJl.Journal_DrillDown`
(`Drilldowns\DD_JL.cs`) builds the actual drilldown payload correctly (`BuildDrilldownList`
iterates every selected row and adds one `JournalsQuerySubmit` per row to the list sent to
the server - that part was never broken), but the separate `consolidated`/`jobName` string
sent alongside the payload was built like this, unconditionally:

```csharp
string consolidated = BuildConsolidatedObject(worksheetName, journalsAddress, _ddType,
    BuildMultiStringSafe(lastValue, lastJournalValues));
```

`lastValue`/`lastJournalValues` are `BuildDrilldownList`'s `out` parameters - they hold
only the LAST cell processed in its `foreach (Range loopCell in journalsRange.Cells)`
loop, overwritten every iteration. So whenever more than one row was selected ("Multi
Select"), the jobName string was built entirely from one arbitrary row's field values
(period/ledger/currency/etc.) as if it were the only row - misleading at best, and exactly
the "parsing the last value... breaks the existing functionality" symptom the user
described.

The correct pattern already existed twice elsewhere in this same subsystem, just never
applied here:
- `DrilldownBl.GetDrilldownTitle` (`DD_BL.cs`): `if (totalCount >= 2) return "Multi
  Select.";` - skips per-value parsing entirely once more than one cell is selected.
- `DrilldownSl.Subledger_DrillDown` (`DD_SL.cs`, **already correct, no fix needed
  there**): `string? multiString = strBuilder.Count >= 2 ? "Multi Select" :
  strBuilder.ElementAtOrDefault(0);` - same idea, one line.

**Fix** (`DD_JL.cs`, `Journal_DrillDown`): compute `multiString` conditionally instead of
unconditionally calling `BuildMultiStringSafe` on the last row's values:

```csharp
string multiString = selectedCount >= 2
    ? "Multi Select."
    : BuildMultiStringSafe(lastValue, lastJournalValues);

string consolidated = BuildConsolidatedObject(worksheetName, journalsAddress, _ddType, multiString);
```

Also cleaned up `BuildWorksheetName` (same file): it took `lastValue`/`lastJournalValues`
params and called `BuildMultiStringSafe` on them in its `else` branch purely to assign the
result to a local that was then discarded (`_ = multiString;`) - dead code, likely a
leftover from an earlier refactor. Simplified its signature to just
`BuildWorksheetName(Excel.Worksheet sheet, int selectedCount)`, dropping the unused
params/dead computation; its actual logic (append `" +"` to the worksheet name suffix when
`selectedCount >= 2`, otherwise don't) was already correct and unchanged.

Applied identically to `GLSense.Addin.Core\Drilldowns\DD_JL.cs` (AIPowered) and
`GLSense\Drilldowns\DD_JL.cs` (FinalWorkingCode) - user explicitly reported this as a bug
in both codebases. `DD_SL.cs` was checked in both codebases per the user's request
("Also same is applicable for DD_SL.cs if needed update here also") - it already has the
correct `Count >= 2 ? "Multi Select" : ...` pattern in both, so **no change was needed
there**.

- **Status**: implemented in both codebases, not yet rebuilt/tested by the user.

---

## 19. GLAbout: version/build date wired to manifest.json, always-local timestamps, clearer instance-compatibility logging (AIPowered-only)

Three related requests in one pass, all scoped to AIPowered - GLAbout is a manifest-
driven feature that doesn't exist in FinalWorkingCode (which hardcodes
`AppConstants.DefaultVersion`/`DefaultCommitDate` instead).

### 19.1 Build Date was always blank - `Context.ReleaseDate` was never actually set

`GLAbout.SetVersionAndBuildDateText()` has always read `ServiceLocator.Version` /
`ServiceLocator.ReleaseDate` (`IGLSenseContext.ReleaseDate` has existed since section 14),
but nothing anywhere in the codebase ever assigned `Context.ReleaseDate` - only
`Context.Version = resolvedVersion` was ever set (`GLSense\AddinModule.cs`,
`AddinModule_OnRibbonLoaded` and `ReloadAddinCore`). So the About window's "Build Date"
line was always blank/`"Unknown"` in every session across this whole engagement.

**Fix**: right after each `GlobalsEx.Context.Version = resolvedVersion;` line in both
methods, added `GlobalsEx.Context.ReleaseDate = GlobalsEx.Context.Paths?.LatestReleaseDate;`
(+ a debug log of both values). `Paths.LatestReleaseDate` is the exact manifest.json entry
`UpdateBootstrapper` just used to resolve `resolvedVersion` (same object, see 19.3 below
for why that's now guaranteed fresh), so Version and ReleaseDate are always describing the
same manifest entry.

### 19.2 manifest.json's `releaseDate` is now always local time, never UTC

`GLSense.Addin.Core\post_build.cmd` wrote `releaseDate` via
`[DateTime]::UtcNow.ToString('yyyy-MM-ddTHH:mm:ssZ')` - UTC, with a trailing `Z` (Zulu/UTC
marker). Per request, changed to `[DateTime]::Now.ToString('yyyy-MM-ddTHH:mm:ss')` - local
time, and the misleading `Z` suffix dropped (keeping it while writing local time would
have been actively wrong, not just imprecise). `PathProvider.CreateDefaultManifestFile()`
(the seed manifest written the very first time no `manifest.json` exists yet) was already
using `DateTime.Now` (local, correct), but a different format (`"dd-MMM-yyyy"`, date-only)
- aligned it to the same `"yyyy-MM-ddTHH:mm:ss"` shape so every manifest.json writer in the
codebase now produces one consistent, parseable, always-local format.

### 19.3 `UpdateBootstrapper` now refreshes the manifest cache before reading it

Discovered while wiring 19.1: `PathProvider.LatestVersion`/`LatestReleaseDate` are cached
`static` fields, only re-parsed from disk when `Refresh()`/`InitializeVersion()` runs (the
constructor calls it once). The `PathProvider` instance behind `GlobalsEx.Context.Paths` is
created exactly once, in `GLSenseContext`'s constructor, at `AddinModule_OnRibbonLoaded`
time. On a manual Reload (`RibReload_OnClick` -> `ReloadAddinCore`), if the user rebuilt
`GLSense.Addin.Core` in between (a very normal thing to do - that's the whole point of the
Reload button), a fresh `manifest.json` with a new version/releaseDate would already be
sitting on disk, but `ExtractLocalZipAndAdopt`'s `paths.LatestVersion` read would still
return the STALE value cached at Excel startup - meaning Reload could adopt/report the
wrong version and (now) the wrong build date too.

**Fix**: `UpdateBootstrapper.ResolveVersionToLoad` now calls `paths.Refresh()` as the very
first thing in its `try` block, before any `Directory.Exists`/`File.Exists`/`Latest*` reads
- guarantees every subsequent read in that method (and, transitively, whatever
`AddinModule.cs` copies into `Context.Version`/`Context.ReleaseDate` right after) reflects
whatever is actually on disk at that exact moment, not a startup-time snapshot. Cheap and
idempotent (`Refresh()` just re-parses the JSON), no behavior change for the common case
where nothing changed between checks.

### 19.4 GLAbout displays the build date in a friendly, parsed format

`SetVersionAndBuildDateText()` now runs `ServiceLocator.ReleaseDate` through a new
`FormatBuildDate()` helper: parses the local `"yyyy-MM-ddTHH:mm:ss"` string with
`DateTimeStyles.None` (deliberately NOT `AssumeUniversal`/`AdjustToUniversal` - it's
already local, parsing should not shift it again) and re-formats as
`"dd-MMM-yyyy hh:mm tt"` (matches FinalWorkingCode's `AppConstants.DefaultCommitDate`
date style, `dd-MMM-yyyy`, with an added time component since build times matter for this
hot-reload dev loop). Falls back to showing the raw string as-is (not "Unknown") if it
doesn't parse, so a legacy/unexpected manifest value is still visible rather than hidden.

### 19.5 Clearer, categorized logging in `CheckInstanceCompatibility`/`CheckUrlCompatibility`

Per request: logs should clearly say *why* an instance check failed - not reachable,
generic error, security issue, or certificate error - instead of one generic bucket.
Previously, `CheckUrlCompatibility`'s `HttpRequestException` catch logged everything as
`"Network error for {url}: ..."` regardless of whether the real cause was a DNS/connection
failure or a TLS/certificate rejection (`StrictCertificateValidator.Validate` already logs
detailed chain-status lines when it rejects a cert, but the outer exception handler here
never connected its own log line to that same category).

**Fix**: added `LogInstanceCheckFailure(url, HttpRequestException ex)` - walks the
exception's `InnerException` chain to the root cause and logs one clearly-labeled line:
- `AuthenticationException`/`CryptographicException`, or a message mentioning
  certificate/SSL/TLS/trust/authentication -> `"Certificate/security error connecting to
  {url}: ..."` (points back at `StrictCertificateValidator`'s own detailed log lines).
- `SocketException` -> `"Instance not reachable at {url}: {SocketErrorCode} - ..."`.
- anything else -> `"Network error connecting to {url}: ..."` (previous generic message,
  kept as the final fallback).

Also split the existing `TaskCanceledException` catch: a genuine timeout now logs
`LogError("Instance not reachable (request timed out) for {url}...")`, while an explicit
`CancellationToken`-driven cancellation (`ex.CancellationToken.IsCancellationRequested`)
logs `LogWarn("...was cancelled...")` instead - a user-cancelled check isn't the same thing
as an unreachable instance. The final catch-all now includes `ex.GetType().Name` in its
message (`"Unexpected error checking instance {url}: {TypeName}: ..."`) so even a truly
unanticipated exception type is identifiable at a glance in the log.

- **Status**: implemented, not yet rebuilt/tested by the user.

---

## 20. Regression: GLConfiguratorViewModel was baking resolved values into formulas instead of preserving References (fixed in **both** AIPowered and FinalWorkingCode)

User-reported regression, explicitly against FinalWorkingCode (and applicable to both, per
"this needs to be applicable for both code bases"): "when building any formula, the top
priority is for the references, but in GLConfiguratorViewModel it is replacing the
references with the [resolved] values."

### Root cause

`GetFieldValue(FieldBinding field)` has two very different jobs pulling in opposite
directions, and section 2.2's fix (Journal Source/Category always disabled) made it always
do the second one:
1. **Business-logic/validation** call sites (`IsJournalValidationSatisfied`, enable-state
   checks, `IsBtJournalsType`, etc.) need the field's actual CURRENT value - if
   Activity/BalanceType/CurrencyType is bound via Reference, code comparing it against
   literal tokens like `"PTD"`/`"Actual"` needs the resolved value, not a raw cell address.
   This is exactly what 2.2 fixed, and that fix is correct and still needed.
2. **Formula-building** call sites (`BuildFormulaArguments`, `GetFinalPeriodValue`,
   `GetBudgetEncumbranceValue`, `GetAccountSegments` - everything that ultimately feeds
   `WriteFormulaToCell`) need the RAW, unresolved reference when one is set, so
   `FormatFormulaArg` (which deliberately leaves a value unquoted when it contains `"!"` or
   `"$"`) writes a live cell reference into the formula instead of a hardcoded literal.
   Baking in today's resolved value here means the formula stops recalculating the moment
   the referenced cell's value changes later - silently defeating the entire point of
   letting a field be bound via Reference instead of the ComboBox.

2.2's fix made `GetFieldValue`'s "Step 1: Highest priority — RefValue" branch always resolve
through `GetRangeValueSafe()` and return the resolved value (falling back to the raw ref
text only if resolution fails) - correct for job #1, but every formula-building call site
in the same file also calls `GetFieldValue()`, so job #2 silently broke at the same time:
every Reference-bound field's formula argument became a quoted, hardcoded snapshot instead
of a live reference.

### Fix

Added a new method, `GetFormulaFieldValue(FieldBinding field)`, used ONLY by the formula-
argument builders: returns the raw, unresolved `field.RefValue` outright when set (Reference
wins, unresolved), and only falls through to `GetFieldValue()`'s ComboValue/multi-select/
model-property handling when `RefValue` is blank. `GetFieldValue()` itself is untouched -
every validation/business-logic call site keeps resolving references, which is still
correct there.

Switched to `GetFormulaFieldValue` at every formula-building call site in
`GLConfiguratorViewModel.cs`:
- `BuildFormulaArguments()`: Ledger, Activity, BalanceType, Currency, CurrencyType,
  ActualFlag, JournalSource, JournalCategory.
- `GetFinalPeriodValue()`: StartDate, EndDate, EndPeriod, Period (all feed
  `CombinePeriod`/the formula directly).
- `GetBudgetEncumbranceValue()`: Budget, Encumbrance.
- `GetAccountSegments()`: AccountAssignment.

Also fixed the "Periods" token inside `CollectAllFieldValues()` (the read-only Parameter
Summary panel, not the formula itself) - it already had its own explicit "prefer
references, then values" comment and calls `FormatSummaryToken` (which has the identical
"prefer Excel references (unquoted) otherwise quote literal values" contract), but was
feeding it `GetFieldValue`'s already-resolved value, so `FormatSummaryToken`'s own
`IsRealRange` check could never see a reference to prefer. Switched
StartDate/EndDate/Period/EndPeriod to `GetFormulaFieldValue` there too so the summary
actually honors its own stated contract. Every OTHER field in that same summary
(`Ledger`/`Activity`/`BalanceType`/etc.) intentionally keeps `GetFieldValue`'s resolved
value - that panel is a preview of current values for everything except Periods, which was
the one place explicitly designed to show a reference instead.

### Audit of every other "reference vs. value" formula-builder (per "check all other areas
where we are writing the format of the Excel cell")

Every other `WriteFormulaToCell`-owning ViewModel with its own `RefValue`-aware
`GetFieldValue` was checked for the same bug pattern in both codebases:
`GLGetPeriodModel.cs`, `GLGetPeriodByYearModel.cs`, `GLDailyRatesViewModel.cs`,
`GLPeriodByDateModel.cs`, `GLPeriodDetails.cs`, `GLSegmentFuncsViewModel.cs` (the last one
also has its own explicit comment block confirming the correct reference-first intent, from
the ACCOUNTTYPE work in section 13). All six already return the raw, unresolved
`refVal.Trim()` directly - none of them were ever changed to resolve-through-Excel the way
`GLConfiguratorViewModel.GetFieldValue` was in section 2.2, so **no code change was needed
in any of them** - `GLConfiguratorViewModel.cs` was the only place this regression existed,
in both codebases.

- **Status**: implemented in both codebases, not yet rebuilt/tested by the user.

---

## 21. Balance Configurator pane: focus stealing, double-click crash, and segment-window resize-on-add (all AIPowered-only)

User report (paraphrased from garbled speech-to-text): with the Balance Configurator task
pane open, (a) typing into an Excel cell types into the pane's own controls instead, (b)
double-clicking a balance-formula cell throws an exception and clears the pane's data
instead of refreshing it, and (c) opening GLSegmentManager and adding values to the right
grid visibly resizes the window on every add, when it should only size once at load. Root
causes and fixes below - all three are independent bugs discovered together.

### 21.1 Keyboard focus stealing from Excel cells

**Architecture**: the Balance Configurator pane is not a normal WPF-hosted-in-WinForms
control. `GLSense.Addin.Core.Views.ConfiguratorPaneHost` creates a real, borderless,
top-level WPF `Window` on its own dedicated WPF thread
(`Utilities.WpfAppManager.InvokeOnWpfThread`), and only hands back its native HWND to
`GLSense\GLConfiguratorPane.cs` (a WinForms `ADXExcelTaskPane`), which then Win32
`SetParent`s that HWND into itself, rewrites its window style bits
(`WS_POPUP`->`WS_CHILD`), and keeps it sized via `MoveWindow`. To make Tab/keyboard
navigation flow into this reparented content at all, `GLConfiguratorPane.EmbedContent()`
calls `AttachThreadInput` between Excel's own main UI thread and Addin.Core's dedicated
WPF thread.

**Root cause**: `AttachThreadInput` merges which window is considered "focused" across the
two attached threads, but does nothing to hand focus back to Excel when the user clicks a
worksheet cell - once the WPF pane content takes Win32 keyboard focus, it simply stays
there, so keystrokes meant for the Excel cell go to whichever pane control last had focus.

**Fix** (`D:\SQLLite_Test\AIPowered\GLSense\GLSense\AddinModule.cs`): added a `SetFocus`
P/Invoke declaration, and inside the existing `adxExcelAppEvents1_SheetSelectionChange`
handler's `if (blpane != null && blpane.Visible)` block (the same block that already calls
`RelaunchPane()`/`ResetPaneReference()` to keep the pane synced to the active cell), added
an unconditional `SetFocus(excelHwnd)` call using `GlobalsEx.Context?.ExcelHandle`. This
runs on the exact same thread `EmbedContent()` originally attached via `AttachThreadInput`,
so it reliably wins the focus hand-back.

- **Status**: implemented (AIPowered only). Reasoned fix grounded in well-documented Win32
  `AttachThreadInput` behavior, but not yet empirically verified against a live repro (no
  live Excel session available in this environment) - flag for the user to specifically
  confirm typing into cells works correctly with the pane open.

### 21.2 Double-click / selection-change cross-thread crash while pane is open

**Repro path from the log** (`GLSense_Logs_22-Jul-2026.log`): selecting a cell while the
Configurator pane is already visible re-invokes `GLBalanceConfigurator.ReLoadConfigurator`
via `adxExcelAppEvents1_SheetSelectionChange`'s "keep pane synced to active cell" logic (the
same block touched in 21.1). Unlike the pane's *initial* open, this re-entry path's `await`
chain is not guaranteed to resume back on the WPF UI thread - this project's WPF host does
not reliably install a `DispatcherSynchronizationContext` (documented in sections 2.4/2.5),
so continuations can land on an arbitrary ThreadPool thread.

**Crash**: `GLConfiguratorViewModel.ProcessSingleLedger` (and structurally identical
methods) mutate `GenericLedgerModel.IsSelected` via reflection
(`GetType().GetProperty("IsSelected").SetValue(...)`) to drive multi-select Ledger state.
This synchronously fires `PropertyChanged` -> `Ledger_PropertyChanged` ->
`UpdateParameterSummary()` -> a WPF `DependencyObject.SetValue` call. When that chain runs
off the UI thread (i.e. during this re-entrant reload), it throws
`InvalidOperationException: The calling thread cannot access this object because a
different thread owns it` inside `Dispatcher.VerifyAccess()` - confirmed by the exact stack
trace in the log (`ProcessSingleLedger` -> `HandleLedgerValue` -> `ProcessLedgerFieldAsync`
-> `ApplyFormulaParamsAsync` -> `LoadConfiguratorAsync` -> `ReLoadConfigurator`). The
unhandled exception is what aborts the reload and leaves the pane's data cleared.

**Fix** (`GLSense.Addin.Core\ViewModels\GLConfiguratorViewModel.cs`): wrapped every
previously-unguarded reflection-based `IsSelected` mutation block in
`await _dispatcher.InvokeAsync(...)`, matching the pattern already established (and
documented in sections 2.4/2.5) for `SetLedgerField`:
- `ProcessEmptyValue()` - the "reset all Ledgers' IsSelected to false" loop.
- `ProcessMultipleLedgers(string value)` - the reset + name-matching + multi-select block.
- `ProcessSingleLedger(string value)` - the reset + name-matching + single-select block
  (captures the matched ledger name in an outer local so the follow-up `SetLedgerField`
  call can branch on it after the dispatched block completes).
- `ProcessEncumbranceField(...)` - the same `EncumbranceModel.IsSelected` reflection
  pattern, in both its multi-select (`;`-delimited) and single-select branches (its
  pre-existing "reset to false" loop was already correctly wrapped; only the
  match-and-select part was missing the wrap).

Audited (via grep for `isSelectedProp`/`GetProperty("IsSelected")`) for any other
unwrapped call site: `SetLedgerField` was already fully wrapped, and the generic
`ApplyRefHelper<T>` static helper is only ever invoked directly from WPF UI events (a
different, always-on-UI-thread code path), so it was deliberately left untouched.

- **Status**: implemented, AIPowered only (per the bug report's explicit scope). Not yet
  ported to FinalWorkingCode - not requested.

### 21.3 GLSegmentManager (and GLSegmentRef/GLSegmentValues) resizing on every grid item add

**Root cause**: all three master-detail segment picker windows
(`GLSegmentManager.xaml.cs`, `GLSegmentRef.xaml.cs`, `GLSegmentValues.xaml.cs`) construct a
shared `SegmentSelectorViewModel` and wire an identical
`DataLoadedAction = () => { ForceSizeToContentResettle(); PumpDispatcherFrame(); }` lambda,
originally meant to fire once after the initial async load populates the (otherwise empty)
grids - see sections 1.4b/1.4c. But `DataLoadedAction` is invoked unconditionally from
`SegmentSelectorViewModel.UpdatePagingAndGrid()`, which is ALSO the shared choke point that
`SelectedItemsRight`'s property setter funnels through on every interactive add/remove to
the right-hand grid - not just the initial load. `SegmentSelectorViewModel`'s own comment
assumes this is a "cheap no-op" on later invocations, but for these windows the right grid's
row count genuinely changes the window's `SizeToContent`-measured size, so every add/remove
was triggering a real, visible resize.

**Fix**: rather than changing the shared ViewModel (which all three windows depend on),
added a `private bool _hasResettledAfterInitialLoad;` guard field to each of the three
Views, and changed each one's `DataLoadedAction` lambda to check-and-set that guard before
calling `ForceSizeToContentResettle()`/`PumpDispatcherFrame()`, so the resettle now only
ever fires once, on the first population after load:

```csharp
DataLoadedAction = () =>
{
    if (_hasResettledAfterInitialLoad) return;
    _hasResettledAfterInitialLoad = true;

    ForceSizeToContentResettle();
    PumpDispatcherFrame();
}
```

Applied identically to all three: `GLSegmentManager.xaml.cs`, `GLSegmentRef.xaml.cs`, and
`GLSegmentValues.xaml.cs` (the latter is `SizeToContent="Manual"` with a fixed width, so the
bug was less visible there, but the same guard was added for consistency since it shares
the exact same ViewModel/lambda pattern).

- **Status**: implemented, AIPowered only (these windows don't exist in FinalWorkingCode in
  this form). Not yet rebuilt/tested by the user.

---

## 22. Three usability/messaging bugs: Insert wording, masked Segment-Function validation messages, missing no-balance-formula guard on Reset (all fixed in **both** AIPowered and FinalWorkingCode)

### 22.1 "Please add values before clicking OK" on screens with no OK button

`GLSegmentValues.xaml.cs` and `GLRollerGroups.xaml.cs` (both codebases) validate that at
least one item was added to the right-hand grid before writing to Excel, but the warning
text said "...before clicking OK." - these windows only have an **Insert** button
(`btnOK`'s `Style="{StaticResource InsertButtonStyle}"` - the field/handler names are
legacy, the visible button says "Insert"), so the message was actively misleading. Changed
the wording to "No items added. Please add values before clicking Insert." in all 4 files
(`GLSegmentValues.xaml.cs`/`GLRollerGroups.xaml.cs` x 2 codebases).

- **Status**: implemented in both codebases, not yet rebuilt/tested by the user.

### 22.2 Segment Functions: specific mandatory-field messages were being overwritten by a generic "Failed to write formula to cell."

**Root cause**: `GLSegmentFuncsViewModel.WriteFormulaToCell` (shared by all formula modes
on the single `GLSegmentFunctions` view - ENABLEDFLAG/SUMMARYFLAG/DESCRIPTION/NEXTSEGMENT/
PREVIOUSSEGMENT/DFF/ACCOUNTTYPE) already had a `ValidateMandatoryFields()` check that raises
a specific message via `ShowWarningAction` for each missing required field ("Segment name is
mandatory.", "Segment value is mandatory.", "Attribute is mandatory for DFF formula.") and
returns `false`. But `GLSegmentFunctions.xaml.cs`'s `FormatAndWriteCell` had an
unconditional `else` branch that ran immediately afterward on any `false` return: it called
`AppOverlayControl.ShowWarning("Failed to write formula to cell.")` regardless of what
`WriteFormulaToCell` had already shown - overwriting the specific, correct message in the
shared toast control with a useless generic one. The user only ever saw the generic message
no matter which field was actually missing.

**Fix**: removed the redundant generic `ShowWarning` call from the `else` branch in both
codebases' `GLSegmentFunctions.xaml.cs` (kept the debug log line) - `WriteFormulaToCell`
already raises the correct, specific message on every one of its failure paths (both
`ValidateMandatoryFields`'s missing-field messages and the catch block's actual-exception
message), so the View no longer needs to (and must not) show a second one.

**Also added**: `ValidateMandatoryFields()` never checked `LedgerField` even though
`FormulaParameters()` always appends the ledger value as a formula argument - so leaving
Ledger unselected silently wrote a formula with an empty ledger argument instead of
warning. Added a `"Ledger is mandatory."` check (both codebases), ordered first to match
the UI's top-to-bottom field order (Ledger, Segment Name, Segment Values, Attributes).

Confirmed the "booleans are not mandatory" part of the request needs no code change: the
Next/Previous Parent/Child checkboxes and the Description "parent value" checkbox
(`IsParentChecked`/`IsChildChecked`/`IsParentValueChecked`) were never validated as
required and still aren't - they're optional flags, not mandatory selections.

- **Status**: implemented in both codebases, not yet rebuilt/tested by the user.

### 22.3 Reset Worksheet/Workbook silently "succeeds" on a sheet/workbook with no balance formulas

**Root cause**: `AddinEntry.ResetBalances` (AIPowered) / `AddinModule.ResetBalances`
(FinalWorkingCode) - backing `RibClearSheet`/`RibClear` - only ever checked for broken
workbook links before proceeding straight to `BalancesReset(...)` for the target sheet(s).
Unlike `BalanceRefresh.ExistsBalanceFormulasAsync` (AIPowered, already guards both Sheet and
Book Refresh/Snapshot) and `ValidateHighlightPreconditions` (FinalWorkingCode, guards the
Highlight feature), Reset never called `CommonFunctions.BalanceFormulaExists` at all, for
either Sheet or Book scope - so resetting a sheet/workbook with no balance formulas just
silently did nothing, with no feedback to the user.

**Fix**: added the identical guard to `ResetBalances` in both codebases, right after the
broken-links check and before the actual reset loop - for `resetType == "Sheet"`, checks
`CommonFunctions.BalanceFormulaExists` on the active sheet only; for `"Book"`, checks it
across every worksheet in the active workbook (matching `ExistsBalanceFormulasAsync`'s own
Sheet-vs-Book branching). If none exist, shows `"No balance formulas found in worksheet
\"{name}\"."` / `"...workbook \"{name}\"."` and returns without touching
`BalancesReset` at all.

- **Status**: implemented in both codebases, not yet rebuilt/tested by the user.

---

## 23. Refresh/Reset: move precondition checks before the progress/wait window appears (both codebases)

Follow-up to section 22.3. The precondition checks were correct in what they checked, but
ran in the wrong place relative to the progress/wait window - the user would see a window
flash up only to be immediately dismissed by an error/warning message. Reordered so the
window only ever appears once every precondition has already passed.

### 23.1 Sheet/Book Refresh (and Snapshot, which shares the same code path)

`BalanceRefresh.RefreshBalancesInternalAsync` (both codebases) used to run in this order:
`InitializeAsync` -> `CreateAndShowProgressWindow` -> `InitializeProgressWindowAsync` ->
`ValidateWorkbookIsSavedAsync` -> `ValidateNoBrokenLinksAsync` -> (snapshot path prompt) ->
`ValidateBalanceFormulasExistAsync` -> proceed. Moved `ValidateWorkbookIsSavedAsync` and
`ValidateBalanceFormulasExistAsync` to run immediately after `InitializeAsync`, BEFORE
`CreateAndShowProgressWindow` - so the window is only created once both "is the workbook
saved" and "do balance formulas exist" have already passed. `ValidateNoBrokenLinksAsync` and
the snapshot-path prompt were left exactly where they were (after the window appears) -
only the workbook-saved and balance-formulas-exist checks were explicitly called out for
reordering. Both validators internally call `MessageProgressWindowAsync`, which is already a
safe no-op when `Win` is still null (its own null-check), so no other change was needed to
make them safe to call pre-window.

Since `RefreshingBalancesAsync("Snapshot", "Sheet"/"Book")` (RibSnapWorksheet/RibSnapWorkbook)
routes through this same `RefreshBalancesInternalAsync` method (only the separate, still-
unwired `SubmitSnapAsync`/`SubmitSnapshotInternalAsync` "submit to server" path does not),
this reorder applies to Snapshot as well as plain Refresh, for both Sheet and Book scope.

### 23.2 Sheet/Book Reset

`AddinEntry.ResetBalances` (AIPowered) / `AddinModule.ResetBalances` (FinalWorkingCode) -
which already gained a balance-formulas-exist guard in section 22.3 - had that guard placed
AFTER `CommonMethods.DisableExcelSettings()` and the wait window's creation, so a
"Reset Balance Formulas" window still flashed up right before the "no balance formulas
found" warning. Moved the balance-formulas-exist check to the very first thing in the
method (before `AppState.Instance.ResetFormulas = true`/`DisableExcelSettings()`/window
creation), showing the warning directly via `CommonFunctions.GLSenseMessage` (no wait window
exists yet to close first) and returning immediately if formulas don't exist. The
broken-links check remains exactly where it was (after the window appears), matching how
23.1 left `ValidateNoBrokenLinksAsync` alone for Refresh.

- **Status**: implemented in both codebases (23.1 and 23.2), not yet rebuilt/tested by the
  user.

---

## 24. Four more user-reported bugs: GLSegmentManager resize (real fix this time), double-click drilldown crash, popup mouse-wheel scrolling, Currency default (all AIPowered-only)

### 24.1 GLSegmentManager still resizing on every left-to-right grid move (section 21.3's fix was incomplete)

Section 21.3's `_hasResettledAfterInitialLoad` guard on `DataLoadedAction` was a real fix for
its own narrow problem (an extra manual `ForceSizeToContentResettle()`/`PumpDispatcherFrame()`
call firing on every grid mutation) but it could never have fixed the user-visible symptom,
because it doesn't touch the actual root cause: this window's `SizeToContent="WidthAndHeight"`
re-measures itself automatically on every layout pass regardless of any manual resettle call.
The real cause: `GLSegmentManager.xaml`'s "Dual DataGrids" `Border` (`Grid.Row="3"`, which sits
in a `Height="*"` row that collapses to Auto-like behavior under `SizeToContent`) never had a
`MaxHeight`, unlike the exact same "Dual DataGrids" `Border` on `GLSegmentValues.xaml` and
`GLRollerGroups.xaml`, which both already use `MaxHeight="450"` (section 52's fix, from an
earlier session) specifically so `dgLeft`/`dgRight`'s own internal `ScrollViewer`s take over
once content exceeds that cap instead of letting the DataGrid (and therefore the window) keep
growing/shrinking with every row added or removed. `GLSegmentManager` was simply missed when
that convention was established (task #9 built it after task #52's pass, and the two were never
reconciled). Fixed by adding the identical `MaxHeight="450"` to `GLSegmentManager.xaml`'s Dual
DataGrids Border - now dgRight's row count no longer affects the window's measured size at all.

- **Status**: implemented (AIPowered only - `GLSegmentManager` doesn't exist in
  FinalWorkingCode), not yet rebuilt/tested by the user.

### 24.2 Double-click drilldown on a balance formula cell throws InvalidCastException

Log evidence (`GLSense_Logs_22-Jul-2026.log`, Context: `DrilldownBl.Balance_Drilldown`):

```
Type: System.InvalidCastException
Message: Unable to cast object of type 'System.__ComObject' to type 'Microsoft.Office.Interop.Excel.WorkbookClass'.
   at Microsoft.Office.Interop.Excel._Application.get_ActiveWorkbook()
   at GLSense.Addin.Core.Utilities.CommonFunctions.SanitizeSheetName(String raw, Workbook wb)
   at GLSense.Addin.Core.Drilldowns.DrilldownBl.BuildDrilldownInfo(Range balanceRange, String ddType)
   at GLSense.Addin.Core.Drilldowns.DrilldownBl.<Balance_Drilldown>d__48.MoveNext()
```

**Root cause**: `CommonFunctions.SanitizeSheetName(string raw, Excel.Workbook? wb = null)`
unconditionally ran `wb ??= ServiceLocator.ExcelApp?.ActiveWorkbook;` at the top of the method,
even though `wb` is only ever actually used in the empty-`raw` branch immediately below it.
Every real caller (`DrilldownBl.BuildDrilldownInfo`, reached from the `SheetBeforeDoubleClick`
chain) always passes a non-empty, already-built `raw` name, so that fetch ran for no reason on
every single call - and in this specific call path (deep in the double-click drilldown's async
chain, itself already several `await`s removed from the original Excel event), fetching
`ActiveWorkbook` fresh crosses the host<->Addin.Core AppDomain boundary via .NET Remoting
(`CrossAppDomainSink.SyncProcessMessage` in the stack trace) and threw the cast exception,
aborting the whole drilldown.

**Fix**: moved the `wb ??= ServiceLocator.ExcelApp?.ActiveWorkbook;` fetch inside the
`if (string.IsNullOrWhiteSpace(raw))` branch that actually needs it, eliminating the
unnecessary (and in this call path, crash-prone) cross-domain COM fetch entirely for the
common non-empty-name case.

- **Status**: implemented (AIPowered only - the crash is specific to the AppDomain-crossing
  architecture; FinalWorkingCode has the identical eager-fetch code in its own
  `SanitizeSheetName`, same file/line, but can't reproduce this crash since it's a monolith
  with no AppDomain boundary to cross - left untouched there, not requested this round).

### 24.3 SuggestAppendComboBox popup: mouse wheel doesn't scroll the item list

**First attempt (real bug, but NOT the actual reported symptom's cause)**: `Themes/Generic.xaml`'s
`SuggestAppendComboBox` template wrapped `PART_ListBox` in its own separate
`<ScrollViewer VerticalScrollBarVisibility="Auto">`, ON TOP OF the `ListBox`'s own internal
`ScrollViewer`. Nesting a `ScrollViewer` around a control that already hosts its own is a
well-known WPF trap (outer `ScrollViewer` sees `PreviewMouseWheel` first and swallows it). This
was fixed by removing the redundant outer `ScrollViewer` (moving the themed `ScrollBar` style
onto `ListBox.Resources`), and ported to FinalWorkingCode's `Themes\Generic.xaml` too (which had
the identical nested-`ScrollViewer` bug, `MaxHeight="300"`/`"250"` instead of AIPowered's
`"240"`/`"200"`). **User then confirmed this did NOT fix the actual issue** - the real bug only
reproduces on comboboxes hosted inside the Balance Configurator's `ADXExcelTaskPane`, never on
regular windows (`GLSegmentValues`' `cmbHierarchy`, an identical control instance, scrolled
correctly the whole time) - proving the nested-`ScrollViewer` template bug, while real, was never
the (sole) cause of the user's actual symptom. See 24.3.1 below for the real root cause and fix.
The nested-`ScrollViewer` removal is still a legitimate, worthwhile fix in its own right (it
would have caused problems eventually) and was left in place in both codebases.

#### 24.3.1 Second attempt (also did not work): WPF class handler for PreviewMouseWheel

Theorized that `WM_MOUSEWHEEL` was being misdelivered to the reparented/hosting parent window
instead of the open `Popup`'s own HWND, and registered one WPF class handler for
`PreviewMouseWheel` on every `Window` via `EventManager.RegisterClassHandler`, forwarding the
wheel delta into the open popup's `ScrollViewer` whenever the event reached a `Window` unhandled.

**User confirmed this also did not fix it.** That result is itself informative: it means the
wheel message isn't reaching *any* WPF `Window`'s routed-event tunnel at all in this hosting
context - not "reaching the wrong window" as first theorized, but not entering WPF's routed-event
system anywhere. No WPF-level event handler, placed anywhere in any window, can catch a message
that never becomes a WPF routed event to begin with. This second attempt was removed rather than
left in place, since it provably does nothing.

#### 24.3.2 Third attempt (fixed the scroll bug, but caused a new regression): global low-level mouse hook

**Root cause identified**: this matches a separately, independently documented Office Add-in/VSTO
bug (WPF ComboBox popup hosted via `ElementHost` inside an Office task pane - see the Microsoft
Q&A thread "Cant select items in a WPF combobox(VSTO Addin) if they are outside the parent
window"). A WPF `Popup` normally captures the mouse (`Mouse.Capture`) while open so that
clicking/scrolling anywhere is correctly attributed to it. That capture/hit-testing mechanism is
Win32-backed and scoped to the *owning top-level window* (Excel's own main window, once the
content is embedded into its window tree) - once the popup's on-screen position extends outside
that owning window's screen rectangle, which is routine for a narrow, docked task pane's
dropdowns, capture/hit-testing for the portion outside those bounds silently breaks.
`GLSegmentValues` and every other non-task-pane window are their own independent top-level
windows, so their popups' bounds are always inside their own owning window - never affected. The
Balance Configurator's content is embedded inside Excel's window tree in both codebases
(AIPowered's HWND-reparenting bridge in `GLConfiguratorPane.cs`; FinalWorkingCode's plain
`ElementHost` in its own `GLConfiguratorPane.cs`) - exactly the case that breaks.

**Fix tried**: bypass the broken pipeline with a low-level, system-wide mouse hook
(`SetWindowsHookEx(WH_MOUSE_LL, ...)`), installed only while a popup from this control was open,
comparing the hook's raw screen coordinates against the open popup's actual on-screen rectangle
and scrolling its `ScrollViewer` directly when the cursor was over it.

**User confirmed this fixed the scroll issue, but introduced a new regression**: opening
`GLSegmentRef` (segment-value picker launched from the Balance Configurator) and clicking OK no
longer closed the window immediately - it stayed open/unresponsive until another window was
clicked - and general responsiveness felt slower. Cause: `WH_MOUSE_LL` hooks always invoke their
callback **on the thread that installed them**, and because the hook itself is global/system-wide
by Win32 design (thread-scoped installation isn't supported for the `_LL` hook types), *every*
mouse-move message on the entire desktop - not just within Excel, not just over this control - had
to round-trip through that one callback on the shared Addin.Core WPF Dispatcher thread before
Windows could continue processing it. Whenever that thread was even briefly busy (closing
`GLSegmentRef`, populating `GLAccountsRef`), mouse and redraw processing backed up system-wide -
exactly the reported symptom. Removed entirely (both codebases) - see 24.3.3 for the replacement.

#### 24.3.3 Real fix: intercept WM_MOUSEWHEEL at the WinForms task-pane host's own WndProc

The wheel message *is* being delivered natively somewhere - just not to any WPF `Window` (ruling
out attempt 24.3.1) and not in a way that requires a global hook to see (ruling out 24.3.2's
justification for going system-wide). The actual delivery point is the WinForms
`ADXExcelTaskPane` host control itself - `GLConfiguratorPane.cs` - a plain WinForms `Control`,
entirely invisible to WPF's routed-event system, which is exactly why the WPF Window class
handler (24.3.1) could never have worked no matter where it was attached.

**Fix** (`GLConfiguratorPane.cs`, both codebases): this control already overrides `WndProc` for
`WM_SIZING`/`WM_WINDOWPOSCHANGING` (the DPI-aware minimum-size enforcement). It now also catches
`WM_MOUSEWHEEL` there, decodes the screen coordinates (`lParam`) and signed wheel delta (high word
of `wParam`), and forwards them:
- **FinalWorkingCode** (monolith, no AppDomain boundary): directly into a new public static method,
  `Controls.SuggestAppendComboBox.TryScrollOpenPopupAtScreenPoint(screenX, screenY, wheelDelta)`.
- **AIPowered** (Addin.Core lives in a separate AppDomain): via a new `IGLSenseAddin` method,
  `TryScrollOpenComboBoxPopup(screenX, screenY, wheelDelta)`, implemented in `AddinEntry.cs` as a
  thin delegation to a new `ConfiguratorPaneHost.TryScrollOpenComboBoxPopup(...)`, which itself
  delegates to the same `SuggestAppendComboBox.TryScrollOpenPopupAtScreenPoint(...)`.

Both call the exact same logic previously used inside the (now-removed) hook callback - checking
whichever `SuggestAppendComboBox` instance currently has `_openInstance` set, testing the raw
screen point against that popup's actual on-screen rectangle (`FrameworkElement.PointToScreen`),
and scrolling its `ListBox`'s internal `ScrollViewer` directly if it's a hit. If something was
scrolled, `WndProc` returns without calling `base.WndProc`, swallowing the message; otherwise it
falls through exactly as before, completely unaffected.

This is a fundamentally narrower fix than the global hook: it only ever intercepts messages
already being delivered to this one specific, already-owned WinForms control, on the thread that
was already going to process them (Excel's own UI thread) - zero impact on any other window,
thread, or application, and zero risk of the input-lag regression from 24.3.2. It's also
independent of whichever window happens to be shown *inside* the task pane at any given time
(currently `GLSegmentManager` as a stand-in for `GLSegmentRef` in AIPowered - see 24.3.4) since the
fix lives at the task-pane-hosting layer, not in any specific hosted window.

- **Status**: **superseded - proven ineffective.** User confirmed the real symptom: fixing
  GLSegmentRef's focus/slow-close regression (by disabling the 24.3.2 hook and relying solely on
  this WndProc interception) caused the combobox scroll issue to resurface. That result proves
  `GLConfiguratorPane.WndProc` never actually received `WM_MOUSEWHEEL` for the popup at all - with
  the 24.3.2 hook disabled, nothing caught the message, meaning this "fix" was silently dead code
  the whole time it looked like it was working (it wasn't; the 24.3.2 hook, briefly left running
  alongside it during testing, was doing all the actual work). Removed entirely from both
  codebases (the `WM_MOUSEWHEEL` `WndProc` branch, the `IGLSenseAddin.TryScrollOpenComboBoxPopup`
  method, and its `ConfiguratorPaneHost`/`AddinEntry` plumbing). See 24.3.5 for the real fix.

#### 24.3.4 (superseded) Note on GLSegmentManager vs. GLSegmentRef (AIPowered)

This note (originally about 24.3.3 living at the task-pane-hosting layer) is superseded along with
24.3.3 itself. Its conclusion still holds under 24.3.5, and more strongly: the real fix now lives
entirely inside `Controls\SuggestAppendComboBox.cs` in both codebases, with **no** changes to
`GLConfiguratorPane.cs`/`ConfiguratorPaneHost.cs` at all - so it is automatically independent of
which window (`GLSegmentManager`, `GLSegmentRef`, or any future replacement) is shown inside the
Balance Configurator's task pane at any given time.

#### 24.3.5 Real fix: keep the low-level hook, but run it on its own dedicated, non-blocking thread

**Diagnosis**: 24.3.2 (global `WH_MOUSE_LL` hook on the shared WPF Dispatcher thread) was the only
attempt that ever actually saw the wheel message - proven by the fact that disabling it (to test
24.3.3's WndProc interception in isolation) made scrolling stop working again entirely. So the
defect in 24.3.2 was never "hooking is the wrong approach" - it was *which thread* the hook ran
on. `WH_MOUSE_LL` hooks always invoke their callback **on the thread that installed them**, and
since low-level hooks are inherently global/system-wide by Win32 design (thread-scoped
installation isn't supported for the `_LL` hook types), *every* mouse-move message on the entire
desktop had to round-trip through that one callback on the busy, shared Addin.Core WPF Dispatcher
thread before Windows could continue - hence the GLSegmentRef slow-close/focus regression whenever
that thread was doing anything else.

**Fix** (`Controls\SuggestAppendComboBox.cs`, both codebases): install the exact same `WH_MOUSE_LL`
hook, but on its own small, dedicated, otherwise-idle background thread (`EnsureMouseHookThreadRunning`
/ `MouseHookThreadProc`) with its own native Win32 message loop (`GetMessage`/`TranslateMessage`/
`DispatchMessage` - required for any low-level hook to receive callbacks at all). This thread does
nothing else for the lifetime of the process, so Windows always finds it responsive regardless of
how busy the main WPF UI thread is - eliminating 24.3.2's regression while keeping its only
proven-effective mechanism.

Since the hook thread is *not* the WPF Dispatcher thread, it must never touch WPF
`DispatcherObject`s directly (`FrameworkElement.PointToScreen`, `ScrollViewer.ScrollToVerticalOffset`,
etc. all enforce thread affinity and would throw `InvalidOperationException` from a foreign
thread). Three-part design to work around this:
1. The open popup's on-screen rectangle is computed **on the WPF thread**, once, when the popup
   opens (`UpdateOpenPopupScreenRect`, called from `Popup.Opened` alongside the existing
   `_openInstance` tracking), and cached in a `volatile` field as a plain-data `ScreenRect`
   (primitive doubles only, not a WPF type) - safe for the hook thread to read without any
   thread-affinity concern, and without needing a lock (reference assignment is atomic).
2. The hook callback (`LowLevelMouseHookCallback`, running on the dedicated thread) does a pure
   numeric point-in-rect test (`ScreenRect.Contains`) against that cached rect - no WPF calls
   anywhere in the hit-test path.
3. Only on an actual hit does it touch WPF at all, and even then only by marshaling the scroll
   itself onto the WPF thread **asynchronously** via `Dispatcher.BeginInvoke` (`ScrollOpenPopup`) -
   the hook thread never blocks waiting for the (possibly busy) UI thread to catch up, which is
   the whole point of decoupling it onto its own thread in the first place.

`ShutdownMouseHook()` unhooks and signals the dedicated thread to exit (`PostThreadMessage` +
`WM_QUIT`, then a bounded `Thread.Join`) - wired into `AddinEntry.Shutdown()` in AIPowered (called
both at real shutdown and before every hot-reload AppDomain swap, so a stale hook pointing at an
unloaded AppDomain's delegate can never linger) and into `AddinModule_AddinBeginShutdown` in
FinalWorkingCode (a monolith with no hot-reload concept, but still good practice if the add-in is
ever disabled/unloaded without Excel closing).

- **Status**: implemented in **both** codebases, **user-confirmed resolved** - both the
  ADXTaskPane combobox scroll issue and the GLSegmentRef focus/slow-close regression are fixed
  together, with no further side effects reported. This closes out the section 24.3 saga.
  Supersedes 24.3.2, 24.3.3, and 24.3.4 (all three removed/superseded - two proven either inert or
  actively harmful, one narrowed to a note). The original nested-`ScrollViewer` template fix
  (24.3) remains in place in both codebases as a legitimate, independent fix in its own right.

#### 24.3.6 Summary (quick reference)

**Issue**: `SuggestAppendComboBox` popups (the custom autosuggest/multi-select dropdown control
used throughout the app) didn't respond to mouse wheel scrolling, but *only* when hosted inside
the Balance Configurator's `ADXExcelTaskPane` - identical instances in ordinary top-level windows
(`GLSegmentValues`, `GLRollerGroups`, etc.) always scrolled correctly.

**Root cause**: a WPF `Popup` opens as its own top-level HWND; Win32 mouse capture/hit-testing for
that HWND depends on the *owning* top-level window's screen bounds, and silently breaks once the
popup renders outside those bounds - routine for a narrow, docked task pane's dropdowns (a
separately, independently documented Office Add-in/VSTO bug: WPF ComboBox popups hosted via
`ElementHost` inside a task pane). As a result, `WM_MOUSEWHEEL` for the popup is never delivered
through any normal Win32/WPF/WinForms routing path - not to any WPF `Window`, not even to the
hosting WinForms control's own `WndProc` - and is only observable via a raw, pre-routing low-level
mouse hook (`SetWindowsHookEx(WH_MOUSE_LL, ...)`).

**Fix** (`Controls\SuggestAppendComboBox.cs`, both codebases): install that low-level hook, but on
its own small, dedicated, permanently-idle background thread (with its own native Win32 message
loop), instead of the shared WPF Dispatcher thread every other UI operation also runs on. The open
popup's on-screen rectangle is precomputed on the WPF thread and cached as plain primitives so the
hook thread's hit-test never touches WPF objects (which would throw due to thread affinity); a hit
is scrolled by marshaling onto the WPF thread asynchronously (`Dispatcher.BeginInvoke`), so the
hook thread itself never blocks. This is what avoids the earlier regression: a naive hook on the
*shared* UI thread makes every mouse-move message on the entire desktop round-trip through that
one thread before Windows can continue, which stalls input/redraw system-wide the moment that
thread is even briefly busy (e.g. closing `GLSegmentRef`) - the classic symptom being a window that
doesn't visibly close until another window is clicked. Three earlier attempts (removing a nested
`ScrollViewer` in `Generic.xaml`, a WPF `Window` class handler, and intercepting in
`GLConfiguratorPane.WndProc`) were all tried and ruled out first; see 24.3-24.3.5 above for the
full diagnostic trail.

### 24.4 Balance Configurator: Currency doesn't default to the ledger's currency

User confirmed (via clarifying question) this is a genuine bug in the existing default-currency
logic, not a request for a different concept. `ApplyDefaultSelections()` already had
`CurrencyField.ComboValue = Currencies.FirstOrDefault(c => c.CurrencyCode ==
AppState.Instance.SelectedLedger.CurrencyCode);` - but `Currencies` is loaded by
`LoadDataAsync(LedgerRecord ledger)` via `repo.GetCurrencies(cubeId, ledger.LedgerId)`, scoped
to the `ledger` PARAMETER passed into `LoadConfiguratorAsync` - a different reference than the
globally-tracked `AppState.Instance.SelectedLedger` (the ribbon's "currently active ledger",
which is not guaranteed to be the same ledger a given formula/cell's Configurator session is
actually working with). Whenever they diverged, the `FirstOrDefault` matched nothing and
`CurrencyField` was silently left blank instead of defaulting.

**Fix**: added a private field `_activeLedger` (`LedgerRecord?`), set from
`LoadConfiguratorAsync`'s own `ledger` parameter right at entry - i.e. captured from the exact
same reference `LoadDataAsync` uses to populate `Currencies` - and changed
`ApplyDefaultSelections` to match `Currencies` against `_activeLedger?.CurrencyCode` instead of
`AppState.Instance.SelectedLedger.CurrencyCode`. This guarantees the currency lookup is always
against the same ledger the `Currencies` collection was actually populated for.

- **Status**: implemented (AIPowered only). FinalWorkingCode's `GLConfiguratorViewModel.cs` has
  the exact same code (same `LoadConfiguratorAsync(bool, LedgerRecord, ...)` signature, same
  `AppState.Instance.SelectedLedger.CurrencyCode` line) and is very likely affected by the same
  latent bug, but was left untouched since not requested for both codebases this round - worth
  porting this same `_activeLedger` fix there if the same symptom is confirmed on that side too.

---

## 25. Audit: other places that could hit the same cross-AppDomain COM marshaling crash as SanitizeSheetName (AIPowered only)

Following section 24.2's fix, searched the rest of `GLSense.Addin.Core` for the same risk
shape: a COM-object-returning property (`ActiveWorkbook`/`ActiveSheet`, NOT a primitive like
`.Address`/`.Name`) fetched fresh via `ServiceLocator.ExcelApp`/an instance's own cached
`Application` reference, from code reached several `await`s deep in an async chain that
originated from an Excel event (double-click drilldown, background HTTP response handling) -
exactly the shape that threw `InvalidCastException: Unable to cast object of type
'System.__ComObject' to type 'WorkbookClass'` crossing the host<->Addin.Core AppDomain
boundary (`CrossAppDomainSink.SyncProcessMessage`) in `SanitizeSheetName`. Six confirmed
instances fixed; a handful of other `ActiveWorkbook`/`ActiveSheet` call sites were reviewed
and deliberately left alone (see "Not changed" below) because they're reached synchronously,
at the very start of a ribbon-click handler, before any `await` has happened yet - the same
risk profile as every other working call site in this codebase, not the crash-prone one.

### 25.1 DD_BL.cs / DD_JL.cs / DD_SL.cs: `HandleBackgroundProcessingAsync`'s `Names.Add`

All three drilldown classes (`DrilldownBl`/`DrilldownJl`/`DrilldownSl`) have a
`HandleBackgroundProcessingAsync(string msg)` method, reached after the drilldown's HTTP
response has already been awaited, that adds a hidden named range as a job-tracking marker
for the "launch process window to check status" follow-up. All three did
`ExcelApp.ActiveWorkbook.Names.Add(...)` - a fresh fetch, this deep in the chain, of exactly
the kind that crashed in `SanitizeSheetName`. Each of these three classes already caches a
`Workbook` reference at construction time, before any async work begins
(`BlWorbook`/`JlWorbook` via `ExcelExternalRef.ResolveRangeWithContext` in the constructor;
`DrilldownSl` has no dedicated field but keeps the same `ExternalResolveResult.Workbook`).
Switched all three to reuse that already-resolved reference instead of re-fetching, wrapped
in a try/catch that logs a warning (non-fatal) on failure, since this only adds a
supplementary tracking range - not the drilldown itself.

### 25.2 CustomDrilldown.cs: `ProcessCustomDrilldownAsync`

Fetched `ServiceLocator.ExcelApp.ActiveWorkbook` fresh to look up `CustomXMLParts` and find
the drilldown metadata "for this sheet" (`ws`, already passed in as a parameter). Switched to
`ws?.Parent as Excel.Workbook` - this both avoids the risky re-fetch AND is more correct: it
needs the workbook that actually owns `ws`, not whatever happens to be "active" by the time
this continuation runs (which can differ if focus moved to a different window during the
preceding awaits).

### 25.3 BalanceHighlighter.cs: `SelectAdaptiveBalanceRange`

Reached after two awaits from the original ribbon click (`FindAdaptiveMemoryCellsFast`, then
`SafelyCloseWaitWindowAsync` in the `finally` block). Used to fetch
`ServiceLocator.ExcelApp.ActiveWorkbook`/`.ActiveSheet` fresh just to `.Activate()` them (a
no-op if focus never moved) before selecting the found range. Switched to
`adaptiveBalanceRange.Worksheet`/`.Worksheet.Parent` - activates the sheet/workbook that
actually owns the found range instead of blindly re-fetching "whatever is active now", fixing
both the crash risk and a latent correctness gap (the old code did nothing useful if focus
*had* actually moved elsewhere, since it activated the wrong sheet/workbook before selecting).

### 25.4 DDDatatoWorksheet.cs: two independent `ActiveWorkbook` fetches per run

`DD_DatetoWorksheet()` fetched `DD_ExcelApp.ActiveWorkbook` near its start, and
`CreateCustomDrilldownXMLPart` (called later in the same run, from `ApplyFormatting`)
independently fetched it AGAIN - two separate risky cross-domain fetches for what should be
the same workbook within one drilldown-to-worksheet operation. Added a cached `_ddWorkbook`
field and changed both call sites to `_ddWorkbook ??= DD_ExcelApp.ActiveWorkbook` so the second
one reuses whatever the first one already fetched, halving the exposure per run. Left the
constructor's public contract (`DDDatatoWorksheet(Excel.Application, object, string, string,
CancellationToken, GLWaitWindow)`) untouched - it's explicitly documented in this file's own
header as preserved exactly across all 3 call sites (DD_BL/DD_JL/DD_SL), and this class
genuinely does need the currently-active workbook here (the destination for the new
drilldown sheet), not a value that could safely be swapped for a differently-scoped cached
reference the way the other fixes above could.

### Not changed (reviewed, low risk)

- `BalanceHighlighter.RibHighlight_OnClick`'s own `ServiceLocator.ExcelApp.ActiveSheet` fetch
  (line 68) - the very first Excel-touching statement of a method invoked directly and
  synchronously from a ribbon click, before any `await`. Same risk profile as
  `ResetBalances`/`DrillCellHighlighter.RibDrillCells_OnClick` and other working call sites.
- `DrillCellHighlighter.cs`'s `ActiveCell`/`ActiveSheet` fetch - same "first statement of a
  ribbon-click-triggered method" shape.
- `GLLOVs.xaml.cs`'s `ActiveWorkbook.Names.*` calls (`CreateNamedRangeAsync`/
  `DeleteNamedRangeAsync`/`NameRangeExists`) - triggered from this WPF window's own button
  clicks (already on the UI thread, not a continuation resumed from an Excel COM event
  callback), and already wrapped in try/catch with logging, matching this codebase's existing
  defensive-non-fatal convention for this kind of call.

- **Status**: all four fixes (25.1-25.4) implemented, AIPowered only (this is an
  AIPowered-architecture-specific crash class - FinalWorkingCode has no AppDomain boundary to
  cross), not yet rebuilt/tested by the user.

---

## 26. Segment Manager (`GLSegmentManager.xaml`/`.xaml.cs` + `SegmentSelectorViewModel.cs`): six layout/UX polish items from user screenshot feedback (AIPowered only - `GLSegmentManager` doesn't exist in FinalWorkingCode, only its older sibling `GLSegmentRef` does)

User sent four annotated screenshots of the Segment Manager window (opened from the Balance
Configurator's `AcctsRef_EditRequested`) with six specific complaints. All six fixed in one
pass; none required touching `GLConfiguratorPane.cs`/`GLBalanceConfigurator`/anything outside
`GLSegmentManager.xaml(.cs)`, `SegmentSelectorViewModel.cs`, and the shared
`DataGridColumnFillHelper.cs`.

### 26.1 Left DataGrid resized/re-squeezed "Is-Summary" column while scrolling

**Symptom**: window opened fine, but scrolling the left segment-values grid (`dgLeft`) made
the window resize and cut off the "Is-Summary" checkbox column.

**Root cause**: `DataGridColumnFillHelper.Refresh()` (shared by every window that uses it -
see its own header comment) computes the "Description" column's fill width, but its old
fallback - "if available space is less than the column's natural content width, leave it at
`DataGridLength.Auto`" - was the bug. Once left in `Auto` mode, a *virtualizing* `DataGrid`
column keeps re-measuring against whichever rows are currently realized in the viewport. As
the user scrolled to rows with longer `Description` text, the column silently re-grew itself
(no `SizeChanged`/`Loaded` event needed - it's a live remeasure), and since this window uses
`SizeToContent="WidthAndHeight"`, the whole window resized to match, squeezing the
fixed-width "Is-Summary" column (which sits after "Description") out of view.

**Fix**: `Refresh()` now always resolves the fill column to a concrete pixel
`DataGridLength` - clamped to a `MinFillColumnWidth` (80px) floor instead of ever falling back
to `DataGridLength.Auto`. A pinned numeric width doesn't re-measure against scrolled-in rows,
so the column (and the window) stays stable regardless of what's scrolled into view; if
content is genuinely wider than the floor, the grid scrolls horizontally instead (matching the
original fallback's intent, just bounded). This is a shared-helper fix, so it also hardens
every other window using `DataGridColumnFillHelper` (GLSegmentValues, GLRollerGroups, GLLOVs,
etc.) against the same latent failure mode, not just GLSegmentManager.

### 26.2 Is-Summary + Value columns should always be fully visible, Description gets the rest

Already effectively true structurally ("Value" and "Is-Summary" are fixed-pixel columns, 100px
and 90px) - the visible symptom was entirely explained by 26.1 above (Description silently
growing past its fair share). No separate change needed once 26.1 landed.

### 26.3 Left "Segments" panel didn't reach down to the footer (dead gap below the list) - three attempts, see 26.3.3 for the one that stuck

**Symptom** (screenshot annotated in red/green): the docked `lstSegments` panel (`Grid.Column
="0"` in the main content area) stopped a good distance above the paging footer, leaving a
visible blank gap, while the detail panel on the right (`Grid.Column="1"`) correctly stretched
all the way down to the footer's top edge. Both columns were direct children of the same
single-row `Grid` and both defaulted to `VerticalAlignment="Stretch"`, so in theory they
should already match - but visibly didn't under this window's `SizeToContent="WidthAndHeight"`.

#### 26.3.1 First attempt (REVERTED - hung and crashed Excel): `ElementName` Height binding

Bound the segments `Border`'s `Height` directly to the detail panel Grid's own `ActualHeight`
(`Height="{Binding ActualHeight, ElementName=DetailPanelGrid}"`, `VerticalAlignment="Top"`).
This looked like a clean, deterministic fix, but the user reported the window now hangs on
open and takes Excel down with it. The app's log file showed nothing thrown at all
(`Logs\GLSense_Logs_<date>.log` recorded only a `WARN Window load cancelled by user` - i.e. the
user force-closing it, not an exception) - consistent with a genuine UI-thread layout hang
rather than a caught .NET error. Root cause: an `ElementName` binding to a sibling's
`ActualHeight`, combined with this window's `SizeToContent="WidthAndHeight"`, creates a
measure/arrange feedback loop - the Border's height depends on the detail Grid's
`ActualHeight`, but that Grid's own resolved height depends on its row's share of the window's
total size, which `SizeToContent` can only compute by first measuring every child *including
this Border* - a genuine layout cycle that never converges (WPF doesn't always throw a "cycle
detected" exception for this shape; it can just spin). Reverted.

#### 26.3.2 Second attempt (REVERTED - still crashed): `"*"` to `"Auto"` row-type change

Changed `RootGrid`'s main-content row and the detail panel's dual-DataGrid row from `Height
="*"` to `Height="Auto"`, reasoning that a `"*"` row absorbs 100% of any leftover height once
`MinHeight="750"` forces the window taller than its natural content, and that switching to
`Auto` would let both columns' default `Stretch` genuinely share the same, correctly-measured
row height with zero bindings. This was a real, safe structural change on its own (no cycle
risk, unlike 26.3.1) - but the user reported the window **still** hung and crashed, and had to
kill Excel again. Whatever combination of factors is actually causing this - possibly something
in this window's multi-layer nested-Grid structure interacting with `SizeToContent`
independently of the specific row-type choice - re-deriving it further wasn't worth the risk of
a fourth failed attempt. This row-type change was superseded by 26.3.3's full structural
rewrite below, not layered on top of it.

#### 26.3.3 Actual fix (user-provided wireframe, full structural rewrite): `Grid.RowSpan` sidebar, no nested Grids, no bindings

The user provided a simple block-diagram wireframe of the intended layout: a single flat
`Segments | (header / DataGrids / Pagination stack)` two-column area, with the Segments column
visually spanning the full height of everything to its right, and one full-width `Buttons
Section` below both columns (Pagination is *not* full-width in this design - it sits directly
under the DataGrids, matching their width).

Rewrote `GLSegmentManager.xaml`'s body from scratch around this: `RootGrid` now has just 3 rows
(title bar, one `Auto` main-content row, one `Auto` buttons row) instead of the previous 4. The
main-content row holds a single flat `Grid` (no more separate nested "detail panel" `Grid` -
that extra nesting layer is gone entirely) with 2 columns (`220` / `*`) and 5 rows (Value/
Reference, Hierarchy, Search, dual DataGrids, Pagination - all `Auto`). The Segments `Border`
is placed at `Grid.Column="0" Grid.Row="0" Grid.RowSpan="5"`, so it spans exactly the combined
height of all 5 rows on the right - this is a plain, hard Grid-arrange guarantee (a `RowSpan`
cell's height is always exactly the sum of its spanned rows' resolved heights), not a binding,
not dependent on `SizeToContent`/`Auto`-vs-`Star` measure-pass quirks, and not layered nested
Grids interacting with each other - the single most standard, lowest-risk WPF technique for
"make a sidebar match a multi-row content column's height," and the one that should have been
reached for from the start instead of the two riskier attempts above. Pagination moved from
being its own full-width root-level footer into row 4 of this same grid (`Grid.Column="1"`
only) so it sits under the DataGrids' width, not under Segments too - matching the wireframe.
The single remaining full-width footer (`Buttons Section` - Clear Defaults/Ok/Close) stays a
separate `RootGrid` row below everything, spanning the whole window width as before.

Width was already fixed (`ColumnDefinition Width="220"`, a literal pixel value immune to
content overflow) and the subtitle text under each segment name already has `TextTrimming
="CharacterEllipsis"` - so "should not expand, show ... if it needs to" was already satisfied
structurally; nothing further changed there.

**If this still hangs/crashes**: do not reach for another `ElementName`/`RelativeSource`
`Height`/`Width` binding against a sibling in this window (26.3.1), and don't assume a plain
`"*"`-to-`"Auto"` row-type swap alone is sufficient (26.3.2 wasn't) - the working design is the
flat single-Grid-with-`RowSpan` structure in 26.3.3. Re-check whether the crash is even layout-
related at all before retrying a layout change - rebuild cleanly, fully close and relaunch
Excel per the Deployment Note below (a stale AppDomain-loaded copy would make ANY fix look like
it "didn't work"), and check `Logs\GLSense_Logs_<date>.log` for anything thrown before assuming
it's the same class of bug again.

#### 26.3.4 The actual likely root cause (found after 26.3.3 also crashed): `DataGridColumnFillHelper`'s `SizeChanged`->`Refresh` loop, not the Grid row structure at all

The user reported 26.3.3 **still** hung and crashed - with a crucial new detail: the window
opens fine, and only starts resizing/hanging *after* that, not immediately. That timing, and the
fact that three completely unrelated row/layout structures (26.3.1's binding, 26.3.2's
`Auto` rows, 26.3.3's `RowSpan`) all failed identically, means the Grid row structure was very
likely never the actual cause - something common to all three attempts is. The two candidates
that stayed constant across all three: this section's own `DataGridColumnFillHelper.cs` change
(26.1's floor-clamp) and the `IsEnabled`/reference-validation additions (26.5). The `IsEnabled`
bindings and `ValidateAndApplyReferenceValue` only fire in response to user interaction with the
Reference field or a pre-existing `PropertyChanged` subscription wired *after* initial load (see
26.5's own text) - neither explains a hang that starts right after open, before the user has
touched anything. `DataGridColumnFillHelper` was the remaining suspect, and a real bug turned up
in it: `EnableFillColumn` wires `grid.SizeChanged += Refresh`, and `Refresh()` itself sets
`fillColumn.Width` twice (once to `Auto` to re-measure, once to a resolved concrete value) with a
`grid.UpdateLayout()` in between - each of those writes can itself provoke another `SizeChanged`
on the same grid (a column width change nudging the grid's own rendered size, especially with
this window's `dgLeft`/`dgRight` sitting in a `"*"`-sized row under `SizeToContent`). Nothing in
the method stopped that from re-entering `Refresh()` before the first call had finished - a
plausible mechanism for exactly the reported symptom (opens fine, then a resize loop that never
settles) regardless of which Grid row-type experiment was active at the time.

**Fix** (in `DataGridColumnFillHelper.cs`, benefits every window using this helper, not just
GLSegmentManager):
- Added a per-grid re-entrancy guard (`static readonly HashSet<DataGrid> _refreshing`) -
  `Refresh()` now returns immediately if it's already running for that same grid, turning any
  potential re-entrant storm into "skip one redundant pass" instead of a runaway loop.
  `finally { _refreshing.Remove(grid); }` always clears the guard even if `Refresh()` throws.
- Added an idempotency check: capture the column's width *before* the `Auto` re-measure
  (`previousWidth`), and only actually reassign `fillColumn.Width` if the newly resolved width
  differs from `previousWidth` by more than half a pixel - otherwise restore `previousWidth`
  unchanged. This avoids a genuinely no-op `DependencyProperty` write that could still ripple
  into another layout pass for zero visual benefit. (First version of this check compared
  against `fillColumn.Width` directly, which is a bug - by that point in the method it had
  *already* been set to `Auto` a few lines earlier, so the check would always see `IsAuto=true`
  and never actually skip anything; fixed to compare against the captured `previousWidth`.)

**Also reverted 26.3.2's rationale but not its symptom**: with 26.3.4's guard in place as the
primary suspected fix, `RootGrid`'s main-content row and the dual-DataGrid row were switched
back from `Auto` to `"*"` - not because `Auto` was wrong on its own, but because `"*"` is this
codebase's own established, already-proven convention for a `BaseWindow` whose content includes
a real `DataGrid` (see CLAUDE.md section 1.5's "legitimate uses of a star row" list -
`GLCubeDetails`, `GLUserConfig`, `GLServerConfiguration`, `GLLOVs`, `GLRollerGroups` all do
exactly this), relying on `BaseWindow`'s own `Window.ContentRendered`-based resettle (section
1.4e) to handle the "collapse until manual resize" symptom generically, window-agnostically -
exactly what the user asked about directly ("why not 32, `*`, 40 or some height to fill the last
row") and the answer is: yes, `"*"` is correct here, matching every sibling DataGrid window; the
buttons-footer row stays `Auto` (not a hardcoded pixel value like "40") because `Auto` already
sizes correctly to whatever the button style's height/padding needs, matching the convention
used everywhere else in this codebase (no window hardcodes a footer row's pixel height).

**Status**: superseded by 26.3.5 below - the user reported this attempt *also* still hung/
crashed, but supplied the crucial missing detail that finally pinned it down.

#### 26.3.5 The actual root cause, confirmed: `DataGridColumnFillHelper.Refresh()` running re-entrantly inside `BaseWindow`'s own resettle/pump sequence

The user's exact words pinned this down precisely: *"after i launched the window it wait for a
milli second (gap issue after close button was noticed here) and adjusted the width to close the
gap and got hanged and excel crashed."* That is a frame-by-frame description of
`BaseWindow.OnLoaded`/`OnContentRendered` (`BaseWindow.cs`) doing exactly what they're designed to
do (CLAUDE.md section 1, 1.4d/1.4e): the window opens with the classic stale `SizeToContent`
first-measurement gap near the title bar, then `ForceSizeToContentResettle()` runs (toggles
`SizeToContent` off/on, nudges `Width`/`Height`, three `UpdateLayout()` calls) followed
immediately by `PumpDispatcherFrame()` - which explicitly pushes a nested `DispatcherFrame` and
pumps **every pending operation at `Background` priority and above** until told to stop. The
crash coincides exactly with that sequence, not with any particular Grid row-type choice - which
finally explains why 26.3.1/26.3.2/26.3.3 (three unrelated layout structures) all failed
identically: none of them touched the actual problem.

`DataGridColumnFillHelper.EnableFillColumn` wires `grid.SizeChanged += (s,e) => Refresh(...)` at
default (`DataBind`/`Normal`-ish, effectively "whenever WPF gets to it") priority - well within
what `PumpDispatcherFrame`'s `Background`-and-above pump will process. `Refresh()` itself calls
`grid.UpdateLayout()` (a forced, synchronous re-measure) and reassigns `fillColumn.Width` twice -
either of which can itself provoke another `SizeChanged` on the same grid. Combined with
`ForceSizeToContentResettle()`'s own three `UpdateLayout()` calls on the *Window* happening
around the same moment, this created a realistic path for `Refresh()` to run synchronously
re-entrant, nested inside the outer window's own in-progress layout pass and/or repeatedly inside
the `PumpDispatcherFrame` pump - a known-fragile combination with `DataGrid`'s layout
implementation, and consistent with "resizing and hanging" rather than an immediate crash on
open. 26.3.4's re-entrancy guard (same-call self-protection) was real and worth keeping, but
didn't stop `Refresh()` from being *invoked* during that dangerous window - it only stopped it
from calling itself recursively once already running.

**Fix** (in `DataGridColumnFillHelper.cs`, benefits every window using this helper): both the
`Loaded` and `SizeChanged` hooks in `EnableFillColumn` now defer the actual `Refresh()` call via
`grid.Dispatcher.BeginInvoke(..., DispatcherPriority.ContextIdle)` instead of invoking it
directly from the event handler. `ContextIdle` is strictly *below* `Background`, so
`PumpDispatcherFrame`'s pump - which only processes `Background` and above - will never execute a
deferred `Refresh()` call; it waits until the dispatcher is genuinely idle, guaranteed to be after
`ForceSizeToContentResettle()`/`PumpDispatcherFrame()` have both fully returned. This matches the
priority `GLSegmentManager.xaml.cs`'s own `Window_Loaded` already uses for its manual `Refresh()`
calls, so every call site into this helper now defers consistently the same way.

### 26.4 Footer area should start higher up / match the user's wireframe

Direct consequence of 26.3.3's rewrite - Segments now spans exactly the same height as the
Value/Hierarchy/Search/DataGrids/Pagination stack via `RowSpan`, Pagination sits under the
DataGrids' width (not full window width), and only the Buttons Section spans the full width,
below both columns - matching the wireframe exactly. No separate fix needed beyond 26.3.3.

### 26.5 Reference mode should disable Hierarchies/Search/both grids, and validate the referenced cell's value

**Before**: selecting a cell Reference for a segment already correctly disabled that
segment's Value textbox (`IsTextEnabled`/`IsRefEditEnabled` via
`SegmentSelectorViewModel.ApplyEnableState`), but the Hierarchy combo, Search controls, and
both DataGrids stayed fully interactive even though none of them apply once a Reference is
driving the segment - and nothing validated what the referenced cell actually contained.

**Fix**:
- `GLSegmentManager.xaml`: the Hierarchy `Border` (`Grid.Row="1"`), Search `Border`
  (`Grid.Row="2"`), and the dual-grid `Border` (`Grid.Row="3"`) all now bind `IsEnabled` to
  the same `SelectedSegment.IsTextEnabled` flag already driving the Value textbox - `false`
  exactly when a Reference is active. WPF's `IsEnabled` cascades to every descendant control,
  so this greys out/disables the Hierarchy combo, Search combo+textbox, both `DataGrid`s, and
  all five transfer buttons with one binding each, no new ViewModel property needed.
- `SegmentSelectorViewModel.cs`, `HandleReferenceChange`: when a segment's `Reference` becomes
  non-blank, added a call to a new `ValidateAndApplyReferenceValue(seg)` method (only acts on
  the currently-selected segment, matching the existing "active segment" guard pattern used
  elsewhere in this class). It resolves the cell via `ExcelApp.Range[seg.Reference].Value2`
  (same pattern as the existing `ResolveSegmentValueText`), then:
  - empty cell -> `ShowWarningAction` toast: "No data in the referenced cell for '{segment}'."
  - non-empty -> parsed via a new `ParseReferenceCellValues` helper, which strips a matching
    leading/trailing `--`/`-` wrapper (some upstream sheets prefix/suffix a value list that
    way) and splits on `,` - then each token is checked against the segment's currently loaded
    valid values (`_allSegmentValues`, case-insensitive). Any token not found -> a toast naming
    which one(s) weren't found. All valid -> mirrored into `seg.Value` (shown in the now-
    disabled Value box, and - via the existing `HandleValueChange`/`SyncActiveSegmentGrid`
    chain - the right-hand "selected values" grid too, exactly as if picked manually).
- **Regression guard**: mirroring the resolved value into `seg.Value` re-triggers
  `HandleValueChange`, which repopulates the right-grid selection (`_selectedRight`), and that
  in turn feeds `UpdateRefWindowState()` (invoked at the end of `HandleReferenceChange` via
  `UpdateMultiRowState()`). `UpdateRefWindowState` used to check `_selectedRight.Any()`
  *before* checking whether a Reference was active, so it would flip `IsTextEnabled` back to
  `true`/`IsRefEditEnabled` back to `false` the instant the right grid got repopulated - silently
  re-enabling the Value box even though the Reference was still driving it. Fixed by checking
  `SelectedSegment.Reference` first in `UpdateRefWindowState`, so Reference always wins
  regardless of what's in the right grid.

### 26.6 Footer visual tidy-up (reduce gaps, consistent order) + pagination wording matched to GLSegmentValues

Originally a cosmetic pass on the two separate root-level footer `Border`s (Paging + Actions):
standardized padding, tightened margins, reduced inter-button gaps. Superseded structurally by
26.3.3's rewrite, which folded the paging/pagination footer into row 4 of the main content grid
(under the DataGrids only, per the user's wireframe) - only the Buttons Section remains a
separate, full-width root-level footer now. Its `Padding="10,8"` and the `Clear Defaults`/`Ok`/
`Close` button gaps (6px each) from this original pass are preserved in the 26.3.3 rewrite.

Follow-up request: make the pagination footer's labels match `GLSegmentValues.xaml`'s own Paging
Footer, not just functionally but verbatim - wording, icons, and layout. Copied that footer's
content into `GLSegmentManager.xaml`'s pagination `Border` exactly: a `FileLinesSolid` icon
before "Per Page:", the "Page X of Y" indicator wrapped in a pill-style `Border`
(`Background="{StaticResource BackgroundBrush}"`, `CornerRadius="4"`), a "Showing:" label before
`PageRangeText` (bold, with a tooltip - "Record range currently displayed in the left grid"), and
a `DatabaseSolid` icon before "Total:". Both `FileLinesSolid`/`DatabaseSolid` come from the same
`iconPacks:PackIconFontAwesome` namespace already imported in this file (used elsewhere for the
title-bar icon), and `BackgroundBrush`/`PrimaryBrush` already exist in `GlobalStyles.xaml` - no
new resources needed.

**Status**: all six items (26.1-26.6) implemented, AIPowered only. 26.3 went through five
attempts before the actual root cause was found and fixed - see 26.3.5 for the real fix
(deferring `DataGridColumnFillHelper.Refresh()` off the `BaseWindow` resettle/pump's priority
band entirely). **User confirmed after rebuild: the hang/crash is fixed** ("Now that issue fixed
here are the observations") - 26.3.5's `ContextIdle`-deferral is the real, verified fix; no
further crash reports as of section 26.7 below.

### 26.7 Post-crash-fix testing feedback: four follow-up items

Raised by the user in the same message confirming the crash was fixed, after actually exercising
the rebuilt window. Items 1 (window loads with all 3 left-DGV columns visible) and the general
"scrolling" case from 26.1 needed no further action - only the four below did.

**Item 2 - switching to a segment with longer Account descriptions re-triggered the resize/
Is-Summary-disappearing symptom, now on segment switch rather than scroll.** Root cause:
`DataGridColumnFillHelper.Refresh()`'s existing "set fill column to `Auto`, measure its natural
width, then resolve back to a concrete pixel width" technique (see 26.1) is only safe as a
*transient* measurement trick - but on a window using `SizeToContent="WidthAndHeight"` (every
`BaseWindow`-derived window here), WPF's `SizeToContent` engine re-measures the window on
**every** layout pass, including the one instant `Refresh()` flips the Description column to
`Auto`. A `DataGridColumn` at `Auto` reports its full natural/unclamped width - for a long
Account description that can be hundreds of pixels wider than the column's eventual resolved
width - so for that one pass the window's desired size balloons and the window visibly grows to
fit it, before `Refresh()` sets the column back to its resolved (smaller) pixel width a moment
later. That growth-then-shrink is exactly what the user described ("made the grid resize which
made the window resize... Is-Summary is gone and horizontal scroll bar appeared") - once the
window had already grown, the subsequent shrink evidently didn't fully re-settle the Description
column's resolved width against the still-larger `grid.ActualWidth`, leaving Is-Summary squeezed
out and a horizontal scrollbar appearing instead.

Two-part fix:

1. `DataGridColumnFillHelper.Refresh()` now freezes the window's `SizeToContent` for its own
   duration: before doing the Auto-measure-then-fix dance, it finds the ancestor `Window`
   (`Window.GetWindow(grid)`) and, if its `SizeToContent` isn't already `Manual`, temporarily
   sets it to `Manual` (freezing the window at its current size); the original `SizeToContent`
   is restored in the `finally` block, **after** the fill column has already been set to its
   final resolved width - so the one re-measure that restoring `SizeToContent` triggers sees
   only the correctly-clamped layout, never the transient `Auto` width. This mirrors
   `BaseWindow.ForceSizeToContentResettle()`'s own established toggle-Manual-then-restore
   pattern (section 1.4), just applied around `Refresh()`'s own Auto-remeasure instead of
   around a window-level resettle. This also means the same fix automatically protects every
   other window/grid `DataGridColumnFillHelper` is wired to, not just `GLSegmentManager`.
2. That freeze only helps if `Refresh()` actually *runs* after a segment switch - but
   `EnableFillColumn` only wires `Refresh()` to `grid.Loaded`/`grid.SizeChanged`, and a segment
   switch reloads `PagedSegmentValues` (`SelectedSegment`'s setter -> `LoadSegmentValuesAsync`
   -> `UpdatePagingAndGrid`) without necessarily changing either grid's `ActualWidth`, so
   `SizeChanged` may never fire for a pure data reload. `UpdatePagingAndGrid()` already invokes
   `DataLoadedAction` on every call (not just the first), and `GLSegmentManager.xaml.cs`'s
   constructor already wires that action - previously *only* to run
   `ForceSizeToContentResettle()`/`PumpDispatcherFrame()`, gated to fire once
   (`_hasResettledAfterInitialLoad`). Added an *ungated* `DataGridColumnFillHelper.Refresh()`
   call (for both `dgLeft`/`dgRight`, deferred at `DispatcherPriority.ContextIdle` like every
   other call site) to the front of that same action, so both grids' fill columns are
   explicitly re-resolved on every segment switch, not just on load/manual resize - with the
   freeze from part 1 guaranteeing this can't reintroduce a resize-then-shrink flash.

**Item 3 - clearing the RefEdit box didn't clear the Value box/grid.**
`SegmentSelectorViewModel.HandleReferenceChange`'s "Reference cleared" branch used to check `if
(!string.IsNullOrWhiteSpace(seg.Value)) { seg.SelectedValues = ParseValueToSelections(seg.Value,
...); }` - i.e. it *restored* `SelectedValues` from whatever `seg.Value` still held, rather than
clearing it. That restore made sense before `ValidateAndApplyReferenceValue` existed (26.5, when
`Value` could only ever get populated by something the user typed directly), but since
`ValidateAndApplyReferenceValue` now mirrors the reference's resolved value **into** `seg.Value`
while a Reference is active, clearing the Reference and then "restoring" from that mirrored,
now-stale `Value` silently repopulated the Value box and right grid with data tied to a reference
that no longer exists. Fixed by making the branch unconditionally clear both: `seg.Value =
string.Empty; seg.SelectedValues.Clear();` before calling `ApplyEnableState(seg)` - clearing
`Value` also fires `HandleValueChange` -> `ClearActiveSegmentGrid`, which empties the visible
right-grid selection for the active segment.

**Item 4 - Segments-list subtitle wording: "Direct value: X" should read "Default: X".**
`Converters.cs`'s `SegmentSummaryConverter` (builds each segment's one-line subtitle in the left
list from `Value`/`Reference`/`SelectedValues.Count`) had a hardcoded `$"Direct value: {value}"`
branch. Changed to `$"Default: {value}"`; the `Reference: ...` / `N value(s) selected` / `No
value set` branches were already correct per the user ("good what we have but not direct") and
were left untouched.

**Item 5 - add a hover tooltip on grid rows showing the hovered row's item/value.** Added a
`ToolTip` `Setter` (via `MultiBinding`/`StringFormat`, since a row's tooltip needs more than one
bound field) to both `dgLeft`'s and `dgRight`'s existing `DataGrid.RowStyle` in
`GLSegmentManager.xaml`: `dgLeft` shows `"{DisplaySegmentValue} - {Description}"`, `dgRight` shows
`"{Value1} / {Value2} - {Segment}"`. Also set `ToolTipService.InitialShowDelay="300"` on both row
styles so the tooltip appears promptly on hover rather than after WPF's default ~1s delay.

**Status (superseded - see 26.8 below)**: the `DataGridColumnFillHelper` SizeToContent-freeze fix
described above was implemented but turned out to be a patch on the wrong layer - it didn't fully
fix item 2, and item 5's row-level `MultiBinding` tooltip didn't render at all. Left in place
(harmless, and still protects every OTHER `DataGridColumnFillHelper`-wired window), but
`GLSegmentManager` itself no longer uses `DataGridColumnFillHelper` at all - see 26.8.

### 26.8 Root cause found by comparison to a proven-working sibling window; architectural fix

After 26.7's fix, the user re-tested and reported the symptom was still there, but with a
critical new data point: **`GLSegmentValues.xaml`** - this window's non-master-detail sibling,
built around the same dual-DataGrid/pagination shape - **never has this problem**. Switching
segments there never resizes the window and "Is-Summary" never disappears. Comparing the two
files line-for-line surfaced the actual root cause, which every attempt in 26.3/26.7 had been
dancing around without naming directly:

- `GLSegmentValues.xaml` is `SizeToContent="Manual"` with an explicit fixed `Width="740"
  Height="700"` (`MinWidth`/`MaxWidth`/`MinHeight`/`MaxHeight` only bound how far the user can
  resize it manually) - the window's size is **never** derived from its content. Its dual
  DataGrid's "Description" and "Segment" columns are plain `Width="*"` - ordinary, native WPF
  star-sizing, with no helper class involved at all.
- `GLSegmentManager.xaml` was `SizeToContent="WidthAndHeight"` - a genuinely content-driven
  window, matching the convention most other DataGrid-containing `BaseWindow`s use (section 1.5).
  That convention is exactly what kept causing trouble here: a content-driven window re-measures
  on **every** layout pass, so anything that transiently changes a `DataGridColumn`'s natural
  width - `DataGridColumnFillHelper`'s own Auto-remeasure trick, a segment switch loading longer
  Account descriptions, even WPF's own internal column layout during an `ItemsSource` swap - can
  transiently balloon the *whole window* before settling back down, and that settle-back-down
  step doesn't reliably leave the DataGrid's columns/scrollbar back in the state they started in.
  26.3's five attempts and 26.7's `SizeToContent`-freeze-inside-`Refresh()` patch were all trying
  to out-maneuver this same underlying instability from inside the content-driven model instead
  of removing the instability's precondition.

**Fix**: converted `GLSegmentManager` to match `GLSegmentValues`' proven approach exactly, rather
than continuing to patch around the content-driven model:

- `SizeToContent="WidthAndHeight"` -> `SizeToContent="Manual"` with explicit `Width="1040"
  Height="820"` added (same `MinWidth="940"`/`MaxWidth="1100"`/`MinHeight="750"`/`MaxHeight="900"`
  bounds kept, now just governing manual resize instead of content-driven sizing).
- `dgLeft`'s "Description" column and `dgRight`'s "Segment" column: `Width="Auto"` ->
  `Width="*"`, matching `GLSegmentValues` exactly. No more `DataGridColumnFillHelper` involved.
- Removed `DataGridColumnFillHelper.EnableFillColumn(...)` calls from the constructor, the
  `DataGridColumnFillHelper.Refresh(...)` calls from `Window_Loaded` and from `DataLoadedAction`
  (added in 26.7, now removed again), and the now-pointless `_hasResettledAfterInitialLoad`
  guard field plus its `ForceSizeToContentResettle()`/`PumpDispatcherFrame()` calls -
  `BaseWindow.OnLoaded`/`OnContentRendered` already skip that machinery entirely whenever
  `SizeToContent == Manual` (its whole purpose is fixing `SizeToContent`-driven windows), so none
  of it does anything for a `Manual` window anyway. This also means the crash-prone
  resettle/pump-dispatcher-frame interaction that caused the ENTIRE hang/crash saga in 26.3 no
  longer applies to this window at all - not patched around, structurally absent.
- `RootGrid`'s main-content row and the dual-DataGrid row stay `Height="*"` - under a *fixed*
  window `Height`, a `"*"` row is completely ordinary WPF star-sizing with no measure risk
  whatsoever (this is the exact same `32 / * / Auto / Auto` `RootGrid` row shape
  `GLSegmentValues.xaml` already uses).

**Item 5 tooltips not appearing, fixed by using a proven pattern instead of debugging the broken
one**: the 26.7 attempt set a `DataGridRow`-level `ToolTip` via a `Style` `Setter` bound through a
`MultiBinding` - this did not render at all. Rather than debug why, switched to the pattern
`GLSegmentValues.xaml` already uses successfully for its own Value/Description columns: a
`Setter` on the per-column `TextBlock`'s `ElementStyle` (`DataGridTextColumn.ElementStyle`),
binding the `ToolTip`'s content to that column's own single field (not a `MultiBinding`). Applied
to `dgLeft`'s Value/Description columns and `dgRight`'s Value1/Value2/Segment columns (dgRight's
tooltips are new - `GLSegmentValues` doesn't have a comparable right-grid tooltip need, but the
same proven mechanism applies).

**Item 5 also asked for tooltips on hovering Segments (the master list, not just the grids)**:
added a `ToolTip` (via `SimpleBrowserToolTip`, the same style already used elsewhere in this
window for the Search box/Is-Summary checkbox/pagination "Showing:") on the `StackPanel` inside
`lstSegments`' `ItemTemplate`, showing the segment's full name plus its current
status/value/reference (reusing `SegmentSummaryConverter`, the same converter that already builds
the visible (and sometimes ellipsis-truncated) subtitle line).

**Status**: implemented, AIPowered only. **User confirmed after rebuild: this fixed the
resize/Is-Summary problem completely** - "that fixed the problem and its showing perfect" - and
also confirmed the per-cell tooltips now render correctly. 26.8 is the standing architecture for
this window going forward; do not reintroduce `SizeToContent="WidthAndHeight"` or
`DataGridColumnFillHelper` here without a strong reason, since this exact combination is what
caused the entire 26.3/26.7 saga.

### 26.9 Two refinements after 26.8's fix confirmed working

**Item 1 - hide the redundant "Segment" column in the right grid, give that width to the left
grid.** The right grid only ever shows values for whichever segment is currently selected (its
`ItemsSource` is scoped to the active segment already), so repeating that segment's name on every
row added nothing. Removed `dgRight`'s "Segment" `DataGridTextColumn` entirely; `Value2` changed
from a fixed `Width="90"` to `Width="*"` so it absorbs the freed-up space instead of leaving a
blank gap. Rebalanced the dual-grid area's outer `Grid.ColumnDefinitions` from `1.2*/54/1.0*` to
`1.8*/54/0.6*` (`MinWidth` `220`->`260` on the left, `220`->`160` on the right, since the right
grid now only needs to fit two ~90px columns) so the left grid (Value/Description/Is-Summary)
gets noticeably more of the freed-up width, per the user's request.

**Item 2 - distinguish "this is still the untouched factory default" from "the user actively
selected this."** Previously `SegmentSummaryConverter` showed `"Default: X"` any time a segment
had a non-empty `Value`, regardless of whether that value was the original default the segment
loaded with or something the user had since changed via the dual-grid. Added `SegmentModel.
IsUserSelected` (bool, default `false`) - set to `true` in exactly one place:
`SegmentSelectorViewModel.UpdateRefWindowState()`'s `_selectedRight.Any()` branch. That branch is
reached *only* from a genuine runtime mutation of the right-hand grid - the `AddSelection`/
`RemoveSelection`/`AddBetweenSelection`/`AddNotBetweenSelection`/`AddExcludeSelection` button
handlers, or a direct edit of the Value textbox (`HandleValueChange` -> `SyncActiveSegmentGrid` ->
`UpdateMultiRowState` -> here) - never from the initial default-value parse
(`InitializeSegment`/`ParseAndSetSegmentValues`, which run before this model's `PropertyChanged`
handler is even subscribed) or from a plain segment switch (`SelectedSegment`'s setter restores
`SelectedValues` directly into `_selectedRight` without going through `UpdateRefWindowState` at
all). `SegmentSummaryConverter` now takes an optional 4th `MultiBinding` value (`IsUserSelected`,
defaults to `false`/"Default" if omitted so the converter still works with only 3 bindings) and
shows `"Selected: X"` instead of `"Default: X"` once it's `true`. Updated both `MultiBinding`s in
`GLSegmentManager.xaml` (the visible subtitle and the 26.8 tooltip) to pass this 4th binding.

**Status**: implemented, AIPowered only. Not yet rebuilt/tested by the user - verified by XML
well-formedness check on the XAML instead, per the Deployment note below.

---

## 27. Three unrelated fixes: GLSegmentManager width tweak, GLSegmentDiscovery busy toast, GLWaitWindow Cancel button restyle (AIPowered only)

### 27.1 GLSegmentManager: reduce total window width by 50px, shift 25px from the left grid to the right grid

Follow-up sizing tweak on top of section 26.8/26.9's fixed-size conversion. `Window` `Width`
`1040` -> `990` (-50px; `MinWidth`/`MaxWidth` left at `940`/`1100` - still within range).
Separately, within the dual-grid area's `Grid.ColumnDefinitions`, moved ~25px of width from the
left grid to the right grid on top of that: `1.8*/54/0.6*` -> `1.71*/54/0.69*`, and `MinWidth`
`260/160` -> `235/185`. The star-ratio arithmetic: at the new `990` width, the dual-grid area has
roughly 696px to split across 2.4 star units (~290px/unit under the old 1.8/0.6 ratio) - shifting
25px each way works out to roughly 1.71\*/0.69\* to preserve the same total. Since this is a fixed
`Width` window now (section 26.8), these are one-time, deterministic numbers, not something that
re-derives itself at different window sizes - if the user asks for another pixel-level nudge
later, redo this same arithmetic rather than guessing at new star values.

### 27.2 GLSegmentDiscovery: show a busy toast while writing/inserting values

`BtnSubmit_Click` (the "Insert" action button, `Style="{StaticResource InsertButtonStyle}"`) does
its actual work - `WriteValuesToExcel()` / `PerformInsertIfNeeded()` - entirely synchronously on
the UI thread via direct Excel Interop calls, and depending on `ValueArray.Length` and whether
cells need to shift (Insert mode vs Overwrite), this can take anywhere from a few milliseconds to
a few seconds with no feedback that anything is happening. Changed `BtnSubmit_Click` to `async
void` and, right after validations pass (`ValidatePrerequisites()`/`ValidateOperationSelected()`),
added `AppOverlayControl.ShowBusyasyn("Writing segment values...")` followed by `await
Dispatcher.Yield(DispatcherPriority.Render)` before calling `WriteValuesToExcel()` - the `Yield`
is what actually lets WPF paint the busy overlay before the (still-synchronous, still
UI-thread-bound) Excel writes start; without it, the overlay's `Visibility` change and the Excel
writes would happen within the same synchronous call stack and the overlay would never actually
render on screen until after the writes had already finished. `await
AppOverlayControl.HideBusyAsync()` added to the `finally` block (guarded by a `busyShown` flag so
it's only called if the overlay was actually shown - e.g. not if validation failed and the method
returned early). `WriteValuesToExcel`/`PerformInsertIfNeeded` and all the Excel Interop code below
them are completely unchanged - this is purely a show/hide bracket around the existing call, using
the same `AppOverlay.ShowBusyasyn`/`HideBusyAsync` API already used by every ViewModel-driven
window's `ShowBusyAction`/`HideBusyAsyncAction` (this window has no ViewModel, so it calls
`AppOverlayControl` directly instead).

**Follow-up (27.2b) - the busy overlay showed but its spinner animation looked frozen**: for a
long enough `ValueArray`, the busy toast appeared (the fix above worked) but visibly stopped
animating once the Excel writes started. Root cause: Excel Interop here is in-process COM (the
add-in runs inside Excel's own process), so `WriteCellValue` calls are genuine synchronous,
blocking function calls with no message pump in between - they can't be moved to `Task.Run`
without breaking COM apartment affinity. Blocking the UI thread in a tight loop for more than a
frame or two starves WPF's composition engine of the periodic hand-off it needs to keep painting
the `AppOverlay`'s busy `Storyboard`, even though that kind of animation is normally
composition-thread-driven - a known WPF gotcha when the UI thread doesn't yield at all for an
extended stretch. Fix (`GLSegmentDiscovery.xaml.cs`): converted the whole write chain to `async
Task` - `WriteValuesToExcel()` -> `WriteValuesToExcelAsync()`, `WriteValuesByDirection()` ->
`WriteValuesByDirectionAsync()`, and all four direction-specific loop methods
(`WriteVerticalForward`/`WriteHorizontalForward`/`WriteVerticalBackward`/`WriteHorizontalBackward`)
- each loop now calls a small `YieldEveryBatch(i)` helper after every cell, which does `await
Dispatcher.Yield(DispatcherPriority.Render)` every `YieldBatchSize` (10) iterations. This is a
pure control-flow change - `WriteCellValue`/`PerformInsertIfNeeded` and every actual Excel COM
call are byte-for-byte unchanged, only the calling methods became `async` and periodically hand
control back to the dispatcher between batches of cells so pending render frames (including the
busy overlay's spinner) get a chance to run.

### 27.3 GLWaitWindow: Cancel button should look like a Close button, not a Cancel button

`BtnCancel` (`Views/GLWaitWindow.xaml`) is functionally a Cancel button (`Click="BtnCancel_Click"`,
cancels the in-progress operation) but was styled with `CancelButtonStyle`. Both
`CancelButtonStyle` and `CloseButtonStyle` (`Themes/GlobalStyles.xaml`) share the same
`DynamicContentButton` base template (light gray `#F8F9FA` background/black text at rest,
`PrimaryBrush` blue background/white text on hover - see that style's own comment), so switching
between them is not itself a strong visual change - `DynamicContentButton`'s own `Content`/
`ToolTip` Setters are irrelevant here too, since this Button already supplies its own local
`Content` (an icon + "Cancel" `TextBlock` in a `StackPanel`) and local `ToolTip`, both of which
override whichever style is applied. The actual bug: that local `StackPanel`'s icon and
`TextBlock` had `Foreground="White"` hardcoded, which rendered as white-on-near-white (invisible
or very low contrast) at rest, only becoming legible once hovered (when the template's `IsMouseOver`
trigger flips the Background to blue) - this is what visually read as "not styled like the other
Close buttons in the app" (which never hardcode Foreground, so their text/icons correctly show
black-at-rest/white-on-hover via the template's own triggers). Fixed by: (1) switching
`Style="{StaticResource CancelButtonStyle}"` -> `Style="{StaticResource CloseButtonStyle}"` so
this button is visually identical to every other Close button in the app, and (2) replacing the
hardcoded `Foreground="White"` on both the icon and the "Cancel" `TextBlock` with `Foreground=
"{Binding Foreground, RelativeSource={RelativeSource AncestorType=Button}}"` - the same
explicit-Foreground-propagation pattern already used elsewhere in this codebase (e.g.
`GLSegmentManager.xaml`'s `DataGridCell` -> `DataGridRow` Foreground binding) instead of relying on
implicit WPF property inheritance, so the icon/text always track whatever the Button's own
(template-driven, hover-reactive) Foreground currently is. The button's actual behavior/wording
(`Content` stays "Cancel", `Click` handler unchanged) is untouched - only its visual styling
changed.

**Status**: all three implemented, AIPowered only. Not yet rebuilt/tested by the user - verified
by XML well-formedness check on the XAML instead, per the Deployment note below.

---

## 28. GLRollerGroups doesn't sync to the ribbon's RibSegS segment picker like GLSegmentValues/GLSegmentManager do (AIPowered only)

**Symptom**: picking a segment in the ribbon's `RibSegS` combo makes `GLSegmentValues` and
`GLSegmentManager` open with that same segment already selected/highlighted in their own
segment list/combo. `GLRollerGroups` (`Views/GLRollerGroups.xaml`/`.xaml.cs`) does not - it
always opens with the first segment selected, regardless of what was picked in `RibSegS`.

**How the other two windows do it**: none of these three windows have a live event
subscription to ribbon changes - they're all opened modally (`AddinEntry.ShowGroupCWindow`
+ `win.ShowDialog()`), so "sync" really just means "read the ribbon's last-picked segment once
at `Window_Loaded`, since the window is reconstructed fresh every time it's opened." The
ribbon side of this: `AddinEntry.SegmentChanged` (wired to the host's `RibSegS_OnChange`) sets
two properties on the `AppState` singleton every time the user picks a `RibSegS` item -
`AppState.Instance.DefaultSegment` (the segment name) and `AppState.Instance.SegmentPickedIndex`
(that segment's index within the same `DataRepository.GetSegments(cubeId, ledgerId)` ordering
used everywhere else, `-1` if nothing valid is picked). `GLSegmentValues`/`GLSegmentManager`
share `SegmentSelectorViewModel`, whose `ProcessSegments` (called from `LoadSegmentsAsync`,
itself called from each window's `Window_Loaded`) ends with `SelectInitialSegment()`:

```csharp
private void SelectInitialSegment()
{
    if (Segments.Count > 0)
    {
        if (AppState.Instance.SegmentPickedIndex >= 0 &&
            AppState.Instance.SegmentPickedIndex < Segments.Count)
        {
            SelectedSegment = Segments[AppState.Instance.SegmentPickedIndex];
        }
        else
        {
            SelectedSegment = Segments[0];
        }
    }
}
```

i.e. it reads `AppState.Instance.SegmentPickedIndex` directly, at the moment the segment list
finishes loading.

**Root cause in `GLRollerGroups`**: its own `ViewModels/SimpleSegmentViewModel.cs` already has
the *identical* selection logic inside `LoadSegmentsAsync` -

```csharp
if (Segments?.Count > 0)
{
    if (SegmentPickedIndex >= 0)
    {
        SelectedSegment = Segments[SegmentPickedIndex];
    }
    else
    {
        SelectedSegment = Segments[0];
    }
}
```

- but here `SegmentPickedIndex` is the ViewModel's **own** `public int SegmentPickedIndex { get;
set; } = -1;` property, not a read of `AppState.Instance.SegmentPickedIndex`. Nothing anywhere
in `GLRollerGroups.xaml.cs` (constructor or `Window_Loaded`) ever assigned
`AppState.Instance.SegmentPickedIndex` into it, so it was permanently stuck at its `-1` default
and the `else` branch (`Segments[0]`) always ran. Checked `FinalWorkingCode\GLSense\
Views\GLRollerGroups.xaml.cs`/`ViewModels\SimpleSegmentViewModel.cs` for comparison - this is
not a porting regression, the original monolith has the exact same dead `SegmentPickedIndex`
property and the exact same missing assignment. Since the task scope is AIPowered-only, only
this project's copy was fixed; `FinalWorkingCode` was left untouched as instructed.

**Fix** (`GLSense.Addin.Core\Views\GLRollerGroups.xaml.cs`, `Window_Loaded`): inside the
existing `if (AppState.Instance.SelectedCube != null && AppState.Instance.SelectedLedger !=
null)` block, immediately before the `await vm.LoadSegmentsAsync(...)` call, added:

```csharp
ServiceLocator.Logger?.LogDebug($"GLRollerGroups.Window_Loaded: syncing vm.SegmentPickedIndex from AppState.Instance.SegmentPickedIndex={AppState.Instance.SegmentPickedIndex} (RibSegS selection).");
vm.SegmentPickedIndex = AppState.Instance.SegmentPickedIndex;
```

This is a one-line logical fix plus a matching `ServiceLocator.Logger?.LogDebug` line (same
style as every other debug line in this file/the other two windows). No changes were needed to
`GLRollerGroups.xaml` (the `cmbSegments` combo is already `SelectedItem="{Binding
SelectedSegment, Mode=TwoWay}"` bound to the same `SelectedSegment` this now correctly
initializes, and `Window_Loaded` already sets `cmbSegments.Text = vm.SelectedSegment.SegmentName`
right after loading). Also updated this file's header comment to note the change (it previously
said "No logic changes vs. the original").

**Follow-up hardening** (`ViewModels\SimpleSegmentViewModel.cs`, `LoadSegmentsAsync`): the
`if (SegmentPickedIndex >= 0)` branch quoted above was, before this fix, permanently dead code
(since nothing ever set `SegmentPickedIndex` above `-1`) - activating it exposed a real gap
versus `SegmentSelectorViewModel.SelectInitialSegment`'s otherwise-identical logic: this branch
was missing the upper-bound half of the check (`&& SegmentPickedIndex < Segments.Count`). Without
it, picking a segment in `RibSegS` while a cube/ledger with a *longer* segment list is active,
then opening `GLRollerGroups` against a *different* cube/ledger with a *shorter* segment list,
could leave `SegmentPickedIndex` pointing past the end of `Segments` and throw an
`IndexOutOfRangeException` instead of gracefully falling back to `Segments[0]`. Added the same
`&& SegmentPickedIndex < Segments.Count` guard `SelectInitialSegment` already uses, so the two
now match exactly.

**Status**: implemented, AIPowered only. There is no Windows/MSBuild toolchain available in
this environment to actually compile or run the add-in, so this was verified by careful manual
review of the edited `Window_Loaded` method (control flow, existing variable/property names,
null-conditional logging pattern) rather than a build - `GLRollerGroups.xaml` itself was not
touched, but was re-checked for XML well-formedness anyway via `python3 -c "import
xml.etree.ElementTree as ET; ET.parse(...)"`, per the Deployment note below. Not yet
rebuilt/tested by the user.

---

## 29. Excel.exe left running as an orphaned background process after closing, then the add-in doesn't show up on the next launch (AIPowered only)

**Symptom reported by the user**: "when i perform some excel related operations and close
the excel the excel application is stuck in background running process and if i open
excel then my addin will not be displayed."

**Investigation**: started from `AddinModule.cs` (`GLSense` host project) and its
lifecycle - the same file section 11 already touched (`CountFormulaCells`). Confirmed via
`AddinModule.Designer.cs` (`grep` for every `+=` event wire-up in `InitializeComponent()`)
that only ribbon `OnClick`/`OnChange` handlers and the seven `adxExcelAppEvents1.*`
Excel-event handlers are subscribed there - **no shutdown/disconnect lifecycle event of any
kind was ever subscribed**, not in the Designer file and not by hand in the constructor.
Cross-checked against `FinalWorkingCode\GLSense\AddinModule.cs` (the proven-working old
monolith, same `AddinExpress.MSO.ADXAddinModule` base class) - it explicitly subscribes
`this.AddinBeginShutdown += AddinModule_AddinBeginShutdown;` in its own constructor, and
that handler unsubscribes every Excel event, stops the `SuggestAppendComboBox` mouse-hook
thread, disposes `FormulaCacheManager`, closes the SQLite connection, and force-releases
every Excel COM RCW it holds (`Marshal.FinalReleaseComObject` + double `GC.Collect()`/
`GC.WaitForPendingFinalizers()`) - i.e. the monolith already had to solve this exact class
of bug once, and `AddinBeginShutdown` is the real, existing `ADXAddinModule` event it uses
to do it (not a guessed/invented API - confirmed by its actual, working usage in that
sibling codebase).

This AIPowered port never carried that subscription over. Checked what that meant in
practice for everything this architecture adds on top of the monolith:

- **`GlobalsEx.Loader` (`AddinDomainLoader`, `GLSense.Loader.Core`)**: `Unload()` (`AppDomain.
  Unload(_domain)`) exists and is fully implemented, but the only caller anywhere in the
  codebase is `AddinModule.ReloadAddinCore` - the manual "Reload" ribbon button's hot-reload
  path. Nothing called it on a genuine Excel close, so the entire `GLSense.Addin.Core` child
  AppDomain - every WPF window it ever created, every cached Excel `Range`/`Worksheet`/
  `Workbook` RCW, `HttpClient` instances, the `SuggestAppendComboBox` mouse-hook background
  thread, the open SQLite connection - was simply abandoned in memory instead of being torn
  down, every single time the user closed Excel normally.
- **`IGLSenseAddin.Shutdown()` (`AddinEntry.cs`)**: already closes the reparented Balance
  Configurator WPF window (`ConfiguratorPaneHost.Close()`), calls `SuggestAppendComboBox.
  ShutdownMouseHook()`, and flushes/disposes `FormulaCacheManager` - its own doc comment
  literally says "would equally apply to a genuine Excel shutdown" - but, like `Unload()`,
  its only caller anywhere was `ReloadAddinCore`.
- **`ServiceLocator.Reset()` (`GLSense.Addin.Core\Infrastructure\ServiceLocator.cs`)**: doc
  comment says "Reset the ServiceLocator (useful for shutdown)" - `grep`'d the whole codebase
  for callers: none. Dead code, clearly intended for exactly this scenario.
- **`ServiceLocator.ExcelApp`**: just forwards `IGLSenseContext.ExcelApp`
  (`GLSenseContext.cs`), which is Excel's own `HostApplication` object handed to the add-in
  by Add-in Express - not something this code created, so (per `PORTING_GUIDE.md` section
  4's own rule: "don't release something you didn't create yourself") it should not be
  force-released the way the monolith's `ReleaseAllComObjectsProperly` releases
  `AppState.Instance.ExcelApp`. Almost every `Range`/`Worksheet`/`Workbook` RCW this codebase
  actually creates lives inside the `GLSense.Addin.Core` AppDomain instead (per section 25's
  audit) - unloading that AppDomain already reclaims all of those, which is a strictly safer
  way to get the same result than manually walking and releasing COM objects by hand across
  an AppDomain boundary.

**Root cause**: `AddinBeginShutdown` was never subscribed, so none of the teardown code
above ever ran on a real Excel close - only via the manual Reload button. This is entirely
consistent with the reported symptom: (1) an abandoned AppDomain with a live background
thread (the mouse hook) and open WPF windows/COM references is exactly what leaves `Excel.
exe` as an orphaned background process with no visible window after the user closes it, and
(2) that orphaned process can still be holding file locks on the shadow-copied DLLs under
`Versions\vX\` (`AddinDomainLoader`'s `ShadowCopyFiles`/`CachePath`), which is exactly the
folder the *next* Excel launch's `UpdateBootstrapper`/`AddinDomainLoader.Load()` needs to
read from - a plausible, though not separately proven in this pass, explanation for why the
add-in then fails to show up on that next launch.

**Fix**:

- `GLSense\AddinModule.cs` (constructor): added `this.AddinBeginShutdown +=
  AddinModule_AddinBeginShutdown;`, matching `FinalWorkingCode`'s proven pattern (same real
  `ADXAddinModule` event, not a new/guessed API).
- `GLSense\AddinModule.cs` (new `AddinModule_AddinBeginShutdown` method, placed right after
  `ReloadAddinCore`): unsubscribes all seven `adxExcelAppEvents1.*` handlers (defensive -
  matches `FinalWorkingCode`'s `UnsubscribeFromAllExcelEvents`, guards against a late Excel
  event firing mid-teardown), then calls the exact same two steps `ReloadAddinCore` already
  uses on its outgoing instance, in the same order and for the same documented reason: `GlobalsEx.Addin?.Shutdown()` (closes the Configurator window, stops the
  mouse-hook thread, flushes/disposes the SQLite cache) followed by `GlobalsEx.Loader?.
  Unload(GlobalsEx.Context)` (unloads the whole child AppDomain). Every step is individually
  try/caught and logged, matching this file's existing style - deliberately reused rather
  than duplicated, since `ReloadAddinCore`'s version of this exact sequence is the only
  already-exercised code path for it in this project.
- `GLSense.Addin.Core\AddinEntry.cs` (`Shutdown()`, appended as the last step): added a
  `ServiceLocator.Reset()` call in its own try/catch, since the method already existed
  "for shutdown" but had no caller. Mostly belt-and-braces (the AppDomain unload that follows
  discards these statics anyway), but costs nothing and finally gives that method a reason to
  exist.
- Deliberately **not changed**: no `Marshal.ReleaseComObject`/`FinalReleaseComObject` calls
  were added anywhere in this pass, and `ServiceLocator.ExcelApp`/`GLSenseContext.ExcelApp`
  are never explicitly released. Per the reasoning above, that object is Excel-owned, not
  add-in-created, and the AppDomain unload already reclaims the RCWs this codebase actually
  creates - manually walking/releasing Excel's own `Application`/`Workbooks` COM objects
  (the way `FinalWorkingCode`'s `ReleaseAllComObjectsProperly` does) was judged higher-risk
  than helpful here (COM release ordering mistakes can hang/crash Excel rather than fix
  anything), so it was deliberately left alone rather than guessed at.

**Residual, honestly-stated risk**: `AppDomain.Unload` "forcibly aborts any thread still
executing inside the domain being unloaded" - this is the same documented, accepted
limitation `ReloadAddinCore`'s own comment already calls out for the manual Reload button,
now also reached on every real Excel close. If a background drilldown/refresh/UDF call is
genuinely still mid-flight at the exact moment Excel starts shutting down, it could be
aborted mid-operation. This was judged an acceptable, much smaller risk than the confirmed,
reproducible-every-time bug of never tearing anything down at all, but it is a real,
non-zero trade-off worth knowing about if a rarer "aborted mid-save" style report shows up
after this fix.

**FinalWorkingCode**: not touched. `FinalWorkingCode\GLSense\AddinModule.cs` already
subscribes `AddinBeginShutdown` correctly (see the investigation above) - this bug is
specific to the AIPowered port simply never having carried that subscription over, not a
pattern that also needs mirroring back into `FinalWorkingCode`.

**Status**: implemented, AIPowered only. There is no Windows/MSBuild toolchain available in
this environment to actually compile or run the add-in - verified by careful manual reading
of `AddinModule.cs`/`AddinModule.Designer.cs`/`AddinEntry.cs`/`ServiceLocator.cs`/
`AddinDomainLoader.cs`/`GLSenseContext.cs`/`WpfAppManager.cs`, and by direct comparison
against `FinalWorkingCode`'s own working `AddinBeginShutdown` usage to confirm the event
name/signature are real rather than guessed. No XAML was touched in this pass, so no
`ET.parse` check was needed. Not yet rebuilt/tested by the user - in particular, the
"add-in doesn't show up on next launch" half of the symptom is a plausible consequence of
the orphaned-process file-lock theory above, but wasn't independently reproduced/confirmed
in this pass; if it persists after this fix, the next thing to check would be whether
`UpdateBootstrapper`/`AddinDomainLoader.Load()` can tolerate a still-locked `Versions\vX\`
folder from a not-yet-fully-exited prior `excel.exe` process.

---

## 30. SuggestAppendComboBox: multi-select popup text sits right against the checkbox (AIPowered only)

**Symptom** (screenshot feedback): in the multi-select variant of `SuggestAppendComboBox`'s
dropdown (checkbox-per-item list, used by `RibSegS` and elsewhere), the item text renders almost
touching the checkbox glyph - no visible gap.

**First attempt (ineffective - corrected below)**: `Controls\SuggestAppendComboBox.cs`'s
`UpdateListBoxTemplate()` has a fallback branch that builds the per-item `CheckBox` in code
(`FrameworkElementFactory`) and runs only `if (TryFindResource(key) is not DataTemplate template)`
- i.e. only when no resource matching `ComponentResourceKey(typeof(SuggestAppendComboBox),
"SuggestMultiSelectItemTemplate")` can be found. Padding was added to that fallback `CheckBox`
first, but the user reported zero visible effect even at `Thickness(40,0,0,0)`, and asked whether
`Themes/Generic.xaml` also needed updating.

**Real cause**: `GLSense.Addin.Core\Themes\Generic.xaml` (merged in via `GlobalStyles.xaml`)
already defines a `DataTemplate` keyed with that exact `ComponentResourceKey` (line 12). WPF's
`TryFindResource(key)` finds THIS template every time, so the `.cs` fallback branch never
executes - the earlier fix was dead code. Generic.xaml's `CheckBox` (inside the `DataTemplate`,
~line 14) had no `Padding` set and no explicit `Style`, so it picks up `GlobalStyles.xaml`'s
implicit (no `x:Key`) `CheckBox` style -> `SafeCheckBoxTemplate`, whose `ContentPresenter` has
`Margin="{TemplateBinding Padding}"`.

**Fix**: added `Padding="6,0,0,0"` directly to the `CheckBox` in
`Themes\Generic.xaml`'s `SuggestMultiSelectItemTemplate` `DataTemplate` - this is the template
that actually renders, and `Padding` flows through `SafeCheckBoxTemplate` correctly. Also updated
the `.cs` fallback's comment to explain it's dead code in practice (kept only as a defensive
default matching the same `6,0,0,0` value, in case the Generic.xaml resource is ever unavailable).

**Status**: RESOLVED - confirmed working by user after rebuild.

---

## 31. Expand All / Expand 1 Level: replaced ribbon menu with GLExpandOptions dialog + added By Columns fill (both codebases)

**Request**: `RibExpandAll`/`RibbonExpand1Level` were two `ADXRibbonButton`s hosted inside
`RibSegmentExpand`, an `ADXRibbonMenu` captioned "Hierarchy" - clicking either dispatched
straight to `SegmentDiscoverer.SegmentAction("HierarchyAll"/"Hierarchy1Level")`, which expands a
selected summary/parent segment value into its children by inserting new rows below it and
bulk-filling column 1 of the selection top-to-bottom (`InsertRowsAndFillData`, hard-coded to
`xlShiftDown` + a single-column `Value2` array write). The ask: collapse the menu down to a
single ribbon button that opens a small options window, where the user picks Expand All vs
Expand 1 Level AND, new, whether to fill By Rows (existing behavior) or By Columns (new -
previously unimplemented).

**New window** - `Views\GLExpandOptions.xaml`/`.xaml.cs` (both codebases): a small
`BaseWindow`/`DpiAwareWindow`-derived dialog (styled like `GLSegmentDiscovery.xaml`'s
card-per-option-group + `ModernRadioButton` pattern) with two independent RadioButton groups -
"Expand Level" (`rbExpandAll`/`rbExpand1Level`, `GroupName="LevelGroup"`) and "Fill Direction"
(`rbByRows`/`rbByColumns`, `GroupName="OrientationGroup"`) - plus Expand/Close buttons. No
ViewModel: `BtnExpand_Click` reads the two RadioButton selections, closes the window immediately,
then fire-and-forgets `SegmentDiscoverer.SegmentAction(actionType, byColumns)` - the window
doesn't stay open during the operation since `SegmentAction` already shows its own `GLWaitWindow`
progress dialog (with cancel support).

**Ribbon change** (`AddinModule.Designer.cs`, both codebases): `RibSegmentExpand` converted from
an `ADXRibbonMenu` to a plain `ADXRibbonButton` (same Id/Image, new SuperTip); its two child
`ADXRibbonButton`s (`RibExpandAll`/`RibbonExpand1Level`) and their designer blocks/field
declarations were removed entirely. `RibSegmentExpand.OnClick` now wires to a new
`RibSegmentExpand_OnClick` handler (replacing the old `RibExpandAll_OnClick`/
`RibbonExpand1Level_OnClick` pair) that opens `GLExpandOptions`:
- AIPowered: `AddinModule.cs`'s `RibSegmentExpand_OnClick` dispatches via
  `_ribbonController.ExecuteAction("ShowExpandOptions")` -> `AddinEntry.cs`'s new
  `ShowExpandOptions()` (does the same lightweight DefaultSegment/SegmentPickedIndex guard
  `ShowSegmentDiscovery()` does, then `ShowGroupCWindow("ShowExpandOptions", () => new
  GLExpandOptions())`).
- FinalWorkingCode: `AddinModule.cs`'s `RibSegmentExpand_OnClick` does the same guard directly,
  then `SafeInvokeWpf(() => { var win = new GLExpandOptions(); win.ShowDialogWithOwner(hwnd); })`
  (no `AddinEntry.cs`-equivalent indirection in this codebase).
- `RibExpandAll`/`RibbonExpand1Level` control-ID references removed from:
  AIPowered's `RibbonControlIds.cs` (the const strings plus all 6 shared enable/disable arrays)
  and its one inline array in `AddinModule.cs`; FinalWorkingCode's `Helpers\RibbonStateHelper.cs`
  (6 separate inline literal arrays - no central control-IDs file in this codebase, so this was 6
  edits instead of 1).
- `RibSegmentExplode` (the sibling "Explode" menu with `RibExpodeAll`/`RibbonExplode1Level`) was
  deliberately left untouched in both codebases - out of scope for this request.

**By Columns fill** - `Utilities\SegmentDiscoverer.cs` (both codebases), orientation threaded via
a new `byColumns` parameter (`SegmentAction(ActionType, byColumns = false)` ->
`FillSegmentHierarchies(byColumns)` -> `ExpandSummaryAccountsAsync(..., byColumns)`):
- `ValidateAreaValuesByColumnAsync` (new): column-wise counterpart of `ValidateAreaValuesAsync` -
  reads across row 1 of the selected area (`area.Cells[1, j]` for `j` in `1..Columns.Count`)
  instead of down column 1, one value per column instead of per row.
- `ExpandSummaryAccountsAsync` (extended): when `byColumns` is true, walks `startCol` across the
  area instead of `startRow` down it, calling the new `InsertHierarchyExpansionByColumn` for each
  summary account found and advancing `startCol` by the inserted-child-count + 1 (mirrors the
  existing `startRow` advance logic exactly, just transposed).
- `InsertHierarchyExpansionByColumn` (new): same child-fetch/progress-message steps as
  `InsertHierarchyExpansion`, hands off to `InsertColumnsAndFillData` instead of
  `InsertRowsAndFillData`.
- `InsertColumnsAndFillData` (new): mirror of `InsertRowsAndFillData` with rows/columns
  transposed - inserts new columns via `xlShiftToRight` (not `xlShiftDown`), bulk-fills the
  anchor row across the new columns with a single `[1, children.Count]` `Value2` array (not a
  `[children.Count, 1]` column write), and if the original area spanned multiple rows
  (`multiRow`), copies the anchor column's other rows across into the new columns' matching rows
  (mirrors the existing multi-column sibling-copy for rows).
- `GetInsertedRowCountAsync` renamed to `GetInsertedChildCountAsync` (orientation-agnostic - just
  returns how many children were inserted, whether as rows or columns; reused by both branches of
  `ExpandSummaryAccountsAsync`).
- `GLSegmentDiscovery.xaml.cs`'s existing `DirectionType`/`WriteConfig` enum (Up/Down/DownAll/
  Right/RightAll/Left) was investigated as a possible reusable mechanism but is a private,
  window-scoped type for an unrelated write path (single scalar values written outward from the
  active cell) - not reused directly, though its per-direction-method dispatch idiom was the
  design precedent for the row/column split above.

**Status**: implemented, both codebases. XML/brace/paren-balance verified via the same
`xml.etree.ElementTree.parse` / bracket-count scripts used throughout this log. Not yet
rebuilt/tested by the user - no Windows/MSBuild toolchain available here.

**Follow-up sync (FinalWorkingCode -> AIPowered)**: user edited FinalWorkingCode's
`GLExpandOptions.xaml`/`.xaml.cs` directly (`btnExpand`'s `Content` changed from `"Expand"` to
`"Insert"` - matching `InsertButtonStyle`'s own default Content, which every other window using
that style just inherits rather than overriding; and a `using GLSense.Helpers;` added). Mirrored
both into AIPowered's copy: `Views\GLExpandOptions.xaml`'s `btnExpand.Content` -> `"Insert"`, and
`Views\GLExpandOptions.xaml.cs` gained `using GLSense.Addin.Core.Helpers;` (this codebase's
equivalent of FinalWorkingCode's `GLSense.Helpers`, per the namespace mapping every other ported
file in this project already follows).

---

## 32. Weekly parity sweep: every FinalWorkingCode file changed in the last 7 days, cross-checked and ported into AIPowered

User asked whether today's `GLLogin.xaml.cs` edits (and everything else changed in
FinalWorkingCode over the past week) had made it into AIPowered. No git repo exists in either
codebase, so the audit used filesystem mtimes (`find -mtime`) plus two parallel research agents
comparing ~19 files for functional parity (given the standard namespace/API mapping conventions:
`GLSense.*` -> `GLSense.Addin.Core.*`, `LogUtility.*` -> `ServiceLocator.Logger?.*`, `AppPaths.*` ->
`ServiceLocator.Paths.*`, `AddinModule.RibbonHelper` -> `ServiceLocator.RibbonController?`,
`AppState.Instance.ExcelApp` -> `ServiceLocator.ExcelApp`, WinForms `MessageBoxIcon`/
`MessageBoxButtons` -> WPF `MessageBoxImage`/`MessageBoxButton`). User's decision on scope: "port
everything found." Nine gaps were found and ported into AIPowered:

1. **`Views\GLLogin.xaml.cs`** - today's edits: `using System.Web;` added;
   `CleanLoginUrl` now trims trailing slashes/backslashes
   (`cleanUrl.TrimEnd('/', '\\')`); `ExtractLoginCookies` now URL-decodes the username cookie
   (`HttpUtility.UrlDecode(c.Value) ?? string.Empty`).
2. **`Views\GLSegmentFunctions.xaml.cs`** - `ResultSection.Visibility = Visibility.Collapsed`
   removed from the `"DFF"`/`"ACCOUNTTYPE"` constructor cases, and `Window_Loaded` now
   pre-populates `txtResult.Text` from the live cell's `Value2` for those two function types
   (previously only did this in FinalWorkingCode).
3. **`Themes\Generic.xaml`** - `SuggestAppendComboBox`'s `IsEnabled="False"` trigger gained
   `<Setter Property="Opacity" Value="0.7"/>` (AIPowered's disabled state was staying at full
   opacity).
4. **`Drilldowns\DD_SL.cs` / `DD_JL.cs` / `DD_BL.cs`** - `BuildApiUrl` was missing the
   `jobDescription` query parameter entirely (only sent `jobName`); all three now send both
   `jobName={type}&jobDescription={encoded}`, matching FinalWorkingCode's REST contract.
5. **`Models\GLJobModel.cs`** - `JobName` renamed to `JobDescription`; `DownloadTooltip`'s 6
   branches restored their emoji prefixes (✅🔄⏳❌🚫) that AIPowered's version had dropped.
6. **`ViewModels\GLSubmittedJobsViewModel.cs`** - `JobRecord` had renamed the API's single
   `description` field into two fields (`name` + `jobName`); reverted back to the single
   `description` field to match FinalWorkingCode's DTO exactly (this was flagged as a
   *potential* `NullReferenceException` risk if the live drilldown-processes API only ever
   returns `description`, not `name`/`jobName` - reverting removes that risk). All read sites
   (`ShouldIncludeJob`, `CreateJobModel`, `GetDrillType`, and one easily-missed
   `job.JobName` reference inside a `DDDatatoWorksheet` call around line 978) updated to match.
7. **`Views\GLBalanceConfigurator.xaml`/`.xaml.cs` + `Helpers\DatePickerTooltipHelper.cs`** -
   FinalWorkingCode's date-range blackout/tooltip feature was missing:
   - `dtpStartDate`/`dtpEndDate` gained `CalendarOpened="DatePicker_CalendarOpenedEx"` in the XAML.
   - `DatePicker_CalendarOpenedEx` (new method in the `.xaml.cs`) reads `vm.Periods` to compute
     the ledger's valid min/max date, sets `DisplayDateStart`/`DisplayDateEnd`, and adds
     `BlackoutDates` ranges outside that window so the calendar visually disables out-of-range
     days.
   - `LoadConfiguratorDataAsync` now calls `dtpStartDate.UpdateTooltip()`/
     `dtpEndDate.UpdateTooltip()` after periods load, so the tooltip's date range reflects the
     ledger actually selected (previously only ever showed the static "Click calendar icon..."
     instruction text set at `OnLoaded`).
   - `DatePickerTooltipHelper.CreateTooltipContent` gained the "Available From:"/"Available
     Until:" (or generic "Available:" fallback) range block, driven by
     `DisplayDateStart`/`DisplayDateEnd` and the tooltip's title text.
8. **`ViewModels\GLLovViewModel.cs`** - the hardcoded Balance Type LOV row was missing `JED`,
   `JEDP`, `JEDU` (`{"PTD","YTD","QTD","PJTD","CTD"}` -> `{"PTD","YTD","QTD","PJTD","CTD","JED",
   "JEDP","JEDU"}`), so the LOV item count for that row matches FinalWorkingCode's 8 instead of 5.
9. **`AppConstants.cs`** - added `BalanceTypePTD/YTD/CTD/JED/JEDP/JEDU`, `ActivityDR/CR`,
   `ActivityFlagActual`, `CurrencyTypeTotal/Entered`, `ActualEncumbranceShort`, `PropLedgerName/
   IsSelected/ShortName`, `DateFormatIso` to match FinalWorkingCode's `AppConstants.cs`. Note:
   `GLConfiguratorViewModel.cs` in AIPowered does **not** yet reference these constants - it uses
   the equivalent hardcoded string literals directly (verified the literal values match exactly,
   e.g. `"PTD"`, `"JED"`, `"A+E"`), so this is a parity/documentation addition, not a behavior
   change; wiring the constants in throughout that ~3000-line file would be a much larger, purely
   cosmetic refactor and was intentionally not attempted. Deliberately **not** copied:
   `DefaultVersion`/`DefaultCommitDate` - AIPowered has its own separate version-centralization
   system (section 14) and these two FinalWorkingCode constants are build-specific, not something
   to blindly mirror.

**Explicitly excluded from this sweep**: `GLSenseExcelFunctions.cs`'s UDF sentinel-value handling
diverges between the two codebases (FinalWorkingCode shows a literal placeholder string in some
error cases; AIPowered surfaces the real Excel error value instead). This was judged to be
AIPowered already behaving *more* correctly than FinalWorkingCode, not a regression, so it was
deliberately left unported.

**Status**: all 9 items ported into AIPowered. XML-well-formedness (`ElementTree.parse`) and
brace-balance verified on every touched file; two files (`GLLogin.xaml.cs`, pre-existing;
`GLBalanceConfigurator.xaml.cs`, pre-existing) show a small paren-count mismatch from the crude
counting method picking up parens inside string literals/comments - confirmed both are pre-existing
and unrelated to the edits made here (each edit was individually paren-balance-checked). Not yet
rebuilt/tested by the user - no Windows/MSBuild toolchain available in this environment.

---

## 33. Windows not centered on screen in the shipped MSI (ported from FinalWorkingCode - fixed in **both** codebases)

Reported as a bug against the FinalWorkingCode MSI build, but `BaseWindow.cs` here has the
identical root cause and was fixed the same way once the FinalWorkingCode fix was confirmed
working.

`CenterWindowInExcel()` (called once from `OnLoaded` when `CenterInExcel` is true) computes
`Left`/`Top` from the window's `ActualWidth`/`ActualHeight` at that instant, positioned
against Excel's window rect. Several things can resize the window *after* that one-time
centering, none of which ever recalculated position:

- `FitToAvailableWorkArea()` (also called once from `OnLoaded`, right after
  `CenterWindowInExcel()`) can shrink `Width`/`Height` to fit the screen's work area.
- `ForceSizeToContentResettle()` - the mechanism built in section 1 to fix the
  SizeToContent blank-gap saga - toggles `SizeToContent` off/on and nudges `Width`/`Height`
  by a full pixel to force a genuine native resize. It's called up to three times per
  window: once synchronously in `OnLoaded` (right after `FitToAvailableWorkArea()`), again
  from `OnContentRendered` (guaranteed post-paint, per section 1.4e), and again from
  `AdjustForDpiChange`'s `autoSized` branch on `WM_DPICHANGED`.

Every one of these changes `Width`/`Height` without ever touching `Left`/`Top`, and a
resize always grows/shrinks anchored at the window's current top-left corner - so each
resettle call silently drifted the window further away from wherever `CenterWindowInExcel()`
originally centered it. This is a direct side effect of section 1's own fix: the more
reliably `ForceSizeToContentResettle()` fires post-paint to fix the blank-gap bug, the more
reliably it also un-centers the window afterward.

Fixed by adding `RecenterAfterSizeChange(previousLeft, previousTop, previousWidth,
previousHeight, newWidth, newHeight)`: both `FitToAvailableWorkArea()` and
`ForceSizeToContentResettle()` now capture position/size before they run, and if the size
actually changed, reposition `Left`/`Top` to preserve the same center point afterward
(clamped to `SystemParameters.WorkArea` so it can't be pushed off-screen). `FitToAvailableWorkArea`
compares `Width`/`Height` directly (it assigns those DPs itself); `ForceSizeToContentResettle`
compares `ActualWidth`/`ActualHeight` instead, since a `SizeToContent`-active window's
`Width`/`Height` DPs aren't reliably kept in sync with its true rendered size, and
`ActualWidth`/`ActualHeight` are only guaranteed fresh right after an `UpdateLayout()` call
(which this method already does several times). The DPI-change path needed no separate
handling - it already re-runs `ForceSizeToContentResettle()` for `autoSized` windows, which
now recenters as part of that same call.

Scoped narrowly: `RecenterAfterSizeChange` is only invoked from these two methods when they
themselves changed the size, so it never interferes with `OnMouseMove`'s manual window-drag
logic (which only touches `Left`/`Top`, never `Width`/`Height`, and isn't called from
either of these methods).

**Status: fixed in both codebases.** See FinalWorkingCode's `CLAUDE.md` (`Utilities\DpiAwareWindow.cs`)
for the original write-up - that codebase uses native WPF `WindowStartupLocation=CenterOwner`
instead of a custom `CenterWindowInExcel()`, but has the exact same
resize-after-centering-without-recentering shape in its own `FitToAvailableWorkArea()`/
`EnsureFitsWorkArea()`.

---

## 34. "COM object separated from its underlying RCW" on Explode All / Explode 1 Level (ported from FinalWorkingCode - fixed in **both** codebases)

`Utilities\SegmentDiscoverer.cs`'s `SegmentAction` caches the source worksheet in a
class-level field (`HrWorksheet = CellActive.Worksheet`) before doing any async work.
Many awaits later, `CreateSingleSheetAsync` reads that cached field
(`HrWorksheet?.Index`, then `HrWorksheet?.Copy(...)`) - but just before that, it calls
`SheetExists`/`GetWorksheetByName` to check whether a sheet with the target name already
exists, and both helpers called `ExcelComHelper.SafeRelease(ws, "Worksheet")`
(`Marshal.FinalReleaseComObject`) on *every* worksheet enumerated that didn't match the
target name - including the source sheet, since the new sheet's sanitized name is
essentially never equal to the source sheet's own name. Since .NET's classic COM
interop layer caches one RCW per underlying COM object (within the same execution
context), the `ws` yielded for that sheet during enumeration is the *same* wrapper as
the cached `HrWorksheet` field - so releasing it there detached `HrWorksheet`'s RCW too,
and the very next lines in the same call threw "COM object that has been separated from
its underlying RCW cannot be used." Fully deterministic on the first child sheet of the
first click, not timing-dependent - and confirmed **not** a cross-AppDomain marshaling
issue (unlike section 25's `SanitizeSheetName` bug): the byte-for-byte identical code in
FinalWorkingCode, which has no AppDomain boundary at all, hits the exact same crash.

Fixed by removing the `SafeRelease` calls from `SheetExists`/`GetWorksheetByName`
entirely - both are lightweight name-lookup helpers with no long-lived object graph to
worry about (unlike Application/Workbook/Worksheets, which is what actually keeps
Excel.exe alive after close, per section 29), so they don't need to force-release
anything, and doing so here was unsafe given `HrWorksheet`'s lifetime. Verified these two
methods are called from nowhere else in the file.

**Status: fixed in both codebases**, user-confirmed working in FinalWorkingCode before
porting here.

## 35. GLSegmentValues/GLSegmentManager: Hierarchy field not cleared when segment changes (ported from FinalWorkingCode - fixed in **both** codebases)

`SegmentSelectorViewModel.SelectedSegment`'s setter kicks off `LoadSegmentValuesAsync()`
for the newly selected segment, which repopulates `HierarchyItems` (the Hierarchy
combo's `ItemsSource`) for the new segment - but nothing cleared `SelectedHierarchy`
itself (the combo's bound `SelectedItem`), so after picking a hierarchy value and then
switching segments, the Hierarchy combo kept showing the previous segment's stale
selection even though it no longer applied. Since `GLSegmentManager` shares this same
`SegmentSelectorViewModel` (see section 26/9), this affected both windows.

Fixed by clearing the `_selectedHierarchy` backing field directly (not through the
`SelectedHierarchy` property setter) whenever `SelectedSegment` actually changes.
Bypassing the property setter is deliberate: setting it normally would also fire
`LoadHierarchySegmentValuesAsync()`, which would run concurrently with (and race
against) the `LoadSegmentValuesAsync()` call already firing for the newly selected
segment - clearing the field directly and raising `OnPropertyChanged` avoids that.

**Status: fixed in both codebases**, user-confirmed working in FinalWorkingCode before
porting here.

---

## Deployment note (important when a fix "doesn't seem to work")

`GLSense.Addin.Core` loads into a separate, shadow-copied AppDomain
(`GLSense.Loader.Core\AddinDomainLoader.cs`) from a **versioned deployment folder**
(`%LOCALAPPDATA%\ORBIT\Excel_Logs\GLSense_Logs_New\Versions\v11.1.0\`), refreshed only
by `GLSense.Addin.Core\post_build.cmd` (a `<PostBuildEvent>` that `xcopy`s
`$(TargetDir)` into that versioned folder). Saving a `.xaml`/`.cs` file does **not**
update the deployed copy - only an actual Rebuild does. And because the AppDomain is
created once when Excel starts, **a running Excel session won't pick up a fresh build
either** - Excel itself must be fully closed and relaunched after rebuilding. When a fix
appears not to have taken effect, rule this out before re-investigating the code.
