# AIPowered UI: remove WPF-UI, sync windows with FinalWorkingCode

## Problem

AIPowered's window layer (`GLSense.Addin.Core\Views\*`) is built on `BaseWindow : Wpf.Ui.Controls.FluentWindow`
(the WPF-UI package). The user has wanted WPF-UI gone for some time but it never actually got removed, and
independently of that, AIPowered's window layouts have drifted from FinalWorkingCode's (the reference/"golden"
UI) through many rounds of incremental, WPF-UI-specific workarounds (see `CLAUDE.md` section 1's whole
`SizeToContent`/blank-gap saga, section 24.3's combo-scroll saga, etc.). The goal is to bring AIPowered's UI
back in sync with FinalWorkingCode's — same layout, same behavior, same visual style — while keeping
AIPowered's actual reason to exist: Excel/UI/business-logic separated into `GLSense` (host) /
`GLSense.Addin.Core` (hot-reloadable) instead of FinalWorkingCode's single monolith, so updates don't require
a full MSI reinstall.

## Current state (confirmed by inspection)

- `BaseWindow` (`GLSense.Addin.Core\Views\BaseWindow.cs`) derives from `Wpf.Ui.Controls.FluentWindow` and uses
  WPF-UI's `SizeToContent`-driven auto-measurement model. `CLAUDE.md` section 1 documents an extensive,
  multi-round fight to keep that model from producing a visible "blank gap" on window open.
- FinalWorkingCode's equivalent base class, `DpiAwareWindow` (`FinalWorkingCode\GLSense\Utilities\DpiAwareWindow.cs`),
  is a plain `Window` with **no** WPF-UI dependency at all. It solves the same "gap on open" problem
  completely differently: `SizeToContent="Manual"` always, with all fit/scale/center math done synchronously
  in `OnSourceInitialized` (before the window is ever shown/painted), plus a shared `WindowLoadingPlaceholder`
  overlay shown until the window's first real frame renders. This is a fundamentally more robust design than
  fighting WPF-UI's auto-measurement, and retires essentially all of `CLAUDE.md` section 1's workarounds.
- The actual WPF-UI footprint in AIPowered is shallow: `BaseWindow`'s base class, `WpfUiBootstrapper.cs`
  (exists solely to hand-feed `FluentWindow` the theme brushes/resources it needs), two `<Reference>` entries
  in `GLSense.Addin.Core.csproj` + one `packages.config` entry, and unused `xmlns:ui` namespace declarations
  in every window XAML (no `<ui:...>` controls remain anywhere — icons were already migrated to FontAwesome
  in an earlier session, see `CLAUDE.md` section 4).
- FinalWorkingCode has zero WPF-UI references anywhere. Its `Themes\GlobalStyles.xaml`/`Themes\Generic.xaml`
  are fully self-contained (define their own brushes from scratch), so no theme-manager class is needed once
  `BaseWindow` no longer derives from `FluentWindow`.
- AIPowered already has ported copies of `WpfAppManager.cs` and `MouseWheelFocusHelper.cs` (both of which
  `DpiAwareWindow` depends on). It is missing `WindowLoadingPlaceholder.cs`, which needs to be added.
- 26 windows derive from `BaseWindow`. Of the windows that exist on both sides, all but one (`GLSegmentManager`)
  have a direct FinalWorkingCode counterpart file to sync against.
- `GLSegmentManager` (`GLSense.Addin.Core\Views\GLSegmentManager.xaml(.cs)`) is an AIPowered-only "master-detail
  redesign" of `GLSegmentRef`, built as a side-by-side trial (see its own header comment and
  `GLBalanceConfigurator.xaml.cs` lines ~207-222, which explicitly documents how to roll the trial back). The
  user has decided to end the trial: **`GLSegmentManager` is retired in favor of `GLSegmentRef`**, which is
  itself already a faithful, previously-completed, still-current port of FinalWorkingCode's `GLSegmentRef.xaml`
  (same `SegmentSelectorViewModel`, same construction pattern as `GLSegmentValues`).
  - Only two real (non-comment) references to `GLSegmentManager` exist outside its own files: the two
    `<Compile>`/`<Page>` entries in `GLSense.Addin.Core.csproj`, and the single instantiation
    `new GLSegmentManager(AcctsRef.Text)` in `GLBalanceConfigurator.xaml.cs`. Every other hit (in
    `GLWaitWindow.xaml`, `GLRollerGroups.xaml.cs`, `GLSegmentValues.xaml.cs`, `GLSegmentRef.xaml.cs`,
    `DataGridColumnFillHelper.cs`, `GlobalStyles.xaml`, `LedgerModel.cs`, `SegmentSelectorViewModel.cs`,
    `AttachmentsDialog.xaml`, `Converters.cs`, `BaseWindow.cs`) is prose in a comment documenting a shared bug
    fix — informational only, not a code dependency, and not touched by this work unless that file is already
    being edited for another reason in its own batch.

## Scope

**In scope:**
1. Replace `BaseWindow`'s implementation with FinalWorkingCode's `DpiAwareWindow` engine (same public class
   name `BaseWindow`, so every `: BaseWindow` declaration across ~26 windows is untouched). Port
   `WindowLoadingPlaceholder.cs`. Retire `WpfUiBootstrapper.cs` in favor of directly merging FinalWorkingCode's
   `GlobalStyles.xaml`/`Generic.xaml` into `Application.Current.Resources`.
2. Remove WPF-UI entirely: the two `<Reference>` entries + `packages.config` entry in
   `GLSense.Addin.Core.csproj`, and the unused `xmlns:ui` declarations across all window XAML files.
3. Retire `GLSegmentManager`: revert `GLBalanceConfigurator.xaml.cs`'s call site to
   `new GLSegmentRef(AcctsRef.Text)` (and clean up the now-stale "TRIAL" comment block around it), delete
   `GLSegmentManager.xaml`/`GLSegmentManager.xaml.cs`, remove its two `csproj` entries.
4. For every remaining shared window, port FinalWorkingCode's XAML wholesale (structure, `SizeToContent`
   mode, explicit Width/Height/Min/Max, DataGrid column widths, etc.) replacing AIPowered's current markup,
   removing now-obsolete WPF-UI-workaround code (`DataGridColumnFillHelper` calls, `ForceSizeToContentResettle`/
   `PumpDispatcherFrame` calls, `OnContentRendered` resettle hooks, etc.) where FinalWorkingCode doesn't need
   them under the new base class.
5. Preserve every AIPowered-specific, non-layout fix that has no FinalWorkingCode counterpart because
   FinalWorkingCode doesn't have the architecture that requires it: cross-AppDomain-safe COM access
   (`CLAUDE.md` §25), `AddinBeginShutdown` teardown (§29), dispatcher-thread marshaling in ViewModels
   (§2.4/§2.5/§21.2), `TryDisableExcelSettings`/calculation-mode guards (§36/§37), the low-level mouse-hook
   combo-scroll fix (§24.3.5), and all Excel-COM/business-logic code in general.

**Out of scope (unless discovered to be load-bearing during a batch):**
- Any change to `GLBalanceConfigurator`'s `UserControl`/`ConfiguratorPaneHost` HWND-reparenting mechanism —
  it doesn't derive from `BaseWindow` and isn't part of the base-class swap; it only inherits new theme
  resources automatically via the `GlobalStyles.xaml`/`Generic.xaml` replacement.
- Cleaning up the historical `GLSegmentManager` comment references listed above, except incidentally.
- Any ribbon/business-logic behavior change not required by the layout/base-class port.

## Target architecture

- `BaseWindow : Window` (was `: FluentWindow`), carrying `DpiAwareWindow`'s full engine: synchronous
  `OnSourceInitialized`-time fit/scale/center (before first paint), `FitToAvailableWorkArea`,
  `RecenterAfterSizeChange`, `EnsureFitsWorkArea`, `WM_DPICHANGED` handling, `WindowLoadingPlaceholder`
  integration, `ShowDialogWithOwner`/`ShowWithOwner`/`SetExcelOwner`. AIPowered-specific adaptations kept:
  `ServiceLocator.Logger` logging (instead of `LogUtility`), Excel handle sourced from `ServiceLocator`
  (instead of a caller-supplied `IntPtr`), existing public property names where they already match
  (`EnableAutoLayoutRefresh`, `EnableExcelCentering`, `EnableEscapeToClose`, `AutoClampToWorkArea`,
  `WorkAreaMargin`, `IconSymbol`).
- New file: `GLSense.Addin.Core\Utilities\WindowLoadingPlaceholder.cs`, ported from FinalWorkingCode, adapted
  to `ServiceLocator.Logger` logging.
- Theme bootstrap: at first `BaseWindow` construction (or an equivalent single init point), merge
  FinalWorkingCode's `GlobalStyles.xaml`/`Generic.xaml` (already the files AIPowered has, now brought in sync
  with FinalWorkingCode's actual resource definitions) into `Application.Current.Resources.MergedDictionaries`
  directly — no theme-manager class, matching how FinalWorkingCode does it.
- `WpfUiBootstrapper.cs` deleted; its one call site in `BaseWindow`'s constructor removed.

## Execution plan (phased, with checkpoints — no build/run access exists in this environment)

**Phase 1 — Foundation + 3 pilots.** Port the `BaseWindow`/`WindowLoadingPlaceholder`/theme-resource changes
described above, remove the WPF-UI package references, and apply the wholesale-XAML-port + code-behind cleanup
to 3 pilot windows chosen to cover the three structural shapes in play:
- `GLWaitWindow` (simplest — static, no grid, validates the base class works at all)
- `GLMessageWindow` (variable-height text — the one window deliberately kept off a fixed-size model)
- `GLSegmentValues` (already `SizeToContent="Manual"` + fixed bounds + dual DataGrid on both sides — lowest-risk
  check that FinalWorkingCode's exact numbers/structure drop in cleanly)

STOP. User rebuilds and tests: does each pilot open without a blank gap, center correctly, resize/DPI-change
correctly, and (for `GLSegmentValues`) scroll/select correctly?

**Phase 2 — Batch A: Period/date family (6).** `GLGetPeriod`, `GLGetPeriodByDate`, `GLGetPeriodByYear`,
`GLGetPeriodDetails`, `GLGetPeriodStartEnd`, `GLDailyRates`. STOP, user tests.

**Phase 3 — Batch B: Login/static utility (5).** `GLAbout`, `GLLoginDetails`, `GLLogin`, `AttachmentsDialog`,
`WebView2PopupWindow`. STOP, user tests.

**Phase 4 — Batch C: Segment/hierarchy (5 windows to sync).** `GLSegmentDiscovery`, `GLSegmentFunctions`,
`GLSegmentRef` (now the sole segment-picker, replacing the retired `GLSegmentManager`), `GLRollerGroups`,
`GLExpandOptions`. This phase also bundles the `GLSegmentManager` retirement (delete files, revert
`GLBalanceConfigurator.xaml.cs` call site, remove `csproj` entries) since it's directly tied to `GLSegmentRef`'s
sync. STOP, user tests — specifically re-testing the Balance Configurator's "Account Assignment(s)" Edit flow,
since that's the call site that changed.

**Phase 5 — Batch D: Data-grid/master windows (5).** `GLCubeDetails`, `GLUserConfig`, `GLServerConfiguration`,
`GLLOVs`, `GLJobsMonitor`. STOP, user tests.

**Phase 6 — Batch E: remaining (1).** `GLDrilldownCustomization`. Also: spot-check `AppOverlay`,
`ExcelRefEditControl`, and `GLBalanceConfigurator`'s `UserControl` for anything that visually assumed the old
WPF-UI-fed brush values, now that the theme dictionaries have been fully replaced. STOP, user does final
end-to-end test.

Each phase after Phase 1 re-verifies that `SegmentSelectorViewModel` (shared by `GLSegmentValues`,
`GLSegmentRef`) still behaves correctly for whichever of its consumers lands in that phase — not assumed safe
just because an earlier consumer tested fine.

## Verification protocol per phase

**Correction after spec approval**: this environment does have a working MSBuild toolchain (Visual Studio 2022
Community, resolved via `vswhere`). `GLSense.Addin.Core.csproj` builds cleanly on its own with
`/p:SignAssembly=false /p:Configuration=Debug` (0 warnings, 0 errors — confirmed by an actual build run; the
`SignAssembly=false` override is needed only because this machine doesn't have the real
`GLSense.Contracts.pfx` signing key, per `CLAUDE.md` §36/§37's own note — verification-only, never used for a
real deliverable build). This means every task in the implementation plan gets a real, automated compile-check
gate: `msbuild GLSense.Addin.Core\GLSense.Addin.Core.csproj /t:Build /p:Configuration=Debug
/p:SignAssembly=false /v:normal /nologo` must report `0 Error(s)` before a task is considered done, not just an
XML-well-formedness/brace-balance read-through.

What automated compilation **cannot** catch: whether a window actually centers correctly, resettles cleanly on
DPI change, avoids the old "blank gap," or behaves correctly interactively at runtime inside Excel — none of
that runs without a real Excel session. So each phase still ends with the user rebuilding (already done as part
of the plan's own build step) and relaunching Excel to exercise that phase's windows; the plan's automated gate
just means whatever reaches that step is guaranteed to compile, so the user's time is spent testing behavior,
not chasing compile errors.

## Risks

- **Shared ViewModel regressions**: `SegmentSelectorViewModel` backs `GLSegmentValues` (Phase 1) and
  `GLSegmentRef` (Phase 4) — a change that looks safe against one can still break the other. Compilation alone
  won't catch a behavioral regression here; both consumers need an explicit re-check when the second one's
  phase lands.
- **Silent behavior loss**: FinalWorkingCode's code-behind is a reference for layout-driving methods only;
  business logic / Excel-COM / AppDomain-safety code specific to AIPowered must be re-verified present after
  each file's wholesale XAML replacement, not just carried over by assumption. This is a manual-review risk,
  not something the compiler flags (removed business logic can still compile cleanly).
- **Runtime-only behavior**: window centering, DPI-change handling, and the "blank gap" fix itself can only be
  confirmed by the user actually running Excel — the build gate proves the code compiles, not that it looks
  or behaves right.
