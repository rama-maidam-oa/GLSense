# AIPowered windows: header/chrome + style + sizing parity with FinalWorkingCode (Phase 2)

## Problem

Phase 1 (see `docs/superpowers/specs/2026-08-19-wpfui-removal-finalworkingcode-sync-design.md`) replaced
`BaseWindow`'s WPF-UI-derived engine with `DpiAwareWindow`'s model and proved it against 3 pilot windows
(`GLWaitWindow`, `GLMessageWindow`, `GLSegmentValues`). Those 3 windows now match FinalWorkingCode visually:
custom header bar with its own close button (native Windows title bar suppressed), FinalWorkingCode's control
styles/colors/fonts, and FinalWorkingCode's exact sizing (`SizeToContent` mode, Width/Height/Min/Max).

The remaining 21 `BaseWindow`-derived windows still have AIPowered's older look: no custom header/close button
(or an inconsistent one), some now-stale WPF-UI-era styling, and layout/sizing that has drifted from
FinalWorkingCode's numbers through the many incremental fixes documented across `CLAUDE.md`. The user wants
every one of them brought back in sync with FinalWorkingCode — "size, height, width, layout, everything."

## Explicit deviation from the Phase 1 spec

Phase 1's own spec planned a future "Phase 4" that retires `GLSegmentManager` in favor of `GLSegmentRef`. The
user has since told this session directly: **do not touch `GLSegmentRef` or `GLSegmentManager` — that is a
separate, final piece to be done on its own.** This spec excludes both from scope entirely, superseding that
part of the Phase 1 spec. Nothing here should assume or depend on that retirement having happened.

## Scope

**In scope — 21 windows**, batched by risk (see Execution plan): `GLAbout`, `AttachmentsDialog`,
`WebView2PopupWindow`, `GLExpandOptions`, `GLGetPeriod`, `GLGetPeriodByDate`, `GLGetPeriodByYear`,
`GLGetPeriodDetails`, `GLGetPeriodStartEnd`, `GLDailyRates`, `GLSegmentDiscovery`, `GLSegmentFunctions`,
`GLLoginDetails`, `GLLOVs`, `GLRollerGroups`, `GLJobsMonitor`, `GLUserConfig`, `GLServerConfiguration`,
`GLCubeDetails`, `GLLogin`, `GLDrilldownCustomization`.

For each window: port FinalWorkingCode's custom header bar (title text, own close button, whatever
per-window icon/accent it uses) and native-title-bar suppression, control styles/colors/fonts, and sizing
(`SizeToContent` mode + explicit Width/Height/Min/Max/DataGrid column widths) — matching the wholesale-XAML-port
approach Phase 1 used for `GLWaitWindow`/`GLMessageWindow`.

**Preserve, do not revert** — every AIPowered-specific fix with no FinalWorkingCode counterpart, because
FinalWorkingCode's architecture doesn't need it:
- Cross-AppDomain-safe COM access patterns (`CLAUDE.md` §25).
- ViewModel dispatcher-thread marshaling (§2.4/§2.5/§21.2).
- `DataGridColumnFillHelper` usage where it is still the *correct* mechanism for a genuinely
  `SizeToContent="WidthAndHeight"` window — but see the note below on `GLSegmentManager`-adjacent windows.
- `SegmentSelectorViewModel`-driven bindings and callbacks (`DataLoadedAction`, etc.) used by `GLRollerGroups`.
- Any other behavior fix cataloged in `CLAUDE.md` sections 2 through 37 for a window in this batch — read that
  window's own history in `CLAUDE.md` before editing it, the same discipline every prior fix in this codebase
  has followed.
- `WatermarkTextBox` style on `GLLOVs`' Comments field, tooltips, ribbon-sync logic, and any other
  window-specific ViewModel/code-behind wiring not related to header/chrome/style/sizing.

For windows with real ViewModel/data-binding complexity (`GLLOVs`, `GLRollerGroups`, `GLJobsMonitor`,
`GLUserConfig`, `GLServerConfiguration`, `GLCubeDetails`, `GLLogin`, `GLDrilldownCustomization`): reconcile the
code-behind by hand rather than blindly overwriting, following pilot 3's (`GLSegmentValues`) precedent from
Phase 1 — read both sides in full before editing.

**Out of scope:**
- `GLSegmentRef`, `GLSegmentManager` — explicitly excluded per the user's direct instruction.
- `GLBalanceConfigurator` (`UserControl`, no window chrome — different hosting model) and `AppOverlay`/
  `ExcelRefEditControl` — untouched, same exclusion Phase 1 already established.
- Any ribbon/business-logic behavior change not required by the layout/chrome/style port.
- Deleting now-dead `GlobalStyles.xaml`/`Generic.xaml` keys as each batch stops needing them — deferred to one
  final prune pass after all 5 batches land (mirrors Phase 1's own Task 8), so no batch risks breaking a
  still-active window's styling mid-effort.

## Execution plan — 5 batches, simplest/lowest-risk first

**Batch 1 (4):** `GLAbout`, `AttachmentsDialog`, `WebView2PopupWindow`, `GLExpandOptions` — simple, mostly
static content, closest in shape to Phase 1's `GLWaitWindow`/`GLMessageWindow` pilots.

**Batch 2 (5):** `GLGetPeriod`, `GLGetPeriodByDate`, `GLGetPeriodByYear`, `GLGetPeriodDetails`,
`GLGetPeriodStartEnd` — near-identical family (`CLAUDE.md` §1.4/§1.4f/§8.4 already establishes their shared
`MinHeight`/`MaxHeight` pattern and confirms them "perfect" post-revert — port chrome/style only, do not
re-litigate their sizing numbers without checking that history first).

**Batch 3 (4):** `GLDailyRates`, `GLSegmentDiscovery`, `GLSegmentFunctions`, `GLLoginDetails` — other static
forms, same low-risk shape as Batch 2 but not a uniform family.

**Batch 4 (4):** `GLLOVs`, `GLRollerGroups`, `GLJobsMonitor`, `GLUserConfig` — DataGrid/ViewModel-heavy,
reconcile-carefully per the Scope section above.

**Batch 5 (4):** `GLServerConfiguration`, `GLCubeDetails`, `GLLogin`, `GLDrilldownCustomization` — highest
complexity (async data load, WebView2, ribbon sync) — last, once the pattern is well-proven across 17 other
windows.

Each batch: port all windows in the batch, run the build-verification command below, commit, then stop for the
user to rebuild and visually check in Excel before the next batch starts.

## Verification protocol per batch

MSBuild is available in this environment (confirmed: `C:\Program Files\Microsoft Visual Studio\2022\Community\
MSBuild\Current\Bin\MSBuild.exe`), so every batch gets a real compile-check gate, not just an XML/brace-balance
read-through:

```
msbuild GLSense.Addin.Core\GLSense.Addin.Core.csproj /t:Build /p:Configuration=Debug /p:SignAssembly=false /v:normal /nologo
```

Must report `0 Error(s)` before a batch is considered done. `SignAssembly=false` is verification-only (this
machine lacks the real `GLSense.Contracts.pfx` signing key) — never used for a real deliverable build. A clean
compile does not prove the visual result is correct (WPF layout/styling bugs mostly aren't compile errors) — the
user's own rebuild + visual check in Excel after each batch remains the actual acceptance test, same as every
other fix in this codebase (see `CLAUDE.md`'s deployment note: a running Excel session won't pick up a fresh
build either — Excel must be fully closed and relaunched).

## Deferred / explicitly not attempted here

- `GLSegmentRef`/`GLSegmentManager` retirement or sync — separate future effort per the user's instruction.
- Final prune of now-dead `GlobalStyles.xaml`/`Generic.xaml` keys — one pass after all 5 batches land.
- Any change to `GLBalanceConfigurator`, `AppOverlay`, `ExcelRefEditControl`.
