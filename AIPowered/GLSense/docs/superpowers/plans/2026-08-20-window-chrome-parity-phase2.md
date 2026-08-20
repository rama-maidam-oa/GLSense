# Window Header/Chrome/Style/Sizing Parity — Phase 2 Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Bring the remaining 21 `BaseWindow`-derived AIPowered windows into full visual parity with their
FinalWorkingCode counterparts — custom header bar + close button (replacing the native title bar), control
styles/colors/fonts, and sizing/layout (`SizeToContent` mode + explicit Width/Height/Min/Max) — while preserving
every AIPowered-specific business-logic fix that has no FinalWorkingCode counterpart.

**Architecture:** Each window's `.xaml` gets FinalWorkingCode's header `Border`/`HeaderBar` style, close button,
control styles, and sizing ported in, following the exact convention already established and shipped by Phase 1's
3 pilots (`GLWaitWindow`, `GLMessageWindow`, `GLSegmentValues`) — see Global Constraints below for the literal
header markup to reuse. Code-behind is reconciled by hand, not overwritten, for every window with real
AIPowered-specific ViewModel/business logic (flagged per-task below).

**Tech Stack:** .NET Framework 4.8.1, WPF, AddinExpress. MSBuild via
`C:\Program Files\Microsoft Visual Studio\2022\Community\MSBuild\Current\Bin\MSBuild.exe`.

**Spec:** `docs/superpowers/specs/2026-08-20-window-chrome-parity-phase2-design.md`

## Global Constraints

- Build verification command for every task: `msbuild GLSense.Addin.Core\GLSense.Addin.Core.csproj /t:Build
  /p:Configuration=Debug /p:SignAssembly=false /v:normal /nologo` — must report `0 Error(s)`. `SignAssembly=false`
  is verification-only (this machine lacks the real `GLSense.Contracts.pfx`); never use it for a real deliverable
  build.
- **Full-sync policy**: per the user's explicit instruction ("size, height, width, layout, everything should be
  in sync"), every window in this plan adopts FinalWorkingCode's exact visible header text, icon, control
  styles/colors, and sizing numbers — even where an older `CLAUDE.md` note calls AIPowered's current value
  "confirmed perfect" or "the user's own fix." Those notes predate Phase 1's `BaseWindow` rewrite (which replaced
  the old WPF-UI/`FluentWindow` sizing engine with the `DpiAwareWindow`-style engine FinalWorkingCode's numbers
  were tuned against) and are treated as superseded for this effort, not violated. The one exception: business
  logic/bindings explicitly called out as "preserve" in a task below — those never get overwritten, regardless of
  which file's XAML looks more "correct."
- **Canonical header markup** (reuse verbatim in every task, substituting only the `Kind`/`Text` values and
  omitting the `Margin` if the window's own rows already carry independently-tuned margins — see `GLSegmentValues`'s
  own header comment for when to omit it):
  ```xml
  <Border Grid.Row="0" Style="{StaticResource HeaderBar}" MouseLeftButtonDown="TitleBar_MouseLeftButtonDown">
      <Grid>
          <Grid.ColumnDefinitions>
              <ColumnDefinition Width="Auto"/>
              <ColumnDefinition Width="*"/>
              <ColumnDefinition Width="Auto"/>
          </Grid.ColumnDefinitions>
          <StackPanel Grid.Column="0" Orientation="Horizontal" VerticalAlignment="Center">
              <iconPacks:PackIconFontAwesome Kind="<IconKind>" Style="{StaticResource LargeIcon}"/>
              <TextBlock Text="<Title>" Style="{StaticResource TitleHeaderTextBlock}"/>
          </StackPanel>
          <Button Grid.Column="2" Style="{StaticResource CustomWindowCloseButtonStyle}" Click="BtnClose_Click"/>
      </Grid>
  </Border>
  ```
  Root element gets `WindowStyle="None"` and (for windows moving off `WidthAndHeight`) `Background="#F8F9FA"`,
  replacing the old `TitleBarGridStyle`/`TitleBarIconStyle`/`TitleBarTextStyle`/`TitleBarCloseButtonStyle`/
  `ExtendsContentIntoTitleBar="True"`/`WindowStyle="SingleBorderWindow"` idiom entirely.
- **Drag-handler policy**: `EnhancedDragDropHelper.cs` does NOT exist anywhere in AIPowered (grep-confirmed empty
  across the whole codebase) even though FinalWorkingCode's own constructors call
  `EnhancedDragDropHelper.EnableWindowDrag(this)` instead of wiring a header `MouseLeftButtonDown`. Every window in
  this plan already has its own working `TitleBar_MouseLeftButtonDown` handler in its `.xaml.cs` (confirmed present
  in all 21 by the batch research) — keep it, wire it onto the new header `Border` exactly as shown above. Never
  add an `EnhancedDragDropHelper` call.
- Standard namespace mapping (same as Phase 1): `GLSense.Views` → `GLSense.Addin.Core.Views`,
  `GLSense.Utilities` → `GLSense.Addin.Core.Utilities`, `GLSense.Controls` → `GLSense.Addin.Core.Controls`,
  `GLSense.Converters` → `GLSense.Addin.Core.Converters`, `GLSense;component` → `GLSense.Addin.Core;component`,
  `LogUtility.*` → `ServiceLocator.Logger?.*`, `AppState.Instance.ExcelApp` → `ServiceLocator.ExcelApp`. Root
  element `<utils:DpiAwareWindow x:Class="GLSense.Views.<Name>">` → `<views:BaseWindow
  x:Class="GLSense.Addin.Core.Views.<Name>">`.
- `GLSegmentRef` and `GLSegmentManager` are out of scope for this entire plan — do not touch either file in any
  task, even incidentally.
- `GLBalanceConfigurator`, `AppOverlay`, `ExcelRefEditControl` are out of scope — do not touch.
- Theme keys used throughout this plan (`HeaderBar`, `LargeIcon`, `TitleHeaderTextBlock`,
  `CustomWindowCloseButtonStyle`, `ContentBorderStyle`, `CompactBorderStyle`, `HeaderLabel`, `SectionHeader`,
  `CompactRefEditControl`, `ModernDataGridColumnHeader`, `ModernDataGridRowWithBorder`, `ModernDataGridCell`,
  `ModernDataGridTextBlock`, `TabItemStyle`/`ModernTabItem`, `CompactCardStyle`, `TextBlockLabelStyle`,
  `CompactCheckBox`, `SimpleBrowserToolTip`) already exist in AIPowered's `GlobalStyles.xaml`/`Generic.xaml` per
  Phase 1 Task 3's additive port. If any single grep for a key used in a task below comes back empty, stop and
  add that key's definition (copied from FinalWorkingCode's `Themes/GlobalStyles.xaml`) before proceeding with
  that task — do not silently skip the style and leave the old inline markup in place.
- Do not remove a `DataGridColumnFillHelper.EnableFillColumn(...)`/`.Refresh(...)` call from a window unless that
  task explicitly says to (only applies when a window's `SizeToContent` is changing away from `WidthAndHeight` to
  `Manual` — the helper becomes actively harmful under `Manual` per `CLAUDE.md` §8.1's finding, established on
  `GLSegmentValues`). Leave it alone on any window staying on `WidthAndHeight`.

---

## Phase 1 (Batch 1): Simple/static windows

### Task 1: Port `GLAbout` header/chrome/sizing

**Files:**
- Modify: `GLSense.Addin.Core\Views\GLAbout.xaml`
- Modify: `GLSense.Addin.Core\Views\GLAbout.xaml.cs`

- [ ] **Step 1: Read both sides in full before editing**

Read `FinalWorkingCode\GLSense\Views\GLAbout.xaml`/`.xaml.cs` and AIPowered's current copies in full. This window
carries real, AIPowered-specific version/build-date/compatibility-check logic (`CLAUDE.md` §8.6, §19) that a
careless port could destroy — this is a reconcile task, not a blind copy.

- [ ] **Step 2: Replace the header**

Apply the canonical header markup from Global Constraints: `Kind="CircleInfoSolid"`, `Text="About Orbit GLSense"`.
Remove the old `TitleBarGridStyle` Grid entirely. Root element: `WindowStyle="None"`.

- [ ] **Step 3: Port sizing**

Change `SizeToContent="WidthAndHeight" MinWidth="450" MaxWidth="650" MinHeight="500" MaxHeight="620"` to
`SizeToContent="Manual" MinWidth="400" MaxWidth="650" MinHeight="450" MaxHeight="600"`. Add
`Background="#F8F9FA"` to the root element. Remove `ExtendsContentIntoTitleBar="True"`.

- [ ] **Step 4: Port control styles**

Change the Company/Product and Instance Compatibility section `Border`s from inline
`Background="{StaticResource WhiteBrush}" CornerRadius="8" BorderThickness="1" BorderBrush="{StaticResource
BorderLightBrush}"` to `Style="{StaticResource ContentBorderStyle}"`. Leave `dgInstances`'s `ColumnHeaderStyle`/
`RowStyle` untouched (already byte-identical to FinalWorkingCode's). Keep the footer's own Close button (`btnClose`,
`CloseButtonStyle`) — already correct per `CLAUDE.md` §6.1, do not remove it just because the header now has its
own close button too (both are intentional).

- [ ] **Step 5: Reconcile code-behind — preserve verbatim**

Do NOT introduce FinalWorkingCode's `{x:Static data:AppConstants.DefaultVersion}` XAML bindings for
`txtVersion`/`txtBuildDate` — AIPowered has no such constants; keep `SetVersionAndBuildDateText()`'s imperative,
manifest-driven approach (`ServiceLocator.Version`/`ServiceLocator.ReleaseDate` + `FormatBuildDate()`, `CLAUDE.md`
§19.1/§19.4) exactly as it is. Keep `xmlFilePath = ServiceLocator.Paths.UrlsDirectory` (not FinalWorkingCode's
`AppPaths.TempUrlsPath`). Keep `LogInstanceCheckFailure`/the categorized-exception `CheckUrlCompatibility` retry
logic (`CLAUDE.md` §19.5) exactly as it is — do not replace with FinalWorkingCode's simpler split. Keep the
`/GLSense.Addin.Core;component/Images/orbit_logo.png` image path (§8.6). Keep the `TitleBar_MouseLeftButtonDown`
handler, rewired onto the new header `Border` per the Global Constraints drag-handler policy.

- [ ] **Step 6: Build verification**

Run: `msbuild GLSense.Addin.Core\GLSense.Addin.Core.csproj /t:Build /p:Configuration=Debug /p:SignAssembly=false
/v:normal /nologo`
Expected: `Build succeeded. 0 Error(s)`.

- [ ] **Step 7: Commit**

```bash
git add GLSense.Addin.Core/Views/GLAbout.xaml GLSense.Addin.Core/Views/GLAbout.xaml.cs
git commit -m "Port GLAbout header/chrome/sizing parity with FinalWorkingCode"
```

---

### Task 2: Port `AttachmentsDialog` header/chrome/sizing

**Files:**
- Modify: `GLSense.Addin.Core\Views\AttachmentsDialog.xaml`
- Modify: `GLSense.Addin.Core\Views\AttachmentsDialog.xaml.cs`

- [ ] **Step 1: Replace the header**

Canonical header markup: `Kind="PaperclipSolid"`, `Text="Select Attachments to Download"`. Root: `WindowStyle="None"`.

- [ ] **Step 2: Port sizing**

Change `SizeToContent="WidthAndHeight"` to `SizeToContent="Manual"`. `MinWidth="340" MaxWidth="600" MinHeight="400"
MaxHeight="700"` already match FinalWorkingCode — leave those 4 numbers unchanged. Change
`WindowStartupLocation="CenterOwner"` to `WindowStartupLocation="Manual"` (matches FinalWorkingCode — `BaseWindow`'s
own `CenterOverExcelOnce()` handles centering under `Manual`, not WPF's native startup-location mechanism). Add
`Background="#F8F9FA"`. Remove `ExtendsContentIntoTitleBar="True"`. Keep `ShowInTaskbar="False"` — FinalWorkingCode's
XAML has no such attribute (defaults to visible in taskbar) but nothing suggests this was a deliberate AIPowered
regression; keep it.

- [ ] **Step 3: Port control styles**

Content `Border`: drop `BorderBrush="{StaticResource BorderLightBrush}"` to match FinalWorkingCode's borderless
`Background="White" CornerRadius="8" BorderThickness="1"` (no visible border color set). Change "Download Selected"
button from `Style="{StaticResource InsertButtonStyle}"` to `Style="{StaticResource DynamicContentButton}"`. Add
the missing `<local:AppOverlay x:Name="AppOverlayControl" Grid.RowSpan="3" .../>` element at the bottom of the root
Grid (currently absent — `BaseWindow.DismissActiveToast`/`IsInteractionOverlayVisible`'s `FindName` lookup silently
no-ops without it).

- [ ] **Step 4: Preserve verbatim**

Keep the `ScrollViewer`'s `MinHeight="220"` (`CLAUDE.md`'s own note: "already uses a deliberate ScrollViewer
MinHeight=220 pattern") — FinalWorkingCode's `ScrollViewer` has no such floor; do not strip it. Keep
`TitleBar_MouseLeftButtonDown`, rewired onto the new header.

- [ ] **Step 5: Build verification**

Run: `msbuild GLSense.Addin.Core\GLSense.Addin.Core.csproj /t:Build /p:Configuration=Debug /p:SignAssembly=false
/v:normal /nologo`
Expected: `Build succeeded. 0 Error(s)`.

- [ ] **Step 6: Commit**

```bash
git add GLSense.Addin.Core/Views/AttachmentsDialog.xaml GLSense.Addin.Core/Views/AttachmentsDialog.xaml.cs
git commit -m "Port AttachmentsDialog header/chrome/sizing parity with FinalWorkingCode"
```

---

### Task 3: Port `WebView2PopupWindow` header/chrome/sizing

**Files:**
- Modify: `GLSense.Addin.Core\Views\WebView2PopupWindow.xaml`
- Modify: `GLSense.Addin.Core\Views\WebView2PopupWindow.xaml.cs`

- [ ] **Step 1: Replace the header**

Canonical header markup: `Kind="CircleInfoSolid"`, `Text="Sign-in"`. Root: `WindowStyle="None"`.

- [ ] **Step 2: Port sizing**

`SizeToContent="Manual" Width="500" Height="650" MinWidth="380" MinHeight="480"` already match FinalWorkingCode
exactly — no numeric changes needed. Add `Background="#F8F9FA"`. Remove `ExtendsContentIntoTitleBar="True"` and
`WindowStyle="SingleBorderWindow"`.

- [ ] **Step 3: Port control styles**

Keep the WebView2 content `Border`'s existing `Style="{StaticResource ContentBorderStyle}"` (already correct).
Follow the same header-row convention (32px-fixed-row → `Auto`-sized `HeaderBar` row, `RootGrid` structure) already
established by `GLWaitWindow`/`GLMessageWindow`/`GLSegmentValues` — do not adopt FinalWorkingCode's own
`Margin="6"`-on-outer-Grid convention if it conflicts with `RootGrid`'s existing per-row margin structure; keep the
row-margin approach the 3 pilots already use.

- [ ] **Step 4: Preserve verbatim**

Keep `TitleBar_MouseLeftButtonDown`, rewired onto the new header (FinalWorkingCode's own file wires no drag
mechanism at all for this specific window — AIPowered's dedicated handler is a strict improvement, keep it). Keep
the header comment explaining `Owner` is set directly by `WebView2NavigationResilience.OnNewWindowRequested`, not
`SetExcelOwner`/`ModalToExcel` — this is architecture-specific and must not be touched. All `CoreWebView2_*` event
handlers and disposal logic are unchanged.

- [ ] **Step 5: Build verification**

Run: `msbuild GLSense.Addin.Core\GLSense.Addin.Core.csproj /t:Build /p:Configuration=Debug /p:SignAssembly=false
/v:normal /nologo`
Expected: `Build succeeded. 0 Error(s)`.

- [ ] **Step 6: Commit**

```bash
git add GLSense.Addin.Core/Views/WebView2PopupWindow.xaml GLSense.Addin.Core/Views/WebView2PopupWindow.xaml.cs
git commit -m "Port WebView2PopupWindow header/chrome/sizing parity with FinalWorkingCode"
```

---

### Task 4: Port `GLExpandOptions` header/chrome/sizing

**Files:**
- Modify: `GLSense.Addin.Core\Views\GLExpandOptions.xaml`
- Modify: `GLSense.Addin.Core\Views\GLExpandOptions.xaml.cs`

- [ ] **Step 1: Replace the header**

Canonical header markup: `Kind="FilterSolid"`, `Text="GLSense Expand Segment Hierarchy"`. Root: `WindowStyle="None"`.

- [ ] **Step 2: Port sizing**

`SizeToContent="WidthAndHeight"` stays unchanged (this window correctly stays on `WidthAndHeight` in
FinalWorkingCode too — small fixed-content options dialog). Change `MinHeight="220"` to `MinHeight="240"`
(FinalWorkingCode's value). `MaxHeight="320" MinWidth="420" MaxWidth="460"` already match — leave unchanged. Add
`Background="#F8F9FA"`. Remove `ExtendsContentIntoTitleBar="True"` and `WindowStyle="SingleBorderWindow"`.

- [ ] **Step 3: Port control styles**

Change both content-row `Border`s from inline `Background="{StaticResource WhiteBrush}" CornerRadius="6"
BorderThickness="1" BorderBrush="{StaticResource BorderLightBrush}" Padding="10,8"` to `Style="{StaticResource
ContentBorderStyle}"`. Change `Label Content="Expand Level :" Foreground="{StaticResource PrimaryBrush}"
FontWeight="SemiBold" MinWidth="110"` to `TextBlock Text="Expand Level : " Style="{StaticResource SectionHeader}"`
(same treatment for the "Fill Direction" label). Change `rbByColumns`'s `Margin="20,0,0,0"` to `Margin="35,0,0,0"`.
Footer `Border`: change to `Background="White" CornerRadius="6" BorderThickness="1" BorderBrush="#E9ECEF"
Margin="0,8,0,0"` (add the missing top margin).

- [ ] **Step 4: Preserve verbatim**

Keep `TitleBar_MouseLeftButtonDown`, rewired onto the new header. `btnExpand`/`btnClose` already use
`InsertButtonStyle`/`CloseButtonStyle` correctly — no change. `BtnExpand_Click`'s
`rbExpand1Level.IsChecked`/`rbByColumns.IsChecked` reads and `SegmentDiscoverer.SegmentAction(actionType,
byColumns)` call are unchanged.

- [ ] **Step 5: Build verification**

Run: `msbuild GLSense.Addin.Core\GLSense.Addin.Core.csproj /t:Build /p:Configuration=Debug /p:SignAssembly=false
/v:normal /nologo`
Expected: `Build succeeded. 0 Error(s)`.

- [ ] **Step 6: Commit**

```bash
git add GLSense.Addin.Core/Views/GLExpandOptions.xaml GLSense.Addin.Core/Views/GLExpandOptions.xaml.cs
git commit -m "Port GLExpandOptions header/chrome/sizing parity with FinalWorkingCode"
```

---

## Phase 2 (Batch 2): Period/date family

Shared context for all 5 tasks below: theme keys `CompactBorderStyle`/`HeaderLabel`/`CompactRefEditControl` replace
each window's inline `Border`/`Label` styling (`ArrowBrush="#2E86AB" ArrowWidth="10" ArrowHeight="8"` added to every
`SuggestAppendComboBox`). Adopt FinalWorkingCode's `Margin="10"` root-Grid + `Auto/*/Auto` column structure and
`Auto`-height header row (replacing the old hardcoded 32px header row) per the Global Constraints header pattern.

### Task 5: Port `GLGetPeriod` header/chrome/sizing

**Files:**
- Modify: `GLSense.Addin.Core\Views\GLGetPeriod.xaml`
- Modify: `GLSense.Addin.Core\Views\GLGetPeriod.xaml.cs`

- [ ] **Step 1: Replace the header**

Canonical header markup: `Kind="CalendarDaysSolid"`, `Text="Get Period"` (a literal, hardcoded title — not bound
to `WindowCaption`, matching FinalWorkingCode's own-header-per-window convention). Root: `WindowStyle="None"`.

- [ ] **Step 2: Port sizing**

Change `SizeToContent="WidthAndHeight" MinWidth="600" MaxWidth="650" MinHeight="345" MaxHeight="450"` to
`SizeToContent="Manual" MinWidth="600" MaxWidth="700" MinHeight="360" MaxHeight="450"`. Add
`Background="#F8F9FA"`. Remove `ExtendsContentIntoTitleBar="True"` and `WindowStyle="SingleBorderWindow"`.

- [ ] **Step 3: Port control styles**

Field-row `Border`s: inline styling → `Style="{StaticResource CompactBorderStyle}"`. `Label`s: inline
`Foreground="{StaticResource PrimaryBrush}" FontWeight="SemiBold"` → `Style="{StaticResource HeaderLabel}"`.
`SuggestAppendComboBox` instances: add `ArrowBrush="#2E86AB" ArrowWidth="10" ArrowHeight="8"`.
`ExcelRefEditControl` instances: add `Style="{StaticResource CompactRefEditControl}"`.

- [ ] **Step 4: Preserve verbatim**

Keep `TitleBar_MouseLeftButtonDown`, rewired onto the new header. `BtnSubmit_Click`/`BtnClose_Click`/`Window_Loaded`
unchanged — already using `ServiceLocator.ExcelApp`/`ServiceLocator.Logger` correctly.

- [ ] **Step 5: Build verification**

Run: `msbuild GLSense.Addin.Core\GLSense.Addin.Core.csproj /t:Build /p:Configuration=Debug /p:SignAssembly=false
/v:normal /nologo`
Expected: `Build succeeded. 0 Error(s)`.

- [ ] **Step 6: Commit**

```bash
git add GLSense.Addin.Core/Views/GLGetPeriod.xaml GLSense.Addin.Core/Views/GLGetPeriod.xaml.cs
git commit -m "Port GLGetPeriod header/chrome/sizing parity with FinalWorkingCode"
```

---

### Task 6: Port `GLGetPeriodByDate` header/chrome/sizing

**Files:**
- Modify: `GLSense.Addin.Core\Views\GLGetPeriodByDate.xaml`
- Modify: `GLSense.Addin.Core\Views\GLGetPeriodByDate.xaml.cs`

- [ ] **Step 1: Replace the header**

Canonical header markup: `Kind="CalendarDaysSolid"`, `Text="Get Period by Date"`. Root: `WindowStyle="None"`.

- [ ] **Step 2: Port sizing**

Change `SizeToContent="WidthAndHeight" MinWidth="620" MaxWidth="660" MinHeight="345" MaxHeight="480"` to
`SizeToContent="Manual" MinWidth="620" MaxWidth="700" MinHeight="360" MaxHeight="480"`. Add
`Background="#F8F9FA"`. Remove `ExtendsContentIntoTitleBar="True"` and `WindowStyle="SingleBorderWindow"`.

- [ ] **Step 3: Port control styles — DO NOT touch the DatePicker**

Same `CompactBorderStyle`/`HeaderLabel`/`ArrowBrush`/`CompactRefEditControl` porting as Task 5, for every field
EXCEPT `dtpDate`. **`dtpDate` (the `DatePicker`) must be left exactly as AIPowered currently has it** —
`Style="{StaticResource ModernDatePicker}"` with `BorderThickness="0" Background="Transparent"` on its inline
`DatePickerTextBox` template. This is an AIPowered-only fix (`CLAUDE.md` §8.5 + §9.1, both scoped AIPowered-only)
that FinalWorkingCode's own file still lacks (it has no `Style` on the `DatePicker` at all, plus a
`BorderThickness="1"` double-border bug AIPowered already fixed). Copying FinalWorkingCode's `dtpDate` markup here
would be a regression, not a sync.

- [ ] **Step 4: Preserve verbatim**

Keep `TitleBar_MouseLeftButtonDown`, rewired onto the new header. No other code-behind changes.

- [ ] **Step 5: Build verification**

Run: `msbuild GLSense.Addin.Core\GLSense.Addin.Core.csproj /t:Build /p:Configuration=Debug /p:SignAssembly=false
/v:normal /nologo`
Expected: `Build succeeded. 0 Error(s)`.

- [ ] **Step 6: Commit**

```bash
git add GLSense.Addin.Core/Views/GLGetPeriodByDate.xaml GLSense.Addin.Core/Views/GLGetPeriodByDate.xaml.cs
git commit -m "Port GLGetPeriodByDate header/chrome/sizing parity with FinalWorkingCode"
```

---

### Task 7: Port `GLGetPeriodByYear` header/chrome/sizing

**Files:**
- Modify: `GLSense.Addin.Core\Views\GLGetPeriodByYear.xaml`
- Modify: `GLSense.Addin.Core\Views\GLGetPeriodByYear.xaml.cs`

- [ ] **Step 1: Replace the header**

Canonical header markup: `Kind="CalendarDaysSolid"`, `Text="Get Period by Year"`. Root: `WindowStyle="None"`.

- [ ] **Step 2: Port sizing**

Change `SizeToContent="WidthAndHeight" MinWidth="600" MaxWidth="650" MinHeight="345" MaxHeight="480"` to
`SizeToContent="Manual" MinWidth="600" MaxWidth="700" MinHeight="360" MaxHeight="480"`. Add
`Background="#F8F9FA"`. Remove `ExtendsContentIntoTitleBar="True"` and `WindowStyle="SingleBorderWindow"`.

- [ ] **Step 3: Port control styles**

Same `CompactBorderStyle`/`HeaderLabel`/`ArrowBrush`/`CompactRefEditControl` porting as Task 5 across
`cmbLedgers`/`cmbPeriodYear`/`cmbPeriodNum`/`txtResult`/`CellReference` (no DatePicker in this window). Leave the
`LabelMinWidth="110"` resource value unchanged — it's a pre-existing AIPowered-only value unrelated to this
chrome/sizing port; do not change it to FinalWorkingCode's `100` (out of scope, would be a layout-metric change
this task isn't chartered to make).

- [ ] **Step 4: Preserve verbatim**

Keep `TitleBar_MouseLeftButtonDown`, rewired onto the new header.

- [ ] **Step 5: Build verification**

Run: `msbuild GLSense.Addin.Core\GLSense.Addin.Core.csproj /t:Build /p:Configuration=Debug /p:SignAssembly=false
/v:normal /nologo`
Expected: `Build succeeded. 0 Error(s)`.

- [ ] **Step 6: Commit**

```bash
git add GLSense.Addin.Core/Views/GLGetPeriodByYear.xaml GLSense.Addin.Core/Views/GLGetPeriodByYear.xaml.cs
git commit -m "Port GLGetPeriodByYear header/chrome/sizing parity with FinalWorkingCode"
```

---

### Task 8: Port `GLGetPeriodDetails` header/chrome/sizing

**Files:**
- Modify: `GLSense.Addin.Core\Views\GLGetPeriodDetails.xaml`
- Modify: `GLSense.Addin.Core\Views\GLGetPeriodDetails.xaml.cs`

- [ ] **Step 1: Replace the header, preserving the dynamic-title binding**

Canonical header markup, but keep `x:Name="TxtGLFormulaType"` on the header `TextBlock` (default `Text="Get
Period"`) instead of a bare literal — this window is shared across multiple formula names and its constructor
(`GLGetPeriodDetails(string FormulaName)`) retitles this exact named element at runtime, same pattern
FinalWorkingCode's own file uses (`x:Name="TxtGLFormulaType" Text="Get Period" Style="{StaticResource
TitleHeaderTextBlock}"`). `Kind="CalendarDaysSolid"`. Root: `WindowStyle="None"`.

- [ ] **Step 2: Port sizing**

Change `SizeToContent="WidthAndHeight" MinWidth="600" MaxWidth="650" MinHeight="300" MaxHeight="450"` to
`SizeToContent="Manual" MinWidth="600" MaxWidth="700" MinHeight="320" MaxHeight="450"`. Add
`Background="#F8F9FA"`. Remove `ExtendsContentIntoTitleBar="True"` and `WindowStyle="SingleBorderWindow"`.

- [ ] **Step 3: Port control styles**

Same `CompactBorderStyle`/`HeaderLabel`/`ArrowBrush`/`CompactRefEditControl` porting as Task 5. Preserve
`x:Name="LblGLFormulaType"` on its dynamically-retitled `Label`/style-ported equivalent — the constructor sets
this at runtime too.

- [ ] **Step 4: Preserve verbatim, verify constructor signature**

Keep `TitleBar_MouseLeftButtonDown`, rewired onto the new header. Confirm `GLGetPeriodDetails(string FormulaName)`
constructor is unchanged and still sets both `TxtGLFormulaType.Text`/`LblGLFormulaType.Content` (or equivalent) —
read the constructor body before finishing this task to confirm the exact property names it sets, since the header
port must not rename either element out from under it.

- [ ] **Step 5: Build verification**

Run: `msbuild GLSense.Addin.Core\GLSense.Addin.Core.csproj /t:Build /p:Configuration=Debug /p:SignAssembly=false
/v:normal /nologo`
Expected: `Build succeeded. 0 Error(s)`.

- [ ] **Step 6: Commit**

```bash
git add GLSense.Addin.Core/Views/GLGetPeriodDetails.xaml GLSense.Addin.Core/Views/GLGetPeriodDetails.xaml.cs
git commit -m "Port GLGetPeriodDetails header/chrome/sizing parity with FinalWorkingCode"
```

---

### Task 9: Port `GLGetPeriodStartEnd` header/chrome/sizing

**Files:**
- Modify: `GLSense.Addin.Core\Views\GLGetPeriodStartEnd.xaml`
- Modify: `GLSense.Addin.Core\Views\GLGetPeriodStartEnd.xaml.cs`

- [ ] **Step 1: Replace the header, preserving the dynamic-title binding**

Same shared-window dynamic-title pattern as Task 8 (`x:Name="TxtGLFormulaType"`/`x:Name="LblGLFormulaType"`,
constructor is `GLGetPeriodStartEnd(string FormulaName)`). Canonical header markup, `Kind="CalendarDaysSolid"`.
Root: `WindowStyle="None"`.

- [ ] **Step 2: Port sizing**

Change `SizeToContent="WidthAndHeight" MinWidth="600" MaxWidth="650" MinHeight="300" MaxHeight="450"` to
`SizeToContent="Manual" MinWidth="600" MaxWidth="700" MinHeight="320" MaxHeight="450"`. Add
`Background="#F8F9FA"`. Remove `ExtendsContentIntoTitleBar="True"` and `WindowStyle="SingleBorderWindow"`.

- [ ] **Step 3: Port control styles**

Same `CompactBorderStyle`/`HeaderLabel`/`ArrowBrush`/`CompactRefEditControl` porting as Task 5. Leave
`ChkIncludeAdjacents`'s `Style="{StaticResource ModernCheckBox}"` unchanged — already correct on both sides.

- [ ] **Step 4: Preserve verbatim, verify constructor signature**

Keep `TitleBar_MouseLeftButtonDown`, rewired onto the new header. Confirm `GLGetPeriodStartEnd(string
FormulaName)` still sets both named header elements correctly after the port, same check as Task 8.

- [ ] **Step 5: Build verification**

Run: `msbuild GLSense.Addin.Core\GLSense.Addin.Core.csproj /t:Build /p:Configuration=Debug /p:SignAssembly=false
/v:normal /nologo`
Expected: `Build succeeded. 0 Error(s)`.

- [ ] **Step 6: Commit**

```bash
git add GLSense.Addin.Core/Views/GLGetPeriodStartEnd.xaml GLSense.Addin.Core/Views/GLGetPeriodStartEnd.xaml.cs
git commit -m "Port GLGetPeriodStartEnd header/chrome/sizing parity with FinalWorkingCode"
```

---

## Phase 3 (Batch 3): Other static forms

### Task 10: Port `GLDailyRates` header/chrome/sizing

**Files:**
- Modify: `GLSense.Addin.Core\Views\GLDailyRates.xaml`
- Modify: `GLSense.Addin.Core\Views\GLDailyRates.xaml.cs`

- [ ] **Step 1: Replace the header**

Canonical header markup: `Kind="MoneyCheckDollarSolid"`, `Text="GLSense Daily Rates"` (literal, not bound to
`WindowCaption`; `TxtFuncName`'s `x:Name` can stay if referenced elsewhere — check code-behind before removing it).
Root: `WindowStyle="None"`.

- [ ] **Step 2: Port sizing**

Change `SizeToContent="WidthAndHeight" MinWidth="640" MaxWidth="660" MinHeight="350" MaxHeight="450"` to
`SizeToContent="Manual" MinWidth="640" MaxWidth="700" MinHeight="400" MaxHeight="450"`. Add
`Background="#F8F9FA"`. Remove `ExtendsContentIntoTitleBar="True"` and `WindowStyle="SingleBorderWindow"`.

- [ ] **Step 3: Port control styles**

Content rows: inline styling → `Style="{StaticResource CompactBorderStyle}"` (Border) / `Style="{StaticResource
HeaderLabel}"` (label). Add `Grid.IsSharedSizeScope="True"` to the outer content `Grid`, and
`SharedSizeGroup="LabelColumn"`/`"EditorColumn"`/`"RefColumn"` to each row's matching `ColumnDefinition`s (ported
from FinalWorkingCode — aligns every row's columns without needing per-row hardcoded widths). `SuggestAppendComboBox`
instances: add `ArrowBrush="#2E86AB" ArrowWidth="10" ArrowHeight="8"`. `ExcelRefEditControl` instances: add
`Style="{StaticResource CompactRefEditControl}"`. Do NOT port FinalWorkingCode's unused, dead
`DailyRatesDatePickerTextBoxStyle` resource (defined but never applied in FinalWorkingCode's own file).

- [ ] **Step 4: DO NOT touch the DatePicker**

Same rule as Task 6: `dtpDate` must keep AIPowered's existing `Style="{StaticResource ModernDatePicker}"` +
`BorderThickness="0" Background="Transparent"` inline-template fix (`CLAUDE.md` §8.5/§9.1/§9.1) — FinalWorkingCode's
own file still has the un-fixed double-border/small-icon markup. Do not copy it.

- [ ] **Step 5: Preserve verbatim**

Keep `TitleBar_MouseLeftButtonDown`, rewired onto the new header (recommended over introducing
`EnhancedDragDropHelper`, which doesn't exist in this codebase). No `ViewModel`/business-logic changes.

- [ ] **Step 6: Build verification**

Run: `msbuild GLSense.Addin.Core\GLSense.Addin.Core.csproj /t:Build /p:Configuration=Debug /p:SignAssembly=false
/v:normal /nologo`
Expected: `Build succeeded. 0 Error(s)`.

- [ ] **Step 7: Commit**

```bash
git add GLSense.Addin.Core/Views/GLDailyRates.xaml GLSense.Addin.Core/Views/GLDailyRates.xaml.cs
git commit -m "Port GLDailyRates header/chrome/sizing parity with FinalWorkingCode"
```

---

### Task 11: Port `GLSegmentDiscovery` header/chrome/sizing

**Files:**
- Modify: `GLSense.Addin.Core\Views\GLSegmentDiscovery.xaml`
- Modify: `GLSense.Addin.Core\Views\GLSegmentDiscovery.xaml.cs`

- [ ] **Step 1: Replace the header**

Canonical header markup: `Kind="CompassSolid"` (FinalWorkingCode's icon — note this differs from AIPowered's
current `IconSymbol="FilterSolid"`), `Text="GLSense Segment Values Discover"`. Root: `WindowStyle="None"`.

- [ ] **Step 2: Port sizing**

Change `SizeToContent="WidthAndHeight" MinWidth="660" MaxWidth="700" MinHeight="345" MaxHeight="520"` to
`SizeToContent="Manual" MinHeight="440" MaxHeight="490" MinWidth="680" MaxWidth="780"`. Add
`Background="#F8F9FA"`. Remove `ExtendsContentIntoTitleBar="True"` and `WindowStyle="SingleBorderWindow"`. Do
**not** add `Loaded="Window_Loaded"` to the root element — AIPowered wires this in the constructor already
(`Loaded += Window_Loaded;`), and FinalWorkingCode's XAML-level `Loaded="Window_Loaded"` would double-subscribe
if both are present.

- [ ] **Step 3: Port control styles**

Content rows: inline styling → `Style="{StaticResource ContentBorderStyle}"` (Border) / `Style="{StaticResource
SectionHeader}"` `TextBlock` (replacing the plain `Label`). `txtSegmentValue`/`txtRangeSelected`: change
`Style="{StaticResource ReadOnlyTextBox}"` to `Style="{StaticResource ModernTextBox}" IsReadOnly="True"`. Port
FinalWorkingCode's yellow (`#FBB41A`) tooltip callout with `LightbulbRegular` icon for `txtSegmentValue`'s tooltip,
replacing the current plain-`TextBlock` tooltip content. Do not port FinalWorkingCode's 2-column root `Grid`
split (`628*`/`69*`) — nothing meaningfully occupies the second column in FinalWorkingCode's own file either
(everything uses `Grid.ColumnSpan="2"`); keep AIPowered's simpler single-column content structure.

- [ ] **Step 4: Preserve verbatim**

Keep `TitleBar_MouseLeftButtonDown`, rewired onto the new header. **Do not touch any code-behind logic** — the
async `WriteValuesToExcelAsync`/`YieldEveryBatch` pattern (`CLAUDE.md` §27.2/§27.2b, a deliberate fix for a frozen
busy-spinner bug) and `ServiceLocator.RibbonController?.GetControlPressed("RibAsFormula")` (required due to the
AppDomain boundary) must survive exactly as they are — do not replace with FinalWorkingCode's synchronous
`PumpDispatcherFrame` pattern or its `AddinModule.CurrentInstance.RibAsFormula.Pressed` direct-ribbon-access. This
task is XAML-only.

- [ ] **Step 5: Build verification**

Run: `msbuild GLSense.Addin.Core\GLSense.Addin.Core.csproj /t:Build /p:Configuration=Debug /p:SignAssembly=false
/v:normal /nologo`
Expected: `Build succeeded. 0 Error(s)`.

- [ ] **Step 6: Commit**

```bash
git add GLSense.Addin.Core/Views/GLSegmentDiscovery.xaml GLSense.Addin.Core/Views/GLSegmentDiscovery.xaml.cs
git commit -m "Port GLSegmentDiscovery header/chrome/sizing parity with FinalWorkingCode"
```

---

### Task 12: Port `GLSegmentFunctions` header/chrome/sizing

**Files:**
- Modify: `GLSense.Addin.Core\Views\GLSegmentFunctions.xaml`
- Modify: `GLSense.Addin.Core\Views\GLSegmentFunctions.xaml.cs`

- [ ] **Step 1: Replace the header, preserving the dynamic-title binding**

Canonical header markup with `x:Name="TxtFuncName"` kept on the header `TextBlock` (retitled per function type by
the constructor's `switch`). `Kind="SlidersSolid"` (FinalWorkingCode's icon, differs from AIPowered's current
`IconSymbol="FilterSolid"`). Root: `WindowStyle="None"`.

- [ ] **Step 2: Port sizing**

Change `SizeToContent="WidthAndHeight" MinWidth="640" MaxWidth="660" MinHeight="350" MaxHeight="520"` to
`SizeToContent="Height" MinHeight="350" MaxHeight="490" MinWidth="640" MaxWidth="700"` — matching FinalWorkingCode's
current live file exactly (supersedes the older `CLAUDE.md` §8.3 prose description of `MinHeight="375"`, per this
codebase's own established rule to prefer the live file over historical notes, set by Phase 1 Task 7). Add
`Background="#F8F9FA"`. Remove `ExtendsContentIntoTitleBar="True"` and `WindowStyle="SingleBorderWindow"`.

- [ ] **Step 3: Port control styles — preserve the 4 conditional-visibility section names exactly**

Content rows: inline styling → `Style="{StaticResource CompactBorderStyle}"` (Border) / `Style="{StaticResource
HeaderLabel}"` (label), matching FinalWorkingCode's tighter `Padding="12,6"`. **The 4 `x:Name`'d elements
`ParentChildSection`, `AttributesSection`, `ResultSection`, `IncludeValuesSection` MUST keep those exact names on
the exact same elements** — the constructor's `switch(funcName)` (identical in both codebases for all 7 cases:
`ENABLEDFLAG`, `SUMMARYFLAG`, `DESCRIPTION`, `NEXTSEGMENT`, `PREVIOUSSEGMENT`, `DFF`, `ACCOUNTTYPE`) toggles their
`Visibility` by these names; renaming or restructuring any of the 4 breaks that switch silently (no compile error —
XAML `x:Name` lookups fail at runtime).

- [ ] **Step 4: Preserve verbatim**

Keep `TitleBar_MouseLeftButtonDown`, rewired onto the new header. No code-behind logic changes — confirmed
functionally identical to FinalWorkingCode's `.xaml.cs` already (only logging plumbing differs). Both codebases
already have the §22.2 fix (no redundant generic `ShowWarning` in `FormatAndWriteCell`) — nothing to preserve there
beyond leaving it alone.

- [ ] **Step 5: Build verification**

Run: `msbuild GLSense.Addin.Core\GLSense.Addin.Core.csproj /t:Build /p:Configuration=Debug /p:SignAssembly=false
/v:normal /nologo`
Expected: `Build succeeded. 0 Error(s)`.

- [ ] **Step 6: Commit**

```bash
git add GLSense.Addin.Core/Views/GLSegmentFunctions.xaml GLSense.Addin.Core/Views/GLSegmentFunctions.xaml.cs
git commit -m "Port GLSegmentFunctions header/chrome/sizing parity with FinalWorkingCode"
```

---

### Task 13: Port `GLLoginDetails` header/chrome/sizing

**Files:**
- Modify: `GLSense.Addin.Core\Views\GLLoginDetails.xaml`
- Modify: `GLSense.Addin.Core\Views\GLLoginDetails.xaml.cs`

- [ ] **Step 1: Replace the header**

Canonical header markup: `Kind="CircleInfoSolid"` (same icon AIPowered already uses — no icon change here),
`Text="GLSense Login Details"`. Name the
`Button` `x:Name="CmdClose"` and the two `TextBox`es `x:Name="TxtUser"`/`x:Name="TxtURL"` to match FinalWorkingCode
(purely additive naming — neither side's code-behind currently references these controls by name, confirmed
risk-free). Root: `WindowStyle="None"`.

- [ ] **Step 2: Port sizing — full sync per Global Constraints policy**

Change `SizeToContent="Width" MinHeight="100" MaxHeight="230" MinWidth="460" MaxWidth="480"` to
`SizeToContent="Manual" MinHeight="260" MaxHeight="400" MinWidth="460" MaxWidth="580"`. This overrides the value
`CLAUDE.md` §1.4 calls "the user's own fix" — per this plan's Global Constraints, that note predates the Phase 1
`BaseWindow` rewrite and FinalWorkingCode's richer content (Step 3) is mechanically taller regardless, so the
larger bounds are necessary, not just a sync-for-its-own-sake change. Add `Background="#F8F9FA"`. Remove
`ExtendsContentIntoTitleBar="True"` and `WindowStyle="SingleBorderWindow"`.

- [ ] **Step 3: Port control styles — richer content structure**

Replace the bare `StackPanel Margin="10,10,10,4"` with two independent bordered rows with FinalWorkingCode's single
`Border Background="White" CornerRadius="6" BorderThickness="1" BorderBrush="#E9ECEF" Padding="20"` wrapping a
2-row `Grid`. Each row becomes an icon-badge treatment: `Border Background="#F8F9FA" CornerRadius="4"` containing a
`PackIconFontAwesome Kind="UserRegular"` (User row) / `Kind="LinkSolid"` (URL row) + `TextBlock Style="{StaticResource
HeaderTextBlock}"`, replacing the plain `Label Content="User :" Foreground="{StaticResource PrimaryBrush}"
FontWeight="SemiBold"`. Port the richer tooltip content showing the actual bound value in a `Background="#F5F5F5"`
code-style block (`TextBlock Text="{Binding LoginUserName, ...}" FontFamily="Segoe UI" Background="#F5F5F5"`),
replacing the current plain static-text tooltip.

- [ ] **Step 4: Preserve verbatim**

Keep `TitleBar_MouseLeftButtonDown`, rewired onto the new header. `BtnClose_Click` calling `Close()` is unchanged —
no other code-behind logic exists in this file beyond a constructor debug log.

- [ ] **Step 5: Build verification**

Run: `msbuild GLSense.Addin.Core\GLSense.Addin.Core.csproj /t:Build /p:Configuration=Debug /p:SignAssembly=false
/v:normal /nologo`
Expected: `Build succeeded. 0 Error(s)`.

- [ ] **Step 6: Commit**

```bash
git add GLSense.Addin.Core/Views/GLLoginDetails.xaml GLSense.Addin.Core/Views/GLLoginDetails.xaml.cs
git commit -m "Port GLLoginDetails header/chrome/sizing parity with FinalWorkingCode"
```

---

## Phase 4 (Batch 4): DataGrid/ViewModel-heavy windows

### Task 14: Port `GLLOVs` header/chrome/sizing

**Files:**
- Modify: `GLSense.Addin.Core\Views\GLLOVs.xaml`
- Modify: `GLSense.Addin.Core\Views\GLLOVs.xaml.cs`

- [ ] **Step 1: Read both sides in full before editing**

This window has real ViewModel-driven data loading (`Loaded="Window_Loaded"` → `vm.LoadDataAsync`) and several
AIPowered-only fixes (`CLAUDE.md` §6.1, §6.2, §8.2, §8.5, §10.1, §10.2, §30) — read `FinalWorkingCode\GLSense\Views\
GLLOVs.xaml`/`.xaml.cs` and AIPowered's current copies in full before editing.

- [ ] **Step 2: Replace the header — including the Ledger combo moving into it**

FinalWorkingCode's header structurally differs from every other window in this plan: the Ledger combo lives
*inside* the header itself, not in the body. Canonical header markup with `Kind="ListCheckSolid"`,
`Text="GLSense LOVs"`, PLUS a `StackPanel Grid.Column="1" HorizontalAlignment="Right"` containing `Label
Content="Ledger:"` + `cmbLedgers` inside the same header `Border` (adjust the header's `Grid.ColumnDefinitions` to
`Auto/Auto/Auto` or `Auto/*/Auto` as needed to fit both the title and the right-aligned Ledger combo before the
close button). Root: `WindowStyle="None"`.

- [ ] **Step 3: Port sizing**

Change `SizeToContent="WidthAndHeight" MinWidth="550" MaxWidth="750" MinHeight="650" MaxHeight="800"` to
`SizeToContent="Manual" MinWidth="550" MaxWidth="750" MinHeight="750"` (no `MaxHeight` — FinalWorkingCode has none).
Add `Background="#F8F9FA"`. Remove `ExtendsContentIntoTitleBar="True"` and `WindowStyle="SingleBorderWindow"`.

- [ ] **Step 4: Port DataGrid column-header template and tooltips — keep AIPowered's live-count binding**

Port FinalWorkingCode's sort-arrow `ColumnHeaderStyle` `ControlTemplate` (currently missing from AIPowered's
version entirely). Port the amber `#FBB41A`/`LightbulbRegular`-icon tooltip callout styling for the "Available
LOVs"/"Items Count" columns, but **keep AIPowered's existing `MultiBinding` for the Items Count tooltip**
(`"{0} item(s) available for '{1}'"`, showing the live count — `CLAUDE.md` §10.2) rather than reverting to
FinalWorkingCode's static "Number of values in this list" text. Change the Comments field's `Label
Style="{StaticResource HeaderLabel}"` — wait, this window's Comments label should move from `TextBlock
Style="{StaticResource HeaderTextBlock}"` to `Label Style="{StaticResource HeaderLabel}"` to match FinalWorkingCode
(confirm `HeaderLabel` exists via grep before applying — Global Constraints' theme-key list already covers this).

- [ ] **Step 5: Remove the now-unneeded DataGridColumnFillHelper call**

Remove `DataGridColumnFillHelper.EnableFillColumn(dgLovs, dgLovs.Columns[0])` from the constructor — per `CLAUDE.md`
§8.1's finding (established on `GLSegmentValues`), this helper is unneeded once the window is on native `Width="*"`
columns under `SizeToContent="Manual"` (both sides already use `2*`/`1*` column widths; FinalWorkingCode's own
constructor has no such call).

- [ ] **Step 6: Preserve verbatim**

Keep the `WatermarkTextBox` style's `MaxWidth="480"` on its `WatermarkText` `TextBlock` (`CLAUDE.md` §10.1) — do
not remove even though the sizing-mode change reduces the original bug's exposure; it's cheap, harmless insurance.
Keep `MinHeight="250" MaxHeight="400"` on the LOVs DataGrid's `Border` (`CLAUDE.md` §6.1/§8.2) — FinalWorkingCode
has no cap here at all. Keep `Loaded="Window_Loaded"` wired on the root element exactly as it is — do NOT change
this window's load-triggering mechanism to FinalWorkingCode's `PrepareAsync()`-called-before-`ShowDialogWithOwner()`
pattern; that is a genuine architectural divergence, not a chrome/sizing concern, and is out of scope for this
task. Keep `TitleBar_MouseLeftButtonDown`, rewired onto the new header.

- [ ] **Step 7: Build verification**

Run: `msbuild GLSense.Addin.Core\GLSense.Addin.Core.csproj /t:Build /p:Configuration=Debug /p:SignAssembly=false
/v:normal /nologo`
Expected: `Build succeeded. 0 Error(s)`.

- [ ] **Step 8: Commit**

```bash
git add GLSense.Addin.Core/Views/GLLOVs.xaml GLSense.Addin.Core/Views/GLLOVs.xaml.cs
git commit -m "Port GLLOVs header/chrome/sizing parity with FinalWorkingCode"
```

---

### Task 15: Port `GLRollerGroups` header/chrome/sizing

**Files:**
- Modify: `GLSense.Addin.Core\Views\GLRollerGroups.xaml`
- Modify: `GLSense.Addin.Core\Views\GLRollerGroups.xaml.cs`

- [ ] **Step 1: Replace the header**

Canonical header markup: `Kind="LayerGroupSolid"`, `Text="Segment Rollup Groups"`. Root: `WindowStyle="None"`.

- [ ] **Step 2: Port sizing**

Change `SizeToContent="WidthAndHeight" MinWidth="800" MaxWidth="940" MinHeight="740" MaxHeight="890"` to
`SizeToContent="Manual" MinWidth="800" MaxWidth="840" MinHeight="790"` (no `MaxHeight`). Add
`Background="#F8F9FA"`. Remove `ExtendsContentIntoTitleBar="True"` and `WindowStyle="SingleBorderWindow"`.

- [ ] **Step 3: Port DataGrid styles, tooltips, and footer control styling**

Port FinalWorkingCode's named `TitleRowCellStyle`/`TitleRowElementStyle` for the left DataGrid (replacing the
current inline/unnamed triggers), plus the amber-callout tooltips on Roller-Group header rows ("Examples: Cash
(header) / Petty Cash, Bank Accounts (items)") and the Description column's `{Binding Description}` tooltip. Port
per-cell tooltips (`{Binding Value1}`/`{Binding Segment}`) on the right DataGrid's `Value`/`Segment` columns. Add
`FontWeight="SemiBold" Foreground="#2E86AB"` to `chkMultipleRows`/`rbOverwrite`/`rbInsert`/`rbByRows`/`rbByColumns`,
plus their amber-callout tooltip content. Restructure the search row to match FinalWorkingCode's single `Auto`-width
`StackPanel` (icon + `cmbSearchType` + `txtSearch MinWidth="280"`) with `ChartPieSolid`/`SearchenginBrands` icons
before each label, replacing the current 3-column stretch layout. Change `excelRefEdit`'s width `180`→`175`; add
`btnOK Width="90"`.

- [ ] **Step 4: Remove the now-unneeded DataGridColumnFillHelper call, if present**

Grep this file's constructor for any `DataGridColumnFillHelper` call; remove it per the same `CLAUDE.md` §8.1
reasoning as Task 14, since this window is moving to `SizeToContent="Manual"`.

- [ ] **Step 5: Preserve verbatim**

Keep `MaxHeight="450"` on the Dual DataGrids `Border` (`CLAUDE.md` §6.1/§24.1) — FinalWorkingCode has no cap here;
this matches the precedent already kept on `GLSegmentValues`'s own Dual DataGrids Border through its Phase 1 port.
The `ScrollViewer` wrap around the Segment/Search fields (`CLAUDE.md` §6.1) needs no separate action — confirmed
FinalWorkingCode's own current file already has this exact wrap, so a faithful port satisfies it automatically.
**Keep the `vm.SegmentPickedIndex = AppState.Instance.SegmentPickedIndex;` line in `Window_Loaded`** (`CLAUDE.md`
§28) — confirmed absent from FinalWorkingCode's current file (an AIPowered-only fix not yet ported the other
direction) — this task must not lose it. Keep `TitleBar_MouseLeftButtonDown`, rewired onto the new header.

- [ ] **Step 6: Build verification**

Run: `msbuild GLSense.Addin.Core\GLSense.Addin.Core.csproj /t:Build /p:Configuration=Debug /p:SignAssembly=false
/v:normal /nologo`
Expected: `Build succeeded. 0 Error(s)`.

- [ ] **Step 7: Commit**

```bash
git add GLSense.Addin.Core/Views/GLRollerGroups.xaml GLSense.Addin.Core/Views/GLRollerGroups.xaml.cs
git commit -m "Port GLRollerGroups header/chrome/sizing parity with FinalWorkingCode"
```

---

### Task 16: Port `GLJobsMonitor` header/chrome/sizing

**Files:**
- Modify: `GLSense.Addin.Core\Views\GLJobsMonitor.xaml`
- Modify: `GLSense.Addin.Core\Views\GLJobsMonitor.xaml.cs`

- [ ] **Step 1: Replace the header**

Canonical header markup: `Kind="SwatchbookSolid"`, `Text="Processed Jobs"` (adopt FinalWorkingCode's exact, shorter
caption text per the Global Constraints full-sync policy — dropping AIPowered's current "GLSense Processed Jobs"
prefix). Root: `WindowStyle="None"`.

- [ ] **Step 2: Port sizing**

Change `SizeToContent="WidthAndHeight" MinWidth="800" MaxWidth="900" MinHeight="700" MaxHeight="850"` to
`SizeToContent="Manual" MinWidth="800" MaxWidth="1000" MinHeight="700"` (no `MaxHeight`). Add
`Background="#F8F9FA"`. Remove `ExtendsContentIntoTitleBar="True"` and `WindowStyle="SingleBorderWindow"`.

- [ ] **Step 3: Port tooltip content and search icon**

Port FinalWorkingCode's amber-callout tooltips on every column (Process ID, Name, Type, Phase, Status, Date,
Download) and every footer button (Delete/Delete All/Download Logs/Download Output(s)/Refresh), plus the
`SearchenginBrands` icon before the "Search:" label. The sort-arrow `ColumnHeaderStyle` `ControlTemplate` already
matches FinalWorkingCode — no change needed there.

- [ ] **Step 4: Remove now-superseded guards**

Remove `DataGridColumnFillHelper.EnableFillColumn(dgJobs, dgJobs.Columns[1])` from the constructor (same `CLAUDE.md`
§8.1 reasoning — FinalWorkingCode's `Name` column is already native `Width="3*"`). Remove `MinHeight="300"` from
the Jobs DataGrid `Border` — this guard (`CLAUDE.md` §6.1) was written against the pre-Phase-1 `SizeToContent=
"WidthAndHeight"` infinite-measure collapse bug, a failure mode the current `BaseWindow` engine (post Phase 1)
doesn't have; FinalWorkingCode has no such floor either. Dropping it matches the full-sync policy; the window's
700px `MinHeight` easily satisfies a 300px floor regardless, so this carries no visible risk.

- [ ] **Step 5: Preserve verbatim**

Keep `TitleBar_MouseLeftButtonDown`, rewired onto the new header. All other constructor/handler wiring
(`ShowWarningAction`/`ShowInfoAction`/`ShowBusyAction`, `DgJobsRow_PreviewMouseLeftButtonDown`,
`UpdateHeaderCheckbox`, `FindAncestor<T>`) already matches FinalWorkingCode — do not touch.

- [ ] **Step 6: Build verification**

Run: `msbuild GLSense.Addin.Core\GLSense.Addin.Core.csproj /t:Build /p:Configuration=Debug /p:SignAssembly=false
/v:normal /nologo`
Expected: `Build succeeded. 0 Error(s)`.

- [ ] **Step 7: Commit**

```bash
git add GLSense.Addin.Core/Views/GLJobsMonitor.xaml GLSense.Addin.Core/Views/GLJobsMonitor.xaml.cs
git commit -m "Port GLJobsMonitor header/chrome/sizing parity with FinalWorkingCode"
```

---

### Task 17: Port `GLUserConfig` header/chrome/sizing

**Files:**
- Modify: `GLSense.Addin.Core\Views\GLUserConfig.xaml`
- Modify: `GLSense.Addin.Core\Views\GLUserConfig.xaml.cs`

- [ ] **Step 1: Read both sides in full before editing**

This window's research pass only covered the first 120 lines of each XAML file (DrillDowns tab + start of Options
tab) given the strength of the pattern established by the other 3 windows in this batch — read
`FinalWorkingCode\GLSense\Views\GLUserConfig.xaml`/`.xaml.cs` and AIPowered's current copies **in full**
(including every remaining tab and the entire `.xaml.cs`) before editing, mirroring Phase 1 Task 7's "read both
sides in full" discipline more strictly than any other task in this plan.

- [ ] **Step 2: Replace the header and add the outer window frame**

Canonical header markup: `Kind="UserGearSolid"`, `Text="GLSense User Preferences"` (caption text already matches
exactly — only the icon and header style change). Root: `WindowStyle="None"`. Wrap the entire content `Grid` in an
outer `Border Background="#F8F9FA" BorderBrush="#DEE2E6" BorderThickness="1"` — a full-window frame FinalWorkingCode
has that AIPowered's `RootGrid` currently lacks entirely (not present on any other window in this batch).

- [ ] **Step 3: Port sizing**

Change `SizeToContent="WidthAndHeight" MinWidth="500" MaxWidth="600" MinHeight="400" MaxHeight="560"` to
`SizeToContent="Manual" MinHeight="430" MaxHeight="520" MinWidth="550" MaxWidth="630"`.

- [ ] **Step 4: Port shared-style keys throughout**

Replace inline `Style` blocks with FinalWorkingCode's named, shared styles: `ColumnHeaderStyle="{StaticResource
ModernDataGridColumnHeader}"`, `RowStyle="{StaticResource ModernDataGridRowWithBorder}"`,
`CellStyle="{StaticResource ModernDataGridCell}"`, `ElementStyle="{StaticResource ModernDataGridTextBlock}"`,
`ItemContainerStyle="{StaticResource TabItemStyle}"` (replacing AIPowered's current `ModernTabItem`),
`Style="{StaticResource CompactCardStyle}"` wrapping each tab's content, `Style="{StaticResource
TextBlockLabelStyle}"` for field labels, `Style="{StaticResource CompactCheckBox}"` for checkboxes. Before applying
each, grep `GlobalStyles.xaml`/`Generic.xaml` to confirm the key exists (per Global Constraints) — if any is
missing, add its definition copied from FinalWorkingCode's `Themes/GlobalStyles.xaml` first.

- [ ] **Step 5: Preserve verbatim**

Keep the `ScrollViewer Grid.Row="1"` wrap around `MainTabControl` — confirmed FinalWorkingCode's own current file
already has this too (`CLAUDE.md` §6.1's fix has converged with the golden reference; no separate preserve action
beyond a faithful port). Keep `TitleBar_MouseLeftButtonDown`, rewired onto the new header. Whatever
`MainTabControl_SelectionChanged`/`CmbOptions_Loaded`/spinner/checkbox handlers exist must be re-verified control
name matches after applying Step 4's style-key swaps (a style change never renames an `x:Name`, but confirm no
handler assumed a specific `Style`-driven visual state).

- [ ] **Step 6: Build verification**

Run: `msbuild GLSense.Addin.Core\GLSense.Addin.Core.csproj /t:Build /p:Configuration=Debug /p:SignAssembly=false
/v:normal /nologo`
Expected: `Build succeeded. 0 Error(s)`.

- [ ] **Step 7: Commit**

```bash
git add GLSense.Addin.Core/Views/GLUserConfig.xaml GLSense.Addin.Core/Views/GLUserConfig.xaml.cs
git commit -m "Port GLUserConfig header/chrome/sizing parity with FinalWorkingCode"
```

---

## Phase 5 (Batch 5): Highest-complexity windows

### Task 18: Port `GLServerConfiguration` header/chrome/sizing (chrome/sizing only — business logic untouched)

**Files:**
- Modify: `GLSense.Addin.Core\Views\GLServerConfiguration.xaml`
- Modify: `GLSense.Addin.Core\Views\GLServerConfiguration.xaml.cs`

- [ ] **Step 1: Read both sides in full before editing**

This window's status/toast UX was deliberately reworked in AIPowered (`CLAUDE.md` §6) and diverges intentionally
from FinalWorkingCode — read both `.xaml.cs` files in full to reconfirm the `UpdateStatus` signature difference
before touching anything.

- [ ] **Step 2: Replace the header**

Canonical header markup: `Kind="DatabaseSolid"`, `Text="Instance Configuration"`. Root: `WindowStyle="None"`. Since
FinalWorkingCode's header `Border` has no `MouseLeftButtonDown` wired in its current file, explicitly add
`MouseLeftButtonDown="TitleBar_MouseLeftButtonDown"` onto the copied header (do not silently lose drag capability).

- [ ] **Step 3: Port sizing**

Change `SizeToContent="WidthAndHeight" MinWidth="640" MaxWidth="720" MinHeight="410" MaxHeight="650"` to
`SizeToContent="Manual" MinWidth="650" MaxWidth="720" MinHeight="500" MaxHeight="650"`. Add
`WindowStartupLocation="CenterOwner"`, `Background="#F8F9FA"`, and a root `Grid Margin="10"` (replacing per-Border
margins with the shared inset, matching FinalWorkingCode's structure). Remove `ExtendsContentIntoTitleBar="True"`
and `WindowStyle="SingleBorderWindow"`.

- [ ] **Step 4: Port DataGrid header template, ScrollViewer wrap, tooltips, and Border styles**

Port FinalWorkingCode's sort-arrow `ColumnHeaderStyle` `ControlTemplate` for `dgInstances` (currently just flat
Setters in AIPowered, no `Template` override — a real column-header-clipping risk, same shape as `CLAUDE.md` §7.1's
original fix). Wrap `dgInstances` in `<ScrollViewer VerticalScrollBarVisibility="Auto"
HorizontalScrollBarVisibility="Disabled">` matching FinalWorkingCode (currently relying only on the DataGrid's own
attached-property scrolling). Port the amber `#FBB41A`/`LightbulbRegular` tip-box tooltip content for the "URL
Name"/"URL Address" columns and `btnUpdate`/`btnDelete`, plus `EditingElementStyle` tooltips shown while editing a
cell (currently missing entirely). Change the DataGrid section `Border` to `Style="{StaticResource
ContentBorderStyle}" Margin="0,0,0,12"` and the action-buttons `Border` to `Background="White" CornerRadius="6"
BorderBrush="#E9ECEF" Padding="8,4"` (both replacing inline `WhiteBrush`/`BorderLightBrush`/`CornerRadius="8"`
values). Keep AIPowered's `btnUpdate` tooltip text ("Add / save server configuration") — do not adopt
FinalWorkingCode's differently-worded tooltip, since the tip-box styling is the visual element being ported, not
the specific copy.

- [ ] **Step 5: Remove the now-unneeded DataGridColumnFillHelper call**

Remove `DataGridColumnFillHelper.EnableFillColumn(dgInstances, dgInstances.Columns[1])` from the constructor — this
call is only correct under `SizeToContent="WidthAndHeight"` (`CLAUDE.md` §8.1); once this window moves to `Manual`
in Step 3, it becomes actively harmful. The "URL Address" column's `Width="2*"` declaration itself is unchanged
(both files already declare it that way — only the runtime override via the helper is being removed).

- [ ] **Step 6: Preserve verbatim — do not touch the reworked status/toast UX**

**Do NOT port FinalWorkingCode's two-argument `UpdateStatus(string message, bool isSuccess)` signature or its
inline `txtStatus` success/error message behavior.** Keep AIPowered's single-argument `UpdateStatus(string
message)` (routes to `AppOverlayControl?.ShowWarning(...)` toast) and `UpdateInstanceCountStatus()` (persistent
"Total configurations loaded : N" count via the `SuccessMessage` style) exactly as they are — this is a deliberate,
already-reasoned AIPowered redesign (`CLAUDE.md` §6), not drift to correct. Keep `UrlInstances_CollectionChanged`/
`Instance_PropertyChanged` calling `UpdateInstanceCountStatus()`, and `SaveConfiguration`'s `HasUnsavedChanges()`/
`_lastSavedSnapshot` guard — none of these have a FinalWorkingCode counterpart and none are touched by a chrome/
sizing-only port. Confirm `SuccessMessage`/`ErrorMessage` styles still resolve from `GlobalStyles.xaml` after the
XAML changes above (used by `UpdateInstanceCountStatus`) — verify at edit time, do not assume.

- [ ] **Step 7: Build verification**

Run: `msbuild GLSense.Addin.Core\GLSense.Addin.Core.csproj /t:Build /p:Configuration=Debug /p:SignAssembly=false
/v:normal /nologo`
Expected: `Build succeeded. 0 Error(s)`.

- [ ] **Step 8: Commit**

```bash
git add GLSense.Addin.Core/Views/GLServerConfiguration.xaml GLSense.Addin.Core/Views/GLServerConfiguration.xaml.cs
git commit -m "Port GLServerConfiguration header/chrome/sizing parity with FinalWorkingCode (business logic untouched)"
```

---

### Task 19: Port `GLCubeDetails` header/chrome/sizing (chrome/sizing only — ribbon-sync/async-load logic untouched)

**Files:**
- Modify: `GLSense.Addin.Core\Views\GLCubeDetails.xaml`
- Modify: `GLSense.Addin.Core\Views\GLCubeDetails.xaml.cs`

- [ ] **Step 1: Read both sides in full before editing**

This window's `Window_Loaded` carries the hard-won AppState-commit-ordering fix from `CLAUDE.md` §3 (the
Ribledger-dropdown-empty saga, one of the longest debugging sagas in this codebase) — read both `.xaml.cs` files in
full and locate every `Grid.Row`/`Grid.GetRow(...)` index reference in the code-behind before touching row
structure, per the recommendation below to leave the row layout flat.

- [ ] **Step 2: Replace the header**

Canonical header markup: `Kind="CubesSolid"`, `Text="GLSense Cube Details"`. Root: `WindowStyle="None"`. Explicitly
add `MouseLeftButtonDown="TitleBar_MouseLeftButtonDown"` onto the copied header (FinalWorkingCode's current file has
none wired there either).

- [ ] **Step 3: Port sizing**

Change `SizeToContent="WidthAndHeight" MinWidth="680" MaxWidth="750" MinHeight="500" MaxHeight="650"` to
`SizeToContent="Manual" MinWidth="800" MaxWidth="950" MinHeight="500" MaxHeight="650"`. Add
`Background="#F8F9FA"` and a root `Grid Margin="10"`. Remove `ExtendsContentIntoTitleBar="True"` and
`WindowStyle="SingleBorderWindow"`. **Leave AIPowered's existing flat 4-row `Grid` structure in place** rather than
adopting FinalWorkingCode's nested (outer 3-row + inner 2-row) structure — the async-grid-load resettle logic tied
to specific row indices (`CLAUDE.md` §1.4b) is a real risk if row indices shift, and a flat structure produces the
same visible layout once sizing/margins match.

- [ ] **Step 4: Port control colors/sizes and fix the stray tooltip text**

Change the cube-selector `Label`'s `Foreground="{StaticResource PrimaryBrush}"` to the literal `Foreground="#2E86AB"`
(same underlying color, matching FinalWorkingCode's hardcoded value). On `cmbCubes`: add `Width="Auto" MaxWidth="500"
ArrowBrush="#2E86AB" ArrowWidth="10" ArrowHeight="8" Background="White" Foreground="#212529" BorderBrush="#E9ECEF"`
— first confirm `SuggestAppendComboBox` actually exposes `ArrowBrush`/`ArrowWidth`/`ArrowHeight` as dependency
properties (grep `Controls\SuggestAppendComboBox.cs`) before adding them; skip silently if the properties don't
exist rather than leaving an XAML property that fails to bind. Change `chkValidateCube`'s `Margin="0,0,20,0"` to
`Margin="0,0,25,0"`. Change `btnValidateCube`'s `MinWidth="90"` to `MinWidth="100"`. Change the footer status
message's `MaxWidth="400"` to `MaxWidth="450"`. **Fix the warning-column tooltip text to read `"Warning"`** — do
NOT copy FinalWorkingCode's literal `"? Warning"` text (a stray, almost certainly mis-encoded character in
FinalWorkingCode's own source) — keep AIPowered's existing clean string instead.

- [ ] **Step 5: Preserve verbatim — do not touch ribbon-sync/async-load logic**

`Window_Loaded`'s `ProcessCubeSelectionNew`/`ProcessCubeSelectionReload` → `LoadCubeLedgers` → `UpdateRibbonForCube`
call ordering (commit `AppState.Instance.SelectedCube`/`SelectedLedger` before `LoadCubeLedgers`, which itself must
run before `UpdateRibbonForCube` — `CLAUDE.md` §3) must not be reordered or touched in any way. Keep
`LoadCubeLedgers`'s debug logging. Confirm `DataGridColumnFillHelper` is genuinely absent from this file's
constructor (already confirmed during research — `dgCubes`'s columns are fixed/star widths with no fill-column
call) so no removal step is needed here, unlike `GLServerConfiguration`/`GLLOVs`/`GLJobsMonitor`. Keep
`TitleBar_MouseLeftButtonDown`.

- [ ] **Step 6: Build verification**

Run: `msbuild GLSense.Addin.Core\GLSense.Addin.Core.csproj /t:Build /p:Configuration=Debug /p:SignAssembly=false
/v:normal /nologo`
Expected: `Build succeeded. 0 Error(s)`.

- [ ] **Step 7: Commit**

```bash
git add GLSense.Addin.Core/Views/GLCubeDetails.xaml GLSense.Addin.Core/Views/GLCubeDetails.xaml.cs
git commit -m "Port GLCubeDetails header/chrome/sizing parity with FinalWorkingCode (ribbon-sync logic untouched)"
```

---

### Task 20: Port `GLLogin` header/chrome/sizing

**Files:**
- Modify: `GLSense.Addin.Core\Views\GLLogin.xaml`
- Modify: `GLSense.Addin.Core\Views\GLLogin.xaml.cs`

- [ ] **Step 1: Replace the header — adopt FinalWorkingCode's exact title/icon per the full-sync policy**

Canonical header markup: `Kind="RightToBracketSolid"` (replacing AIPowered's current `IconSymbol="KeySolid"`),
`Text="Orbit GLSense Login"` (replacing AIPowered's current `WindowCaption="GLSense Login"` — a real, deliberate
product-copy change per the Global Constraints full-sync policy, not an oversight). Wire the header's close button
to `CloseButton_Click` (AIPowered's existing title-bar-specific handler — keep the two-handler split rather than
consolidating to FinalWorkingCode's single `BtnClose_Click` for both header and footer, to avoid touching working
code-behind). Root: `WindowStyle="None"`.

- [ ] **Step 2: Port sizing — adopt FinalWorkingCode's growable-width policy**

Change `SizeToContent="WidthAndHeight" MinWidth="450" MaxWidth="450" MinHeight="750" MaxHeight="900"` to
`SizeToContent="Manual" Width="540" MinWidth="480" MinHeight="750"` (no `MaxWidth`/`MaxHeight` — FinalWorkingCode
leaves both axes free to grow past their minimums). This replaces AIPowered's fixed-width policy with
FinalWorkingCode's growable one, per the full-sync policy — flag to the user during the batch's post-build visual
check specifically because a WebView2 login page's internal layout may have assumed the old fixed 450px viewport;
watch for any visual clipping/reflow in the WebView2 content area during that check. Add `Background="#F8F9FA"`
and a root `Grid Margin="10"`. Remove `ExtendsContentIntoTitleBar="True"` and `WindowStyle="SingleBorderWindow"`.

- [ ] **Step 3: Port label/card/button styles**

Replace the Instance/URL label rows' `Style="LabelWithIconStyle"`/`LabelIconStyle`/`LabelTextStyle` treatment with
FinalWorkingCode's pill-style boxes: `Border Background="#F8F9FA" CornerRadius="4" Padding="8,6"` containing
`iconPacks:PackIconFontAwesome Style="{StaticResource SmallIcon}" Foreground="#2E86AB"` + `TextBlock
Style="{StaticResource HeaderTextBlock}"`. Keep the form section wrapper on AIPowered's existing
`Style="LoginCardStyle"` (confirm it renders identically to FinalWorkingCode's inline `Border Background="White"
CornerRadius="6" BorderBrush="#E9ECEF" Padding="15"` before deciding — if `LoginCardStyle`'s definition already
produces this exact result, keep using the shared style rather than reverting to inline values). Add
`ArrowBrush="#2E86AB" ArrowWidth="10" ArrowHeight="8"` and `Background="White" Foreground="#212529"` to `cmbServer`.
Add `Height="32"` to `txtServer`. Change its tooltip `Style` from `ChromeStyleToolTip` to `SimpleBrowserToolTip`.
Change the footer's `btnClose` from a `Grid`-with-spacer-column `Style="DynamicContentButton" Content="Close"` to a
plain `StackPanel HorizontalAlignment="Right"` with `Style="CloseButtonStyle"` (icon-only close button, no
`Content="Close"` text) — wire it to the existing `BtnClose_Click` handler, unchanged.

- [ ] **Step 4: Preserve verbatim**

Do not touch `CleanLoginUrl`'s trailing-slash trim or `ExtractLoginCookies`'s URL-decoding (`CLAUDE.md` §32 item 1)
— this task is XAML-only and neither is reachable from anything changed above. Do not touch the `_resilience`
WebView2 attach/detach lifecycle. Keep `TitleBar_MouseLeftButtonDown`, rewired onto the new header, alongside the
existing `CloseButton_Click` (title bar) / `BtnClose_Click` (footer) split.

- [ ] **Step 5: Build verification**

Run: `msbuild GLSense.Addin.Core\GLSense.Addin.Core.csproj /t:Build /p:Configuration=Debug /p:SignAssembly=false
/v:normal /nologo`
Expected: `Build succeeded. 0 Error(s)`.

- [ ] **Step 6: Commit**

```bash
git add GLSense.Addin.Core/Views/GLLogin.xaml GLSense.Addin.Core/Views/GLLogin.xaml.cs
git commit -m "Port GLLogin header/chrome/sizing parity with FinalWorkingCode"
```

---

### Task 21: Port `GLDrilldownCustomization` header/chrome/sizing

**Files:**
- Modify: `GLSense.Addin.Core\Views\GLDrilldownCustomization.xaml`
- Modify: `GLSense.Addin.Core\Views\GLDrilldownCustomization.xaml.cs`

- [ ] **Step 1: Verify the FontAwesome icon name before using it**

FinalWorkingCode's header uses `Kind="GearsSolid"` (plural); AIPowered's current `IconSymbol="GearSolid"`
(singular) — these are two different icon names, and only one may actually exist in the installed
`MahApps.Metro.IconPacks.FontAwesome` package version. Before writing the header markup, check which name is valid:
search the installed package's `PackIconFontAwesomeKind` enum (e.g. via
`GLSense.Addin.Core\obj\Debug\*.cs`-generated references, or by grepping any other window's XAML already using a
`Gears`/`Gear`-prefixed icon, or by attempting a build after choosing one and treating an XAML parse/resource
error as the signal to try the other). Use whichever name is confirmed valid; if both are valid, prefer
FinalWorkingCode's `GearsSolid` per the full-sync policy.

- [ ] **Step 2: Replace the header**

Canonical header markup with the icon name confirmed in Step 1, `Text="GLSense Drilldown Customization"` (title
text already matches exactly on both sides). Root: `WindowStyle="None"`. Wire
`MouseLeftButtonDown="TitleBar_MouseLeftButtonDown"` — keep AIPowered's existing handler name (confirmed present
via grep) even though FinalWorkingCode's own file names its handler `Header_MouseLeftButtonDown`; do not rename
AIPowered's handler to match.

- [ ] **Step 3: Port sizing**

Change `SizeToContent="WidthAndHeight" MinWidth="1250" MaxWidth="1300" MinHeight="750" MaxHeight="900"` to
`SizeToContent="Manual" MinWidth="1250" MaxWidth="1500" MinHeight="850"` (no `MaxHeight` — FinalWorkingCode leaves
this axis unbounded, reasonable for a full-bleed WebView2 host where more vertical space is harmless). Add
`Background="#F8F9FA"` and a root `Grid Margin="10"`. Remove `ExtendsContentIntoTitleBar="True"` and
`WindowStyle="SingleBorderWindow"`.

- [ ] **Step 4: Port minor style values**

Change the WebView2 `Border`'s `Margin="10,8,10,8"` to `Margin="0,2,0,0"` (the outer `Grid Margin="10"` from Step 3
now supplies the outer inset this margin previously provided directly) — keep `Style="{StaticResource
ContentBorderStyle}"` unchanged, already correct on both sides. Change `SaveInfoMessage`'s `FontSize="13"` to
`FontSize="12"`.

- [ ] **Step 5: Preserve verbatim**

No AIPowered-only business logic is at risk here — `BtnSaveLocally_Click`'s cancellation handling and
`LogRawJson`-on-failure logic is untouched by this XAML-only task regardless. Keep `TitleBar_MouseLeftButtonDown`
(Step 2) and the existing `BtnClose_Click` unchanged.

- [ ] **Step 6: Build verification**

Run: `msbuild GLSense.Addin.Core\GLSense.Addin.Core.csproj /t:Build /p:Configuration=Debug /p:SignAssembly=false
/v:normal /nologo`
Expected: `Build succeeded. 0 Error(s)`.

- [ ] **Step 7: Commit**

```bash
git add GLSense.Addin.Core/Views/GLDrilldownCustomization.xaml GLSense.Addin.Core/Views/GLDrilldownCustomization.xaml.cs
git commit -m "Port GLDrilldownCustomization header/chrome/sizing parity with FinalWorkingCode"
```

---

## Handoff to the user

After Task 21's commit, all 21 windows are chrome/style/sizing-synced with FinalWorkingCode, with every
AIPowered-specific business-logic fix cataloged in `CLAUDE.md` sections 2 through 37 preserved. Remaining follow-up
work, explicitly deferred (per the spec):
- `GLSegmentRef`/`GLSegmentManager` sync or retirement — separate future effort.
- Final prune of now-dead `GlobalStyles.xaml`/`Generic.xaml` keys — one pass after this plan's own execution,
  mirroring Phase 1's Task 8.

The user should rebuild and visually check each phase's windows in Excel before the next phase starts, per this
codebase's established deployment note: a running Excel session will not pick up a fresh build — Excel must be
fully closed and relaunched after each phase's rebuild.
