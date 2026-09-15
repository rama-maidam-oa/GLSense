# Floating Balance Configurator Window (Option A) Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Add a new, non-modal, native WPF `Window` that hosts the existing Balance Configurator UI as a floating alternative to the current docked `ADXExcelTaskPane`, without touching the task pane at all, so the two can be compared side by side before deciding whether to replace the pane.

**Architecture:** The task pane's rendering bugs (DPI/monitor-drag glitches, duplicate/ghost separator lines while resizing) trace back to hosting WPF content inside a WinForms `ElementHost` embedded in an Add-in Express `ADXExcelTaskPane` — a classic WPF/WinForms "Airspace" interop problem, compounded by four redundant, uncoordinated resize-enforcement code paths in `GLConfiguratorPane.cs` with no explicit repaint calls. A genuine top-level WPF `Window` has none of this: it is pure WPF end to end, and this codebase already has a mature, battle-tested `DpiAwareWindow` base class (used by ~15 other windows) that gets DPI handling, work-area clamping, and owner-focus restoration for free. The plan wraps the *existing, unmodified* `GLBalanceConfigurator` `UserControl` (the same one the task pane already hosts) inside a new `DpiAwareWindow`-derived `GLBalanceConfiguratorWindow`, wires a new sibling ribbon button to open it, and hardens the one genuinely new risk this introduces — keeping OS keyboard focus on the floating window (not Excel's active cell) after the user clicks into Excel to pick a cell reference.

**Tech Stack:** .NET Framework 4.8.1, WPF (old-style `.csproj` — every new file needs an explicit `<Compile>`/`<Page>` entry, no implicit globbing), Add-in Express 10.1.4703.0 (ribbon buttons are designer-generated `ADXRibbonButton` fields in `AddinModule.Designer.cs`, not raw ribbon XML), MahApps.Metro.IconPacks for icons. **No automated UI test framework exists anywhere in this repo** — every prior fix in `FinalWorkingCode\CLAUDE.md` is verified by an MSBuild build plus manual interactive verification in Excel, not unit tests. This plan follows that same established verification method: each task ends with a build-verify step, and the final task is a structured manual QA checklist instead of a test suite.

**Spec:** This plan document — written directly from a live investigation of the current codebase (see file:line references inline in each task) rather than a separate spec file, since the "spec" here is simply "reuse the existing Balance Configurator UI unmodified, in a new floating window, without disturbing the task pane."

## Global Constraints

- **FinalWorkingCode only.** Do not touch anything under `AIPowered\`.
- **Do not modify, remove, or disable the existing task pane** (`GLConfiguratorPane.cs`, `GLConfiguratorPane.Designer.cs`, the `RibFSG` ribbon button/handler, `AppState.BalancePane`/`displayConfigurator`). The new window is strictly additive — a new file, a new ribbon button, and small additive extensions to `AddinModule.cs`/`AppState.cs`/`ExcelWindowPositioning.cs`/`ExcelRefEditControl.xaml.cs`.
- **Reuse `GLBalanceConfigurator` (the `UserControl` in `Views\GLBalanceConfigurator.xaml`/`.xaml.cs`) completely unchanged.** It is already self-contained: its own `AppOverlayControl` (busy/warning/toast), its own `GLConfiguratorViewModel` construction (`ExcelApp = AppState.Instance.ExcelApp.Application`), and its own `ReLoadConfigurator()`/`ResetCellReference()`/`OnCloseRequested` — all already correctly scoped to the control itself, not the task pane. Its constructor's `parentPane` parameter is optional (`GLBalanceConfigurator(GLConfiguratorPane parentPane = null)`) and already null-safe everywhere it's used (`Views\GLBalanceConfigurator.xaml.cs:96-99`, `:286`), so hosting it with no parent pane requires zero changes to that file.
- **This codebase is single-threaded for all WPF/Excel UI** — confirmed via `Utilities\WpfAppManager.cs`: `EnsureApplication()` creates the WPF `Application` object *on whatever thread calls it* (Excel's own main STA thread, via ribbon-click handlers), it never spawns a separate thread. There is therefore no `AttachThreadInput`-across-threads class of bug possible here — the "typing leaks into the Excel cell" risk is a **Win32 foreground/keyboard-focus activation** issue (a non-modal window not truly holding OS focus after a programmatic hand-off), not a cross-thread message-pump issue. Task 2 addresses exactly that.
- Every new `.xaml`/`.xaml.cs` file needs explicit `<Compile>` and `<Page>` entries added to `GLSense\GLSense.csproj` (old-style project, confirmed via `Views\GLCubeDetails.xaml`'s existing entries at `GLSense.csproj:255-257,377-380`) — MSBuild will not pick the files up otherwise, and the build will fail loudly with "type not found" if this is forgotten.
- Follow the existing `DpiAwareWindow`-derived window convention exactly (see `Views\GLCubeDetails.xaml:1-21`): `WindowStyle="None"`, `ResizeMode="CanResize"`, custom header `Border` with `HeaderBar`/`TitleHeaderTextBlock`/`CustomWindowCloseButtonStyle` styles from the shared `GlobalStyles.xaml` resource dictionary, `EnhancedDragDropHelper.EnableWindowDrag(this)` in the constructor (chromeless windows have no native title bar to drag by).

---

### Task 1: New floating window shell hosting the existing Balance Configurator UI

**Files:**
- Create: `GLSense\Views\GLBalanceConfiguratorWindow.xaml`
- Create: `GLSense\Views\GLBalanceConfiguratorWindow.xaml.cs`
- Modify: `GLSense\GLSense.csproj`
- Modify: `GLSense\AppState.cs:96-97`

**Interfaces:**
- Produces: `GLSense.Views.GLBalanceConfiguratorWindow : GLSense.Utilities.DpiAwareWindow`, with:
  - `public void ShowFloating(IntPtr excelHwnd)` — shown/used by Task 3
  - `public System.Threading.Tasks.Task RelaunchWindow()` — used by Task 4
  - `public System.Threading.Tasks.Task ResetWindowReference()` — used by Task 4
  - Named element `ConfiguratorControl` of type `GLSense.Views.GLBalanceConfigurator` (the existing, unmodified `UserControl`)
- Produces: `GLSense.AppState.Instance.BalanceWindow` property of type `GLBalanceConfiguratorWindow`
- Consumes: `GLSense.Utilities.DpiAwareWindow` (`ShowWithOwner`, `DisableAutoSizing` — both already exist, `Utilities\DpiAwareWindow.cs:58,169-195`), `GLSense.Views.GLBalanceConfigurator` (already exists, unmodified), `GLSense.Helpers.EnhancedDragDropHelper.EnableWindowDrag` (already used by `GLCubeDetails.xaml.cs:45`)

- [ ] **Step 1: Create the window XAML**

```xml
<utils:DpiAwareWindow x:Class="GLSense.Views.GLBalanceConfiguratorWindow"
                      xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
                      xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
                      xmlns:local="clr-namespace:GLSense.Views"
                      xmlns:utils="clr-namespace:GLSense.Utilities"
                      xmlns:iconPacks="clr-namespace:MahApps.Metro.IconPacks;assembly=MahApps.Metro.IconPacks.FontAwesome"
                      Background="#F8F9FA" SizeToContent="Manual"
                      MinWidth="600" MaxWidth="1000" MinHeight="300" MaxHeight="900"
                      Width="650" Height="700"
                      WindowStyle="None" ResizeMode="CanResize" WindowStartupLocation="CenterOwner"
                      Loaded="Window_Loaded">
    <utils:DpiAwareWindow.Resources>
        <ResourceDictionary>
            <ResourceDictionary.MergedDictionaries>
                <ResourceDictionary Source="pack://application:,,,/GLSense;component/Themes/GlobalStyles.xaml"/>
            </ResourceDictionary.MergedDictionaries>
        </ResourceDictionary>
    </utils:DpiAwareWindow.Resources>
    <Grid Margin="0">
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
        </Grid.RowDefinitions>
        <Border Grid.Row="0" Style="{StaticResource HeaderBar}">
            <Grid>
                <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="Auto"/>
                    <ColumnDefinition Width="*"/>
                    <ColumnDefinition Width="Auto"/>
                </Grid.ColumnDefinitions>
                <StackPanel Grid.Column="0" Orientation="Horizontal" VerticalAlignment="Center">
                    <iconPacks:PackIconFontAwesome Kind="CalculatorSolid" Style="{StaticResource LargeIcon}"/>
                    <TextBlock Text="Balance Configurator" Style="{StaticResource TitleHeaderTextBlock}"/>
                </StackPanel>
                <Button Grid.Column="2" Style="{StaticResource CustomWindowCloseButtonStyle}" Click="BtnClose_Click"/>
            </Grid>
        </Border>
        <local:GLBalanceConfigurator x:Name="ConfiguratorControl" Grid.Row="1"/>
    </Grid>
</utils:DpiAwareWindow>
```

`MinWidth`/`MinHeight`/`MaxWidth`/`MaxHeight` here are plain WPF `Window` properties — native `ResizeMode="CanResize"` enforces them automatically with no extra code, unlike the task pane's manual `WM_SIZING`/`SetBoundsCore` RECT-mutation hacks in `GLConfiguratorPane.cs:277-340` (which is exactly the code implicated in the ghosting bug). `MinWidth="600"`/`MinHeight="300"` are the task pane's own floor verbatim (`GLConfiguratorPane.cs:17-18`'s `_minWidthDip`/`_minHeightDip`). The pane itself has **no** `MaximumSize` at all (confirmed via grep — only `MinimumSize` exists in `GLConfiguratorPane.cs`), so `MaxWidth="1000"`/`MaxHeight="900"` have no pane precedent to mirror; they're a reasonable starting ceiling per the original request ("keep the min and max... so it should [not] be [resized] beyond this size"). Adjust the exact numbers later once the window has been used interactively — nothing else in this plan depends on the specific values.

- [ ] **Step 2: Create the code-behind**

```csharp
using GLSense.Helpers;
using GLSense.Utilities;
using System;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Input;
using System.Windows.Interop;

namespace GLSense.Views
{
    public partial class GLBalanceConfiguratorWindow : DpiAwareWindow
    {
        public GLBalanceConfiguratorWindow()
        {
            LogUtility.LogDebug("GLBalanceConfiguratorWindow.ctor invoked");

            // Deliberately leave DisableAutoSizing at its default (false) - it looks
            // tempting to set true here to stop DpiAwareWindow's auto-fit/recenter passes
            // from fighting a user's manual drag-resize, but that flag is an all-or-nothing
            // switch (Utilities\DpiAwareWindow.cs:54-58): it ALSO disables the WM_DPICHANGED
            // handler's ApplyScaleTransform call (:452,485), which is the actual mechanism
            // that rescales WPF content when this window moves to a different-DPI monitor.
            // Setting it true would silently reintroduce the exact DPI-glitch bug this
            // window exists to fix. None of the ~15 existing floating windows in this repo
            // set this flag (confirmed via grep) - the fixed MinWidth/MaxWidth/MinHeight/
            // MaxHeight in this window's XAML are plain WPF Window properties, already
            // enforced natively by ResizeMode="CanResize" independent of this flag, so
            // there is nothing to disable here.

            InitializeComponent();
            EnhancedDragDropHelper.EnableWindowDrag(this);

            ConfiguratorControl.OnCloseRequested += () => Close();
        }

        /// <summary>
        /// Shows this window non-modally owned by Excel's main window, then forces real
        /// OS foreground/keyboard-focus onto it - see the ForceSetForegroundWindow doc
        /// comment in ExcelWindowPositioning.cs for why a plain Activate() alone is not
        /// always enough here.
        /// </summary>
        public void ShowFloating(IntPtr excelHwnd)
        {
            LogUtility.LogDebug("GLBalanceConfiguratorWindow.ShowFloating invoked");
            ShowWithOwner(excelHwnd);
            ReclaimForegroundFocus();
        }

        private void ReclaimForegroundFocus()
        {
            try
            {
                Activate();
                var hwnd = new WindowInteropHelper(this).Handle;
                if (hwnd != IntPtr.Zero)
                {
                    ExcelWindowHelper.ForceSetForegroundWindow(hwnd);
                }
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex, "GLBalanceConfiguratorWindow.ReclaimForegroundFocus");
            }
        }

        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            LogUtility.LogDebug("GLBalanceConfiguratorWindow.Window_Loaded invoked");

            // A freshly-shown non-modal window can be the active window yet still have
            // no WPF element holding keyboard focus - explicitly move focus into the
            // content so the very first keystroke after opening lands in a field here,
            // not wherever WPF keyboard focus last was.
            ReclaimForegroundFocus();
            MoveFocus(new TraversalRequest(FocusNavigationDirection.First));
        }

        private void BtnClose_Click(object sender, RoutedEventArgs e)
        {
            LogUtility.LogDebug("GLBalanceConfiguratorWindow.BtnClose_Click invoked");
            Close();
        }

        public Task RelaunchWindow() => ConfiguratorControl.ReLoadConfigurator();

        public Task ResetWindowReference()
        {
            GLBalanceConfigurator.ResetCellReference();
            return Task.CompletedTask;
        }
    }
}
```

`ExcelWindowHelper.ForceSetForegroundWindow` is made `public` in Task 2 below — this step references it, but Task 2 must land first (or in the same pass) for this to compile.

- [ ] **Step 3: Register the new files in the project**

In `GLSense\GLSense.csproj`, next to the existing `GLCubeDetails.xaml`/`.xaml.cs` entries (`GLSense.csproj:255-257` for `<Compile>`, `:377-380` for `<Page>`):

```xml
<Compile Include="Views\GLBalanceConfiguratorWindow.xaml.cs">
  <DependentUpon>GLBalanceConfiguratorWindow.xaml</DependentUpon>
</Compile>
```

```xml
<Page Include="Views\GLBalanceConfiguratorWindow.xaml">
  <SubType>Designer</SubType>
  <Generator>MSBuild:Compile</Generator>
</Page>
```

- [ ] **Step 4: Add the `AppState` field**

In `GLSense\AppState.cs`, add `using GLSense.Views;` to the top `using` block (alongside the existing `using GLSense.Interfaces;`/`using GLSense.Models;`/`using GLSense.Utilities;` at lines 1-3), then in the "Balance Configurator Pane" section (`AppState.cs:95-97`):

```csharp
//Balance Configurator Pane
public GLConfiguratorPane BalancePane { get; set; }
public bool displayConfigurator { get; set; } = false;

//Balance Configurator floating window (Option A prototype - task pane above is untouched)
public GLBalanceConfiguratorWindow BalanceWindow { get; set; }
```

- [ ] **Step 5: Build-verify**

Run:
```
& "C:\Program Files\Microsoft Visual Studio\2022\Community\MSBuild\Current\Bin\MSBuild.exe" "D:\SQLLite_Test\GLSense\FinalWorkingCode\GLSense.sln" /t:GLSense /p:Configuration=Debug /p:RegisterForComInterop=false /v:minimal /nologo
```
Expected: build succeeds (this task alone won't be exercisable yet — nothing calls `new GLBalanceConfiguratorWindow()` until Task 3 — so this step only confirms the new files compile and are wired into the project correctly).

- [ ] **Step 6: Commit**

```bash
git add GLSense/Views/GLBalanceConfiguratorWindow.xaml GLSense/Views/GLBalanceConfiguratorWindow.xaml.cs GLSense/GLSense.csproj GLSense/AppState.cs
git commit -m "GLSense: Add floating Balance Configurator window shell (Option A prototype)"
```

---

### Task 2: Harden OS foreground/keyboard-focus reclaim after Excel steals it

**Files:**
- Modify: `GLSense\Helpers\ExcelWindowPositioning.cs:101`
- Modify: `GLSense\Views\ExcelRefEditControl.xaml.cs:1-9,210-223`

**Interfaces:**
- Produces: `GLSense.Helpers.ExcelWindowHelper.ForceSetForegroundWindow(IntPtr hWnd)` — now `public` (was `private`), used by Task 1's `GLBalanceConfiguratorWindow.ReclaimForegroundFocus()` and by this task's own `ExcelRefEditControl.BtnEdit_Click` fix.

**Why this task exists:** `ExcelRefEditControl.BtnEdit_Click` (`Views\ExcelRefEditControl.xaml.cs:129-224`) already has the right *shape* for reclaiming focus after Excel's native `InputBox(Type:=8)` cell-picker closes — it does `Window.GetWindow(this)` to find its hosting `Window` (line 132), disables it while the InputBox is up, then in the `finally` block calls `parentWindow.Activate()` (line 215). Today this branch is dead for the task pane (`Window.GetWindow(this)` returns `null` for `ElementHost`-hosted content, so it falls through to the weaker `hostContainer.Focus()` path instead), but it becomes the *live* path the moment `GLBalanceConfigurator` is hosted inside a real `Window` — which is exactly Task 1's change. A plain `Window.Activate()` is not reliably enough on its own: `Helpers\ExcelWindowPositioning.cs:86-100`'s own doc comment explains Windows silently refuses a bare `SetForegroundWindow` call whenever another process currently holds the foreground-activation right (already reproduced once in this codebase, for reactivating Excel's own window after a long operation) — the fix there is `AttachThreadInput` to temporarily borrow that right. This task reuses that exact, already-proven helper instead of hoping `Activate()` alone is enough for the new window.

- [ ] **Step 1: Promote `ForceSetForegroundWindow` to public**

In `GLSense\Helpers\ExcelWindowPositioning.cs`, change:

```csharp
private static void ForceSetForegroundWindow(IntPtr hWnd)
```

to:

```csharp
/// <summary>
/// Reliably brings <paramref name="hWnd"/> to the foreground. A plain
/// SetForegroundWindow call is not enough on its own: Windows silently refuses it
/// (it just flashes the taskbar icon instead) whenever some other process currently
/// holds the "foreground activation" right - attaching this thread's input queue to
/// whichever thread currently owns the foreground window temporarily grants that
/// right. Generic (not Excel-specific) despite living on this class - reused by
/// GLBalanceConfiguratorWindow to reclaim focus after Excel's native cell-picker
/// InputBox closes.
/// </summary>
public static void ForceSetForegroundWindow(IntPtr hWnd)
```

(the existing doc comment above the old signature can be deleted since the new one above supersedes it and generalizes the explanation beyond "Excel's own window").

- [ ] **Step 2: Use it in `ExcelRefEditControl.BtnEdit_Click`'s `Window`-host reactivation path**

In `GLSense\Views\ExcelRefEditControl.xaml.cs`, add `using System.Windows.Interop;` to the top `using` block (alongside the existing `using GLSense.Helpers;` at line 1), then change the `finally` block at lines 210-222 from:

```csharp
finally
{
    if (windowDisabled && parentWindow != null)
    {
        parentWindow.IsEnabled = true;
        parentWindow.Activate();
    }
    else if (hostDisabled && hostContainer != null)
    {
        hostContainer.IsEnabled = true;
        hostContainer.Focus();
    }
}
```

to:

```csharp
finally
{
    if (windowDisabled && parentWindow != null)
    {
        parentWindow.IsEnabled = true;
        parentWindow.Activate();

        var parentHwnd = new WindowInteropHelper(parentWindow).Handle;
        if (parentHwnd != IntPtr.Zero)
        {
            ExcelWindowHelper.ForceSetForegroundWindow(parentHwnd);
        }
    }
    else if (hostDisabled && hostContainer != null)
    {
        hostContainer.IsEnabled = true;
        hostContainer.Focus();
    }
}
```

This only changes behavior for genuine `Window`-hosted content (the new floating window going forward) — the `hostContainer.Focus()` branch used by today's `ElementHost`-hosted task pane is untouched.

- [ ] **Step 3: Build-verify**

Run the same MSBuild command as Task 1 Step 5. Expected: build succeeds.

- [ ] **Step 4: Commit**

```bash
git add GLSense/Helpers/ExcelWindowPositioning.cs GLSense/Views/ExcelRefEditControl.xaml.cs
git commit -m "GLSense: Reuse the proven AttachThreadInput foreground-reclaim helper for window-hosted cell-reference pickers"
```

---

### Task 3: Ribbon button to open the floating window (task pane's own button untouched)

**Files:**
- Modify: `GLSense\AddinModule.Designer.cs` (near lines 53, 384, 389-399, 1344)
- Modify: `GLSense\AddinModule.cs` (near line 2730, after `RibFSG_OnClick`)

**Interfaces:**
- Consumes: `GLSense.Views.GLBalanceConfiguratorWindow` (Task 1), `AppState.Instance.BalanceWindow` (Task 1), `AppState.Instance.ExcelApp.Hwnd` (existing, same value `RibGetCube_OnClick` already uses at `AddinModule.cs:2913`), `SafeInvokeWpf` (existing, `AddinModule.cs:423-444`)

- [ ] **Step 1: Add the new ribbon button field and designer wiring**

In `GLSense\AddinModule.Designer.cs`, add the field declaration next to the existing one (`AddinModule.Designer.cs:1344`):

```csharp
public AddinExpress.MSO.ADXRibbonButton RibFSG;
public AddinExpress.MSO.ADXRibbonButton RibFSGWindow;
```

In `InitializeComponent()`, next to where `RibFSG` is instantiated (`AddinModule.Designer.cs:53`):

```csharp
this.RibFSG = new AddinExpress.MSO.ADXRibbonButton(this.components);
this.RibFSGWindow = new AddinExpress.MSO.ADXRibbonButton(this.components);
```

Add it to the same ribbon box `RibFSG` already lives in (`AddinModule.Designer.cs:384`), so it appears right next to the existing "Balance" button:

```csharp
this.adxRibbonBox4.Controls.Add(this.RibFSG);
this.adxRibbonBox4.Controls.Add(this.RibFSGWindow);
this.adxRibbonBox4.Controls.Add(this.RibLiveCalc);
```

And its own configuration block, right after the existing `RibFSG` block (`AddinModule.Designer.cs:389-399`):

```csharp
// 
// RibFSGWindow
// 
this.RibFSGWindow.Caption = "Balance (Window)";
this.RibFSGWindow.Id = "adxRibbonButton_a1b2c3d4e5f6a7b8c9d0e1f2a3b4c5d6";
this.RibFSGWindow.Image = 18;
this.RibFSGWindow.ImageList = this.ImageList_16X16;
this.RibFSGWindow.ImageTransparentColor = System.Drawing.Color.Transparent;
this.RibFSGWindow.Ribbons = AddinExpress.MSO.ADXRibbons.msrExcelWorkbook;
this.RibFSGWindow.ScreenTip = "Balance (Floating Window)";
this.RibFSGWindow.SuperTip = "Opens the Balance Configurator as a floating window instead of the docked task pane (preview)";
this.RibFSGWindow.OnClick += new AddinExpress.MSO.ADXRibbonOnAction_EventHandler(this.RibFSGWindow_OnClick);
```

`Id` must be a value not already used elsewhere in this file (`adxRibbonButton_a1b2c3d4e5f6a7b8c9d0e1f2a3b4c5d6` is a placeholder unique hex string — grep `AddinModule.Designer.cs` for it before finalizing to be certain it doesn't collide with an id added by someone else in the meantime).

- [ ] **Step 2: Add the click handler**

In `GLSense\AddinModule.cs`, right after `RibFSG_OnClick` (`AddinModule.cs:2692-2730`):

```csharp
private void RibFSGWindow_OnClick(object sender, IRibbonControl control, bool pressed)
{
    LogUtility.LogDebug("RibFSGWindow_OnClick clicked.");
    try
    {
        SafeInvokeWpf(() =>
        {
            var win = new GLBalanceConfiguratorWindow();
            AppState.Instance.BalanceWindow = win;
            win.Closed += (s, e) =>
            {
                if (AppState.Instance.BalanceWindow == win)
                {
                    AppState.Instance.BalanceWindow = null;
                }
            };
            win.ShowFloating((IntPtr)AppState.Instance.ExcelApp.Hwnd);
        });
    }
    catch (Exception ex)
    {
        LogUtility.LogException(ex);
    }
}
```

This deliberately mirrors `RibGetCube_OnClick`'s pattern (`AddinModule.cs:2907-2915`) — a fresh instance per click, shown via the `DpiAwareWindow` owner-aware `Show*WithOwner` family — except non-modal (`ShowFloating`/`ShowWithOwner`, not `ShowDialogWithOwner`), since the whole point of this window is that the user can still click Excel cells while it's open, exactly like the task pane today. `AppState.Instance.BalanceWindow` is cleared on `Closed` so Task 4's `SheetSelectionChange` handler can tell a closed window from an open one without holding a stale reference.

- [ ] **Step 3: Build-verify**

Run the same MSBuild command as Task 1 Step 5. Expected: build succeeds.

- [ ] **Step 4: Manual smoke test**

Open Excel with the add-in loaded, click the new "Balance (Window)" ribbon button (next to the existing "Balance" button). Expected: a floating window titled "Balance Configurator" appears, centered over Excel, with the same fields as the task pane. Click the existing "Balance" button too — expected: the docked task pane still opens/toggles exactly as before, completely unaffected.

- [ ] **Step 5: Commit**

```bash
git add GLSense/AddinModule.Designer.cs GLSense/AddinModule.cs
git commit -m "GLSense: Add ribbon button to open the floating Balance Configurator window"
```

---

### Task 4: Live-formula-cell auto-refresh parity with the task pane

**Files:**
- Modify: `GLSense\AddinModule.cs:1999-2027`

**Interfaces:**
- Consumes: `GLBalanceConfiguratorWindow.RelaunchWindow()`/`ResetWindowReference()` (Task 1), `AppState.Instance.BalanceWindow` (Task 1)

**Why this task exists:** Today, selecting a cell that contains an existing `GLSense_GetBalance` formula while the task pane is open auto-reloads the pane to show that formula's settings (`adxExcelAppEvents1_SheetSelectionChange`, `AddinModule.cs:1999-2027`, calling `RelaunchPane()`), and selecting any other cell resets the pane's tracked cell reference (`ResetPaneReference()`). The floating window needs the identical behavior for feature parity — "all the functionalities should work" — since a non-modal window leaves the user free to click around Excel while it's open, same as the pane.

- [ ] **Step 1: Extend `adxExcelAppEvents1_SheetSelectionChange`**

In `GLSense\AddinModule.cs`, replace the method body (`AddinModule.cs:1999-2027`):

```csharp
private void adxExcelAppEvents1_SheetSelectionChange(object sender, object sheet, object range)
{
    try
    {
        if (!AppState.Instance.IsLoginCompleted) return;
        if (range is not Excel.Range rng) return;
        if (rng.Rows.Count != 1 || rng.Columns.Count != 1) return;

        LogUtility.LogDebug($"SheetSelectionChange fired. Sheet={(sheet as Excel.Worksheet)?.Name}, Cell={rng.Address}");

        bool isBalanceFormulaCell = TryGetSingleCellFormula(rng, out string formula) &&
            formula.IndexOf(AppConstants.glBal, StringComparison.OrdinalIgnoreCase) >= 0;

        AppState.Instance.BalancePane = GetPaneInstance();
        if (AppState.Instance.BalancePane != null && AppState.Instance.BalancePane.Visible)
        {
            if (isBalanceFormulaCell)
            {
                _ = AppState.Instance.BalancePane.RelaunchPane();
            }
            else
            {
                _ = AppState.Instance.BalancePane.ResetPaneReference();
            }
        }

        if (AppState.Instance.BalanceWindow != null && AppState.Instance.BalanceWindow.IsVisible)
        {
            if (isBalanceFormulaCell)
            {
                _ = AppState.Instance.BalanceWindow.RelaunchWindow();
            }
            else
            {
                _ = AppState.Instance.BalanceWindow.ResetWindowReference();
            }
        }
    }
    catch (Exception ex)
    {
        LogUtility.LogException(ex, "adxExcelAppEvents1_SheetSelectionChange");
    }
}
```

This factors the formula check out to `isBalanceFormulaCell` (computed once instead of inline inside the pane's own `if`, since it's now needed by both branches) — the pane's own behavior is otherwise byte-for-byte identical to before. `Window.IsVisible` is the WPF equivalent of the WinForms `.Visible` already used for the pane check, and is `true` from `Show()` until `Close()`/`Hide()`.

- [ ] **Step 2: Build-verify**

Run the same MSBuild command as Task 1 Step 5. Expected: build succeeds.

- [ ] **Step 3: Manual smoke test**

With the floating window open (via the new ribbon button) and pointed at some blank cell, click on a different cell that already contains a `GLSense_GetBalance(...)` formula. Expected: the floating window's fields reload to reflect that formula's settings, matching what the task pane already does for the identical action.

- [ ] **Step 4: Commit**

```bash
git add GLSense/AddinModule.cs
git commit -m "GLSense: Floating Balance Configurator window auto-refreshes on formula-cell selection, matching the task pane"
```

---

### Task 5: Full manual verification pass

**Files:** none (verification only)

No automated test suite exists for this WPF/COM-interop UI in this repo (see Tech Stack note above) — this task is the verification-before-completion step, run interactively in Excel with the add-in loaded (Debug build from Task 1-4's build-verify steps).

- [ ] **Step 1: Build the full solution one more time**

```
& "C:\Program Files\Microsoft Visual Studio\2022\Community\MSBuild\Current\Bin\MSBuild.exe" "D:\SQLLite_Test\GLSense\FinalWorkingCode\GLSense.sln" /p:Configuration=Debug /p:RegisterForComInterop=false /v:minimal /nologo
```
Expected: full solution builds with no errors.

- [ ] **Step 2: DPI/monitor-drag smoke test (the original reported bug)**

Open the floating window via the new ribbon button. Drag the Excel window (with the floating window still open) from the primary monitor to a second monitor with a different DPI/scaling setting, or change Windows display scaling while the window is open. Expected: no ghosting, no duplicate/stuck rendering — this is native WPF `Window` DPI handling via the existing `DpiAwareWindow.AdjustForDpiChange`/`WM_DPICHANGED` path (`Utilities\DpiAwareWindow.cs:446-521`), the same mechanism every other floating window in this app already relies on successfully.

- [ ] **Step 3: Min/Max resize smoke test**

Drag-resize the floating window from each edge and corner. Expected: it stops cleanly at the `MinWidth`/`MinHeight`/`MaxWidth`/`MaxHeight` set in Task 1's XAML, with no duplicate/ghost border or separator lines during the drag (native WPF resize, none of the task pane's `WM_SIZING`/`SetBoundsCore` RECT-mutation code is involved).

- [ ] **Step 4: Keyboard-focus round trip via "Select Excel Cell"**

On any field with an `ExcelRefEditControl` (e.g. the Cell Reference field), click "Select Excel Cell", pick a cell in Excel's native picker prompt, and let it complete. Immediately start typing (e.g. Tab to another field and type). Expected: keystrokes land in the floating window's field, not in the Excel cell that was just picked.

- [ ] **Step 5: Keyboard-focus round trip via a plain Excel click (no "Select Excel Cell")**

With the floating window open, click directly on any Excel cell (not using the "Select Excel Cell" button at all), then click back into a text field in the floating window and type. Expected: keystrokes land in the field.

- [ ] **Step 6: Full end-to-end Insert flow**

Fill out the floating window's fields for a valid balance formula (Ledger, Activity, Balance Type, Period, etc.), provide a cell reference, click Insert. Expected: the formula is written into the target cell exactly as the task pane would produce it (same `GLConfiguratorViewModel`/`WriteFormulaToCell` code path, since `GLBalanceConfigurator` itself is unmodified). Click Cancel on a fresh instance — expected: the window closes with no formula written.

- [ ] **Step 7: Reopen and formula-cell auto-refresh**

Close the floating window, reopen it via the ribbon button while the active cell already has a balance formula. Expected: it loads that formula's settings on open (`GLBalanceConfigurator.OnLoaded` already calls `ReLoadConfigurator()` unconditionally, `Views\GLBalanceConfigurator.xaml.cs:145-179`). Then, with it still open, select a different formula cell — expected: it reloads live (Task 4).

- [ ] **Step 8: Coexistence smoke test**

Open the docked task pane (existing "Balance" button) and the floating window (new "Balance (Window)" button) at the same time. Expected: no crash, no cross-talk — they are two independent instances of `GLBalanceConfigurator`, each with its own `GLConfiguratorViewModel`, so editing one must not affect the other's currently-displayed fields (though both will react to the same `SheetSelectionChange`/formula-cell events independently, which is expected/correct).

- [ ] **Step 9: Record the outcome**

Note any failures found in steps 2-8 back to the person who requested this feature before deciding whether to proceed toward replacing the task pane — that decision is explicitly out of scope for this plan (per the original request: "Once testing with the native window works then we can decide").
