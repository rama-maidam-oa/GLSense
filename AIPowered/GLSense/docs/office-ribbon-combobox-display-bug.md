# Excel ribbon combo box doesn't display text set from code (Office 16.0.20527.20088)

**Status:** open. Being reported to Add-in Express, and to Microsoft through IT. The
workaround is in AIPowered 11.1.2 and verified (see "Workaround" below).
**First seen:** 26-Oct-2026. **Office build:** 16.0.20527.20088, Microsoft 365 Apps,
Current Channel, 64-bit, installed 21-Oct-2026.

## Impact

GLSense's ribbon has two combo boxes: **Ledger** (`Ribledger`) and **Segment**
(`RibSegS`). After a user picks a cube and clicks OK, the add-in fills the Ledger box and
selects the first ledger as the default. On the affected Office build the box stays
blank, although the ledger *is* selected internally and everything that uses it still
works. The user can't see which ledger is active until they open the dropdown.

This isn't a GLSense regression:
- Every GLSense build shows it, including FinalWorkingCode 11.1.1 and an older
  released MSI that worked before the Office update.
- A minimal Add-in Express sample with no GLSense code reproduces it
  (`ADXComboRepro`, see "Repro").

Every customer whose Office updates to this build (or later, until Microsoft fixes it)
should be expected to see it. The same applies to any Office add-in that sets a ribbon
combo box's text from code.

## Behaviour

In a new Excel session, a ribbon combo box ignores text set from code until the user has
opened its dropdown once by hand. It doesn't matter whether the list had items when it
was opened. After that first open, text set from code is displayed every time for the
rest of the session.

## Evidence

All gathered on Excel 16.0.20527.20088, Add-in Express 10.1.4703.0:

- **Excel receives the right value.** Logging the control's `PropertyChanging` event
  shows Excel asking for `Text` and receiving the correct ledger name. Nothing is drawn.
- **Excel doesn't load the item list until the dropdown is opened.** There are no
  `ItemCount` / item requests for the combo box before that.
- **What didn't fix it:**
  - refreshing the box by its correct RibbonX id (`IRibbonUI.InvalidateControl`), or
    the whole ribbon (`IRibbonUI.Invalidate`);
  - hiding and re-showing the box;
  - running every update on Excel's main thread;
  - giving every item a unique id;
  - a fresh `Excel15.xlb`.
- **A drop-down doesn't help either.** With `ADXRibbonDropDown`, Excel requests the items
  but never requests `SelectedItemIndex` / `SelectedItemID`, so the selection set from
  code is never shown.
- **Opening the dropdown once from code works.** Using UI Automation
  (`ExpandCollapsePattern.Expand`, then `Collapse`) has the same effect as the user
  opening it, and from then on the text set from code is displayed.

## Repro

`D:\SQLLite_Test\ADXComboRepro`. A one-file Add-in Express module with a combo box, a
drop-down, and Load / Select next / Clear / Prime buttons. Every value Add-in Express
hands to Excel is logged to `%TEMP%\ComboRepro.log`. The full source is also in
`ComboRepro_SourceCode.txt`, for sending past mail filters that block zip files.

1. Fresh Excel, ComboRepro tab, dropdown never opened: click **Load items + select
   first**. The combo box stays blank.
2. Open and close the dropdown once, then click Load again. The text is displayed.
3. Restart Excel, click **Prime (open + close combo)**, then Load. The text is
   displayed.

## Workaround

Prime each combo box once per Excel window by opening and closing it through UI
Automation, before GLSense sets its text. Constraints found while building the repro:

- **Run UI Automation on a worker thread.** Opening the dropdown needs Excel's main
  thread to respond, so calling it from that thread deadlocks. Hand the text update
  back to the main thread afterwards.
- **The tab must be showing.** UI Automation only sees ribbon controls on the selected
  tab.
- **Expect a flicker.** The dropdown visibly flashes open and closed once.

### GLSense implementation (AIPowered 11.1.2, verified 26-Oct-2026)

- `GLSense.Addin.Core\Helpers\RibbonComboPrimer.cs`:
  - `PrimeOnce(excelHwnd, accessibleName, dispatcher, then)` finds the ribbon combo by
    its accessible name, opens it, waits 150 ms, and closes it, on a worker thread.
  - It then runs `then` on Excel's main thread.
  - It tracks the window handles already primed, so each Excel window is primed once.
  - If the combo isn't found or anything fails, it logs a warning and allows a retry in
    that window. It never throws.
- `AddinEntry.ShowCubeDetails` calls `PrimeLedgerComboIfNeeded()` right after the Cube
  dialog closes, when a ledger is selected. The user is on the Orbit GLSense tab at that
  point, so no tab switching is needed. The callback clears the Ledger text, sets the
  selected ledger again, and refreshes the box, so Excel redraws it.
- Only **Ledger** (`Ribledger`) is primed. It's the only ribbon combo where GLSense sets
  a default from code. Segment (`RibSegS`) starts blank.
- The code lives in Addin.Core, so it ships through hot reload, with no host or MSI
  change.

**Verified on Office 16.0.20527.20088:**
- In a fresh Excel window, the first cube selection shows the ledger. The dropdown
  flashes once.
- The next cube selection in the same window shows its ledger with no flash.

**Expected but not yet tested:**
- A new workbook window (it has its own ribbon) primes once again on its first cube
  selection.
- After a hot reload, the window primes once again, because the list of primed windows
  starts empty.

**Rejected alternative:** priming at ribbon load. At that point Excel is on the Home
tab, where UI Automation can't see GLSense's controls. Doing it there would mean
switching tabs at every Excel start, and it still wouldn't cover workbook windows
opened later. Opening the box while it is disabled (before login) does still prime it.

## Reporting

- **Add-in Express:** email with the repro steps, the logged callbacks and the
  environment, plus `ComboRepro_SourceCode.txt` and `ComboRepro.log`.
- **Microsoft:** a Microsoft 365 support case raised by IT from the admin center. Also
  ask IT whether the update can be held or rolled back for affected users.
- Remove the workaround once a fixed Office build is confirmed.
