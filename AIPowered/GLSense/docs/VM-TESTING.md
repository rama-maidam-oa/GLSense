# Testing GLSense and XLEdge (AIPowered) on a test VM

Use the VM to test what ships (the MSI or a DLL-only release) on a clean machine. Keep Debug builds and fast
iteration on the dev machine. Releases themselves: see `DEPLOYMENT.md`.

## One-time VM setup

1. Windows, Office 365 (signed in, activated), WebView2 Runtime, and network or VPN access to the test servers
   (QA01, development). Check that the login page opens in a browser on the VM.
2. Install Claude Code on the VM and copy the whole `D:\SQLLite_Test\ExcelSmoke` folder to it, in any location.
   The scripts find `SmokeLib.ps1` next to themselves. Screenshots go to `.\out`, or to `$env:SMOKE_OUT` if it's
   set (point it at the shared folder).
3. For the login script, save the test accounts once as user environment variables, numbered per account:
   `setx GLSENSE_TEST_USER_1 '<user>'` and `setx GLSENSE_TEST_PASSWORD_1 '<password>'`, then `_2`, `_3`, and so
   on. Single quotes keep `$` and `&` in passwords literal. Open a new terminal afterwards. Pick the account with
   `Step4c-Login.ps1 -Account 2`. Without `-Account`, the script reads the unnumbered `GLSENSE_TEST_USER` and
   `GLSENSE_TEST_PASSWORD`. Never put credentials in a script. They're stored as plain text in the VM's user
   profile and in its snapshots, so use test accounts only.
4. Create a shared folder between the host and the VM for builds going in and logs/screenshots coming out.
5. Take a snapshot now: **"Clean - Office only"**. Revert to it before every MSI test, so every install is a real
   first install.
6. Optional: a second snapshot or VM with **32-bit** Office (see CLAUDE.md §108), and one at 125% or 150% display
   scaling.

## Running the tests

- Claude Code has to run **inside** the VM. UI Automation only works on an unlocked desktop on the same machine as
  Excel; PowerShell Remoting from the host has no desktop.
- Keep the VM window open and not minimized while scripts run. A minimized RDP session stops drawing and UI
  Automation fails. Don't touch the mouse during a run: the scripts click for real.
- `Register-DebugAddin.ps1` is for the dev machine only: it needs Add-in Express installed. On the VM, install
  the MSI.

## Which path for which fix

| The fix changes | On the VM |
|---|---|
| Core only (`GLSense.Addin.Core`, a DLL-only release) | Install the current MSI once. Then **Reload > Offline**, pointing at a folder with the new `manifest.json` + zip (copied from `GLSense.Addin.Core\SetupFiles\Release\Manifest\`). Script: `Step3-ReloadOffline.ps1 -Folder <that folder> -Tag <name>`. |
| Host (`GLSense\...`: ribbon, task pane, Reload/Version History windows, shutdown) | Revert the snapshot, install the new MSI. |
| Already deployed to development | Log in (`Step4b-PickInstance.ps1`, `Step4c-Login.ps1 -Account <n>`), then **Reload > Online** (`Step5-ReloadOnline.ps1`). |

To save signing operations, test core-only fixes with Reload > Offline. Only build a new MSI when the host changed.

## Per-fix checklist

1. Note the OISR, the build commit, and the path from the table above.
2. Revert the snapshot if you're installing an MSI.
3. Reproduce the bug report's steps. Turn Debug on (ribbon) before the steps when you need the full trace.
4. Copy back:
   - GLSense log: `%LOCALAPPDATA%\ORBIT\Excel_Logs\GLSense_Logs\Logs\`
   - XLEdge log: `%LOCALAPPDATA%\ORBIT\Excel_Logs\XLEdge_Logs\` (`XLEdge_Logs_<date>.log`)
   - screenshots (`out\`, or `SMOKE_OUT`).
5. Record the result in the fix's CLAUDE.md entry ("Tested on VM: ...").

## What only the VM can show (run these at least once per MSI release)

- **First install:** the add-in loads, and Version History has one "Install" row. Close and reopen Excel: it loads
  again (§40.2).
- **Upgrade:** install the previous MSI, use it once, install the new one. Version History rows are kept (§62).
- **Uninstall:** Add/Remove Programs. `%LOCALAPPDATA%\Orbit Analytics\GLSense` is gone (§62).
- **Excel exit:** close Excel; `EXCEL.EXE` leaves Task Manager within a few seconds, also after a drilldown
  (§71, §95). Compare with the dev machine, which has Bitdefender.
- **32-bit Office:** the Balance Configurator pane shows its controls (§108).
- **Scaling:** windows open centered over Excel and aren't cut off at 125%/150%.
- **Both add-ins installed:** XLEdge login/logout follows into GLSense (§91).

## Scripts available

`Step1-OpenTab` (GLSense ribbon), `Step2-VersionHistory` / `Step2b-VersionHistoryTips`, `Step3-ReloadOffline`,
`Step4-LoginLook` / `Step4b-PickInstance` / `Step4c-Login`, `Step5-ReloadOnline`. They drive GLSense only; for
XLEdge, use `SmokeLib.ps1`'s helpers (`Start-ExcelWindow`, `Find-Element`, `Click-Element`, `Save-Screenshot`)
in a new step script.
