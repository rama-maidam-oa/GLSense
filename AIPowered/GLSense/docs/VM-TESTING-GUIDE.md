# Testing a fix in the ExcelTest VM: step by step

This is the hands-on guide. For which test path to use and what to check, see `VM-TESTING.md`.

The setup in short:

| Where | What |
|---|---|
| Laptop | Code, builds, git, Visual Studio. You **build** here. |
| VM `ExcelTest` | Office 365, Claude Code, the smoke scripts in `C:\ExcelSmoke`. You **test** here. |
| Checkpoint `Clean - Office only` | Windows + Office + WebView2 + Claude Code + test accounts, no GLSense or XLEdge installed. |

Commands marked **Laptop (admin)** run in an *Administrator* PowerShell on the laptop. Commands marked **VM** run in
PowerShell or in Claude Code inside the VM.

---

## 1. Start a test session

1. **Laptop (admin)**, only if you're testing an MSI: go back to the clean state:
   ```powershell
   Restore-VMCheckpoint -VMName ExcelTest -Name 'Clean - Office only' -Confirm:$false
   ```
   This throws away everything done in the VM since the checkpoint, so copy out any logs you still need first.
   Skip this step for a core-only (Reload) test on an already installed add-in.
2. Open **Hyper-V Manager** and double-click `ExcelTest`. Start it if it's off.
3. In the connect dialog, click **Show Options > Local Resources > More...**, check drive **D:**, and click
   **Connect**. The laptop's D: drive is then `\\tsclient\D` inside the VM. If no dialog appears, use
   **View > Enhanced session**.
4. Sign in to Windows in the VM.
5. **VM**: open Windows Terminal and start Claude Code:
   ```powershell
   cd C:\ExcelSmoke
   claude
   ```
   It starts without asking you to sign in (the token is already set).

Keep the VM window open and not minimized while tests run. Don't move the mouse in the VM during a scripted run: the
scripts click for real.

---

## 2. Move the latest code into the VM

You never copy source code into the VM. You build on the laptop and copy the **build output**. What you copy depends on
what the fix changed.

### 2a. Core-only fix (most fixes): manifest + zip, then Reload > Offline

Use this when the fix is only in `GLSense.Addin.Core` (or `XLEdge.Addin.Core`). The add-in must already be installed in
the VM.

1. **Laptop**: build the core project in **Debug**. Debug creates the same zip and manifest as Release but uses no
   signing operations.
   - GLSense output: `D:\SQLLite_Test\GLSense\AIPowered\GLSense\GLSense.Addin.Core\SetupFiles\Debug\Manifest\`
   - XLEdge output: `D:\SQLLite_Test\XLEdge\BusinessLayers\11.1.3\AIPowered\XLEdge\XLEdge.Addin.Core\SetupFiles\Debug\Manifest\`

   Each holds exactly one `manifest.json` and one `v<version>_<date>.zip`.
2. Copy that folder into the VM, into a **new folder per build** (Offline mode needs exactly one manifest + zip pair in
   the folder). Either way works:

   **Option A, from the VM (drag and drop):** open `\\tsclient\D\SQLLite_Test\GLSense\AIPowered\GLSense\GLSense.Addin.Core\SetupFiles\Debug\Manifest`
   in Explorer and copy both files to, for example, `C:\VM-In\OISR-22875-1\`.

   **Option B, from the laptop (admin):**
   ```powershell
   $src = 'D:\SQLLite_Test\GLSense\AIPowered\GLSense\GLSense.Addin.Core\SetupFiles\Debug\Manifest'
   $dst = 'C:\VM-In\OISR-22875-1'   # new folder per build
   Get-ChildItem $src -File | ForEach-Object {
       Copy-VMFile -Name ExcelTest -SourcePath $_.FullName -DestinationPath "$dst\$($_.Name)" -FileSource Host -CreateFullPath -Force
   }
   ```
3. **VM**, in Excel: **GLSense ribbon > Reload > Offline > Browse...**, pick `C:\VM-In\OISR-22875-1`, then **Reload**.
   Or let Claude do it (see section 3): `Step3-ReloadOffline.ps1 -Folder C:\VM-In\OISR-22875-1 -Tag OISR-22875`.

   Offline only accepts a release **newer** than the one loaded. Every build gets a fresh release date, so a new build
   is always accepted. To go back to an older build, use **Version History** instead.

### 2b. Host change: install the MSI

Use this when the fix touches the host project (`GLSense\...`: ribbon, task pane, Reload/Version History windows,
shutdown), or for any first-install/upgrade/uninstall test.

1. **Laptop**: build the setup project in **Release** (this signs the host DLLs).
   - GLSense: `D:\SQLLite_Test\GLSense\AIPowered\GLSense\GLSense\OrbitGLSenseSetups\Release\OrbitGLSense.msi`
   - XLEdge: `D:\SQLLite_Test\XLEdge\BusinessLayers\11.1.3\AIPowered\XLEdge\XLEdge\OrbitXLEdgeSetups\Release\OrbitXLEdge.msi`
2. **Laptop (admin)**: revert the checkpoint (section 1, step 1).
3. Copy the MSI in:
   ```powershell
   Copy-VMFile -Name ExcelTest -SourcePath 'D:\SQLLite_Test\GLSense\AIPowered\GLSense\GLSense\OrbitGLSenseSetups\Release\OrbitGLSense.msi' `
       -DestinationPath 'C:\VM-In\OrbitGLSense.msi' -FileSource Host -CreateFullPath -Force
   ```
   (or drag it from `\\tsclient\D\...` in the VM).
4. **VM**: close Excel, double-click `C:\VM-In\OrbitGLSense.msi`, then open Excel.

### 2c. Already deployed to the development server

Nothing to copy. In the VM, log in to the development server and use **Reload > Online**.

### 2d. Updated smoke scripts

If you change a script in `D:\SQLLite_Test\ExcelSmoke` on the laptop, copy it in again:
```powershell
Get-ChildItem D:\SQLLite_Test\ExcelSmoke -Filter *.ps1 | ForEach-Object {
    Copy-VMFile -Name ExcelTest -SourcePath $_.FullName -DestinationPath "C:\ExcelSmoke\$($_.Name)" -FileSource Host -Force
}
```

---

## 3. What to type in Claude Code (VM)

Claude Code in the VM knows nothing about the fix, the repo or Jira. Tell it the ticket, what changed, which build to
load and the steps to reproduce. It drives Excel with the scripts in `C:\ExcelSmoke` (`SmokeLib.ps1` has
`Start-ExcelWindow`, `Find-Element`, `Click-Element`, `Save-Screenshot`).

**First check that automation works:**
```
Run .\Step1-OpenTab.ps1 and tell me whether the GLSense ribbon tab opened. Show me the screenshot.
```

**Load a core build and log in:**
```
Load the build in C:\VM-In\OISR-22875-1 with .\Step3-ReloadOffline.ps1 -Folder C:\VM-In\OISR-22875-1 -Tag OISR-22875.
Then pick the QA01 instance with .\Step4b-PickInstance.ps1 -Instance QA01 and log in with .\Step4c-Login.ps1 -Account 1.
Tell me the loaded version from the log.
```

**Reproduce a bug and check the fix (write the steps out):**
```
Test OISR-22875. Bug: a segment value typed as 11$9 in the Account Assignment window reopened as a cell reference.
Expected after the fix: it reopens as a value.
Steps:
1. Open Book1, select A1, open the Balance Configurator from the GLSense ribbon.
2. Pick cube "101 standard view imp", ledger "Vision Operations (USA)".
3. In Account Assignment, type 11$9 for Account, click OK, Insert.
4. Close the configurator, select A1 again, reopen it.
Check: the Account value shows 11$9 and the Reference box is empty.
Use the SmokeLib helpers, take a screenshot at each step, and tell me pass or fail.
Afterwards show me the last 30 lines of the GLSense log.
```

**MSI checks (after installing the MSI):**
```
GLSense was just installed from the MSI. Check:
1. Excel opens and the GLSense ribbon tab is there.
2. Version History shows one "Install" row.
3. Close Excel. Report how many seconds EXCEL.EXE takes to leave Task Manager.
4. Reopen Excel and confirm the add-in loads again.
```

**Collect results:**
```
Copy today's GLSense log and the out\ screenshots to \\tsclient\D\VM-Out\OISR-22875\.
```

Tips:
- One ticket per Claude session. Start a new session (`/clear`) for the next ticket.
- If a script can't find a control, ask Claude to take a screenshot and list the window's controls before guessing.
- Ask for screenshots: you see exactly what Claude saw.

---

## 4. Get results back to the laptop

`Copy-VMFile` only copies laptop to VM. To get files out, use the redirected drive from inside the VM:

- Logs: `%LOCALAPPDATA%\ORBIT\Excel_Logs\GLSense_Logs\Logs\` and `%LOCALAPPDATA%\ORBIT\Excel_Logs\XLEdge_Logs\`
- Screenshots: `C:\ExcelSmoke\out\`

```powershell
$out = '\\tsclient\D\VM-Out\OISR-22875'
New-Item -ItemType Directory -Force $out | Out-Null
Copy-Item "$env:LOCALAPPDATA\ORBIT\Excel_Logs\GLSense_Logs\Logs\*.log" $out
Copy-Item C:\ExcelSmoke\out\* $out
```

To send screenshots straight to the laptop during a run, set this once in the VM (open a new terminal afterwards):
```powershell
setx SMOKE_OUT '\\tsclient\D\VM-Out'
```
This only works while you're connected with drive D: redirected.

---

## 5. Finish

1. Write the result in the fix's CLAUDE.md entry on the laptop ("Tested on VM: ...").
2. Close Excel in the VM, then shut down the VM (or leave it running for the next test).
3. Before the next MSI test, revert the checkpoint again (section 1, step 1).

---

## Quick reference

| I want to... | Do this |
|---|---|
| Start clean | Laptop (admin): `Restore-VMCheckpoint -VMName ExcelTest -Name 'Clean - Office only' -Confirm:$false` |
| Test a core-only fix | Build core in Debug, copy `SetupFiles\Debug\Manifest` to a new `C:\VM-In\...` folder, Reload > Offline |
| Test a host fix | Build the setup in Release, revert the checkpoint, copy and install the MSI |
| Use another test account | `.\Step4c-Login.ps1 -Account 2` |
| See the laptop's D: in the VM | `\\tsclient\D` (enhanced session, drive D: redirected) |
| Get logs out | Copy to `\\tsclient\D\VM-Out\...` from the VM |
| Claude asks to sign in | The token is missing: check `$env:CLAUDE_CODE_OAUTH_TOKEN` in a new terminal |

## Troubleshooting

| Problem | Fix |
|---|---|
| `\\tsclient\D` is missing | Reconnect: Hyper-V Manager > Connect > Show Options > Local Resources > More > D: > Connect. |
| Scripts fail to click | The VM window was minimized or the screen locked. Restore it, set screen/sleep to Never, rerun. |
| Reload > Offline says "no manifest.json + zip pair" | The folder has none or more than one pair. Use a new folder with exactly one of each. |
| Reload > Offline says no update | The build isn't newer than the loaded one. Rebuild, or use Version History. |
| `Copy-VMFile` fails | Run it in an Administrator PowerShell on the laptop; the VM must be running. |
| `claude` not recognized | Open a new terminal; if still missing, check `$env:USERPROFILE\.local\bin\claude.exe` exists. |
