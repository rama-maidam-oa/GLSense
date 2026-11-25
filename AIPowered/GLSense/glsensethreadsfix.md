# GLSense threading audit: handoff (stopped 23-Nov-2026 ~20:45 for a PC restart; previous stop 20-Nov 18:55)

When this file is passed to Claude: read it, then continue at **"Next step"**. Don't redo finished work.
Read **"Checkpoint 24-Nov-2026 evening"** first: it overrides older state below where they differ.

## Checkpoint 25-Nov-2026: moved to 11.1.3 (overrides everything below)
- At the user's request, all of `threading-audit` was brought to `11.1.3` (code + this file + fix log), then the
  `threading-audit` branch was deleted. Work continues on `11.1.3` (worktree `D:\SQLLite_Test\worktrees\GLSense-11.1.3`).
- Fix-log numbers on 11.1.3: 128 = batch 1, 130 = K1/K2, **136** = C1 busy guard (was 133), **137** = E1 (was 134),
  **138** = E7 (was 135), **139** = L1/L3/L10 (was 136). 133-135 on 11.1.3 are OISR-23017/OISR-22854.
- **Batches 4-6 (E1, E7, L1/L3/L10) are still NOT VM-verified** - they are now on the delivered branch; VM-test them
  next (plan in "Checkpoint 24-Nov-2026 evening" below). The staged VM build (24-Nov 09:07) is older: rebuild + re-stage.

## Checkpoint 24-Nov-2026 evening (overrides everything below)
- **Plan changed by the user:** code every batch first, VM-test them afterwards one by one.
- **Done (all build-verified, Debug unsigned; batches 4-6 NOT VM-verified):**
  - `c93e1aa` merge of `11.1.3` into `threading-audit`: gzip refresh (OISR-22975, §132), installer signing (§131),
    OISR-22998 (§129). The busy guard entry was renumbered §131 -> **§133**.
  - `474b022` batch 4, E1 logged-out double-click (§134, host, needs an MSI).
  - `f3da269` batch 5, E7 batch cache (§135, core).
  - `abed826` batch 6, L1/L3/L10 load hardening (§136, needs an MSI).
- **VM test order when the VM is free** (the user's FWC 11.1.1 test is using it now):
  1. Rebuild Debug, stage to `D:\VM-Stage\GLSense-Debug`, check the staged DLLs are this build.
  2. `Threading-LedgerRefresh` run: covers E1 (step 5), gzip Sheet/Book Refresh, busy guard, Insert LOVs.
  3. E7: needs a test (refresh with a formula the batch can't answer, then check the cache keeps the old value).
  4. L1/L3/L10: L1 by staging a corrupt zip in `AddinCore\Manifest\`; L3/L10 only by forced failure (or a normal
     load proving nothing broke).
  5. Gzip needs the matching server change on QA02 - check before blaming the client.
- **Still open:** L8 (user will discuss later, keep as is); merge into 11.1.3 after the VM pass. C3 closed: by design (user, 25-Nov-2026).

## Checkpoint 24-Nov-2026 (overrides 23-Nov below)
- **Batch 2 (K1/K2) PASSED** in the VM (run 22, 24-Nov): 2x Insert LOVs, run 2 reused `1_LOV`; no errors;
  Excel exited 9 s after Quit. Fix log §130 written and committed. Run 21 failed only on a harness bug
  (`Worksheets.Add()` puts the LOV sheet first) - fixed in `Threading-LedgerRefresh.ps1` (`$lovSheet`).
- ExcelSmoke harness fixes (23-Nov list + `$lovSheet`) committed on ExcelSmoke `master`.
- **Next:** ask the user before batch 3 (C1 busy guard, still uncommitted in `AddinEntry.cs`/`GLExpandOptions.xaml.cs`),
  and before merging `threading-audit` into 11.1.3 (see "Merge into 11.1.3").

## Checkpoint 23-Nov-2026 (PC was slow for every app; user restarted it)
- **Batch 2 (K1/K2, `d22ab78`) is STILL NOT VM-verified.** Runs 18-20 below never reached its LOV check.
- **Uncommitted in GLSense, branch `threading-audit` (keep, don't revert):** batch 3 = C1 busy guard,
  `GLSense.Addin.Core/AddinEntry.cs` (`RunExclusive` + 27 call sites) and `Views/GLExpandOptions.xaml.cs`.
  Not built into the staged VM build, not VM-tested. Commit only after batch 2 passes and the user says go.
- **Uncommitted in ExcelSmoke (`master`), harness fixes - commit them once a run passes:**
  - `SmokeLib.ps1`: `Start-ExcelWindow` waits 300 s (was 90); Close-OfficePromo also ends UserOOBEBroker/
    WebExperienceHost(App)/ApplicationFrameHost when in front; Assert-ExcelInFront prints pid + name.
  - `Step4-LoginLook.ps1`: waits up to 45 s for the login window. `Step4b-PickInstance.ps1`: pick retried.
  - `Tests/Threading-LedgerRefresh.ps1`: LOV window closed via `WindowPattern.Close()` + `Wait-NoWindow`;
    validation type read with `[System.__ComObject].InvokeMember('Type', GetProperty, ...)` because
    PowerShell's COM binder returned $null for `.Validation.Type` (run 18).
  - `Host/Run-VMTest.ps1`: `-WaitMinutes 20` (boot + auto-logon limit, was 10); `finally` turns the VM off.
- **11.1.3 has `94d3506` (OISR-22998, §129)**, committed and pushed; Jira comment posted. Also ported
  (uncommitted) to branch `ribbon-compact` for that session. Reverted on `threading-audit`.
- **User rules added 23-Nov:**
  - After EVERY VM run (pass or fail): report the key result lines, then delete that run's folder so
    `D:\VM-Out` ends empty (rm -rf from bash; Remove-Item on D:\VM-Out itself is blocked).
  - When a batch completes, ASK the user before starting the next batch.
- **VM runs on 23-Nov:**
  | Run | Result |
  |---|---|
  | 18 | Logged in; cube "102 EBS to Blaze" (1 ledger only, so no real ledger switch); Sheet + Book Refresh OK; LOV run 1 "LOV copied successfully" to D5 (dropdown visible), but the harness read of the validation type returned null -> InvokeMember fix above. No ERROR/Exception; Excel exited 65 s after Quit. |
  | 19 (x3) | Two attempts: VM not reachable in 10 min -> wait raised to 20. Third: VM up after ~80 min; Install passed; Step1 failed "No usable Excel workbook window within 90s"; SmokeRun then hung until stopped. |
  | 20 | VM logon 10 min; first step 20 min after task start; SmokeRun killed mid-Install-Build at 20:38 (exit 0xC000013A = terminated from outside). |
  The whole PC was slow; runs 19-20 are environment failures, not product results.
- **Next after the restart:** check the PC/VM are responsive (host free memory/disk, VM boots in a few
  minutes), then rerun the batch-2 test (step 2 of "Next step" below; no re-stage needed). If it fails
  again for environment reasons, ask the user (options: fix the VM, or test K1 by hand in Excel - a second
  "Insert LOVs" into an existing `<ledgerId>_LOV` sheet).
Read CLAUDE.md §95 and §128 first. Rules from the user:
- **No regressions.** Test every batch in the VM before committing it.
- Commit only on branch `threading-audit` (GLSense repo `D:\SQLLite_Test\GLSense`). Never push.
- Commit messages: no Claude/AI attribution.
- Commit the ExcelSmoke test changes on ExcelSmoke `master`.
- Every fix gets a CLAUDE.md fix-log entry that includes a "Same-issue sweep:" line.

## State at the stop
- GLSense repo, branch `threading-audit`:
  - Batch 1 (R1 + C2) is committed as `324cbb6`, with fix log §128.
  - Batch 2 (K1 + K2) is **committed but not VM-verified yet**. The commit holds `CommonFunctions.cs`, `CommonMethods.cs` and this file, and was pushed to `origin/threading-audit` before the restart. Fix log §130 isn't written yet (§129 is OISR-22998 on 11.1.3); add it once the VM run passes.
- ExcelSmoke repo, `master` (local only, no remote). Two commits before the restart:
  - One holds `Tests/Threading-LedgerRefresh.ps1` (Insert LOVs section, run twice, checks validation type 3; tests K1) and `Step4b-PickInstance.ps1` (retries the QA02 pick up to 3 times, checking that the URL box fills; run 10 failed on this flake).
  - The other is a separate commit of `Register-DebugAddin.ps1`'s `-Dll` parameter, from an earlier session.
- Debug build `v11.1.3_2026-11-20T15-31-10` is staged in `D:\VM-Stage\GLSense-Debug`. It contains batch 2.

## What batch 2 changes (K1 + K2)
- **K1, `CommonFunctions.SheetExistsAsync`:** stop calling `FinalReleaseComObject` on each sheet in the loop.
  - .NET keeps one COM wrapper per sheet per AppDomain, so releasing it here detaches the same sheet a running refresh or drilldown holds (§34).
  - The caller, `GLLOVs.CreateLOVSheetAsync`, then re-fetches the sheet by name and gets the detached wrapper. A second Insert LOVs for the same ledger likely failed.
- **K2, `CommonMethods.CanAccessWorksheet`:** stop calling `FinalReleaseComObject` on `ActiveWorkbook`.
  - A missing sheet is now caught as a `COMException` and makes the method return false. Its only user is the `Get_GLSense_MultiFormulas` diagnostic, so a missing sheet now skips that diagnostic instead of aborting the refresh. This behavior change is intended.

## VM runs for batch 2: none has tested it yet
| Run | Folder | What happened |
|---|---|---|
| 10 | `D:\VM-Out\run-20261120-153211` | The QA02 pick in the login window didn't stick, so the login page never loaded. The open login window then stopped Excel from quitting. Fixed in `Step4b-PickInstance.ps1` (retry). |
| 11 | `run-20261120-181657` | The VM didn't become reachable within 10 minutes. |
| 12 | `run-20261120-183613` | Same: the VM didn't become reachable. Hyper-V says ExcelTest is "Running / Operating normally". The restart is expected to fix this. |

## Next step (start here after the restart)
1. Confirm the Hyper-V VM `ExcelTest` comes up. The test run restores the 'Clean - Automation' checkpoint itself.
2. Rerun the batch-2 VM test. The build is already staged; only re-stage it if `D:\VM-Stage\GLSense-Debug` is gone. Run it in the background:
   `cd D:\SQLLite_Test\ExcelSmoke; .\Host\Run-VMTest.ps1 -CopyIn D:\VM-Stage\GLSense-Debug -Tests Threading-LedgerRefresh`
   To re-stage first: `.\Host\Stage-Build.ps1 -Product GLSense`. That needs a Debug build: MSBuild `GLSense.Build\GLSense.Build.csproj -p:Configuration=Debug -p:RegisterForComInterop=false`.
3. Read the newest `D:\VM-Out\run-*`:
   - `results.txt`: PASS/FAIL, plus the "LOV run 1/2" lines.
   - `screens\thr-lov-1.png` and `thr-lov-2.png`.
   - `Excel_Logs\...\GLSense_Logs_*.log`: check for ERROR or Exception, and for "separated" (a detached COM wrapper).
   - Close-Excel's result: Excel must exit after Quit.
4. If a step fails because of the test itself (UI timing, prompts), fix the script and rerun. If the product fails, analyze it and tell the user.
5. On PASS:
   - Add fix-log entry **§130** to CLAUDE.md. Cover K1/K2 and why, the VM verification, and the sweep: other `SafeRelease` calls (`rngFind`, `cell`) release objects the code created itself, so they're fine.
   - CLAUDE.md is CRLF. Keep it CRLF; don't use a Python rewrite.
   - Commit CLAUDE.md §130 on `threading-audit`. The code is already committed. Push only if the user asks.
   - If the VM finds a product bug in batch 2, fix it in a new commit. Don't rewrite the pushed commit.
   - Report to the user, then start batch 3.

## Merge into 11.1.3 (decided 23-Nov-2026)
11.1.3 is the branch delivered to Orbit. Once batch 2 passes the VM test, merge `threading-audit` into `11.1.3`
as ONE merge commit (not cherry-picks), in a temporary 11.1.3 worktree (11.1.3 isn't checked out anywhere).
- 11.1.3 already has `94d3506` (OISR-22998, fix log §129) that `threading-audit` lacks.
- CLAUDE.md tail conflicts: keep §128 = threading batch 1, §129 = OISR-22998, §130 = K1/K2.
- Push only when the user says so. Later batches follow the same rule: VM-tested first, then merged.

## Remaining batches (agreed order; each one small, VM-tested, then committed)
3. **C1 busy guard:** `AddinEntry.OnRibbonAction`. Drilldowns, refresh, discover and reset are fire-and-forget. A second click overwrites static window, cancel-token and sheet fields.
   - Fix: one `Interlocked` "operation running" flag, which also covers K5.
   - Decide which actions it covers. Long operations only; not windows that open modally.
4. **E1 double-click login check:** host `AddinModule.cs:~1606`.
   - When logged out, double-clicking any formula starts the precedents drilldown and blocks in-cell editing. FinalWorkingCode checks login here.
   - Fix: `if (!IsLoggedIn) return;`. This is a **host change and needs an MSI**, so tell the user.
5. **E7:** `UdfDispatcher.cs:~567`, `HandleBatchCalculation`. A batch refresh caches "Click Refresh..." over the last good value. Fix: don't cache that result.
6. **Loading hardening:**
   - **L1** `UpdateBootstrapper.cs:~133`: a bad zip in `Manifest\` stops loading even when a good release is installed. Log it and fall through to the release-history lookup.
   - **L3** `AddinModule.OnRibbonLoaded` (host): no try/catch, so Office may disable the add-in. Wrap it.
   - **L10** `RemoteLoader.cs:~46`: replace `asm.GetTypes()` with `asm.GetType("GLSense.Addin.Core.AddinEntry", true)`.
   - L1 and L10 are Loader.Core; L3 is host. Host changes need an MSI.
- **C3, CLOSED - by design (user, 25-Nov-2026), no change:** Range Refresh deliberately turns Live Calc on; it is only triggered from the Balance Configurator. `RangeRefresher.cs:~136`. `SingleRefresh` stays true after a Range Refresh, so Live Calc stays on while the toggle shows off. FinalWorkingCode does the same. Resetting it immediately may be too early.
- **L8, OPEN - user will discuss later, keep as is:** audit item **L8** (DPI context in `WpfAppManager`) is a **false positive**. The code comment says the setting is intentional.
- **Later, lower priority (from the audit, unverified):** L2, L4, L5, L6, L7, L9, R2, R3, C4–C11, E2–E6, K3–K6, and the speed items. The full table is in the original transcript:
  `C:\Users\RamaMaidam_2fc\.claude\projects\D--SQLLite-Test-GLSense-AIPowered-GLSense\cf449579-1018-4509-94de-d5b9d3c9c3c0.jsonl`, line 505. Search for "I found 33 gaps".

## Still open from before (ask the user first)
- OISR-22872: the Jira comment and the MSI release build.

## VM notes
- Laptop scripts are in `D:\SQLLite_Test\ExcelSmoke\Host`. Results land in `D:\VM-Out\run-*`.
- Account: user smoke, server QA02. Never print the `GLSENSE_TEST_*` variables.
- Known flakes: an empty instance list (now retried), the Copilot and SCOOBE prompts (handled in SmokeLib/Run-All), and the "Clear sheet?" prompt (exact text).
