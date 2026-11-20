# GLSense threading audit: handoff (stopped 20-Nov-2026 18:55 for a system restart)

When this file is passed to Claude: read it, then continue at **"Next step"**. Don't redo finished work.
Read CLAUDE.md §95 and §128 first. Rules from the user:
- **No regressions.** Test every batch in the VM before committing it.
- Commit only on branch `threading-audit` (GLSense repo `D:\SQLLite_Test\GLSense`). Never push.
- Commit messages: no Claude/AI attribution.
- Commit the ExcelSmoke test changes on ExcelSmoke `master`.
- Every fix gets a CLAUDE.md fix-log entry that includes a "Same-issue sweep:" line.

## State at the stop
- GLSense repo, branch `threading-audit`:
  - Batch 1 (R1 + C2) is committed as `324cbb6`, with fix log §128.
  - Batch 2 (K1 + K2) is **committed but not VM-verified yet**. The commit holds `CommonFunctions.cs`, `CommonMethods.cs` and this file, and was pushed to `origin/threading-audit` before the restart. Fix log §129 isn't written yet; add it once the VM run passes.
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
   - Add fix-log entry **§129** to CLAUDE.md. Cover K1/K2 and why, the VM verification, and the sweep: other `SafeRelease` calls (`rngFind`, `cell`) release objects the code created itself, so they're fine.
   - CLAUDE.md is CRLF. Keep it CRLF; don't use a Python rewrite.
   - Commit CLAUDE.md §129 on `threading-audit`. The code is already committed. Push only if the user asks.
   - If the VM finds a product bug in batch 2, fix it in a new commit. Don't rewrite the pushed commit.
   - Report to the user, then start batch 3.

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
- **C3, needs the user's decision before any change:** `RangeRefresher.cs:~136`. `SingleRefresh` stays true after a Range Refresh, so Live Calc stays on while the toggle shows off. FinalWorkingCode does the same. Resetting it immediately may be too early.
- **Tell the user:** audit item **L8** (DPI context in `WpfAppManager`) is a **false positive**. The code comment says the setting is intentional.
- **Later, lower priority (from the audit, unverified):** L2, L4, L5, L6, L7, L9, R2, R3, C4–C11, E2–E6, K3–K6, and the speed items. The full table is in the original transcript:
  `C:\Users\RamaMaidam_2fc\.claude\projects\D--SQLLite-Test-GLSense-AIPowered-GLSense\cf449579-1018-4509-94de-d5b9d3c9c3c0.jsonl`, line 505. Search for "I found 33 gaps".

## Still open from before (ask the user first)
- OISR-22872: the Jira comment and the MSI release build.

## VM notes
- Laptop scripts are in `D:\SQLLite_Test\ExcelSmoke\Host`. Results land in `D:\VM-Out\run-*`.
- Account: user smoke, server QA02. Never print the `GLSENSE_TEST_*` variables.
- Known flakes: an empty instance list (now retried), the Copilot and SCOOBE prompts (handled in SmokeLib/Run-All), and the "Clear sheet?" prompt (exact text).
