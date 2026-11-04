# Excel stays running after close (hot-reload / AppDomain add-ins)

Diagnosis and fix from GLSense AIPowered (11.1.3, 04-Nov-2026). It is written so another
session can apply it to an add-in with the same design, such as XLEdge's AIPowered port. GLSense
details: `CLAUDE.md` section 95 and `PORTING_GUIDE.md` section 1b.

## The architecture this applies to

- A thin **host** add-in (Add-in Express `ADXAddinModule`, ADX's own AppDomain) never reloads.
- The real logic lives in a **core** assembly loaded into a separate AppDomain. It is unloaded and
  reloaded for hot reload, and unloaded at Excel shutdown.
- The host calls the core through .NET Remoting (`MarshalByRefObject` contracts), always on
  Excel's main (STA) thread: ribbon clicks and Excel events. The core is started fire-and-forget:
  `_ = SomeAsyncFlow();`.

## Symptom

- After the user closes Excel, `EXCEL.EXE` stays in Task Manager with **no window, indefinitely**.
  Excel starts `splwow64.exe` a few seconds after the close; that is a red herring.
- Opening a new Excel while the old one lives can make the add-in not load.
- **The add-in's own shutdown completes normally:** core shutdown runs, the AppDomain unloads,
  and the log says teardown is complete. Nothing in the log points at the problem.
- **Only some features trigger it.** In GLSense, login, opening windows and writing values into
  cells were fine. Any drilldown that wrote a sheet triggered it, with Save or No.
- **The old single-AppDomain build (FinalWorkingCode) never had it.**

Not the cause (all checked):
- `splwow64.exe`, the save prompt, and other add-ins such as Adobe PDFMaker.
- Antivirus: Bitdefender adds a ~60-75 s delay on every Excel exit, add-in or not, but it ends.

## Root cause

**Excel is called from a ThreadPool (MTA) thread.** The host calls into the core on Excel's main
thread, but across the AppDomain boundary there is **no `SynchronizationContext`**. So every
async flow the core starts resumes after its first `await` on a ThreadPool thread, and keeps
using Excel from there. COM then creates cross-apartment proxies to Excel in the process MTA.
The MTA never shuts down, so those references are never released, and Excel waits for them as an
automation server after its window closes.

The single-AppDomain build runs these flows with a `SynchronizationContext` on Excel's main
thread, so its awaits resume there. That's why it never stuck.

**Related problem found during the regression pass:** the core used the **host's** Excel object
across the AppDomain boundary (`context.ExcelApp`). Excel objects returned through it are
remoting copies. When the core's AppDomain already holds a generic wrapper for the same COM
object, the copy can't be cast and fails with
`InvalidCastException: Unable to cast __ComObject to WorkbookClass`. Seen on
`Application.ActiveWorkbook`, especially after a hot reload.

## Fix (three parts)

### 1. Install a main-thread SynchronizationContext at every host -> core entry point

Make sure a WPF `Application` (and so a `Dispatcher`) is created once on Excel's main thread, in
the core's `Initialize`. Then, at the top of every entry point that can start async work:

```csharp
// Called first thing in OnRibbonAction / OnExcelEvent (and any other host -> core entry point
// that starts async work). The host calls these on Excel's main thread.
private static void EnsureExcelSynchronizationContext()
{
    try
    {
        var dispatcher = System.Windows.Application.Current?.Dispatcher; // created on Excel's main thread
        if (dispatcher == null || dispatcher.HasShutdownStarted || !dispatcher.CheckAccess())
            return;
        if (SynchronizationContext.Current is System.Windows.Threading.DispatcherSynchronizationContext)
            return;
        SynchronizationContext.SetSynchronizationContext(
            new System.Windows.Threading.DispatcherSynchronizationContext(dispatcher));
    }
    catch (Exception ex) { /* log a warning, never throw across the boundary */ }
}
```

Before adding it, check every `.Wait()`, `.Result` and `GetAwaiter().GetResult()` in the core
for deadlocks. A blocking wait on Excel's main thread, for a task whose continuation now needs
that thread, deadlocks. In GLSense all of them waited on `Task.Run` bodies (no captured context),
or ran on a dedicated STA thread, so none were affected.

Trade-off: work after an `await` now runs on Excel's UI thread, as in the single-AppDomain
build. Move heavy pure-.NET work, such as JSON parsing and building cell arrays, into
`Task.Run`, with no Excel objects inside.

### 2. Safety net for code that still runs off the main thread

```csharp
internal static void RunOnExcelThread(Action action)
{
    var dispatcher = System.Windows.Application.Current?.Dispatcher;
    if (dispatcher != null && !dispatcher.CheckAccess() && !dispatcher.HasShutdownStarted)
        dispatcher.Invoke(action);   // synchronous Invoke - never `await InvokeAsync`
    else
        action();
}
```

Use it for Excel calls made from shared helpers that can be reached from anywhere, such as
"disable/enable ScreenUpdating, DisplayAlerts, EnableEvents" around batch operations.

### 3. Give the core its own Excel object instead of the host's

Contract (host -> core):
```csharp
// IGLSenseContext (contract)
IntPtr AcquireExcelUnknown();   // AddRef'd IUnknown; the caller must Marshal.Release it

// Host implementation (runs in the host AppDomain)
public IntPtr AcquireExcelUnknown()
    => ExcelApp == null ? IntPtr.Zero : Marshal.GetIUnknownForObject(ExcelApp);
```

Core: create it once in `Initialize`, on Excel's main thread. Fall back to the host's object if
the host is older and doesn't have the method; this matters when an online update can load a
newer core into an older host.
```csharp
IntPtr unknown = IntPtr.Zero;
try
{
    unknown = context.AcquireExcelUnknown();
    _localExcelApp = (Excel.Application)Marshal.GetObjectForIUnknown(unknown);
}
catch (Exception ex) { /* log a warning, keep using context.ExcelApp */ }
finally { if (unknown != IntPtr.Zero) Marshal.Release(unknown); }

public static Excel.Application ExcelApp => _localExcelApp ?? context.ExcelApp as Excel.Application;
```

Release it in the core's shutdown/reset, which runs before the AppDomain unloads, both at Excel
close and at hot reload:
```csharp
try { if (_localExcelApp != null) Marshal.FinalReleaseComObject(_localExcelApp); } catch { }
_localExcelApp = null;
GC.Collect(); GC.WaitForPendingFinalizers(); GC.Collect(); GC.WaitForPendingFinalizers();
```

Never pass Excel objects across the host/core boundary. Pass primitives instead, such as the
sheet name or a range address string.

### Also fixed in the same pass: ribbon state after hot reload

The ribbon belongs to the host and survives a reload; the reloaded core starts logged out. If
the core's `Initialize` only switches the ribbon to "logged out", the old cube/ledger captions and
combo items stay visible. Clear them in `Initialize` exactly as Logout does.

## How it was diagnosed (reuse when it happens again)

1. **Isolate by feature.** Run open → one feature → close, one feature at a time, and note which
   ones leave `EXCEL.EXE` behind. Compare with the single-AppDomain build. A tiny PowerShell
   watcher that polls `Get-Process EXCEL` (window handle and exit time) saves a lot of guessing.
2. **Managed view: ClrMD.** Use `Microsoft.Diagnostics.Runtime` in a scratch console app; attach
   with `DataTarget.AttachToProcess(pid, suspend: true)`. List AppDomains, live RCWs
   (`obj.HasRuntimeCallableWrapper`) with `GCRoot` paths, and `System.Runtime.Remoting.ServerIdentity`
   objects. Here it showed that the core AppDomain was unloaded and only Add-in Express's own two
   wrappers were left, as in the working build. So the extra reference was **not** a .NET object.
3. **Native view: cdb + SOS.** Install with `winget install Microsoft.WinDbg`; `cdb.exe` is in the
   package's `amd64` folder. Attach non-invasively:
   `cdb -pv -p <pid> -c ".loadby sos clr; ~0 kn; ~* kc 12; !threads; !DumpDomain; !RCWCleanupList; qd"`.
   The main thread was idle in `EXCEL!IdleLoop::MainLoopWaitMessage`: Excel wasn't hung, it was
   waiting for references.
4. **Find off-thread Excel calls.** A temporary diagnostic logged the thread id, apartment and
   caller whenever the core's `ExcelApp` was read off Excel's main thread, plus the thread id at
   the start of the sheet write. The drilldown's sheet write showed `apt=MTA pool=True`.
   Afterwards the same log line showed thread 1 (STA), and Excel exited within ~3 s.

Ruled out by experiment (didn't help):
- 20 garbage collections in the host after shutdown.
- Moving only the ScreenUpdating/DisplayAlerts/EnableEvents calls to the main thread. The rest of
  the sheet write still ran on the MTA.

## Verification checklist

Close Excel after each item and confirm it leaves Task Manager within a few seconds (allow up to
~75 s on machines with Bitdefender):

- open, then close (no login);
- login, then close;
- each window-only feature, then close;
- each drilldown type, written to a sheet, then close: answer No to Save once and Save once;
- sheet/book refresh, then close; reset, then close;
- download of a large queued output (process/jobs window), then close. The window must stay in
  front of Excel during the download, and Excel must get focus back afterwards;
- hot reload while logged in: the ribbon must show the logged-out state, with no old cube,
  ledger or segment;
- hot reload, then a drilldown: no `__ComObject` to `WorkbookClass` error, and Excel exits after
  close.
