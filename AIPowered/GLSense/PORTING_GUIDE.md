# GLSense migration - porting rules

Ground rules every file ported from `FinalWorkingCode\GLSense` into
`GLSense.Addin.Core` (or the host `GLSense` project) must follow. These exist because
the new architecture adds two things the old monolith didn't have: an AppDomain
boundary (hot reload) and a WPF-UI (`BaseWindow`) shell instead of the old
`DpiAwareWindow`/MahApps shell. Both introduce failure modes that don't show up until
runtime, which is exactly the class of bug ("cross-threading, dispatcher invoke,
object reference not set, RCW issues") that's easy to miss without a compiler in the
loop, so they're written down here instead of re-derived per file.

## 1. Thread affinity (STA / Dispatcher)

- Excel COM objects (`Excel.Application`, `Range`, `Worksheet`, ...) are STA-affine.
  They must only be touched from the thread that created the COM apartment - which is
  Excel's main UI thread. Ribbon clicks and `ADXExcelAppEvents` events already run on
  that thread, so code called directly from `AddinModule.cs` or synchronously from
  `AddinEntry.OnRibbonAction`/`OnExcelEvent` is fine as-is.
- The WPF `Application`/`Dispatcher` for `GLSense.Addin.Core` is created by
  `WpfAppManager.EnsureApplication()`, which runs the first time `AddinEntry.Initialize`
  fires - on Excel's main thread. So WPF and Excel COM share the same physical thread in
  this add-in. That means you usually do **not** need `Dispatcher.Invoke` just to touch
  a WPF window from ribbon/event code.
- You **do** need `WpfAppManager.InvokeOnWpfThread(...)` whenever code might run on a
  different thread than the one that created the Dispatcher - background work started
  with `Task.Run`, a `HttpClient` continuation, a timer callback, etc. `InvokeOnWpfThread`
  already checks `Dispatcher.CheckAccess()` and calls the action directly if you're
  already on the right thread, so wrapping defensively costs nothing.
- Never call `.Result` or `.Wait()` on a `Task` from the STA/Dispatcher thread if
  anything in that Task's chain will eventually need to get back onto the same thread
  (e.g. via `Dispatcher.Invoke`) - that's a classic WPF/STA deadlock. Use `await`, or
  fire-and-forget with the exception already caught inside the async method (see below).

## 2. Fire-and-forget across the AppDomain boundary

- `IGLSenseAddin` methods (`OnRibbonAction`, `OnExcelEvent`, ...) are synchronous
  `void`/`bool` by contract, because they're called via .NET Remoting from `AddinModule`.
  Async work triggered from inside them (e.g. `DrillCellHighlighter.RibCellHighlight_OnClick`)
  has to be started with a discarded `Task` (`_ = SomeAsyncMethod();`).
- Every method invoked this way **must** catch everything internally (including
  `OperationCanceledException`) and log rather than rethrow. An unobserved exception
  from a discarded Task can crash the process on `.NET Framework` (unlike newer .NET
  where it's just logged) - do not skip this.
- Never let an exception cross back out of `AddinEntry`/`RibbonController`/`GLSenseContext`
  into `AddinModule` unhandled - remoting exceptions surface as ugly, hard-to-diagnose
  COM/remoting errors to the end user instead of a clean log entry.

## 3. Passing data across the AppDomain boundary

- Only pass values that are `[Serializable]` or `MarshalByRefObject` through
  `IGLSenseAddin`/`IGLSenseContext`/`IRibbonController` calls. Excel COM objects
  (RCWs) are special-cased by the CLR and marshal fine; plain framework/library POCOs
  (including AddinExpress's own event-args classes like `ADXExcelSheetBeforeEventArgs`)
  generally are **not** serializable and will throw `SerializationException` at the
  boundary.
- Prefer extracting primitives (sheet name, range address as a string, a bool) in the
  host **before** crossing into `GlobalsEx.Addin`, rather than passing the live object
  through. `OnExcelEvent(string eventName, object[] args)` follows this rule - keep
  following it if you add more events or ribbon actions with payloads.

## 4. COM RCW lifetime

- Every `Range`/`Worksheet`/etc. you get back from `SpecialCells`, `Union`, `Intersect`,
  etc. is a distinct RCW. Don't call `Marshal.ReleaseComObject` on something you're
  still going to read from later in the same method (use-after-release throws
  `COMException`/`InvalidComObjectException`), and don't release something you didn't
  create yourself (e.g. `ActiveSheet`, `ActiveCell` - Excel owns those).
- `BalanceFormulaExists` releases the `formulaCells` range it explicitly asked for via
  `SpecialCells` in a `finally` block - that's the pattern: release ranges you fetched
  purely to scan, once you're done scanning them, and nothing else.
- Multi-area ranges: never index `.Cells[i]` directly on a range that might have more
  than one `Area` (e.g. anything from `SpecialCells`) - it's unreliable across areas in
  Excel COM automation. Always `foreach (Range area in rng.Areas) foreach (Range cell in
  area.Cells)`. (This was the actual bug behind the RibCellHighlight regression - see
  `DrillCellHighlighter.CollectDependentsAddresses`.)
- Wrap every COM call site in `try/catch (COMException)` where "the object might already
  be gone" is plausible (a sheet deleted mid-operation, a workbook closed by the user
  while a background task is still running). Treat that as "nothing to do here", not a
  fatal error.

## 5. Null-guards specific to this project

- `ServiceLocator.*` throws `InvalidOperationException` if accessed before
  `ServiceLocator.Initialize` has run. That only happens once, in `AddinEntry.Initialize`,
  so this should never actually trigger in ported code - but if you're writing something
  that could run before `Initialize` (unlikely), guard with `ServiceLocator.IsInitialized`.
- `ServiceLocator.ExcelApp`, `.ActiveWorkbook`, `.ActiveSheet`, `.ActiveCell` can all be
  `null` (no workbook open, or the active sheet is a chart sheet, not a worksheet).
  Always null-check / pattern-match (`is Excel.Worksheet ws`) rather than casting
  directly - a bad cast throws `InvalidCastException`, a null active sheet throws
  `NullReferenceException` a line later if unchecked.
- `AppState.Instance` in this project only has `IsLoginCompleted` right now (unlike the
  old monolith's much bigger `AppState`). Don't assume other properties exist just
  because the old code referenced `AppState.Instance.X` - check what's actually here,
  and add to `AppState` deliberately rather than copying the old shape wholesale.

## 6. WPF-UI / BaseWindow specifics

- Every ported window derives from `BaseWindow` (not the old `DpiAwareWindow`).
  `BaseWindow` already handles PerMonitorV2 DPI awareness, work-area clamping, and
  setting Excel as the window owner (via `ServiceLocator.ExcelHandle` +
  `ModalToExcel`) - don't re-implement any of that in the derived window.
- Drag: use a dedicated title-bar `MouseLeftButtonDown` handler calling `DragMove()`
  (see `GLLogin.xaml`/`GLWaitWindow.xaml`), not a window-wide override. `BaseWindow`
  currently *also* has a window-wide `OnMouseLeftButtonDown`/`OnMouseMove` override that
  does this - that's a known conflict flagged separately (it will intercept clicks meant
  for buttons/inputs). Don't copy that pattern into new windows; when you get to
  `BaseWindow` itself, remove the window-wide override.
- Sizing: use `SizeToContent="Height"` + `Width="<fixed>"` and **only** `Auto`-sized
  rows in the root layout (no `*` rows) for anything that isn't meant to fill the whole
  work area. A `*` row combined with an unset/auto Height produced the original
  GLWaitWindow height bug - don't reintroduce it.
- Any control property you update from code *after* the window has loaded (title text,
  status text, etc.) must be a real `DependencyProperty` (like `Title`) or set directly
  on a named control (`txtTitle.Text = ...`) - not a plain CLR property with no change
  notification bound in XAML. A XAML `{Binding SomeProperty}` against a plain
  auto-property only reads it once; it will not pick up later changes.
- Icons: this project still references MahApps IconPacks (`iconPacks:PackIconFontAwesome`)
  alongside WPF-UI's own `ui:SymbolIcon`. Both work today (`AppOverlay.xaml` uses the
  former, `GLLogin.xaml`'s title bar uses the latter) - pick whichever matches the
  surrounding UI you're porting, but if you use `ui:SymbolIcon`, verify the exact
  `Wpf.Ui.Controls.SymbolRegular` enum member name exists (many "obvious" names like
  `Clock24` or `Dismiss24` don't - `Timer24`/`DismissCircle24` do). Don't guess; a wrong
  name is a compile error, not a typo you'll spot by reading the XAML.

## 7. General port procedure for a file

1. Read the original in `FinalWorkingCode\GLSense`. Note every static it touches
   (`AppState.Instance.X`, `LogUtility.*`, `AddinModule.CurrentInstance`, ...).
2. Re-point each one: `AppState.Instance.ExcelApp` -> `ServiceLocator.ExcelApp`,
   `LogUtility.*` -> `ServiceLocator.Logger.*`, `AddinModule.CurrentInstance.*` -> the
   equivalent via `ServiceLocator`/`GlobalsEx` (should not be needed inside
   `GLSense.Addin.Core` at all - that type isn't referenceable from there).
3. Check every dependency the file has is either already ported or gets ported in the
   same pass - a half-ported feature (view exists, view-model doesn't) is worse than not
   starting it, because it looks done in a file listing.
4. Apply sections 1-6 above.
5. Add the new files to `GLSense.Addin.Core.csproj` (`<Compile>`/`<Page>` items - MSBuild
   for this project type does not glob automatically).
6. Wire the call site: the `AddinModule.cs` ribbon stub becomes a one-line
   `_ribbonController?.ExecuteAction(RibbonControlIds.RibXxx);`, and
   `AddinEntry.OnRibbonAction`'s switch gets a matching `case "RibXxx":`.
