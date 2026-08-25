// GLSense.Addin.Core/AddinEntry.cs
using GLSense.Addin.Core.Caching;
using GLSense.Addin.Core.Common;
using GLSense.Addin.Core.Controls;
using GLSense.Addin.Core.Drilldowns;
using GLSense.Addin.Core.Infrastructure;
using GLSense.Addin.Core.Models;
using GLSense.Addin.Core.Repositories;
using GLSense.Addin.Core.Utilities;
using GLSense.Addin.Core.Views;
using GLSense.Contracts;
using System;
using System.Data;
using System.IO;
using System.Linq;
using System.Net.NetworkInformation;
using System.Reflection;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using Excel = Microsoft.Office.Interop.Excel;

namespace GLSense.Addin.Core
{
    public class AddinEntry : MarshalByRefObject, IGLSenseAddin
    {
        private IGLSenseContext _ctx;

        // Without this, .NET Remoting's default lease (5 min initial lifetime + 2 min
        // renewal per call) eventually expires this instance on the server side after a
        // period without any ribbon/UDF activity. GlobalsEx.Addin (the host's cached proxy
        // to this object - see RibbonController.ExecuteAction) then permanently throws
        // "Object '...' has been disconnected or does not exist at the server" on every
        // subsequent call, regardless of which action is invoked, since the underlying
        // remote object is simply gone - no reload/reconnect happens automatically. Every
        // other cross-AppDomain MarshalByRefObject in this project already returns null
        // here for the same reason (RibbonController, GLSenseContext, PathProvider, Logger,
        // RemoteLoader) - this was the one left over from the original port.
        public override object InitializeLifetimeService()
        {
            return null; // Unlimited lifetime
        }

        public object ExecuteUdf(string functionName, object[] args)
        {
            int argCount = args?.Length ?? 0;
            try
            {
                ServiceLocator.Logger?.LogDebug($"ExecuteUdf: function='{functionName}' argCount={argCount}");

                object result = Udf.UdfDispatcher.Execute(functionName, args);

                ServiceLocator.Logger?.LogDebug($"ExecuteUdf: function='{functionName}' returned '{result}'");
                return result;
            }
            catch (Exception ex)
            {
                // UdfDispatcher.Execute already catches everything internally and returns a
                // sentinel, so reaching this catch means something failed outside that (e.g.
                // ServiceLocator not yet initialized). Never let an exception escape back
                // across the AppDomain boundary into Excel - that would surface as a
                // cryptic COM error instead of a clean #VALUE!/#N/A.
                ServiceLocator.Logger?.LogException(ex, $"ExecuteUdf('{functionName}', argCount={argCount}): unhandled error");
                return UdfSentinels.XlErrorGettingData;
            }
        }

        public void Initialize(IGLSenseContext context)
        {
            _ctx = context;

            try
            {
                // Initialize ServiceLocator - this is the ONLY place where context is set
                ServiceLocator.Initialize(_ctx);

                // No handler anywhere previously caught a truly unhandled exception in
                // THIS AppDomain (only WPF-dispatcher-thread exceptions are covered
                // elsewhere) - a background Task or COM callback thread throwing
                // unhandled here would just silently crash/vanish with nothing in the
                // log. Registered as early as possible. The ADX shell's own AppDomain
                // has an identical hook registered separately in its AddinModule
                // constructor - exceptions in either domain need their own hook.
                AppDomain.CurrentDomain.UnhandledException += OnAppDomainUnhandledException;

                // Environment snapshot (GLSense/Excel/OS/DPI/culture/machine info) now lives
                // in Logger.BuildLogHeader/AppendEnvironmentSnapshot, written once per log
                // FILE (per day) via NLog's own Header mechanism, instead of once per Excel
                // session here - see that method's header comment for why.
                //
                // NLog's FileTarget only creates the physical file (and writes Header) on
                // its first actual log write - the Logger instance behind ServiceLocator.Logger
                // was already constructed (inside GLSenseContext's constructor, before this
                // Initialize call) but never itself writes anything. This line guarantees that
                // first write happens on every Excel open, so the file (and, once per day, the
                // header) always gets created even if nothing else logs during the session
                // (e.g. DebugMode off and the SQLite check below takes the "IsInitialized"
                // branch, whose LogDebug is a no-op without it).
                ServiceLocator.Logger?.LogInfo("GLSense session started.");

                //Initializing SQLite database
                // 1. Ensure DB file + tables exist

                var db = ServiceLocator.Database;
                if (db.IsInitialized)
                {
                    ServiceLocator.Logger.LogDebug("Database is ready for use");
                }
                else
                {
                    ServiceLocator.Logger?.LogWarn("Initialize: SQLite database reported IsInitialized=false.");
                }

                // Set up assembly resolver in this domain
                AppDomain.CurrentDomain.AssemblyResolve += (sender, args) =>
                {
                    var assemblyName = new AssemblyName(args.Name);

                    // Satellite resource assemblies (e.g. GLSense.Addin.Core.resources,
                    // MahApps.Metro.IconPacks.FontAwesome.resources) are proactively probed
                    // by the CLR's culture-fallback mechanism for every assembly that could
                    // have localized resources, regardless of whether one actually ships -
                    // neither this project nor its third-party packages ship any, so this
                    // always "fails" here, which is expected and harmless (the CLR falls
                    // back to the assembly's own embedded default resources).
                    // RemoteLoader.ResolveAssemblyInDomain already logs this once at Debug -
                    // skip re-logging a misleading "could not resolve" line here for the
                    // same expected non-event, so it doesn't look like a real failure in
                    // the log.
                    if (assemblyName.Name.EndsWith(".resources", StringComparison.OrdinalIgnoreCase) ||
                        assemblyName.Name.EndsWith(".g.resources", StringComparison.OrdinalIgnoreCase))
                    {
                        return null;
                    }

                    var assemblyPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, $"{assemblyName.Name}.dll");

                    if (File.Exists(assemblyPath))
                    {
                        ServiceLocator.Logger?.LogDebug($"Initialize.AssemblyResolve: resolved '{assemblyName.Name}' -> '{assemblyPath}'");
                        return Assembly.LoadFrom(assemblyPath);
                    }

                    ServiceLocator.Logger?.LogDebug($"Initialize.AssemblyResolve: could not resolve '{args.Name}' (expected path '{assemblyPath}' not found).");
                    return null;
                };

                // Ensure WPF Application exists - ONLY CALL THIS ONCE
                WpfAppManager.EnsureApplication();

                // Pay WindowLoadingPlaceholder's one-time HWND-creation/first-paint cost now,
                // during Excel's own ribbon-load idle time (deferred to ApplicationIdle
                // priority inside WarmUpInBackground itself), rather than letting the very
                // first real window shown in this Excel session pay it synchronously on its
                // own OnSourceInitialized/ShowMatching call.
                WindowLoadingPlaceholder.WarmUpInBackground();

                // Check if ribbon controller is already available
                if (ServiceLocator.RibbonController != null)
                {
                    if (!LoggedIn())
                    {
                        ServiceLocator.Context.RibbonController.SetState("LoggedOut");
                    }
                    _ctx.Logger.LogDebug("RibbonController already available");
                }
                else
                {
                    _ctx.Logger.LogDebug("RibbonController will be available when ribbon loads");
                }
            }
            catch (Exception ex)
            {
                // Initialize is the very first thing the host calls after creating this
                // AppDomain - if it throws unhandled, the add-in fails to load with zero
                // visibility into why. _ctx.Logger is used here (rather than
                // ServiceLocator.Logger) since ServiceLocator.Initialize(_ctx) itself is
                // inside the try and may not have completed yet.
                try { _ctx?.Logger?.LogException(ex, "AddinEntry.Initialize: Error"); } catch { /* logging itself must never throw */ }
            }
        }

        // Last-resort catch for an exception unhandled anywhere else in this AppDomain.
        // Flushes every still-open action buffer first, so the debug trace leading up to
        // the crash survives even though the buffer's own owning LogScope will never
        // Dispose() normally after this.
        private void OnAppDomainUnhandledException(object sender, UnhandledExceptionEventArgs e)
        {
            try
            {
                var ex = e.ExceptionObject as Exception;
                if (ex != null)
                {
                    ServiceLocator.Logger?.LogException(ex, "AppDomain.UnhandledException");
                }
                else
                {
                    ServiceLocator.Logger?.LogError($"AppDomain.UnhandledException with non-Exception payload: {e.ExceptionObject}");
                }

                ServiceLocator.Logger?.FlushDebugLogs("unhandled exception");
            }
            catch
            {
                // This handler must never itself throw - there's nothing left to catch it.
            }
        }

        // Environment snapshot (GLSense/Excel/OS/DPI/culture/machine info) now lives in
        // Logger.BuildLogHeader/AppendEnvironmentSnapshot (GLSense.Shared), written once
        // per log FILE (per day) via NLog's own Header mechanism, instead of once per
        // Excel session here - see that method's header comment for why.

        private bool LoggedIn()
        {
            if (!AppState.Instance.IsLoginCompleted)
            {
                return false;
            }
            return true;
        }
        public void OnRibbonAction(string action, object parameter)
        {
            // Single entry-point log for every ribbon button/control dispatched through this
            // method (all ~50 cases below) - covers "which button fired" for the whole switch
            // without repeating a LogDebug in every case. parameter is logged too since several
            // cases (LedgerChanged/SegmentChanged/RunSnapshot/RunDrilldownExternal/etc.) carry
            // meaningful payloads across the AppDomain boundary.
            ServiceLocator.Logger?.LogDebug($"OnRibbonAction: action='{action}' parameter='{parameter}'");

            try
            {
            switch (action)
            {
                case "Login":
                    Login();
                    break;
                case "ShowMessage":
                    string filePath = $@"C:\Users\RamaMaidam_2fc\Downloads\Job.log"; ;
                    SimpleMessage(filePath);
                    break;
                case "ShowMessage1":
                    string filePath1 = $@"C:\Users\RamaMaidam_2fc\Downloads\LargeText.txt";
                    SimpleMessage(filePath1);
                    break;
                case "RibCellHighlight":
                    // Fire-and-forget: RibCellHighlight_OnClick is async but OnRibbonAction is
                    // a synchronous void contract method (it's called across the AppDomain
                    // boundary from AddinModule). The method already catches and logs every
                    // exception internally (including OperationCanceledException), so it's
                    // safe to discard the Task here.
                    _ = DrillCellHighlighter.RibCellHighlight_OnClick();
                    break;
                case "Logout":
                    // Fire-and-forget for the same reason as RibCellHighlight above - Logout
                    // makes an async API call (applogout). Exceptions are caught internally.
                    _ = Logout();
                    break;
                case "ShowCubeDetails":
                    // RibGetCube_OnClick - opens GLCubeDetails (Group B). Modal, so no
                    // fire-and-forget needed; mirrors Login()'s dispatch pattern.
                    ShowCubeDetails();
                    break;
                case "LedgerChanged":
                    // Ribledger_OnChange - the host passes the newly selected ledger name
                    // as `parameter` (a string; the only thing that can safely cross the
                    // AppDomain boundary here - see PORTING_GUIDE.md). Fire-and-forget for
                    // the same reason as Logout/RibCellHighlight above.
                    _ = LedgerChanged(parameter as string);
                    break;

                // ---- Group C (Segment/Period pickers) ----
                // ExecuteAction(buttonId) takes no payload (see RibbonController.
                // ExecuteAction), so each ribbon button that needs a distinct formulaName/
                // mode gets its own action-name case here, dispatching into one shared
                // launcher method per window family (mirrors the old monolith's
                // LaunchPeriodDetails/LaunchPeriodStarEnd/LaunchSegmentWindow grouping)
                // rather than one near-duplicate method per button.
                case "ShowPeriod":
                    ShowPeriod();
                    break;
                case "ShowPeriodByYear":
                    ShowPeriodByYear();
                    break;
                case "ShowPeriodByDate":
                    ShowPeriodByDate();
                    break;
                case "ShowPeriodNum":
                    ShowPeriodDetails("NUM");
                    break;
                case "ShowPeriodQtr":
                    ShowPeriodDetails("QTR");
                    break;
                case "ShowPeriodYear":
                    ShowPeriodDetails("YEAR");
                    break;
                case "ShowPeriodStart":
                    ShowPeriodStartEnd("START");
                    break;
                case "ShowPeriodEnd":
                    ShowPeriodStartEnd("END");
                    break;
                case "ShowSegmentEnabledFlag":
                    ShowSegmentWindow("ENABLEDFLAG");
                    break;
                case "ShowSegmentSummaryFlag":
                    ShowSegmentWindow("SUMMARYFLAG");
                    break;
                case "ShowSegmentDescription":
                    ShowSegmentWindow("DESCRIPTION");
                    break;
                case "ShowNextSegment":
                    ShowSegmentWindow("NEXTSEGMENT");
                    break;
                case "ShowPreviousSegment":
                    ShowSegmentWindow("PREVIOUSSEGMENT");
                    break;
                case "ShowSegmentDFF":
                    ShowSegmentWindow("DFF");
                    break;
                case "ShowSegmentAccountType":
                    ShowSegmentWindow("ACCOUNTTYPE");
                    break;
                case "ShowDailyRate":
                    ShowDailyRate();
                    break;

                // ---- Group D (Segment/Period discoverers) ----
                // RibSegProperty/RibExpandAll/RibbonExpand1Level/RibExpodeAll/
                // RibbonExplode1Level/RibDiscoverPeriod are simple one-liners in the old
                // monolith's host (SegmentDiscoverer.SegmentAction(...)/PeriodsDiscoverer.
                // FillPeriods() called directly from RibXxx_OnClick). SegmentDiscoverer/
                // PeriodsDiscoverer now live in this project instead, so each becomes a
                // fire-and-forget dispatch here, same reasoning as RibCellHighlight/Logout
                // above (both methods catch+log every exception internally).
                case "RibSegProperty":
                    _ = SegmentDiscoverer.SegmentAction("Property");
                    break;
                case "RibExpodeAll":
                    _ = SegmentDiscoverer.SegmentAction("ExplodeAll");
                    break;
                case "RibbonExplode1Level":
                    _ = SegmentDiscoverer.SegmentAction("Explode1Level");
                    break;
                case "RibDiscoverPeriod":
                    _ = PeriodsDiscoverer.FillPeriods();
                    break;

                // RigSegDiscover_OnClick - unlike the 6 cases above, the old host body
                // did real validation (DefaultSegment/SegmentPickedIndex + active-cell
                // value checks) before opening GLSegmentDiscovery. GLSegmentDiscovery now
                // lives in this project's AppDomain, so the host can no longer `new` it
                // directly - all of that validation + the window launch moved here.
                case "ShowSegmentDiscovery":
                    ShowSegmentDiscovery();
                    break;

                // RibSegmentExpand_OnClick - was two ribbon buttons (RibExpandAll/
                // RibbonExpand1Level) hosted inside a "Hierarchy" menu, each dispatching
                // straight to SegmentDiscoverer.SegmentAction("HierarchyAll"/
                // "Hierarchy1Level"). Now a single button that opens GLExpandOptions,
                // where the user picks level + fill direction before SegmentAction is
                // invoked from the dialog itself.
                case "ShowExpandOptions":
                    ShowExpandOptions();
                    break;

                // RibSegS_OnChange - the host only passes the selected segment name
                // string (see SegmentChanged's own summary for why the RibSegS.Items/
                // index computation could not cross the AppDomain boundary as-is).
                case "SegmentChanged":
                    SegmentChanged(parameter as string);
                    break;

                // ---- Group E (Drilldowns) ----
                // RibDrillJobs/RibDDConfiguration open BaseWindow-derived windows, so they
                // reuse ShowGroupCWindow (Group C's shared modal-dispatch helper) exactly
                // like the 7 Group C pickers above - no new mechanism needed.
                case "ShowJobsMonitor":
                    ShowJobsMonitor();
                    break;
                case "ShowDrilldownCustomization":
                    ShowDrilldownCustomization();
                    break;
                case "DeleteDrilldownCustomization":
                    DeleteDrilldownCustomization();
                    break;

                // RibBalanceDD/RibBalanceJournalDD/RibBalanceSubLedgerDD/RibTotaDD - old
                // monolith's RunBalanceDrilldownAsync(ddType), one shared method
                // parameterized by mode rather than 4 near-duplicates (same grouping
                // already used for ShowPeriodDetails/ShowPeriodStartEnd/ShowSegmentWindow
                // above). Fire-and-forget for the same reason as RibCellHighlight/Logout -
                // RunBalanceDrilldown catches and logs every exception internally.
                case "RunBalanceDrilldownBL":
                    _ = RunBalanceDrilldown("BL");
                    break;
                case "RunBalanceDrilldownBLJL":
                    _ = RunBalanceDrilldown("BL_JL");
                    break;
                case "RunBalanceDrilldownBLSL":
                    _ = RunBalanceDrilldown("BL_SL");
                    break;
                case "RunBalanceDrilldownUF":
                    _ = RunBalanceDrilldown("UF");
                    break;

                // RibJournalDD/RibSubledgerDD - old monolith's RibJournalDD_OnClick/
                // RibSubledgerDD_OnClick bodies, moved here verbatim (fire-and-forget,
                // same reasoning as the balance drilldowns above).
                case "RibJournalDD":
                    _ = RunJournalDrilldown();
                    break;
                // RibBalancesDDToSubLedger/RibBalancesDDToUnified - old monolith's
                // RibBalancesDDToSubLedger_OnClick/RibBalancesDDToUnified_OnClick (both
                // just called RunDrilldownAsync("BLDD_SL"/"BLDD_UF") - same DrilldownJl
                // mechanism as RibJournalDD, just a different ddType routing to a
                // different REST endpoint, see DD_JL.cs BuildApiUrl). Previously missing
                // end-to-end; restored as part of the completeness pass.
                case "RibBalancesDDToSubLedger":
                    _ = RunJournalDrilldown("BLDD_SL");
                    break;
                case "RibBalancesDDToUnified":
                    _ = RunJournalDrilldown("BLDD_UF");
                    break;
                case "RibSubledgerDD":
                    _ = RunSubledgerDrilldown();
                    break;

                // ---- Group F (Refresh/Clear/Highlight/Hide-Rows) ----
                // RibRefreshAll/RibRefreshBook - old monolith's one-liners
                // (RibRefreshAll_OnClick/RibRefreshBook_OnClick =>
                // BalanceRefresh.RefreshingBalancesAsync("Refresh", "Sheet"/"Book")), moved
                // here verbatim now that BalanceRefresh.cs (and its transitive
                // BulkRefreshProcess/ExcelFormulaGenerator/DataTableBuilder/
                // BalanceNormalizer dependencies) are ported. Fire-and-forget, same
                // reasoning as RunBalanceDrilldown/RunJournalDrilldown above - the async
                // methods catch and log every exception internally.
                case "RibRefreshAll":
                    _ = BalanceRefresh.RefreshingBalancesAsync("Refresh", "Sheet");
                    break;
                case "RibRefreshBook":
                    _ = BalanceRefresh.RefreshingBalancesAsync("Refresh", "Book");
                    break;

                // RibRefreshRange - extracted into Drilldowns\RangeRefresher.cs (its own
                // class rather than folded here - see that file's header comment for why).
                case "RibRefreshRange":
                    _ = RangeRefresher.RibRefreshRange_OnClick();
                    break;

                // RibClearSheet/RibClear - old monolith's ResetBalances(resetType)/
                // BalancesReset(sheetName), small enough to fold directly into this class
                // (mirrors Group D's SegmentChanged-sized methods) - see ResetBalances/
                // BalancesReset below.
                case "RibClearSheet":
                    _ = ResetBalances("Sheet");
                    break;
                case "RibClear":
                    _ = ResetBalances("Book");
                    break;

                // RibHighlight - extracted into Drilldowns\BalanceHighlighter.cs.
                case "RibHighlight":
                    _ = BalanceHighlighter.RibHighlight_OnClick();
                    break;

                // RibHideRows/RibUnHideRows - extracted into
                // Drilldowns\RowVisibilityProcessor.cs (RowProcessor/HideRowProcessor/
                // UnhideRowProcessor hierarchy).
                case "RibHideRows":
                    _ = RowVisibilityProcessor.RibHideRows_OnClick();
                    break;
                case "RibUnHideRows":
                    _ = RowVisibilityProcessor.RibUnHideRows_OnClick();
                    break;

                // ---- Group G (Snapshot/Job submission) ----
                // ExecuteAction(buttonId) takes no payload (see the Group C comment above),
                // so RibSnapSubmit/RibSnapShot/RibSnapWorksheet/RibSnapWorkbook all go
                // straight through GlobalsEx.Addin?.OnRibbonAction(...) from the host
                // (like Ribledger_OnChange/RibSegS_OnChange) instead of _ribbonController.
                // ExecuteAction, since each one needs to carry a bool (or a small
                // pipe-delimited payload for RunSnapshot) across the AppDomain boundary -
                // all of which are primitives/strings, safe to cross via .NET Remoting.
                case "SnapSubmitToggled":
                    // Old monolith: RibSnapSubmit_OnClick => AppState.Instance.SnapshotJob
                    // = RibSnapSubmit.Pressed.
                    AppState.Instance.SnapshotJob = parameter is bool pressedSubmit && pressedSubmit;
                    break;

                case "RunSnapshot":
                    // Old monolith: RibSnapShot_OnClick - see RunSnapshot(string) below.
                    _ = RunSnapshot(parameter as string);
                    break;

                case "SnapWorksheetToggled":
                    // Old monolith: RibSnapWorksheet_OnClick => guard + RibSnapWorkbook.
                    // Pressed = !pressed. See ToggleSnapMode below.
                    ToggleSnapMode("RibSnapWorkbook", parameter is bool pressedWs && pressedWs);
                    break;

                case "SnapWorkbookToggled":
                    // Old monolith: RibSnapWorkbook_OnClick => guard + RibSnapWorksheet.
                    // Pressed = !pressed. See ToggleSnapMode below.
                    ToggleSnapMode("RibSnapWorksheet", parameter is bool pressedWb && pressedWb);
                    break;

                // ---- Group H (Balance Configurator pane + LOVs/Roller/Account) ----
                // RibLiveCalc_OnClick - old monolith's one-liner (AppState.Instance.
                // SingleRefresh = RibLiveCalc.Pressed), same reasoning as SnapSubmitToggled
                // above (bool payload, host has no local action registered so this comes
                // straight through OnRibbonAction).
                case "SingleRefreshToggled":
                    AppState.Instance.SingleRefresh = parameter is bool pressedLiveCalc && pressedLiveCalc;
                    break;

                // ShowLOVs/ShowRollerGroups/ShowSegmentValues (RibLOVs/RibRollerGroup/
                // RibAccount) - the task-pane (GLConfiguratorPane/RibFSG) stays entirely
                // host-side (see AddinModule.RibFSG_OnClick), but these three ribbon
                // buttons open standalone modal dialogs, so they reuse ShowGroupCWindow
                // exactly like the Group C pickers. GLLOVs/GLRollerGroups/GLSegmentValues
                // (+ their ViewModels) are now ported - this is the last remaining Group H
                // content port. The host's RibLOVs_OnClick/RibRollerGroup_OnClick/
                // RibAccount_OnClick already call _ribbonController.ExecuteAction(
                // "ShowLOVs"/"ShowRollerGroups"/"ShowSegmentValues") and needed no further
                // changes once these 3 cases were wired up here.
                case "ShowLOVs":
                    ShowGroupCWindow("ShowLOVs", () => new GLLOVs());
                    break;
                case "ShowRollerGroups":
                    ShowGroupCWindow("ShowRollerGroups", () => new GLRollerGroups());
                    break;
                case "ShowSegmentValues":
                    ShowGroupCWindow("ShowSegmentValues", () => new GLSegmentValues());
                    break;

                // ---- Group I (Config/Debug/About/Help) ----
                // RibUserConfig/Riburl/RibAbout open standalone modal dialogs - same
                // ShowGroupCWindow reuse as every other Group C/H dialog above.
                case "ShowUserConfig":
                    ShowGroupCWindow("ShowUserConfig", () => new GLUserConfig());
                    break;
                case "ShowServerConfiguration":
                    ShowGroupCWindow("ShowServerConfiguration", () => new GLServerConfiguration());
                    break;
                case "ShowAbout":
                    ShowGroupCWindow("ShowAbout", () => new GLAbout());
                    break;
                case "ShowLoginDetails":
                    // RibDBL1_OnAction - old monolith's tiny read-only "current login"
                    // dialog. GLLoginDetails binds straight to AppState.Instance via
                    // {x:Static} in XAML, so ShowGroupCWindow needs no extra setup here.
                    ShowGroupCWindow("ShowLoginDetails", () => new GLLoginDetails());
                    break;

                // RibDebug_OnClick/RibVersionCheck_OnClick - old monolith's one-liners
                // (AppState.Instance.DebugLogs = RibDebug.Pressed / AppState.Instance.
                // VersionCheck = RibVersionCheck.Pressed), same reasoning as
                // SingleRefreshToggled above (bool payload, host has no local action
                // registered so this comes straight through OnRibbonAction).
                case "DebugLogsToggled":
                    bool debugPressed = parameter is bool pressedDebug && pressedDebug;
                    _ctx.DebugMode=debugPressed;
                    AppState.Instance.DebugLogs = debugPressed;
                    if (debugPressed)
                    {
                        ServiceLocator.Logger?.LogDebug("Debug session started. Detailed traces will be written until disabled.");
                    }
                    else
                    {
                        ServiceLocator.Logger?.FlushDebugLogs("Debug session ended");
                    }
                    break;
                case "VersionCheckToggled":
                    AppState.Instance.VersionCheck = parameter is bool pressedVersionCheck && pressedVersionCheck;
                    break;

                // RibHelp_OnClick - old monolith built the help URL from AppState.Instance.
                // LoginUrl/LoginToken (both live in this AppDomain) and called
                // Process.Start directly; Process.Start has no Excel-COM/UI-thread
                // affinity requirement, so this whole thing can stay here rather than
                // crossing back into the host. Fire-and-forget for the same reason as
                // Logout/RibCellHighlight above - ShowHelp catches every exception
                // internally.
                case "ShowHelp":
                    ShowHelp();
                    break;

                // ---- Double-click/hyperlink drilldown dispatch (resolves the
                // SheetBeforeDoubleClick/SheetFollowHyperlink TODOs left by Group E) ----
                // Host-side classification (AddinModule.adxExcelAppEvents1_
                // SheetBeforeDoubleClick/adxExcelAppEvents1_SheetFollowHyperlink) inspects
                // the live Range/Worksheet/Hyperlink the Excel event handed it (formula
                // text, sheet-name patterns, hyperlink ScreenTip/SubAddress) - that can't
                // cross the AppDomain boundary, so only the classification RESULT (plain
                // strings) comes through here.
                case "RunDrilldownExternal":
                    // payload: "ddType|external" (ddType one of BL/JL/SL/EP - see
                    // RunDrilldownByExternalAddress below). Pipe-delimited, same
                    // convention as Group G's RunSnapshot ("mode|isSubmit").
                    if (parameter is string ddPayload)
                    {
                        string[] ddParts = ddPayload.Split(new[] { '|' }, 2);
                        if (ddParts.Length == 2)
                        {
                            _ = RunDrilldownByExternalAddress(ddParts[0], ddParts[1]);
                        }
                    }
                    break;

                case "RunCustomDrilldown":
                    // payload: string[3] { tableSheetName, cellExternalAddress, headerLabel }.
                    // Not pipe-delimited like the others - headerLabel is free-form report
                    // text and could contain '|', so this crosses as a small serializable
                    // array instead (same mechanism OnExcelEvent's object[] args already
                    // relies on).
                    if (parameter is string[] cdArgs && cdArgs.Length == 3)
                    {
                        _ = CustomDrilldown.RunCustomDrilldown(cdArgs[0], cdArgs[1], cdArgs[2]);
                    }
                    break;

                case "RunJournalAttachmentFlow":
                    // payload: journalHeaderId (as string, already converted host-side from
                    // the hyperlink cell's Value2).
                    if (parameter is string journalHeaderIdText)
                    {
                        // Fully qualified: GLSense.Addin.Core.Models.JournalAttachments (a
                        // plain data model) also resolves here via the Models namespace
                        // import, making the bare name ambiguous.
                        _ = Drilldowns.JournalAttachments.RunJournalAttachmentFlow(journalHeaderIdText);
                    }
                    break;
            }
            }
            catch (Exception ex)
            {
                // Never let an exception escape across the AppDomain boundary back into
                // AddinModule/Excel - mirrors the same fail-open pattern OnExcelEvent/
                // ExecuteUdf use below. Most cases above already dispatch into methods that
                // catch+log internally (fire-and-forget Tasks), but a handful of simple
                // synchronous cases (e.g. SnapSubmitToggled's AppState assignment) have no
                // try/catch of their own - this is the safety net for those.
                try { ServiceLocator.Logger?.LogException(ex, $"OnRibbonAction('{action}') failed"); } catch { /* logging itself must never throw */ }
            }
        }

        /// <summary>
        /// Double-click drilldown dispatch - old monolith's RunDrilldown(cellRange, ddType)/
        /// RunBalancePrecedentDrilldown, reunified here as one small switch rather than
        /// duplicating DrilldownBl/DrilldownJl/DrilldownSl construction (those classes are
        /// exactly the same ones RunBalanceDrilldown/RunJournalDrilldown/
        /// RunSubledgerDrilldown below already use for the ribbon-button-triggered
        /// drilldowns - "BL"/"JL"/"SL" here are the SAME ddType values RunDrilldown used in
        /// the old monolith). Deliberately NOT reusing RunBalanceDrilldown/
        /// RunJournalDrilldown/RunSubledgerDrilldown as-is: those re-derive the external
        /// address from ServiceLocator.ExcelApp.Selection, which is correct for a
        /// ribbon-button click (the user already selected the cell) but NOT for a
        /// double-click - Excel does not necessarily move Selection to the double-clicked
        /// cell before SheetBeforeDoubleClick fires, so the host builds the external
        /// address from e.Target itself and passes it through explicitly here instead.
        /// "EP" (Excel Precedents) is the old monolith's RunBalancePrecedentDrilldown
        /// branch (double-clicking a cell with a formula that is NOT a balance formula -
        /// walks precedent cells looking for balance formulas instead of drilling a
        /// balance cell itself).
        /// </summary>
        private static async Task RunDrilldownByExternalAddress(string ddType, string external)
        {
            try
            {
                switch (ddType)
                {
                    case "BL":
                        await new DrilldownBl(ServiceLocator.ExcelApp, external, "BL").ProcessBLDrilldown();
                        break;
                    case "JL":
                        await new DrilldownJl(ServiceLocator.ExcelApp, external).ProcessJLDrilldown();
                        break;
                    case "SL":
                        await new DrilldownSl(ServiceLocator.ExcelApp, external).ProcessSLDrilldown();
                        break;
                    case "EP":
                        await new DrilldownXlPrecedents(external).ProcessEPDrilldown();
                        break;
                    default:
                        ServiceLocator.Logger?.LogDebug($"RunDrilldownByExternalAddress: unknown drilldown type '{ddType}'");
                        break;
                }
            }
            catch (Exception ex)
            {
                ServiceLocator.Logger?.LogException(ex, $"RunDrilldownByExternalAddress({ddType}): Error");
            }
        }

        private void ShowHelp()
        {
            try
            {
                if (!AppState.Instance.IsLoginCompleted)
                {
                    CommonFunctions.GLSenseMessage("Please log in to the instance.", MessageBoxImage.Exclamation);
                    return;
                }

                if (!string.IsNullOrEmpty(AppState.Instance.LoginToken))
                {
                    string helpUrl = AppState.Instance.LoginUrl + "/web/public/redirect-help/Excel_GLSense.htm?jwtParam=" + AppState.Instance.LoginToken;
                    ServiceLocator.Logger?.LogDebug(helpUrl);
                    System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo
                    {
                        FileName = helpUrl,
                        UseShellExecute = true
                    });
                }
            }
            catch (Exception ex)
            {
                ServiceLocator.Logger?.LogException(ex, "ShowHelp");
            }
        }
        private void SafeInvokeWpf(System.Action action)
        {
            if (action == null) return;
            try
            {
                WpfAppManager.InvokeOnWpfThread(() =>
                {
                    try
                    {
                        action();
                    }
                    catch (Exception ex)
                    {
                        _ctx.Logger.LogException(ex, "SafeInvokeWpf: action() failed on WPF thread");
                    }
                });
            }
            catch (Exception ex)
            {
                _ctx.Logger.LogException(ex, "SafeInvokeWpf: InvokeOnWpfThread failed");
            }
        }
        //"C:\Users\RamaMaidam_2fc\Downloads\Job.log"
        private void SimpleMessage(string samplePath)
        {
            bool loggedIn = LoggedIn();

            //if (loggedIn)
            //{

            //    if (File.Exists(samplePath))
            //    {
            //        string content = File.ReadAllText(samplePath);
            //        MessageWindowCls.GLSenseMessage(content, MessageBoxIcon.Information, MessageBoxButtons.OK);
            //    }
            //    else
            //    {
            //        MessageWindowCls.GLSenseMessage($"File not found: {samplePath}", MessageBoxIcon.Information, MessageBoxButtons.OK);
            //    }

            //}
            //else
            //{
            //    MessageWindowCls.GLSenseMessage($"You have not logged in...", MessageBoxIcon.Information, MessageBoxButtons.OK);
            //}
        }
        private void Login()
        {
            try
            {
                ServiceLocator.Logger?.LogDebug("Login: Starting login process...");

                // Use InvokeOnWpfThread from your WpfAppManager
                WpfAppManager.InvokeOnWpfThread(() =>
                {
                    try
                    {
                        ServiceLocator.Logger?.LogDebug("Login: Creating GLLogin window...");

                        var win = new GLLogin
                        {
                            CenterInExcel = true,
                            ModalToExcel = true,  // CRITICAL: Makes it modal to Excel
                            ShowInTaskbar = false  // Don't show in taskbar since it's owned by Excel
                        };
                        // WindowCaption/IconSymbol are already set in GLLogin.xaml itself
                        // ("GLSense Login" / "Key24") - no need to set them again here.

                        // Show as dialog - this will be modal to Excel. GLLogin's own
                        // WebView2 navigation-completed flow drives everything from here:
                        // cookie extraction, the /finance-cubes API call, and the
                        // PartialLoggedIn ribbon-state transition on success. It closes
                        // itself (Close()) once cubes have been fetched and cached.
                        //
                        // Note: AppState.Instance.IsLoginCompleted is deliberately NOT set
                        // here (a previous draft of this method set it unconditionally,
                        // which was wrong). In the original project it only becomes true
                        // once a cube+ledger is chosen in GLCubeDetails (Group B/Cube-Ledger
                        // selection) - a bare successful sign-in only reaches
                        // "PartialLoggedIn", not "LoggedIn".
                        win.ShowDialog();

                        ServiceLocator.Logger?.LogDebug("Login: Login dialog closed.");
                    }
                    catch (Exception ex)
                    {
                        ServiceLocator.Logger?.LogException(ex, "Login: ShowDialog error");
                    }
                });
            }
            catch (Exception ex)
            {
                ServiceLocator.Logger?.LogException(ex, "Login: Error");
            }
        }

        /// <summary>
        /// RibGetCube ribbon action (Group B) - opens the cube/ledger selection window.
        /// Mirrors Login()'s dispatch pattern exactly (WpfAppManager.InvokeOnWpfThread +
        /// ShowDialog()); GLCubeDetails drives everything else itself (cube/ledger fetch,
        /// FillResponsibilitiesAsync, ribbon combo population, FinalizeLogin).
        /// </summary>
        private void ShowCubeDetails()
        {
            try
            {
                ServiceLocator.Logger?.LogDebug("ShowCubeDetails: Opening GLCubeDetails window...");

                WpfAppManager.InvokeOnWpfThread(() =>
                {
                    try
                    {
                        var win = new GLCubeDetails
                        {
                            CenterInExcel = true,
                            ModalToExcel = true,
                            ShowInTaskbar = false
                        };

                        win.ShowDialog();

                        ServiceLocator.Logger?.LogDebug("ShowCubeDetails: Dialog closed.");
                    }
                    catch (Exception ex)
                    {
                        ServiceLocator.Logger?.LogException(ex, "ShowCubeDetails: ShowDialog error");
                    }
                });
            }
            catch (Exception ex)
            {
                ServiceLocator.Logger?.LogException(ex, "ShowCubeDetails: Error");
            }
        }

        /// <summary>
        /// Group C (Segment/Period pickers) - RibPeriod ribbon action. Opens GLGetPeriod
        /// (period + numeric offset picker). Mirrors ShowCubeDetails()'s dispatch pattern
        /// exactly (WpfAppManager.InvokeOnWpfThread + ShowDialog()); GLGetPeriod drives
        /// everything else itself (ledger/period load, existing-formula round-trip parse
        /// in its own Window_Loaded, WriteFormulaToCell on submit).
        /// </summary>
        private void ShowPeriod()
        {
            ShowGroupCWindow("ShowPeriod", () => new GLGetPeriod());
        }

        /// <summary>
        /// Group C - RibPeriodbyYear ribbon action. Opens GLGetPeriodByYear (period year +
        /// period num picker).
        /// </summary>
        private void ShowPeriodByYear()
        {
            ShowGroupCWindow("ShowPeriodByYear", () => new GLGetPeriodByYear());
        }

        /// <summary>
        /// Group C - RibPeriodbyDate ribbon action. Opens GLGetPeriodByDate (date + ledger
        /// + numeric offset picker).
        /// </summary>
        private void ShowPeriodByDate()
        {
            ShowGroupCWindow("ShowPeriodByDate", () => new GLGetPeriodByDate());
        }

        /// <summary>
        /// Group C - shared launcher for RibPeriodNum/RibPeriodQtr/RibPeriodYear (old
        /// monolith's LaunchPeriodDetails(FuncName)). Opens GLGetPeriodDetails with the
        /// formulaName that selects which of GLSense_GetPeriodNum/Quarter/Year gets
        /// written on submit.
        /// </summary>
        private void ShowPeriodDetails(string formulaName)
        {
            ShowGroupCWindow($"ShowPeriodDetails({formulaName})", () => new GLGetPeriodDetails(formulaName));
        }

        /// <summary>
        /// Group C - shared launcher for RibPeriodStart/RibPeriodEnd (old monolith's
        /// LaunchPeriodStarEnd(FuncName)). Opens GLGetPeriodStartEnd with the formulaName
        /// that selects GLSense_GetPeriodStart vs. GLSense_GetPeriodEnd.
        /// </summary>
        private void ShowPeriodStartEnd(string formulaName)
        {
            ShowGroupCWindow($"ShowPeriodStartEnd({formulaName})", () => new GLGetPeriodStartEnd(formulaName));
        }

        /// <summary>
        /// Group C - shared launcher for RibSegmentEnabledFlag/RibSummaryFlag/RibSegment/
        /// RibNextSegment/RibPreviousSegment/RibSegmentDFF (old monolith's
        /// LaunchSegmentWindow(FuncName)). Opens GLSegmentFunctions with the funcName that
        /// selects which of the 6 GLSense_GetSegment*/GetNextSegment/GetPreviousSegment
        /// formulas gets written on submit.
        /// </summary>
        private void ShowSegmentWindow(string funcName)
        {
            ShowGroupCWindow($"ShowSegmentWindow({funcName})", () => new GLSegmentFunctions(funcName));
        }

        /// <summary>
        /// Group C - RibDailyRate ribbon action. Opens GLDailyRates (currency + conversion
        /// type + date picker).
        /// </summary>
        private void ShowDailyRate()
        {
            ShowGroupCWindow("ShowDailyRate", () => new GLDailyRates());
        }

        /// <summary>
        /// Shared dispatch helper for all 7 Group C picker windows - every one of them is a
        /// BaseWindow-derived modal dialog opened the same way ShowCubeDetails() opens
        /// GLCubeDetails (Group B), so this collapses that identical
        /// WpfAppManager.InvokeOnWpfThread + CenterInExcel/ModalToExcel/ShowInTaskbar +
        /// ShowDialog() boilerplate into one place instead of repeating it 7 times.
        /// </summary>
        private void ShowGroupCWindow(string actionLabel, Func<BaseWindow> createWindow)
        {
            try
            {
                ServiceLocator.Logger?.LogDebug($"{actionLabel}: Opening window...");

                WpfAppManager.InvokeOnWpfThread(() =>
                {
                    try
                    {
                        var win = createWindow();
                        win.CenterInExcel = true;
                        win.ModalToExcel = true;
                        win.ShowInTaskbar = false;

                        win.ShowDialog();

                        ServiceLocator.Logger?.LogDebug($"{actionLabel}: Dialog closed.");
                    }
                    catch (Exception ex)
                    {
                        ServiceLocator.Logger?.LogException(ex, $"{actionLabel}: ShowDialog error");
                    }
                });
            }
            catch (Exception ex)
            {
                ServiceLocator.Logger?.LogException(ex, $"{actionLabel}: Error");
            }
        }

        /// <summary>
        /// Ribledger ribbon-combo OnChange (Group B) - lets the user switch ledgers
        /// directly from the ribbon without reopening GLCubeDetails. Port of the old
        /// monolith's Ribledger_OnChange/ValidatePreconditions/ValidateLedgerChange/
        /// PerformLedgerChangeAsync/UpdateRibbonUI, collapsed into one method here since
        /// this project has no per-window code-behind for a ribbon combo handler.
        /// Re-pointed vs. the original: AddinModule.CurrentInstance.Ribledger/RibSegS
        /// (direct ADX control access) -> ServiceLocator.RibbonController.SetComboText/
        /// SetComboItems/ClearComboItems; LogUtility.* -> ServiceLocator.Logger.*;
        /// WinForms MessageBoxIcon/MessageBoxButtons -> WPF MessageBoxImage/MessageBoxButton
        /// (CommonFunctions.GLSenseMessage's actual signature in this project).
        /// AppState.Instance.DefaultSegment/SegmentPickedIndex (segment-picker session
        /// state) are deliberately NOT touched here - they're independent of which ledger
        /// is selected and are owned by RibSegS_OnChange/AddinEntry.SegmentChanged
        /// (Group D) instead.
        /// Balance Configurator pane relaunch (Group H, resolved) happens at the end of
        /// the Task.Run body below via RelaunchConfiguratorPaneIfVisible(), same mechanism
        /// as GLCubeDetails.ConfiguratorRelaunch().
        /// </summary>
        private async Task LedgerChanged(string ledgerName)
        {
            if (string.IsNullOrWhiteSpace(ledgerName))
                return;

            if (!AppState.Instance.IsLoginCompleted || AppState.Instance.SelectedCube == null)
            {
                CommonFunctions.GLSenseMessage("Please log in first.", MessageBoxImage.Exclamation);
                return;
            }

            var ledger = AppState.Instance.SelectedCube.Ledgers?.FirstOrDefault(l => l.LedgerName == ledgerName);
            if (ledger == null)
            {
                CommonFunctions.GLSenseMessage("Ledger not found.", MessageBoxImage.Exclamation);
                return;
            }

            bool shouldContinue = true;
            bool sheetClear = false;

            if (AppState.Instance.SelectedLedger?.CoaId != ledger.Coaid)
            {
                var choice = CommonFunctions.GLSenseMessage(
                    "Different chart of account detected. Clear sheet?",
                    MessageBoxImage.Question,
                    MessageBoxButton.YesNoCancel);

                shouldContinue = choice != MessageBoxResult.Cancel;
                sheetClear = choice == MessageBoxResult.Yes;

                if (!shouldContinue)
                {
                    // Revert the ribbon combo back to the previously selected ledger.
                    ServiceLocator.RibbonController?.SetComboText("Ribledger", AppState.Instance.SelectedLedger?.LedgerName ?? string.Empty);
                    return;
                }
            }

            using var ctsHelper = new Helpers.CancellationHelper();
            var token = ctsHelper.GetToken();
            GLWaitWindow win = null;

            try
            {
                // WpfAppManager.InvokeOnWpfThread only takes an Action (no return value,
                // see CommonFunctions.GLSenseMessage for the same pattern), so capture the
                // created window from inside the delegate.
                WpfAppManager.InvokeOnWpfThread(() =>
                {
                    win = new GLWaitWindow(ctsHelper);
                    win.SetProcessTitle("Switching Ledger");
                    win.SetProcessMessage(sheetClear ? "Clearing sheet..." : "Loading ledger data...");
                    win.Show();
                    win.StartMonitoring();
                });

                // Cross-thread COM note: this Task.Run body touches ServiceLocator.ExcelApp
                // (Cells.Clear) and, via UpdateRibbonAndSelectionForLedger, the ribbon's
                // ADX controls - both live on Excel's main STA apartment, not this
                // ThreadPool thread. That's safe here because Excel's main thread returned
                // immediately from Ribledger_OnChange (LedgerChanged was fired as
                // fire-and-forget) and keeps pumping its message loop, so COM will marshal
                // these calls back to that apartment automatically; it will not deadlock.
                // This mirrors the original monolith's PerformLedgerChangeAsync, which used
                // the identical Task.Run-wrapped-COM-access pattern. Kept inside one
                // Task.Run (rather than dispatching each COM call individually) to match
                // that proven behavior; revisit if this group ever needs to run in a
                // stricter threading model.
                await Task.Run(async () =>
                {
                    if (sheetClear)
                    {
                        token.ThrowIfCancellationRequested();
                        if (ServiceLocator.ExcelApp?.ActiveSheet is Excel.Worksheet sheet)
                        {
                            sheet.Cells.Clear();
                        }
                        await win.Dispatcher.InvokeAsync(() => win.SetProcessMessage("Loading ledger data..."));
                    }

                    token.ThrowIfCancellationRequested();

                    await CommonFunctions.FillResponsibilitiesAsync(ledger.LedgerId, AppState.Instance.SelectedCube.CubeId, token);

                    token.ThrowIfCancellationRequested();
                    UpdateRibbonAndSelectionForLedger(ledger);

                    // Group H (resolved): old monolith's ledger-change relaunch
                    // (`if (BalancePane != null && BalancePane.Visible) _ =
                    // BalancePane.RelaunchPane();`). GLConfiguratorPane is host-only, so
                    // this goes through IRibbonController - the host does the Visible
                    // check before deciding whether to call back into
                    // RelaunchConfiguratorPane.
                    ServiceLocator.RibbonController?.RelaunchConfiguratorPaneIfVisible();
                }, token);
            }
            catch (OperationCanceledException)
            {
                ServiceLocator.Logger?.LogWarn("Ledger change operation cancelled by user!");
            }
            catch (Exception ex)
            {
                ServiceLocator.Logger?.LogException(ex, "LedgerChanged: Error");
                CommonFunctions.GLSenseMessage(ex.Message, MessageBoxImage.Error);
            }
            finally
            {
                if (win != null)
                {
                    await win.Dispatcher.InvokeAsync(() => win.RequestClose());
                }
            }
        }

        /// <summary>
        /// Repopulates the RibSegS ribbon combo for the newly selected ledger and updates
        /// AppState.Instance.SelectedLedger - the two things the old UpdateRibbonUI did,
        /// minus the DefaultSegment/SegmentPickedIndex fields (see LedgerChanged's summary).
        /// </summary>
        private static void UpdateRibbonAndSelectionForLedger(LedgerRecord ledger)
        {
            ServiceLocator.RibbonController?.ClearComboItems("RibSegS");

            var repository = new DataRepository();
            var segs = repository.GetSegments(AppState.Instance.SelectedCube.CubeId, ledger.LedgerId);
            // .ToList() matters here, not just style: SetComboItems is a cross-AppDomain
            // IRibbonController call, and the lazy WhereSelectEnumerableIterator .Select()
            // produces on its own isn't [Serializable] - passing it directly throws a
            // SerializationException during remoting argument marshaling.
            ServiceLocator.RibbonController?.SetComboItems("RibSegS", segs.Select(s => s.SegmentName).ToList());

            AppState.Instance.SelectedLedger = repository.GetLedgers(AppState.Instance.SelectedCube.CubeId)
                ?.FirstOrDefault(l => l.LedgerId == ledger.LedgerId);
        }

        /// <summary>
        /// Group D - RigSegDiscover_OnClick. Port of the old monolith's host body
        /// (AddinModule.cs lines ~2321-2371): validates DefaultSegment/SegmentPickedIndex
        /// and the active cell's value against the selected segment's values, then opens
        /// GLSegmentDiscovery modally. Re-pointed vs. the original: AppState.Instance.
        /// ExcelApp -> ServiceLocator.ExcelApp; WinForms MessageBoxIcon/MessageBoxButtons
        /// -> WPF MessageBoxImage; new GLSegmentDiscovery().ShowDialogWithOwner(hwnd) ->
        /// ShowGroupCWindow(...), the same shared BaseWindow dispatch helper every other
        /// Group C picker window already uses (GLSegmentDiscovery is BaseWindow-derived
        /// now, not DpiAwareWindow, so it fits that helper directly).
        /// </summary>
        private void ShowSegmentDiscovery()
        {
            try
            {
                if (string.IsNullOrEmpty(AppState.Instance.DefaultSegment) || AppState.Instance.SegmentPickedIndex < 0)
                {
                    CommonFunctions.GLSenseMessage("Please select a segment from the dropdown.", MessageBoxImage.Exclamation);
                    return;
                }

                var activeCell = ServiceLocator.ExcelApp?.ActiveCell;
                var activeCellValue = activeCell?.Value2;
                string activeCellText = activeCellValue?.ToString().Trim();

                if (string.IsNullOrEmpty(activeCellText))
                {
                    CommonFunctions.GLSenseMessage("Active cell is empty. Please select a cell with a segment value.", MessageBoxImage.Exclamation);
                    return;
                }

                var repository = new DataRepository();
                var segments = repository.GetSegments(AppState.Instance.SelectedCube.CubeId, AppState.Instance.SelectedLedger.LedgerId);
                var segment = segments?.FirstOrDefault(s => s.SegmentName == AppState.Instance.DefaultSegment);
                if (segment != null)
                {
                    var segmentValues = DataRepository.GetSegmentValues(segment);

                    if (segmentValues != null)
                    {
                        bool isValidValue = segmentValues.Any(v => string.Equals(v.SegmentValue?.Trim(), activeCellText, StringComparison.OrdinalIgnoreCase));
                        if (isValidValue)
                        {
                            ShowGroupCWindow("ShowSegmentDiscovery", () => new GLSegmentDiscovery());
                        }
                        else
                        {
                            CommonFunctions.GLSenseMessage($"The value in the active cell \"{activeCellText}\" does not match any of the values for the selected segment \"{segment.SegmentName}\" . Please select a cell with a valid segment value.", MessageBoxImage.Exclamation);
                        }
                    }
                    else
                    {
                        CommonFunctions.GLSenseMessage($"Failed in fetching segment values for the selected segment \"{segment.SegmentName}\" .", MessageBoxImage.Exclamation);
                    }
                }
                else
                {
                    CommonFunctions.GLSenseMessage("Selected segment not found. Please re-select the segment from the dropdown.", MessageBoxImage.Exclamation);
                }
            }
            catch (Exception ex)
            {
                ServiceLocator.Logger?.LogException(ex, "ShowSegmentDiscovery: Error");
            }
        }

        /// <summary>
        /// Group D - RibSegmentExpand_OnClick. Opens GLExpandOptions (Expand All/1 Level +
        /// By Rows/By Columns RadioButtons), replacing the old RibExpandAll/
        /// RibbonExpand1Level ribbon-menu pair. Only a lightweight DefaultSegment/
        /// SegmentPickedIndex guard happens here - SegmentDiscoverer.SegmentAction
        /// (invoked from the dialog's own Expand button) already does the full
        /// active-cell/segment-value validation, exactly as it did when these were two
        /// separate ribbon buttons.
        /// </summary>
        private void ShowExpandOptions()
        {
            try
            {
                if (string.IsNullOrEmpty(AppState.Instance.DefaultSegment) || AppState.Instance.SegmentPickedIndex < 0)
                {
                    CommonFunctions.GLSenseMessage("Please select a segment from the dropdown.", MessageBoxImage.Exclamation);
                    return;
                }

                ShowGroupCWindow("ShowExpandOptions", () => new GLExpandOptions());
            }
            catch (Exception ex)
            {
                ServiceLocator.Logger?.LogException(ex, "ShowExpandOptions: Error");
            }
        }

        /// <summary>
        /// Group E (Drilldowns) - RibDrillJobs ribbon action. Opens GLJobsMonitor (the
        /// "Processed Jobs" background-jobs monitor, backed by
        /// ViewModels\GLSubmittedJobsViewModel.cs). Old monolith's RibDrillJobs_OnClick
        /// opened it via SafeInvokeWpf + ShowDialogWithOwner(hwnd); this reuses
        /// ShowGroupCWindow exactly like every Group C picker window instead, since
        /// GLJobsMonitor is BaseWindow-derived here (not DpiAwareWindow).
        /// </summary>
        private void ShowJobsMonitor()
        {
            ShowGroupCWindow("ShowJobsMonitor", () => new GLJobsMonitor());
        }

        /// <summary>
        /// Group E (Drilldowns) - RibDDConfiguration ribbon action. Opens
        /// GLDrilldownCustomization (WebView2-embedded drilldown launcher configuration
        /// page, ported earlier this pass). Same ShowGroupCWindow reuse as
        /// ShowJobsMonitor() above.
        /// </summary>
        private void ShowDrilldownCustomization()
        {
            ShowGroupCWindow("ShowDrilldownCustomization", () => new GLDrilldownCustomization());
        }

        /// <summary>
        /// Ported from FinalWorkingCode\GLSense\AddinModule.cs's RibDDDeleteConfiguration_OnClick
        /// (single-project monolith, does everything inline there). Deletes the saved
        /// DRILLDOWNMETADATA CustomXMLPart for the currently selected cube
        /// (Common\DrilldownMetadataXmlStore.cs::Delete), letting the user remove a locally
        /// saved drilldown customization (GLDrilldownCustomization's "Save Locally" button)
        /// without having to save a new one in its place. Unlike ShowDrilldownCustomization
        /// above, this doesn't open a window - it just performs the delete and reports the
        /// result via CommonFunctions.GLSenseMessage, same as other simple ribbon actions.
        /// </summary>
        private void DeleteDrilldownCustomization()
        {
            ServiceLocator.Logger?.LogDebug("AddinEntry.DeleteDrilldownCustomization invoked");

            if (AppState.Instance.SelectedCube == null)
            {
                CommonFunctions.GLSenseMessage("No cube selected. Please select a cube first.", MessageBoxImage.Exclamation, MessageBoxButton.OK);
                return;
            }

            long cubeId = AppState.Instance.SelectedCube.CubeId;

            try
            {
                var wb = ServiceLocator.ExcelApp?.ActiveWorkbook;
                bool deleted = DrilldownMetadataXmlStore.Delete(wb, cubeId);

                if (deleted)
                {
                    ServiceLocator.Logger?.LogDebug($"AddinEntry.DeleteDrilldownCustomization: deleted saved drilldown customization for cubeId={cubeId}.");
                    CommonFunctions.GLSenseMessage("Saved drilldown customization deleted successfully.", MessageBoxImage.Information, MessageBoxButton.OK);
                }
                else
                {
                    ServiceLocator.Logger?.LogDebug($"AddinEntry.DeleteDrilldownCustomization: no saved drilldown customization found for cubeId={cubeId}.");
                    CommonFunctions.GLSenseMessage("No saved drilldown customization exists for the current cube.", MessageBoxImage.Exclamation, MessageBoxButton.OK);
                }
            }
            catch (Exception ex)
            {
                ServiceLocator.Logger?.LogException(ex, "AddinEntry.DeleteDrilldownCustomization");
                CommonFunctions.GLSenseMessage("Failed to delete the saved drilldown customization.", MessageBoxImage.Error, MessageBoxButton.OK);
            }
        }

        /// <summary>
        /// Group E (Drilldowns) - shared launcher for RibBalanceDD/RibBalanceJournalDD/
        /// RibBalanceSubLedgerDD/RibTotaDD (old monolith's RunBalanceDrilldownAsync(ddType)).
        /// Reads the current Excel selection directly off ServiceLocator.ExcelApp (this
        /// project's AppState has no ExcelApp field - same gap DD_BL.cs documents),
        /// builds the external range address, and fires off DrilldownBl.ProcessBLDrilldown()
        /// for the given ddType ("BL"/"BL_JL"/"BL_SL"/"UF"). Called fire-and-forget from
        /// OnRibbonAction (see the 4 "RunBalanceDrilldown*" cases above) - exceptions are
        /// caught and logged here so nothing escapes back across the AppDomain boundary.
        /// </summary>
        private static async Task RunBalanceDrilldown(string ddType)
        {
            try
            {
                Excel.Range rng = ServiceLocator.ExcelApp?.Selection as Excel.Range;
                string external = Helpers.ExcelExternalRef.BuildExternalAddress(rng);
                var runProcess = new DrilldownBl(ServiceLocator.ExcelApp, external, ddType);
                await runProcess.ProcessBLDrilldown();
            }
            catch (Exception ex)
            {
                ServiceLocator.Logger?.LogException(ex, $"RunBalanceDrilldown({ddType}): Error");
            }
        }

        /// <summary>
        /// Group E (Drilldowns) - RibJournalDD/RibBalancesDDToSubLedger/
        /// RibBalancesDDToUnified ribbon actions (old monolith's RibJournalDD_OnClick/
        /// RibBalancesDDToSubLedger_OnClick/RibBalancesDDToUnified_OnClick bodies, all of
        /// which just called RunDrilldownAsync(ddType) with "JL"/"BLDD_SL"/"BLDD_UF"
        /// respectively). The ddType parameter defaults to "JL" so the plain RibJournalDD
        /// case (and the "JL" double-click-dispatch path in RunDrilldownByExternalAddress)
        /// are unaffected. Fire-and-forget from OnRibbonAction, same reasoning as
        /// RunBalanceDrilldown above.
        /// </summary>
        private static async Task RunJournalDrilldown(string ddType = "JL")
        {
            try
            {
                Excel.Range rng = ServiceLocator.ExcelApp?.Selection as Excel.Range;
                string external = Helpers.ExcelExternalRef.BuildExternalAddress(rng);
                var runProcess = new DrilldownJl(ServiceLocator.ExcelApp, external, ddType);
                await runProcess.ProcessJLDrilldown();
            }
            catch (Exception ex)
            {
                ServiceLocator.Logger?.LogException(ex, $"RunJournalDrilldown({ddType}): Error");
            }
        }

        /// <summary>
        /// Group E (Drilldowns) - RibSubledgerDD ribbon action (old monolith's
        /// RibSubledgerDD_OnClick body, moved here verbatim). Fire-and-forget from
        /// OnRibbonAction, same reasoning as RunBalanceDrilldown above.
        /// </summary>
        private static async Task RunSubledgerDrilldown()
        {
            try
            {
                Excel.Range rng = ServiceLocator.ExcelApp?.Selection as Excel.Range;
                string external = Helpers.ExcelExternalRef.BuildExternalAddress(rng);
                var runProcess = new DrilldownSl(ServiceLocator.ExcelApp, external);
                await runProcess.ProcessSLDrilldown();
            }
            catch (Exception ex)
            {
                ServiceLocator.Logger?.LogException(ex, "RunSubledgerDrilldown: Error");
            }
        }

        /// <summary>
        /// Group G (Snapshot/Job submission) - RibSnapShot ribbon action (old monolith's
        /// RibSnapShot_OnClick body). The host can't read AppState.Instance.SnapshotJob
        /// (it lives in this AppDomain), so instead of re-deriving mode/isSubmit here it
        /// receives both, exactly as the old handler read them at click time off
        /// RibSnapWorksheet.Pressed / RibSnapSubmit.AsRibbonCheckBox.Pressed - packed into
        /// one pipe-delimited string ("Sheet|True"/"Book|False") since OnRibbonAction's
        /// parameter can only carry a single primitive/serializable value across the
        /// AppDomain boundary. Fire-and-forget from OnRibbonAction, same reasoning as
        /// RunBalanceDrilldown above - BalanceRefresh's own methods catch and log
        /// everything internally, but this wraps them too for safety.
        /// </summary>
        private static async Task RunSnapshot(string payload)
        {
            try
            {
                if (string.IsNullOrEmpty(payload)) return;

                string[] parts = payload.Split('|');
                string mode = parts.Length > 0 && !string.IsNullOrEmpty(parts[0]) ? parts[0] : "Sheet";
                bool isSubmit = parts.Length > 1 && bool.TryParse(parts[1], out bool parsed) && parsed;

                AppState.Instance.SnapshotJob = isSubmit;

                if (isSubmit)
                {
                    await BalanceRefresh.SubmitSnapAsync(mode);
                }
                else
                {
                    await BalanceRefresh.RefreshingBalancesAsync("Snapshot", mode);
                }
            }
            catch (Exception ex)
            {
                ServiceLocator.Logger?.LogException(ex, $"RunSnapshot('{payload}'): Error");
            }
        }

        /// <summary>
        /// Group G (Snapshot/Job submission) - shared body for RibSnapWorksheet_OnClick/
        /// RibSnapWorkbook_OnClick (old monolith: each one guards on IsLoginCompleted/
        /// SelectedCube, then forces the OTHER ribbon toggle to the opposite state so the
        /// two behave as a mutually-exclusive pair). The host can't touch the sibling
        /// ADXRibbonItem directly from here (it's a host-only type across the AppDomain
        /// boundary), so this reuses ServiceLocator.RibbonController.SetControlPressed -
        /// the same reflection-backed API Group B/C already use for ribbon combo/control
        /// state - to flip it instead.
        /// </summary>
        /// <param name="otherControlName">"RibSnapWorkbook" or "RibSnapWorksheet" - the
        /// sibling toggle to force to the opposite of <paramref name="pressed"/>.</param>
        private void ToggleSnapMode(string otherControlName, bool pressed)
        {
            if (!AppState.Instance.IsLoginCompleted || AppState.Instance.SelectedCube == null)
                return;

            ServiceLocator.RibbonController?.SetControlPressed(otherControlName, !pressed);
        }

        /// <summary>
        /// Group D - RibSegS_OnChange. DESIGN NOTE: the old host body computed
        /// AppState.Instance.SegmentPickedIndex by finding the picked ADXRibbonItem's
        /// index within RibSegS.Items - a host-only ADX collection that cannot cross the
        /// AppDomain boundary. Rather than trying to marshal that index across, the host
        /// now passes only the selected segment name (a string; see LedgerChanged above
        /// for the identical reasoning with Ribledger_OnChange), and this method re-derives
        /// the index from the exact same DataRepository.GetSegments(...) call that
        /// UpdateRibbonAndSelectionForLedger used to populate the RibSegS combo in the
        /// first place - guaranteeing the ordering (and therefore the index) matches what
        /// the ribbon combo actually displayed.
        /// </summary>
        private void SegmentChanged(string segmentName)
        {
            if (string.IsNullOrEmpty(segmentName))
            {
                AppState.Instance.DefaultSegment = string.Empty;
                AppState.Instance.SegmentPickedIndex = -1;
                return;
            }

            try
            {
                var segments = AppState.Instance.SelectedCube != null && AppState.Instance.SelectedLedger != null
                    ? new DataRepository().GetSegments(AppState.Instance.SelectedCube.CubeId, AppState.Instance.SelectedLedger.LedgerId)
                    : null;

                var match = segments?
                    .Select((s, idx) => new { Segment = s, Index = idx })
                    .FirstOrDefault(x => x.Segment.SegmentName == segmentName);

                if (match != null)
                {
                    AppState.Instance.DefaultSegment = segmentName;
                    AppState.Instance.SegmentPickedIndex = match.Index;
                }
                else
                {
                    // Selected segment wasn't found in the freshly-fetched list (e.g. the
                    // ledger changed between the ribbon combo being populated and this
                    // event firing) - fall back to the same "nothing picked" state as the
                    // null/empty case above.
                    AppState.Instance.DefaultSegment = string.Empty;
                    AppState.Instance.SegmentPickedIndex = -1;
                }
            }
            catch (Exception ex)
            {
                ServiceLocator.Logger?.LogException(ex, "SegmentChanged: Error");
                AppState.Instance.DefaultSegment = string.Empty;
                AppState.Instance.SegmentPickedIndex = -1;
            }
        }

        /// <summary>
        /// Logout ribbon action - counterpart to Login() above. Calls the sibling
        /// XLEdge add-in's logoff (best-effort, if present), then hits the server's
        /// applogout endpoint, resets AppState, and returns the ribbon to "LoggedOut".
        /// Async because it makes a network call; OnRibbonAction dispatches it
        /// fire-and-forget (see the "Logout" case above) since the IGLSenseAddin
        /// contract method itself must stay synchronous.
        /// </summary>
        private async System.Threading.Tasks.Task Logout()
        {
            try
            {
                ServiceLocator.Logger?.LogDebug("Logout: Starting logout process...");

                try
                {
                    object edgeAddin = Helpers.EdgeAddinHelper.GetEdgeAddinInstance();
                    edgeAddin?.GetType().InvokeMember("LogoffFromAddin", BindingFlags.InvokeMethod, null, edgeAddin, Array.Empty<object>());
                }
                catch (Exception ex)
                {
                    ServiceLocator.Logger?.LogException(ex, "Exception while logging out from XLEdge via GLSense!");
                }

                // Group H (resolved): old monolith's AddinModule.HideTaskPanes() call
                // before logging off. GLConfiguratorPane/its TaskPaneInstances collection
                // are host-only ADX constructs, so this goes through
                // IRibbonController.HideAllTaskPanes() instead of reaching into
                // AddinModule directly.
                try
                {
                    ServiceLocator.RibbonController?.HideAllTaskPanes();
                }
                catch (Exception ex)
                {
                    ServiceLocator.Logger?.LogException(ex, "Logout: HideAllTaskPanes failed");
                }

                string message;
                try
                {
                    var apiUrl = $"{AppState.Instance.LoginUrl}{AppConstants.WebSecure}applogout";
                    ServiceLocator.Logger?.LogDebug(apiUrl);

                    var result = await Helpers.ApiHelper.ServerAPI(apiUrl, "Form", "", "GET", System.Threading.CancellationToken.None);
                    ServiceLocator.Logger?.LogDebug($"Logout API response: {result}");

                    var logout = Helpers.ApiResponseHelper.Parse<System.Text.Json.JsonElement>(result, Helpers.JsonGlobals.Options);
                    message = logout.IsSuccess
                        ? (logout.Value.TryGetProperty("message", out var msgEl) ? msgEl.GetString() : null) ?? "Successfully logged out."
                        : logout.ErrorMessage;
                }
                catch (Exception ex)
                {
                    ServiceLocator.Logger?.LogException(ex, "Logout: applogout call failed");
                    message = "An unexpected error occurred during logout.";
                }

                AppState.Instance.Reset();

                // AppState.Reset() only clears the in-memory SelectedCube/SelectedLedger
                // model state - it doesn't touch the ribbon UI (by design, AppState is
                // pure data). Without these, RibGetCube/Ribledger kept showing whatever
                // cube/ledger was selected before logout. Matches FinalWorkingCode's
                // GLSenseLogout(): Ribledger.Items.Clear(); Ribledger.Text = "";
                // RibGetCube.Caption = "Cube: Select Cube";
                ServiceLocator.RibbonController?.SetControlLabel("RibGetCube", "Cube : Select Cube");
                ServiceLocator.RibbonController?.ClearComboItems("Ribledger");
                ServiceLocator.RibbonController?.SetComboText("Ribledger", string.Empty);
                ServiceLocator.RibbonController?.ClearComboItems("RibSegS");

                ServiceLocator.RibbonController?.SetState("LoggedOut");

                CommonFunctions.GLSenseMessage(message, MessageBoxImage.Information);

                ServiceLocator.Logger?.LogDebug("Logout: Logout completed, state reset.");
            }
            catch (Exception ex)
            {
                ServiceLocator.Logger?.LogException(ex, "Logout: Error");
            }
        }
        /// <summary>
        /// Generic Excel Application-event dispatcher - see IGLSenseAddin.OnExcelEvent for the
        /// contract this must honor (args are primitives only, never a live COM/ADX object;
        /// never let an exception escape back across the AppDomain boundary into AddinModule).
        ///
        /// SheetActivate/SheetChange/SheetSelectionChange/SheetBeforeDoubleClick/
        /// SheetFollowHyperlink are logging-only here by design - their real behavior is
        /// live-Excel-COM-dependent (inspecting the active worksheet, a changed cell's
        /// formula, etc.) and stays entirely host-side in AddinModule.cs
        /// (ApplySheetActiveState, the SheetChange "reapply LoggedIn state" check, the
        /// double-click/hyperlink drilldown dispatch - see the comments on those handlers
        /// for details). WorkbookActivate and WorkbookBeforeSave DO have real logic here -
        /// see their cases below (previously these two were mistakenly left as
        /// logging-only stubs during the initial port, alongside a comment incorrectly
        /// claiming they had no real logic to carry over - see the old monolith's
        /// adxExcelAppEvents1_WorkbookActivate/WorkbookBeforeSave in AddinModule.cs for what
        /// was actually being ported here).
        /// </summary>
        public bool OnExcelEvent(string eventName, object[] args)
        {
            try
            {
                switch (eventName)
                {
                    case "SheetActivate":
                        ServiceLocator.Logger?.LogDebug($"OnExcelEvent(SheetActivate): Sheet={GetArg(args, 0)}");
                        return true;

                    case "SheetChange":
                        ServiceLocator.Logger?.LogDebug($"OnExcelEvent(SheetChange): Sheet={GetArg(args, 0)}, Range={GetArg(args, 1)}");
                        return true;

                    case "SheetSelectionChange":
                        ServiceLocator.Logger?.LogDebug($"OnExcelEvent(SheetSelectionChange): Sheet={GetArg(args, 0)}, Range={GetArg(args, 1)}");
                        return true;

                    case "SheetBeforeDoubleClick":
                        ServiceLocator.Logger?.LogDebug($"OnExcelEvent(SheetBeforeDoubleClick): Sheet={GetArg(args, 0)}, Range={GetArg(args, 1)}");
                        // Real double-click drilldown dispatch stays host-side (AddinModule.cs) -
                        // always allow the default Excel double-click behavior to proceed.
                        return true;

                    case "SheetFollowHyperlink":
                        ServiceLocator.Logger?.LogDebug($"OnExcelEvent(SheetFollowHyperlink): Sheet={GetArg(args, 0)}, Hyperlink={GetArg(args, 1)}");
                        return true;

                    case "WorkbookActivate":
                        {
                            string workbookName = GetArg(args, 0) as string;
                            ServiceLocator.Logger?.LogDebug($"OnExcelEvent(WorkbookActivate): Workbook={workbookName}");

                            // Port of the old monolith's adxExcelAppEvents1_WorkbookActivate:
                            // reapply the LoggedIn ribbon state and resync the cube/ledger/
                            // segment combos whenever the user switches to a different open
                            // workbook, so the ribbon reflects THIS workbook's last-known
                            // selection rather than whatever the previously-active workbook
                            // left behind.
                            if (AppState.Instance.IsLoginCompleted)
                            {
                                ServiceLocator.RibbonController?.SetState("LoggedIn");
                                SyncRibbonSelectionWithAppState();
                            }

                            return true;
                        }

                    case "WorkbookBeforeSave":
                        {
                            string workbookName = GetArg(args, 0) as string;
                            bool saveAsUi = GetArg(args, 1) is bool asUi && asUi;
                            ServiceLocator.Logger?.LogDebug(
                                $"OnExcelEvent(WorkbookBeforeSave): Workbook={workbookName}, SaveAsUI={saveAsUi}, " +
                                $"CacheInitialized={FormulaCacheManager.Instance.IsInitialized}, HasChanges={FormulaCacheManager.Instance.HasChanges}");

                            // Port of the old monolith's adxExcelAppEvents1_WorkbookBeforeSave:
                            // persist the formula cache to SQLite when the user saves the
                            // workbook. Never cancel the save on failure - PersistToDatabase
                            // already logs and swallows its own errors.
                            if (FormulaCacheManager.Instance.IsInitialized && FormulaCacheManager.Instance.HasChanges)
                            {
                                FormulaCacheManager.Instance.PersistToDatabase();
                            }

                            return true;
                        }

                    default:
                        ServiceLocator.Logger?.LogDebug($"OnExcelEvent: unrecognized eventName '{eventName}'");
                        return true;
                }
            }
            catch (Exception ex)
            {
                // Fail-open: never block Excel or throw back across the AppDomain boundary
                // just because one of these handlers hit an unexpected error.
                ServiceLocator.Logger?.LogException(ex, $"OnExcelEvent: unexpected error handling '{eventName}'");
                return true;
            }
        }

        private static object GetArg(object[] args, int index)
        {
            return args != null && index < args.Length ? args[index] : null;
        }

        /// <summary>
        /// Port of the old monolith's AddinModule.SyncRibbonSelectionWithAppState(): after
        /// login, or whenever WorkbookActivate fires, repopulates the cube/ledger/segment
        /// ribbon combos from AppState.Instance.SelectedCube/SelectedLedger/DefaultSegment.
        /// Re-pointed vs. the original: direct ADXRibbonItem/Ribledger/RibSegS/RibGetCube
        /// manipulation -> IRibbonController.SetComboItems/SetComboText/ClearComboItems/
        /// SetControlLabel (this project can't reference the host's ribbon controls
        /// directly - see IRibbonController.cs's header for the full rationale).
        /// </summary>
        private static void SyncRibbonSelectionWithAppState()
        {
            if (!AppState.Instance.IsLoginCompleted)
                return;

            try
            {
                var cube = AppState.Instance.SelectedCube;
                var ledger = AppState.Instance.SelectedLedger;

                if (cube == null)
                    return;

                ServiceLocator.RibbonController?.SetControlLabel("RibGetCube", $"Selected Cube : {cube.CubeName}");

                // Ribledger: clear-then-maybe-refill was NOT atomic - ClearComboItems ran
                // unconditionally, but SetComboItems (the refill) only ran if
                // cube.Ledgers != null. If that check ever failed (stale/incomplete cube
                // reference, timing hiccup, etc.) the combo was left with its Text set to
                // the selected ledger's name but ZERO items in the dropdown - exactly the
                // "shows one value, dropdown pops up empty" symptom reported. Fixed to be
                // atomic: only clear+refill+set-text together when we actually have data,
                // otherwise leave whatever Ribledger currently shows untouched instead of
                // blanking it. Also logs the ledger count so a recurrence is diagnosable
                // from the logs instead of guesswork.
                var ledgerCount = cube.Ledgers?.Count ?? 0;
                ServiceLocator.Logger?.LogDebug($"SyncRibbonSelectionWithAppState: cube={cube.CubeName}, ledgerCount={ledgerCount}, selectedLedger={ledger?.LedgerName}");
                if (ledgerCount > 0)
                {
                    // .ToList() matters here, not just style: SetComboItems is a
                    // cross-AppDomain IRibbonController call, and the lazy
                    // OrderBy().Select() iterator chain isn't [Serializable] - passing it
                    // directly throws a SerializationException during remoting argument
                    // marshaling.
                    ServiceLocator.RibbonController?.SetComboItems("Ribledger", cube.Ledgers.OrderBy(l => l.LedgerName).Select(l => l.LedgerName).ToList());
                    ServiceLocator.RibbonController?.SetComboText("Ribledger", ledger?.LedgerName ?? string.Empty);

                    // Belt-and-suspenders (see identical comment in GLCubeDetails.LoadCubeLedgers):
                    // force a per-control invalidate right where the items were set, instead
                    // of relying solely on the later blanket _ribbon.Invalidate() from
                    // SetState("LoggedIn").
                    ServiceLocator.RibbonController?.Invalidate("Ribledger");
                }
                else
                {
                    ServiceLocator.Logger?.LogWarn($"SyncRibbonSelectionWithAppState: cube.Ledgers was empty/null for '{cube.CubeName}' - leaving Ribledger untouched instead of blanking it.");
                }

                // RibSegS: same atomicity fix as Ribledger above.
                if (ledger != null)
                {
                    var repository = new DataRepository();
                    var segs = repository.GetSegments(cube.CubeId, ledger.LedgerId).ToList();
                    if (segs.Count > 0)
                    {
                        // .ToList() matters here, not just style: SetComboItems is a
                        // cross-AppDomain IRibbonController call, and the lazy
                        // WhereSelectEnumerableIterator .Select() produces on its own isn't
                        // [Serializable] - passing it directly throws a SerializationException
                        // during remoting argument marshaling.
                        ServiceLocator.RibbonController?.SetComboItems("RibSegS", segs.Select(s => s.SegmentName).ToList());

                        if (!string.IsNullOrWhiteSpace(AppState.Instance.DefaultSegment))
                        {
                            ServiceLocator.RibbonController?.SetComboText("RibSegS", AppState.Instance.DefaultSegment);
                        }
                        ServiceLocator.RibbonController?.Invalidate("RibSegS");
                    }
                    else
                    {
                        ServiceLocator.RibbonController?.ClearComboItems("RibSegS");
                        ServiceLocator.RibbonController?.SetComboText("RibSegS", string.Empty);
                    }
                }
                else
                {
                    ServiceLocator.RibbonController?.ClearComboItems("RibSegS");
                    ServiceLocator.RibbonController?.SetComboText("RibSegS", string.Empty);
                }
            }
            catch (Exception ex)
            {
                ServiceLocator.Logger?.LogException(ex, "SyncRibbonSelectionWithAppState: Failed to sync ribbon selection with app state.");
            }
        }

        /// <summary>
        /// RibClearSheet/RibClear - old monolith's ResetBalances(resetType)/
        /// BalancesReset(sheetName), small enough to fold directly into this class (see the
        /// "RibClearSheet"/"RibClear" cases in OnRibbonAction above). Re-pointed vs. the
        /// original: AppState.Instance.ExcelApp -> ServiceLocator.ExcelApp; the
        /// GLWaitWindow/CancellationHelper setup follows the same inline pattern
        /// LedgerChanged already uses elsewhere in this file (this project has no
        /// CreateAndShowWaitWindow/MessageWaitWindowAsync/ShowErrorMessageAsync helper
        /// methods - those were FinalWorkingCode-only conveniences); WinForms
        /// MessageBoxIcon/MessageBoxButtons -> WPF MessageBoxImage.
        /// </summary>
        private async Task ResetBalances(string resetType)
        {
            using var ctsHelper = new Helpers.CancellationHelper();
            var token = ctsHelper.GetToken();
            GLWaitWindow win = null;

            try
            {
                ServiceLocator.Logger?.LogDebug($"ResetBalances invoked. ResetType={resetType}");

                Excel.Worksheet activeSheet = ServiceLocator.ExcelApp?.ActiveSheet as Excel.Worksheet;

                // Regression fix: check whether the target sheet/workbook has any balance
                // formulas at all FIRST, before doing anything else (no DisableExcelSettings,
                // no progress window) - previously this check (see CommonFunctions.
                // BalanceFormulaExists, the same helper BalanceRefresh.ExistsBalanceFormulasAsync
                // uses for Refresh/Snapshot) ran only after the progress window was already
                // shown, so the user would see a "Reset Balance Formulas" window flash up
                // right before immediately being told there was nothing to reset. Now it's
                // the very first thing that happens, so the window only ever appears once
                // there's actually something to do.
                bool balancesExist = resetType == "Sheet"
                    ? activeSheet != null && CommonFunctions.BalanceFormulaExists(activeSheet.Name)
                    : ServiceLocator.ExcelApp?.ActiveWorkbook?.Worksheets != null &&
                      ServiceLocator.ExcelApp.ActiveWorkbook.Worksheets
                          .Cast<Excel.Worksheet>()
                          .Any(sheet => CommonFunctions.BalanceFormulaExists(sheet.Name));

                if (!balancesExist)
                {
                    string scope = resetType == "Sheet"
                        ? $"worksheet \"{activeSheet?.Name}\""
                        : $"workbook \"{ServiceLocator.ExcelApp?.ActiveWorkbook?.Name}\"";
                    ServiceLocator.Logger?.LogWarn($"ResetBalances: no balance formulas found in {scope} - nothing to reset.");
                    CommonFunctions.GLSenseMessage(
                        $"No balance formulas found in {scope}.",
                        MessageBoxImage.Warning);
                    return;
                }

                AppState.Instance.ResetFormulas = true;
                if (!CommonMethods.TryDisableExcelSettings("AddinEntry.ResetBalances"))
                    return;

                WpfAppManager.InvokeOnWpfThread(() =>
                {
                    win = new GLWaitWindow(ctsHelper);
                    win.SetProcessTitle("Reset Balance Formulas");
                    win.SetProcessMessage("Resetting balances...");
                    win.Show();
                    win.StartMonitoring();
                });

                await win.Dispatcher.InvokeAsync(() => win.SetProcessMessage("Checking if there are any broken links in the workbook."));
                string brokenLinks = CommonFunctions.WorkbookBrokenLinks();
                if (!string.IsNullOrWhiteSpace(brokenLinks))
                {
                    CommonFunctions.GLSenseMessage(
                        $"The workbook has broken links: {Environment.NewLine}\"{brokenLinks}\"{Environment.NewLine}Please fix them.",
                        MessageBoxImage.Error);
                    return;
                }

                token.ThrowIfCancellationRequested();

                if (resetType == "Sheet")
                {
                    if (activeSheet != null)
                    {
                        BalancesReset(activeSheet.Name);
                        token.ThrowIfCancellationRequested();
                    }
                }
                else
                {
                    if (ServiceLocator.ExcelApp?.ActiveWorkbook?.Worksheets != null)
                    {
                        foreach (Excel.Worksheet sheet in ServiceLocator.ExcelApp.ActiveWorkbook.Worksheets)
                        {
                            token.ThrowIfCancellationRequested();
                            sheet.Activate();
                            BalancesReset(sheet.Name);
                        }
                    }
                    activeSheet?.Activate();
                }
            }
            catch (OperationCanceledException)
            {
                ServiceLocator.Logger?.LogWarn("Reset Balances operation was canceled by the user.");
            }
            catch (Exception ex)
            {
                ServiceLocator.Logger?.LogException(ex, "ResetBalances: Error");
            }
            finally
            {
                CommonMethods.TryEnableExcelSettings("AddinEntry.ResetBalances");
                AppState.Instance.ResetFormulas = false;
                if (win != null)
                {
                    await win.Dispatcher.InvokeAsync(() => win.RequestClose());
                }
            }
        }

        private static void BalancesReset(string sheetName)
        {
            try
            {
                if (!(ServiceLocator.ExcelApp?.ActiveWorkbook?.Worksheets[sheetName] is Excel.Worksheet ws))
                    return;

                // Forces Excel to recalculate all formulas in the sheet by toggling the
                // calculation mode - resets any cached balance values without looping
                // through every cell.
                ws.EnableCalculation = false;
                ws.EnableCalculation = true;

                string cleanSheetName = sheetName.Replace("'", "");

                AppState.Instance.CalculatedBalances?
                    .AsEnumerable()?
                    .Where(row => string.Equals(
                        row.Field<string>("excelSheet").Replace("'", ""),
                        cleanSheetName,
                        StringComparison.OrdinalIgnoreCase))
                    .ToList()
                    .ForEach(row => row["cache"] = false);
            }
            catch (Exception ex)
            {
                ServiceLocator.Logger?.LogException(ex, $"BalancesReset: Exception encountered resetting balances in worksheet '{sheetName}'");
            }
        }

        /// <summary>
        /// IGLSenseAddin.Shutdown() - called by the host (AddinModule.ReloadAddinCore) on
        /// the OUTGOING instance right before its AppDomain is unloaded (hot-reload), and
        /// would equally apply to a genuine Excel shutdown. Tears down this instance's own
        /// WPF-side state - closes the reparented Balance Configurator window/HWND so the
        /// host's task pane is never left holding a handle into a domain that no longer
        /// exists - and flushes any pending formula-cache writes to SQLite so a reload/
        /// shutdown never silently drops dirty cache entries.
        /// </summary>
        public void Shutdown()
        {
            try
            {
                ServiceLocator.Logger?.LogDebug("AddinEntry.Shutdown invoked.");
                ServiceLocator.Logger?.FlushDebugLogs("add-in shutting down");

                try
                {
                    Views.ConfiguratorPaneHost.Close();
                }
                catch (Exception ex)
                {
                    ServiceLocator.Logger?.LogException(ex, "Shutdown: ConfiguratorPaneHost.Close failed");
                }

                try
                {
                    // ADXTaskPane mouse-wheel fix (CLAUDE.md section 24.3.5): tears down the
                    // dedicated background thread + WH_MOUSE_LL hook so neither lingers
                    // across a hot-reload swap or real shutdown - a stale hook pointing at
                    // an unloaded AppDomain's delegate would be a real crash/leak risk.
                    SuggestAppendComboBox.ShutdownMouseHook();
                }
                catch (Exception ex)
                {
                    ServiceLocator.Logger?.LogException(ex, "Shutdown: SuggestAppendComboBox.ShutdownMouseHook failed");
                }

                try
                {
                    // Terminates the WPF Application/Dispatcher started by
                    // WpfAppManager.EnsureApplication() (ShutdownMode.OnExplicitShutdown
                    // means nothing else ever stops it) - see WpfAppManager.Shutdown's own
                    // comment for why leaving this running is the most likely reason
                    // Excel.exe lingers as an orphaned process after a real close.
                    Utilities.WpfAppManager.Shutdown();
                }
                catch (Exception ex)
                {
                    ServiceLocator.Logger?.LogException(ex, "Shutdown: WpfAppManager.Shutdown failed");
                }

                try
                {
                    if (FormulaCacheManager.Instance.IsInitialized && FormulaCacheManager.Instance.HasChanges)
                    {
                        FormulaCacheManager.Instance.PersistToDatabase();
                    }
                    FormulaCacheManager.Instance.Dispose();
                }
                catch (Exception ex)
                {
                    ServiceLocator.Logger?.LogException(ex, "Shutdown: FormulaCacheManager flush/dispose failed");
                }

                // Bug fix (CLAUDE.md section 29): ServiceLocator.Reset() already existed
                // ("Reset the ServiceLocator (useful for shutdown)") but was never actually
                // called from anywhere in the codebase - this Shutdown() method is the
                // intended caller. Clears the cached IGLSenseContext/SQLiteHelper
                // references held in this AppDomain's static state. Mostly belt-and-braces
                // given AddinDomainLoader.Unload() (called right after this method returns
                // - see AddinModule.ReloadAddinCore / AddinModule_AddinBeginShutdown on the
                // host side) unloads this entire AppDomain anyway, which would discard
                // these statics regardless - but doing it explicitly here, while the
                // context/logger are still valid, means this runs before the unload rather
                // than being an implicit side effect of it, and costs nothing. Deliberately
                // last: every log line above this still needs ServiceLocator.Logger to work.
                try
                {
                    ServiceLocator.Reset();
                }
                catch (Exception ex)
                {
                    // Logger may already be unavailable at this point (Reset() just
                    // cleared it) - fall back silently rather than risk a
                    // NullReferenceException/InvalidOperationException on the way out.
                    System.Diagnostics.Debug.WriteLine("Shutdown: ServiceLocator.Reset failed: " + ex);
                }
            }
            catch (Exception ex)
            {
                ServiceLocator.Logger?.LogException(ex, "Shutdown: unexpected error");
            }
        }

        /// <summary>
        /// IGLSenseAddin.CreateConfiguratorPaneContent() - see ConfiguratorPaneHost.cs's
        /// header comment for the full HWND-reparenting rationale. Thin delegation only;
        /// idempotent (returns the existing handle if already created).
        /// </summary>
        public IntPtr CreateConfiguratorPaneContent()
        {
            return Views.ConfiguratorPaneHost.CreateContent();
        }

        /// <summary>
        /// IGLSenseAddin.RelaunchConfiguratorPane() - old monolith's
        /// GLConfiguratorPane.RelaunchPane(). Thin delegation only.
        /// </summary>
        public void RelaunchConfiguratorPane()
        {
            Views.ConfiguratorPaneHost.Relaunch();
        }

        /// <summary>
        /// IGLSenseAddin.ResetConfiguratorPaneReference() - old monolith's
        /// GLConfiguratorPane.ResetPaneReference(). Thin delegation only.
        /// </summary>
        public void ResetConfiguratorPaneReference()
        {
            Views.ConfiguratorPaneHost.ResetReference();
        }

        /// <summary>
        /// IGLSenseAddin.CloseConfiguratorPaneContent() - tears down the hosted configurator
        /// Window (used on Shutdown/logoff and before a hot-reload swap). Thin delegation
        /// only.
        /// </summary>
        public void CloseConfiguratorPaneContent()
        {
            Views.ConfiguratorPaneHost.Close();
        }

    }
}