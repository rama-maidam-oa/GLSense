using AddinExpress.MSO;
using GLSense.Contracts;
using GLSense.Loader.Core;
using System;
using System.ComponentModel;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace GLSense
{
    /// <summary>
    ///   Add-in Express Add-in Module
    /// </summary>
    [GuidAttribute("6BD53241-8482-437F-8345-C1707BF0DD31"), ProgId("GLSense.AddinModule")]
    public partial class AddinModule : AddinExpress.MSO.ADXAddinModule
    {
        // Win32 interop for returning keyboard focus to Excel's own main window after the
        // user clicks a worksheet cell while the Balance Configurator task pane is visible -
        // see adxExcelAppEvents1_SheetSelectionChange below for why this is needed.
        [DllImport("user32.dll")] private static extern IntPtr SetFocus(IntPtr hWnd);

        public AddinModule()
        {
            // No handler anywhere previously caught a truly unhandled exception in THIS
            // (host/ADX-shell) AppDomain - only WPF-dispatcher-thread exceptions are
            // covered elsewhere. Addin.Core's AppDomain has an identical hook registered
            // separately in AddinEntry.Initialize - exceptions in either domain need
            // their own hook, since UnhandledException does not cross AppDomain
            // boundaries. Registered before anything else so it's active for the
            // earliest possible failure.
            AppDomain.CurrentDomain.UnhandledException += OnAppDomainUnhandledException;

            Application.EnableVisualStyles();
            InitializeComponent();
            // Please add any initialization code to the AddinInitialize event handler

            // Bug fix (CLAUDE.md section 29, 2026-07-22): this AIPowered port never
            // subscribed to ADXAddinModule's own AddinBeginShutdown event, even though
            // FinalWorkingCode's proven-working AddinModule.cs (the old monolith) wires
            // the exact same event in its own constructor for the exact same reason -
            // see FinalWorkingCode\GLSense\AddinModule.cs's constructor and its
            // AddinModule_AddinBeginShutdown handler. Without this subscription, NONE of
            // the teardown code this project already has - AddinEntry.Shutdown() (closes
            // the reparented Balance Configurator WPF window, stops the
            // SuggestAppendComboBox mouse-hook background thread, flushes/disposes the
            // SQLite formula cache) and AddinDomainLoader.Unload() (unloads the whole
            // GLSense.Addin.Core child AppDomain) - ever ran on a genuine Excel close;
            // both were only ever reachable via the manual "Reload" ribbon button
            // (AddinModule.ReloadAddinCore). On real shutdown the child AppDomain, and
            // everything alive inside it (WPF windows, cached workbook/range COM RCWs,
            // HttpClient instances, the mouse-hook thread, an open SQLite connection),
            // was simply abandoned - the most likely reason Excel.exe lingers as an
            // orphaned background process after its visible window closes, and why a
            // fresh Excel launch afterwards can fail to load the add-in cleanly (the
            // orphaned process can still hold file locks on the shadow-copied DLLs under
            // Versions\vX\, which UpdateBootstrapper/AddinDomainLoader need to read/write
            // on the next launch).
            this.AddinBeginShutdown += AddinModule_AddinBeginShutdown;
        }

        private void OnAppDomainUnhandledException(object sender, UnhandledExceptionEventArgs e)
        {
            try
            {
                var ex = e.ExceptionObject as Exception;
                if (ex != null)
                    GlobalsEx.Context?.Logger?.LogException(ex, "AppDomain.UnhandledException (host)");
                else
                    GlobalsEx.Context?.Logger?.LogError($"AppDomain.UnhandledException (host) with non-Exception payload: {e.ExceptionObject}");

                GlobalsEx.Context?.Logger?.FlushDebugLogs("unhandled exception");
            }
            catch { /* This handler must never itself throw */ }
        }

        #region Add-in Express automatic code
 
        // Required by Add-in Express - do not modify
        // the methods within this region
 
        public override System.ComponentModel.IContainer GetContainer()
        {
            if (components == null)
                components = new System.ComponentModel.Container();
            return components;
        }
 
        [ComRegisterFunctionAttribute]
        public static void AddinRegister(Type t)
        {
            AddinExpress.MSO.ADXAddinModule.ADXRegister(t);
        }
 
        [ComUnregisterFunctionAttribute]
        public static void AddinUnregister(Type t)
        {
            AddinExpress.MSO.ADXAddinModule.ADXUnregister(t);
        }
 
        public override void UninstallControls()
        {
            base.UninstallControls();
        }

        #endregion

        public static new AddinModule CurrentInstance 
        {
            get
            {
                return AddinExpress.MSO.ADXAddinModule.CurrentInstance as AddinModule;
            }
        }

        public Excel._Application ExcelApp
        {
            get
            {
                return (HostApplication as Excel._Application);
            }
        }
        private RibbonController _ribbonController;
        private void AddinModule_OnRibbonLoaded(object sender, IRibbonUI ribbon)
        {
            IntPtr hwnd = new IntPtr(ExcelApp.Hwnd);

            // Create context with ExcelApp
            GlobalsEx.Context = new GLSenseContext(ExcelApp);
            GlobalsEx.Loader = new AddinDomainLoader();

            _ribbonController = new RibbonController(
                    AddinModule.CurrentInstance,
                    ribbon,
                    GlobalsEx.Context?.Logger
                );

            if (GlobalsEx.Context != null)
            {
                GlobalsEx.Context.SetRibbonController(_ribbonController);
            }

            GlobalsEx.Context.ExcelHandle = hwnd;

            // Decide which version of Addin.Core to load - and get it onto disk if it
            // isn't there yet - BEFORE creating the AppDomain. See UpdateBootstrapper
            // (GLSense.Loader.Core) for the full decision tree: extracts a zip sitting in
            // the Manifest folder alongside manifest.json if present, otherwise reuses
            // whatever's already installed in Versions\vX\ if that's present - folder-only,
            // no network step (see CLAUDE.md section 17). This replaces what used to be a
            // hardcoded GlobalsEx.Context.Version = "11.1.0" - that string had no
            // connection to the manifest at all, so nothing previously read from
            // manifest.json ever actually influenced which DLLs got loaded.
            GlobalsEx.Context.Logger?.LogDebug("AddinModule_OnRibbonLoaded: resolving version to load via UpdateBootstrapper.");
            string resolvedVersion = new UpdateBootstrapper().ResolveVersionToLoad(GlobalsEx.Context);

            if (string.IsNullOrEmpty(resolvedVersion))
            {
                GlobalsEx.Context.Logger?.LogError("AddinModule_OnRibbonLoaded: UpdateBootstrapper could not resolve a version to load (no zip/manifest.json in the Manifest folder and no usable local install). Skipping Addin.Core load - the add-in will be unavailable until this is resolved and Excel is restarted.");
                MessageBox.Show(
                    "GLSense could not find a usable add-in version. Make sure GLSense.Addin.Core has been " +
                    "built at least once (its post_build.cmd publishes a zip + manifest.json into the " +
                    "Manifest folder). The add-in will not load this session - check the logs, then restart " +
                    "Excel once resolved.",
                    "GLSense",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
                return;
            }

            GlobalsEx.Context.Version = resolvedVersion;

            // GLAbout's "Build Date" reads Context.ReleaseDate via ServiceLocator - this was
            // never actually set anywhere before (the property existed on IGLSenseContext but
            // nothing ever assigned it), so the About window always showed a blank/"Unknown"
            // build date. Paths.LatestReleaseDate is the same manifest.json entry
            // UpdateBootstrapper just read to resolve resolvedVersion (paths.Refresh() inside
            // ResolveVersionToLoad guarantees it reflects the actual current file on disk), so
            // this is always in sync with the version that's about to load.
            GlobalsEx.Context.ReleaseDate = GlobalsEx.Context.Paths?.LatestReleaseDate;
            GlobalsEx.Context.Logger?.LogDebug($"AddinModule_OnRibbonLoaded: resolvedVersion={resolvedVersion}, releaseDate={GlobalsEx.Context.ReleaseDate}");

            // Load Addin.Core

            GlobalsEx.Addin = GlobalsEx.Loader.Load(GlobalsEx.Context);
        }

        // Guards RibReload_OnClick against re-entrancy (e.g. an accidental rapid
        // double-click) - not a true multi-thread lock, since Excel/ADX only ever raises
        // ribbon OnClick and Excel Application events on the single main STA thread, one
        // at a time, so this handler can never genuinely run concurrently with itself or
        // any other handler in this file.
        private bool _reloadInProgress;

        private void RibReload_OnClick(object sender, IRibbonControl control, bool pressed)
        {
            GlobalsEx.Context?.Logger?.LogDebug($"RibReload_OnClick fired (pressed={pressed})");
            if (_reloadInProgress) return;

            // Hot-reload has no old-monolith equivalent to port from (a single-AppDomain
            // add-in always required a full Excel restart to pick up code changes) - this
            // is new infrastructure. Destructive/risky enough (see the caveat below) that
            // it warrants an explicit confirmation rather than firing on a single click.
            var confirm = MessageBox.Show(
                "This reloads GLSense.Addin.Core.dll from disk without restarting Excel." +
                Environment.NewLine + Environment.NewLine +
                "Make sure you've rebuilt it first (its post_build.cmd copies the new DLL " +
                "into the versions folder this reads from)." +
                Environment.NewLine + Environment.NewLine +
                "Any drilldown/refresh/snapshot currently in progress will be interrupted - " +
                "this does not wait for background work to finish before unloading. Continue?",
                "Reload GLSense Add-in",
                MessageBoxButtons.YesNo,
                MessageBoxIcon.Warning);

            if (confirm != DialogResult.Yes) return;

            _reloadInProgress = true;
            try
            {
                ReloadAddinCore();
            }
            finally
            {
                _reloadInProgress = false;
            }
        }

        /// <summary>
        /// Hot-reload orchestration: tear down the current Addin.Core instance/AppDomain
        /// and load a fresh one, re-pointing GlobalsEx.Addin so every ribbon/Excel-event
        /// handler in this file (which only ever go through GlobalsEx.Addin?./GlobalsEx.
        /// Context?. null-conditionals) transparently picks up the new instance - see the
        /// comment on adxExcelAppEvents1's region above for why AddinModule itself is
        /// never reloaded and why this swap-a-pointer approach is safe for a COM event
        /// sink Excel holds a long-lived reference to.
        ///
        /// Known limitation (deliberately not solved here - would need an in-flight-
        /// operation counter threaded through every async entry point in Addin.Core,
        /// which is a much larger change than "add a reload button"): AppDomain.Unload
        /// forcibly aborts any thread still executing inside the domain being unloaded.
        /// If a background Task.Run from a drilldown/refresh/UDF is genuinely mid-flight
        /// (as opposed to just idle) when this runs, it can be aborted mid-operation,
        /// including potentially mid-COM-call. The confirmation dialog in RibReload_OnClick
        /// exists specifically to surface this risk to the user rather than hide it -
        /// reload only when nothing else is actively running.
        /// </summary>
        private void ReloadAddinCore()
        {
            try
            {
                GlobalsEx.Context?.Logger?.LogDebug("Reload requested via ribbon (RibReload).");

                var oldAddin = GlobalsEx.Addin;
                var loader = GlobalsEx.Loader;

                // 1. Let the outgoing instance tear down its own WPF-side state (currently:
                //    closes the reparented Balance Configurator window/HWND) BEFORE the
                //    AppDomain unloads, so the host's task pane is never left holding a
                //    handle into a domain that no longer exists.
                try
                {
                    oldAddin?.Shutdown();
                }
                catch (Exception ex)
                {
                    GlobalsEx.Context?.Logger?.LogException(ex, "ReloadAddinCore: old instance Shutdown failed");
                }

                // 2. Null the pointer before unloading. Every call site in this file already
                //    goes through GlobalsEx.Addin?./GlobalsEx.Context?. null-conditionals, so
                //    this leaves the add-in transparently "temporarily unavailable" instead
                //    of holding a reference into a domain mid-unload.
                GlobalsEx.Addin = null;

                // 3. Unload the old AppDomain - releases the old GLSense.Addin.Core.dll (and
                //    its shadow-copy) so a freshly rebuilt copy in the versions folder can be
                //    picked up by the next Load().
                loader?.Unload(GlobalsEx.Context);

                // 3b. Re-run UpdateBootstrapper before reloading. Post-build no longer
                //     xcopies a fresh DLL straight into Versions\vX\ (see
                //     GLSense.Addin.Core\post_build.cmd / CLAUDE.md section 17) - the ONLY
                //     way new DLLs reach that folder now is UpdateBootstrapper extracting a
                //     zip that post-build drops directly into the local Manifest folder
                //     (folder-only flow, no network/local-host involved - see CLAUDE.md
                //     section 17). Without this call, Reload would just re-load whatever was
                //     already sitting in Versions\vX\ from the last successful bootstrap -
                //     i.e. it would silently keep reloading stale code after a rebuild,
                //     never picking up new changes.
                string resolvedVersion = new UpdateBootstrapper().ResolveVersionToLoad(GlobalsEx.Context);
                if (string.IsNullOrEmpty(resolvedVersion))
                {
                    GlobalsEx.Context?.Logger?.LogError("ReloadAddinCore: UpdateBootstrapper could not resolve a version to reload (no zip/manifest.json in the Manifest folder and no usable local install).");
                    MessageBox.Show(
                        "Reload failed - no usable add-in version was found. Make sure GLSense.Addin.Core " +
                        "has been rebuilt (its post_build.cmd publishes a zip + manifest.json into the " +
                        "Manifest folder), then try again.",
                        "Reload GLSense Add-in",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Error);
                    return;
                }
                GlobalsEx.Context.Version = resolvedVersion;
                GlobalsEx.Context.ReleaseDate = GlobalsEx.Context.Paths?.LatestReleaseDate;
                GlobalsEx.Context?.Logger?.LogDebug($"ReloadAddinCore: resolvedVersion={resolvedVersion}, releaseDate={GlobalsEx.Context.ReleaseDate}");

                // 4. Load a fresh AppDomain + AddinEntry instance from whatever
                //    UpdateBootstrapper just resolved/extracted, and re-point
                //    GlobalsEx.Addin. Reuses the SAME GlobalsEx.Loader/GlobalsEx.Context
                //    instances (host-side, never torn down) - AddinDomainLoader.Load only
                //    needs calling again, no other host-side re-initialization
                //    (RibbonController, ExcelHandle) is required since none of that lives
                //    in the domain being replaced.
                GlobalsEx.Addin = loader?.Load(GlobalsEx.Context);

                if (GlobalsEx.Addin != null)
                {
                    GlobalsEx.Context?.Logger?.LogDebug("Reload complete - GlobalsEx.Addin re-pointed to a fresh instance.");
                }
                else
                {
                    GlobalsEx.Context?.Logger?.LogError("Reload failed - GlobalsEx.Addin is null after Load(). The add-in is unavailable until Excel is restarted.");
                    MessageBox.Show(
                        "Reload failed - the add-in could not be loaded. Check the logs. " +
                        "Excel will need to be restarted to recover.",
                        "Reload GLSense Add-in",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Error);
                }
            }
            catch (Exception ex)
            {
                GlobalsEx.Context?.Logger?.LogException(ex, "ReloadAddinCore");
                MessageBox.Show(
                    $"Reload failed: {ex.Message}{Environment.NewLine}Excel may need to be restarted to recover.",
                    "Reload GLSense Add-in",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// ADXAddinModule's own shutdown lifecycle event - fires once when Excel is
        /// genuinely closing (also fires if the user disables the add-in from Excel's
        /// COM Add-ins dialog without closing Excel). This is distinct from
        /// ReloadAddinCore's hot-reload path above: AddinModule itself (this class,
        /// the ADXExcelAppEvents sink, the ribbon) is NEVER unloaded/reloaded (see the
        /// big comment above the "Excel Application events" region below for why), so
        /// this handler only needs to tear down the CHILD AppDomain/instance - reusing
        /// the exact same two calls ReloadAddinCore already makes on its outgoing
        /// instance, in the same order and for the same reason (let Addin.Core close its
        /// own WPF-side state and flush its SQLite cache BEFORE the AppDomain that owns
        /// them is unloaded out from under it), rather than duplicating that logic here.
        ///
        /// Before this fix, this event was never subscribed at all (see the constructor
        /// above) - meaning on a real Excel close, GlobalsEx.Addin.Shutdown() and
        /// GlobalsEx.Loader.Unload() never ran, and the child AppDomain (with every WPF
        /// window, COM RCW, HttpClient, background thread, and open SQLite connection it
        /// held) was simply abandoned. See CLAUDE.md section 29 for the full
        /// investigation - this was the most likely root cause of Excel.exe lingering as
        /// an orphaned background process after closing, and of the add-in then failing
        /// to load on the next Excel launch (the orphaned process can still hold file
        /// locks on the shadow-copied DLLs under Versions\vX\).
        /// </summary>
        private void AddinModule_AddinBeginShutdown(object sender, EventArgs e)
        {
            try
            {
                GlobalsEx.Context?.Logger?.LogDebug("AddinModule_AddinBeginShutdown fired - tearing down GLSense.Addin.Core before Excel exits.");
                GlobalsEx.Context?.Logger?.FlushDebugLogs("add-in shutting down");

                // Defensive: stop forwarding any further Excel events into the add-in
                // while it's mid-teardown. AddinModule/adxExcelAppEvents1 are never
                // unloaded, but explicitly unsubscribing here matches FinalWorkingCode's
                // own proven AddinModule_AddinBeginShutdown
                // (UnsubscribeFromAllExcelEvents) and guarantees no late-arriving Excel
                // event tries to call into GlobalsEx.Addin after Shutdown() has already
                // run below (Excel can still raise a handful of events - e.g.
                // WorkbookBeforeSave on an autosave - in the brief window before it fully
                // exits).
                try
                {
                    if (adxExcelAppEvents1 != null)
                    {
                        adxExcelAppEvents1.SheetActivate -= adxExcelAppEvents1_SheetActivate;
                        adxExcelAppEvents1.SheetBeforeDoubleClick -= adxExcelAppEvents1_SheetBeforeDoubleClick;
                        adxExcelAppEvents1.SheetChange -= adxExcelAppEvents1_SheetChange;
                        adxExcelAppEvents1.SheetFollowHyperlink -= adxExcelAppEvents1_SheetFollowHyperlink;
                        adxExcelAppEvents1.SheetSelectionChange -= adxExcelAppEvents1_SheetSelectionChange;
                        adxExcelAppEvents1.WorkbookActivate -= adxExcelAppEvents1_WorkbookActivate;
                        adxExcelAppEvents1.WorkbookBeforeSave -= adxExcelAppEvents1_WorkbookBeforeSave;
                    }
                }
                catch (Exception ex)
                {
                    GlobalsEx.Context?.Logger?.LogException(ex, "AddinModule_AddinBeginShutdown: unsubscribing adxExcelAppEvents1 failed");
                }

                // 1. Let the current instance tear down its own WPF-side state (closes
                //    the reparented Balance Configurator window/HWND, stops the
                //    SuggestAppendComboBox mouse-hook thread, flushes/disposes the SQLite
                //    formula cache) BEFORE the AppDomain unloads - same reasoning as
                //    ReloadAddinCore's step 1 above.
                try
                {
                    GlobalsEx.Addin?.Shutdown();
                }
                catch (Exception ex)
                {
                    GlobalsEx.Context?.Logger?.LogException(ex, "AddinModule_AddinBeginShutdown: GlobalsEx.Addin.Shutdown failed");
                }

                // 2. Null the pointer before unloading - same reasoning as
                //    ReloadAddinCore's step 2 (every call site already goes through
                //    GlobalsEx.Addin?. null-conditionals).
                GlobalsEx.Addin = null;

                // 3. Unload the AppDomain - releases GLSense.Addin.Core.dll's shadow copy
                //    and everything still alive inside that domain (WPF windows, cached
                //    Excel COM RCWs, HttpClient instances, timers/threads) so Excel's own
                //    process shutdown isn't left waiting on (or silently abandoning) any
                //    of it.
                try
                {
                    GlobalsEx.Loader?.Unload(GlobalsEx.Context);
                }
                catch (Exception ex)
                {
                    GlobalsEx.Context?.Logger?.LogException(ex, "AddinModule_AddinBeginShutdown: GlobalsEx.Loader.Unload failed");
                }

                GlobalsEx.Loader = null;

                GlobalsEx.Context?.Logger?.LogDebug("AddinModule_AddinBeginShutdown: teardown complete.");
            }
            catch (Exception ex)
            {
                // Never let an exception escape a COM lifecycle event back into
                // Excel/ADX - same rule PORTING_GUIDE.md section 2 states for
                // AddinEntry/RibbonController call sites.
                GlobalsEx.Context?.Logger?.LogException(ex, "AddinModule_AddinBeginShutdown: unexpected error");
            }
        }

        private void RibLogin_OnClick(object sender, IRibbonControl control, bool pressed)
        {
            GlobalsEx.Context?.Logger?.LogDebug($"RibLogin_OnClick fired (pressed={pressed})");
            _ribbonController?.ExecuteAction("Login");
        }

        private void RibLogout_OnClick(object sender, IRibbonControl control, bool pressed)
        {
            GlobalsEx.Context?.Logger?.LogDebug($"RibLogout_OnClick fired (pressed={pressed})");
            _ribbonController?.ExecuteAction("Logout");
        }

        private void RibDBL1_OnAction(object sender, IRibbonControl control, bool pressed)
        {
            GlobalsEx.Context?.Logger?.LogDebug($"RibDBL1_OnAction fired (pressed={pressed})");
            // Old monolith: opened GLLoginDetails modally via SafeInvokeWpf. GLLoginDetails
            // now lives in Addin.Core - see AddinEntry's "ShowLoginDetails" case.
            _ribbonController?.ExecuteAction("ShowLoginDetails");
        }

        private void RibGetCube_OnClick(object sender, IRibbonControl control, bool pressed)
        {
            GlobalsEx.Context?.Logger?.LogDebug($"RibGetCube_OnClick fired (pressed={pressed})");
            // No parameter to pass across the AppDomain boundary, so this goes through
            // the same ExecuteAction(buttonId)->OnRibbonAction(action, null) fallback path
            // as RibLogin/RibLogout/RibCellHighlight above (see RibbonController.ExecuteAction).
            _ribbonController?.ExecuteAction("ShowCubeDetails");
        }

        private void Ribledger_OnChange(object sender, IRibbonControl Control, string text)
        {
            GlobalsEx.Context?.Logger?.LogDebug($"Ribledger_OnChange fired (text={text})");
            // ExecuteAction(buttonId) takes no payload, so - like the Excel Application
            // event handlers above (OnExcelEvent) - this goes straight through the mutable
            // GlobalsEx.Addin hot-swap pointer to pass the selected ledger name across the
            // AppDomain boundary (a string marshals fine; see PORTING_GUIDE.md's boundary
            // rules). Fire-and-forget: AddinEntry.OnRibbonAction is synchronous, but the
            // actual ledger-switch work it kicks off is async and self-contained.
            GlobalsEx.Addin?.OnRibbonAction("LedgerChanged", text);
        }

        private void RibAccount_OnClick(object sender, IRibbonControl control, bool pressed)
        {
            GlobalsEx.Context?.Logger?.LogDebug($"RibAccount_OnClick fired (pressed={pressed})");
            // Old monolith: opened GLSegmentValues modally via SafeInvokeWpf.
            // GLSegmentValues now lives in Addin.Core - see AddinEntry.ShowSegmentValues().
            // (Group H content port - wired once GLSegmentValues/SegmentSelectorViewModel
            // land in Addin.Core.)
            _ribbonController?.ExecuteAction("ShowSegmentValues");
        }

        private void RibRollerGroup_OnClick(object sender, IRibbonControl control, bool pressed)
        {
            GlobalsEx.Context?.Logger?.LogDebug($"RibRollerGroup_OnClick fired (pressed={pressed})");
            // Old monolith: opened GLRollerGroups modally via SafeInvokeWpf.
            // GLRollerGroups now lives in Addin.Core - see AddinEntry.ShowRollerGroups().
            // (Group H content port - wired once GLRollerGroups/SimpleSegmentViewModel
            // land in Addin.Core.)
            _ribbonController?.ExecuteAction("ShowRollerGroups");
        }

        private void RibLOVs_OnClick(object sender, IRibbonControl control, bool pressed)
        {
            GlobalsEx.Context?.Logger?.LogDebug($"RibLOVs_OnClick fired (pressed={pressed})");
            // Old monolith: opened GLLOVs modally via SafeInvokeWpf. GLLOVs now lives in
            // Addin.Core - see AddinEntry.ShowLOVs(). (Group H content port - wired once
            // GLLOVs/GLLovViewModel land in Addin.Core.)
            _ribbonController?.ExecuteAction("ShowLOVs");
        }

        private void RibFSG_OnClick(object sender, IRibbonControl control, bool pressed)
        {
            GlobalsEx.Context?.Logger?.LogDebug($"RibFSG_OnClick fired (pressed={pressed})");
            // Old monolith: RibFSG_OnClick toggled the GLConfiguratorPane task pane
            // in/out of view (create-on-first-use, otherwise flip Visible). The task
            // pane itself (GLConfiguratorPane) is a host-only WinForms/ADXExcelTaskPane
            // type, so this whole method stays host-side exactly like the original -
            // only blpane.RelaunchPane() (below) actually crosses into Addin.Core, via
            // the HWND-reparenting bridge (see GLConfiguratorPane.cs). DisplayConfigurator
            // is the host-only reentrancy guard replacing the old AppState.Instance.
            // displayConfigurator (see GLConfiguratorPane.cs's doc comment on why it no
            // longer needs to cross the AppDomain boundary).
            try
            {
                GLConfiguratorPane.DisplayConfigurator = true;

                var blpane = GetPaneInstance();

                if (blpane != null)
                {
                    blpane.Visible = !blpane.Visible;
                    if (blpane.Visible)
                    {
                        _ = blpane.RelaunchPane();
                    }
                }
                else
                {
                    adxExcelTaskPanesCollectionItem1.Position = AddinExpress.XL.ADXExcelTaskPanePosition.Right;
                    blpane = (GLConfiguratorPane)adxExcelTaskPanesCollectionItem1.CreateTaskPaneInstance();
                    if (blpane != null)
                    {
                        blpane.Show();
                        blpane.Visible = true;
                    }
                }
            }
            catch (Exception ex)
            {
                GlobalsEx.Context?.Logger?.LogException(ex, "RibFSG_OnClick");
            }
            finally
            {
                GLConfiguratorPane.DisplayConfigurator = false;
            }
        }

        /// <summary>
        /// Host-only lookup of the current GLConfiguratorPane task-pane instance (old
        /// monolith's AddinModule.GetPaneInstance(), used by RibFSG_OnClick, the
        /// SheetSelectionChange handler below, and RibbonController.HideConfiguratorPane/
        /// RelaunchConfiguratorPaneIfVisible - see RibbonController.cs). Public because
        /// RibbonController (Group H) already calls it via `_addinModule?.GetPaneInstance()`.
        /// </summary>
        public GLConfiguratorPane GetPaneInstance()
        {
            try
            {
                return (GLConfiguratorPane)adxExcelTaskPanesCollectionItem1.TaskPaneInstance;
            }
            catch (Exception ex)
            {
                GlobalsEx.Context?.Logger?.LogException(ex, "GetPaneInstance");
                return null;
            }
        }

        private void RibLiveCalc_OnClick(object sender, IRibbonControl control, bool pressed)
        {
            GlobalsEx.Context?.Logger?.LogDebug($"RibLiveCalc_OnClick fired (pressed={pressed})");
            // Old monolith: AppState.Instance.SingleRefresh = RibLiveCalc.Pressed. AppState
            // now lives in Addin.Core's AppDomain; `pressed` is the same value
            // RibLiveCalc.Pressed would read at this point, and bool crosses fine via
            // OnRibbonAction (same pattern as Group G's toggle buttons).
            GlobalsEx.Addin?.OnRibbonAction("SingleRefreshToggled", pressed);
        }

        private void RibSegS_OnChange(object sender, IRibbonControl Control, string text)
        {
            GlobalsEx.Context?.Logger?.LogDebug($"RibSegS_OnChange fired (text={text})");
            // ExecuteAction(buttonId) takes no payload, so - like Ribledger_OnChange
            // above - this goes straight through the mutable GlobalsEx.Addin hot-swap
            // pointer to pass the selected segment name across the AppDomain boundary (a
            // string marshals fine). AddinEntry.SegmentChanged re-derives the
            // ribbon-combo index from DataRepository.GetSegments(...) instead of trying
            // to marshal an index computed from ADX's RibSegS.Items (host-only type).
            GlobalsEx.Addin?.OnRibbonAction("SegmentChanged", text);
        }

        private void RigSegDiscover_OnClick(object sender, IRibbonControl control, bool pressed)
        {
            GlobalsEx.Context?.Logger?.LogDebug($"RigSegDiscover_OnClick fired (pressed={pressed})");
            // Validation (DefaultSegment/SegmentPickedIndex + active cell value) and the
            // GLSegmentDiscovery window itself now live in Addin.Core's AddinEntry.
            // ShowSegmentDiscovery() - GLSegmentDiscovery can't be `new`'d from the host
            // since it lives in the hot-reload AppDomain.
            _ribbonController?.ExecuteAction("ShowSegmentDiscovery");
        }

        private void RibSegProperty_OnClick(object sender, IRibbonControl control, bool pressed)
        {
            GlobalsEx.Context?.Logger?.LogDebug($"RibSegProperty_OnClick fired (pressed={pressed})");
            // SegmentDiscoverer.SegmentAction("Property") in the old monolith.
            _ribbonController?.ExecuteAction("RibSegProperty");
        }

        private void RibSegmentExpand_OnClick(object sender, IRibbonControl control, bool pressed)
        {
            GlobalsEx.Context?.Logger?.LogDebug($"RibSegmentExpand_OnClick fired (pressed={pressed})");
            // Was a "Hierarchy" menu with two items (RibExpandAll/RibbonExpand1Level ->
            // SegmentDiscoverer.SegmentAction("HierarchyAll"/"Hierarchy1Level") directly).
            // Now a single button that opens GLExpandOptions, where the user picks the
            // level (All/1 Level) and fill direction (Rows/Columns) before
            // SegmentDiscoverer.SegmentAction gets called from that dialog.
            _ribbonController?.ExecuteAction("ShowExpandOptions");
        }

        private void RibExpodeAll_OnClick(object sender, IRibbonControl control, bool pressed)
        {
            GlobalsEx.Context?.Logger?.LogDebug($"RibExpodeAll_OnClick fired (pressed={pressed})");
            // SegmentDiscoverer.SegmentAction("ExplodeAll") in the old monolith.
            _ribbonController?.ExecuteAction("RibExpodeAll");
        }

        private void RibbonExplode1Level_OnClick(object sender, IRibbonControl control, bool pressed)
        {
            GlobalsEx.Context?.Logger?.LogDebug($"RibbonExplode1Level_OnClick fired (pressed={pressed})");
            // SegmentDiscoverer.SegmentAction("Explode1Level") in the old monolith.
            _ribbonController?.ExecuteAction("RibbonExplode1Level");
        }

        private void RibDiscoverPeriod_OnClick(object sender, IRibbonControl control, bool pressed)
        {
            GlobalsEx.Context?.Logger?.LogDebug($"RibDiscoverPeriod_OnClick fired (pressed={pressed})");
            // PeriodsDiscoverer.FillPeriods() in the old monolith.
            _ribbonController?.ExecuteAction("RibDiscoverPeriod");
        }

        private void RibRefreshAll_OnClick(object sender, IRibbonControl control, bool pressed)
        {
            GlobalsEx.Context?.Logger?.LogDebug($"RibRefreshAll_OnClick fired (pressed={pressed})");
            // BalanceRefresh.RefreshingBalancesAsync("Refresh", "Sheet") in the old monolith.
            _ribbonController?.ExecuteAction("RibRefreshAll");
        }

        private void RibRefreshBook_OnClick(object sender, IRibbonControl control, bool pressed)
        {
            GlobalsEx.Context?.Logger?.LogDebug($"RibRefreshBook_OnClick fired (pressed={pressed})");
            // BalanceRefresh.RefreshingBalancesAsync("Refresh", "Book") in the old monolith.
            _ribbonController?.ExecuteAction("RibRefreshBook");
        }

        private void RibRefreshRange_OnClick(object sender, IRibbonControl control, bool pressed)
        {
            GlobalsEx.Context?.Logger?.LogDebug($"RibRefreshRange_OnClick fired (pressed={pressed})");
            // RangeRefresher.RibRefreshRange_OnClick() (Group F extraction) in Addin.Core.
            _ribbonController?.ExecuteAction("RibRefreshRange");
        }

        private void RibClearSheet_OnClick(object sender, IRibbonControl control, bool pressed)
        {
            GlobalsEx.Context?.Logger?.LogDebug($"RibClearSheet_OnClick fired (pressed={pressed})");
            // AddinEntry.ResetBalances("Sheet") (Group F extraction) in Addin.Core.
            _ribbonController?.ExecuteAction("RibClearSheet");
        }

        private void RibClear_OnClick(object sender, IRibbonControl control, bool pressed)
        {
            GlobalsEx.Context?.Logger?.LogDebug($"RibClear_OnClick fired (pressed={pressed})");
            // AddinEntry.ResetBalances("Book") (Group F extraction) in Addin.Core.
            _ribbonController?.ExecuteAction("RibClear");
        }

        private void RibHighlight_OnClick(object sender, IRibbonControl control, bool pressed)
        {
            GlobalsEx.Context?.Logger?.LogDebug($"RibHighlight_OnClick fired (pressed={pressed})");
            // BalanceHighlighter.RibHighlight_OnClick() (Group F extraction) in Addin.Core.
            _ribbonController?.ExecuteAction("RibHighlight");
        }

        private void RibCellHighlight_OnClick(object sender, IRibbonControl control, bool pressed)
        {
            GlobalsEx.Context?.Logger?.LogDebug($"RibCellHighlight_OnClick fired (pressed={pressed})");
            _ribbonController?.ExecuteAction(RibbonControlIds.RibCellHighlight);
        }

        private void RibHideRows_OnClick(object sender, IRibbonControl control, bool pressed)
        {
            GlobalsEx.Context?.Logger?.LogDebug($"RibHideRows_OnClick fired (pressed={pressed})");
            // RowVisibilityProcessor.RibHideRows_OnClick() (Group F extraction) in Addin.Core.
            _ribbonController?.ExecuteAction("RibHideRows");
        }

        private void RibUnHideRows_OnClick(object sender, IRibbonControl control, bool pressed)
        {
            GlobalsEx.Context?.Logger?.LogDebug($"RibUnHideRows_OnClick fired (pressed={pressed})");
            // RowVisibilityProcessor.RibUnHideRows_OnClick() (Group F extraction) in Addin.Core.
            _ribbonController?.ExecuteAction("RibUnHideRows");
        }

        private void RibSnapShot_OnClick(object sender, IRibbonControl control, bool pressed)
        {
            GlobalsEx.Context?.Logger?.LogDebug($"RibSnapShot_OnClick fired (pressed={pressed})");
            // Old monolith read mode/isSubmit directly off RibSnapWorksheet.Pressed and
            // RibSnapSubmit.AsRibbonCheckBox.Pressed at click time (both host-only ADX
            // reads), so that's preserved here; only the crossing mechanism changed -
            // ExecuteAction(buttonId) takes no payload, so this goes straight through
            // GlobalsEx.Addin?.OnRibbonAction (like Ribledger_OnChange/RibSegS_OnChange)
            // with both values packed into one pipe-delimited string. See AddinEntry.
            // RunSnapshot(string) for the Addin.Core side.
            string mode = RibSnapWorksheet.Pressed ? "Sheet" : "Book";
            ADXRibbonCheckBox ribbonCheckBox = RibSnapSubmit.AsRibbonCheckBox;
            bool isSubmit = ribbonCheckBox != null && ribbonCheckBox.Pressed;

            GlobalsEx.Addin?.OnRibbonAction("RunSnapshot", $"{mode}|{isSubmit}");
        }

        private void RibSnapWorksheet_OnClick(object sender, IRibbonControl control, bool pressed)
        {
            GlobalsEx.Context?.Logger?.LogDebug($"RibSnapWorksheet_OnClick fired (pressed={pressed})");
            // Old monolith: guard + RibSnapWorkbook.Pressed = !pressed, both moved into
            // AddinEntry.ToggleSnapMode (the AppState.Instance guard fields live in the
            // Addin.Core AppDomain now, so the host can't check them directly).
            GlobalsEx.Addin?.OnRibbonAction("SnapWorksheetToggled", pressed);
        }

        private void RibSnapWorkbook_OnClick(object sender, IRibbonControl control, bool pressed)
        {
            GlobalsEx.Context?.Logger?.LogDebug($"RibSnapWorkbook_OnClick fired (pressed={pressed})");
            // Mirror of RibSnapWorksheet_OnClick above - see AddinEntry.ToggleSnapMode.
            GlobalsEx.Addin?.OnRibbonAction("SnapWorkbookToggled", pressed);
        }

        private void RibSnapSubmit_OnClick(object sender, IRibbonControl control, bool pressed)
        {
            GlobalsEx.Context?.Logger?.LogDebug($"RibSnapSubmit_OnClick fired (pressed={pressed})");
            // Old monolith: AppState.Instance.SnapshotJob = RibSnapSubmit.Pressed - just
            // forwards the click-time pressed state across the AppDomain boundary now.
            GlobalsEx.Addin?.OnRibbonAction("SnapSubmitToggled", pressed);
        }

        private void RibBalanceDD_OnClick(object sender, IRibbonControl control, bool pressed)
        {
            GlobalsEx.Context?.Logger?.LogDebug($"RibBalanceDD_OnClick fired (pressed={pressed})");
            // Old monolith: RunBalanceDrilldownAsync("BL") - reads Selection/ExcelApp
            // entirely on the Addin.Core side, so no parameter needs to cross the
            // AppDomain boundary (see AddinEntry.RunBalanceDrilldown).
            _ribbonController?.ExecuteAction("RunBalanceDrilldownBL");
        }

        private void RibBalanceJournalDD_OnClick(object sender, IRibbonControl control, bool pressed)
        {
            GlobalsEx.Context?.Logger?.LogDebug($"RibBalanceJournalDD_OnClick fired (pressed={pressed})");
            // Old monolith: RunBalanceDrilldownAsync("BL_JL").
            _ribbonController?.ExecuteAction("RunBalanceDrilldownBLJL");
        }

        private void RibBalanceSubLedgerDD_OnClick(object sender, IRibbonControl control, bool pressed)
        {
            GlobalsEx.Context?.Logger?.LogDebug($"RibBalanceSubLedgerDD_OnClick fired (pressed={pressed})");
            // Old monolith: RunBalanceDrilldownAsync("BL_SL").
            _ribbonController?.ExecuteAction("RunBalanceDrilldownBLSL");
        }

        private void RibJournalDD_OnClick(object sender, IRibbonControl control, bool pressed)
        {
            GlobalsEx.Context?.Logger?.LogDebug($"RibJournalDD_OnClick fired (pressed={pressed})");
            // Old monolith: new DrilldownJl(ExcelApp, external).ProcessJLDrilldown() -
            // now dispatched via AddinEntry.RunJournalDrilldown().
            _ribbonController?.ExecuteAction("RibJournalDD");
        }

        private void RibBalancesDDToSubLedger_OnClick(object sender, IRibbonControl control, bool pressed)
        {
            GlobalsEx.Context?.Logger?.LogDebug($"RibBalancesDDToSubLedger_OnClick fired (pressed={pressed})");
            // Old monolith: RunDrilldownAsync("BLDD_SL") -> new DrilldownJl(ExcelApp,
            // external, "BLDD_SL").ProcessJLDrilldown() - now dispatched via
            // AddinEntry.RunJournalDrilldown("BLDD_SL"). Previously dropped during the
            // DD_JL.cs port (ddType parameter was hardcoded to "JL"); restored as part of
            // the end-to-end completeness pass.
            _ribbonController?.ExecuteAction("RibBalancesDDToSubLedger");
        }

        private void RibBalancesDDToUnified_OnClick(object sender, IRibbonControl control, bool pressed)
        {
            GlobalsEx.Context?.Logger?.LogDebug($"RibBalancesDDToUnified_OnClick fired (pressed={pressed})");
            // Old monolith: RunDrilldownAsync("BLDD_UF") -> new DrilldownJl(ExcelApp,
            // external, "BLDD_UF").ProcessJLDrilldown() - now dispatched via
            // AddinEntry.RunJournalDrilldown("BLDD_UF"). See note above.
            _ribbonController?.ExecuteAction("RibBalancesDDToUnified");
        }

        private void RibSubledgerDD_OnClick(object sender, IRibbonControl control, bool pressed)
        {
            GlobalsEx.Context?.Logger?.LogDebug($"RibSubledgerDD_OnClick fired (pressed={pressed})");
            // Old monolith: new DrilldownSl(ExcelApp, external).ProcessSLDrilldown() -
            // now dispatched via AddinEntry.RunSubledgerDrilldown().
            _ribbonController?.ExecuteAction("RibSubledgerDD");
        }

        private void RibTotaDD_OnClick(object sender, IRibbonControl control, bool pressed)
        {
            GlobalsEx.Context?.Logger?.LogDebug($"RibTotaDD_OnClick fired (pressed={pressed})");
            // Old monolith: RunBalanceDrilldownAsync("UF").
            _ribbonController?.ExecuteAction("RunBalanceDrilldownUF");
        }

        private void RibDDConfiguration_OnClick(object sender, IRibbonControl control, bool pressed)
        {
            GlobalsEx.Context?.Logger?.LogDebug($"RibDDConfiguration_OnClick fired (pressed={pressed})");
            // Old monolith: opened GLDrilldownCustomization modally via SafeInvokeWpf.
            // GLDrilldownCustomization now lives in Addin.Core - see AddinEntry.
            // ShowDrilldownCustomization().
            _ribbonController?.ExecuteAction("ShowDrilldownCustomization");
        }

        private void RibDDDeleteConfiguration_OnClick(object sender, IRibbonControl control, bool pressed)
        {
            GlobalsEx.Context?.Logger?.LogDebug($"RibDDDeleteConfiguration_OnClick fired (pressed={pressed})");
            // FinalWorkingCode's RibDDDeleteConfiguration_OnClick does the delete + message
            // inline (single-project monolith); here it's dispatched through to Addin.Core
            // the same way RibDDConfiguration_OnClick above dispatches to
            // ShowDrilldownCustomization() - see AddinEntry.DeleteDrilldownCustomization().
            _ribbonController?.ExecuteAction("DeleteDrilldownCustomization");
        }

        private void RibDrillJobs_OnClick(object sender, IRibbonControl control, bool pressed)
        {
            GlobalsEx.Context?.Logger?.LogDebug($"RibDrillJobs_OnClick fired (pressed={pressed})");
            // Old monolith: opened GLJobsMonitor modally via SafeInvokeWpf. GLJobsMonitor
            // now lives in Addin.Core - see AddinEntry.ShowJobsMonitor().
            _ribbonController?.ExecuteAction("ShowJobsMonitor");
        }

        private void RibSegmentEnabledFlag_OnClick(object sender, IRibbonControl control, bool pressed)
        {
            GlobalsEx.Context?.Logger?.LogDebug($"RibSegmentEnabledFlag_OnClick fired (pressed={pressed})");
            // LaunchSegmentWindow("ENABLEDFLAG") in the old monolith - see AddinEntry.
            // ShowSegmentWindow(string) for the shared dispatch (Group C).
            _ribbonController?.ExecuteAction("ShowSegmentEnabledFlag");
        }

        private void RibSummaryFlag_OnClick(object sender, IRibbonControl control, bool pressed)
        {
            GlobalsEx.Context?.Logger?.LogDebug($"RibSummaryFlag_OnClick fired (pressed={pressed})");
            // LaunchSegmentWindow("SUMMARYFLAG") in the old monolith.
            _ribbonController?.ExecuteAction("ShowSegmentSummaryFlag");
        }

        private void RibSegment_OnClick(object sender, IRibbonControl control, bool pressed)
        {
            GlobalsEx.Context?.Logger?.LogDebug($"RibSegment_OnClick fired (pressed={pressed})");
            // LaunchSegmentWindow("DESCRIPTION") in the old monolith.
            _ribbonController?.ExecuteAction("ShowSegmentDescription");
        }

        private void RibNextSegment_OnClick(object sender, IRibbonControl control, bool pressed)
        {
            GlobalsEx.Context?.Logger?.LogDebug($"RibNextSegment_OnClick fired (pressed={pressed})");
            // LaunchSegmentWindow("NEXTSEGMENT") in the old monolith.
            _ribbonController?.ExecuteAction("ShowNextSegment");
        }

        private void RibPreviousSegment_OnClick(object sender, IRibbonControl control, bool pressed)
        {
            GlobalsEx.Context?.Logger?.LogDebug($"RibPreviousSegment_OnClick fired (pressed={pressed})");
            // LaunchSegmentWindow("PREVIOUSSEGMENT") in the old monolith.
            _ribbonController?.ExecuteAction("ShowPreviousSegment");
        }

        private void RibSegmentDFF_OnClick(object sender, IRibbonControl control, bool pressed)
        {
            GlobalsEx.Context?.Logger?.LogDebug($"RibSegmentDFF_OnClick fired (pressed={pressed})");
            // LaunchSegmentWindow("DFF") in the old monolith.
            _ribbonController?.ExecuteAction("ShowSegmentDFF");
        }

        private void RibSegmentAccountType_OnClick(object sender, IRibbonControl control, bool pressed)
        {
            GlobalsEx.Context?.Logger?.LogDebug($"RibSegmentAccountType_OnClick fired (pressed={pressed})");
            // LaunchSegmentWindow("ACCOUNTTYPE") in the old monolith.
            _ribbonController?.ExecuteAction("ShowSegmentAccountType");
        }

        private void RibPeriod_OnClick(object sender, IRibbonControl control, bool pressed)
        {
            GlobalsEx.Context?.Logger?.LogDebug($"RibPeriod_OnClick fired (pressed={pressed})");
            // Opens GLGetPeriod (period + numeric offset picker).
            _ribbonController?.ExecuteAction("ShowPeriod");
        }

        private void RibPeriodbyDate_OnClick(object sender, IRibbonControl control, bool pressed)
        {
            GlobalsEx.Context?.Logger?.LogDebug($"RibPeriodbyDate_OnClick fired (pressed={pressed})");
            // Opens GLGetPeriodByDate (date + ledger + numeric offset picker).
            _ribbonController?.ExecuteAction("ShowPeriodByDate");
        }

        private void RibPeriodbyYear_OnClick(object sender, IRibbonControl control, bool pressed)
        {
            GlobalsEx.Context?.Logger?.LogDebug($"RibPeriodbyYear_OnClick fired (pressed={pressed})");
            // Opens GLGetPeriodByYear (period year + period num picker).
            _ribbonController?.ExecuteAction("ShowPeriodByYear");
        }

        private void RibPeriodNum_OnClick(object sender, IRibbonControl control, bool pressed)
        {
            GlobalsEx.Context?.Logger?.LogDebug($"RibPeriodNum_OnClick fired (pressed={pressed})");
            // LaunchPeriodDetails("NUM") in the old monolith - see AddinEntry.
            // ShowPeriodDetails(string) for the shared dispatch (Group C).
            _ribbonController?.ExecuteAction("ShowPeriodNum");
        }

        private void RibPeriodQtr_OnClick(object sender, IRibbonControl control, bool pressed)
        {
            GlobalsEx.Context?.Logger?.LogDebug($"RibPeriodQtr_OnClick fired (pressed={pressed})");
            // LaunchPeriodDetails("QTR") in the old monolith.
            _ribbonController?.ExecuteAction("ShowPeriodQtr");
        }

        private void RibPeriodYear_OnClick(object sender, IRibbonControl control, bool pressed)
        {
            GlobalsEx.Context?.Logger?.LogDebug($"RibPeriodYear_OnClick fired (pressed={pressed})");
            // LaunchPeriodDetails("YEAR") in the old monolith.
            _ribbonController?.ExecuteAction("ShowPeriodYear");
        }

        private void RibPeriodStart_OnClick(object sender, IRibbonControl control, bool pressed)
        {
            GlobalsEx.Context?.Logger?.LogDebug($"RibPeriodStart_OnClick fired (pressed={pressed})");
            // LaunchPeriodStarEnd("START") in the old monolith - see AddinEntry.
            // ShowPeriodStartEnd(string) for the shared dispatch (Group C).
            _ribbonController?.ExecuteAction("ShowPeriodStart");
        }

        private void RibPeriodEnd_OnClick(object sender, IRibbonControl control, bool pressed)
        {
            GlobalsEx.Context?.Logger?.LogDebug($"RibPeriodEnd_OnClick fired (pressed={pressed})");
            // LaunchPeriodStarEnd("END") in the old monolith.
            _ribbonController?.ExecuteAction("ShowPeriodEnd");
        }

        private void RibDailyRate_OnClick(object sender, IRibbonControl control, bool pressed)
        {
            GlobalsEx.Context?.Logger?.LogDebug($"RibDailyRate_OnClick fired (pressed={pressed})");
            // Opens GLDailyRates.
            _ribbonController?.ExecuteAction("ShowDailyRate");
        }

        private void Riburl_OnClick(object sender, IRibbonControl control, bool pressed)
        {
            GlobalsEx.Context?.Logger?.LogDebug($"Riburl_OnClick fired (pressed={pressed})");
            // Old monolith: opened GLServerConfiguration modally via SafeInvokeWpf.
            // GLServerConfiguration now lives in Addin.Core - see AddinEntry's
            // "ShowServerConfiguration" case (Group I).
            _ribbonController?.ExecuteAction("ShowServerConfiguration");
        }

        private void RibUserConfig_OnClick(object sender, IRibbonControl control, bool pressed)
        {
            GlobalsEx.Context?.Logger?.LogDebug($"RibUserConfig_OnClick fired (pressed={pressed})");
            // Old monolith: opened GLUserConfig modally via SafeInvokeWpf. GLUserConfig
            // now lives in Addin.Core - see AddinEntry's "ShowUserConfig" case (Group I).
            _ribbonController?.ExecuteAction("ShowUserConfig");
        }

        private void RibDebug_OnClick(object sender, IRibbonControl control, bool pressed)
        {
            // Old monolith: AppState.Instance.DebugLogs = pressed, plus a start/flush log
            // message. AppState now lives in Addin.Core's AppDomain; `pressed` is the same
            // value RibDebug.Pressed would read here, and bool crosses fine via
            // OnRibbonAction (same pattern as RibLiveCalc/RibSnapWorksheet).
            GlobalsEx.Addin?.OnRibbonAction("DebugLogsToggled", pressed);
            GlobalsEx.Context?.Logger?.LogDebug($"RibDebug_OnClick fired (pressed={pressed})");

            // Diagnostic requested directly: dump the REAL Ribledger/RibSegS controls'
            // live state right now, using direct field access on this host-side class (no
            // reflection, no AppDomain crossing - this IS the actual object Excel's Ribbon
            // renders). Click Debug again right after reproducing "dropdown won't open" to
            // get an instant answer: does the live control still hold the items we set at
            // that exact moment, or has something cleared it by then?
            DumpRibbonComboState();
        }

        private void DumpRibbonComboState()
        {
            try
            {
                string ledgerCaptions = "";
                foreach (var item in Ribledger.Items)
                {
                    var captionProp = item.GetType().GetProperty("Caption");
                    string caption = captionProp?.GetValue(item)?.ToString() ?? "?";
                    ledgerCaptions += (ledgerCaptions.Length > 0 ? ", " : "") + caption;
                }
                // LogWarn, not LogDebug: LogDebug is a no-op the instant DebugMode is
                // false (Logger.LogDebug checks "if (!DebugMode) return;" before writing
                // anything), and RibDebug_OnClick's OnRibbonAction("DebugLogsToggled", ...)
                // call right before this flips that flag - so on a "turn debug OFF" click
                // this dump would silently vanish (exactly what happened in the previous
                // repro: the dump ran, but produced nothing, because logging had just been
                // switched off a few lines earlier). LogWarn/LogError don't check DebugMode
                // at all, so this always gets written regardless of the toggle's direction.
                GlobalsEx.Context?.Logger?.LogWarn($"[RibbonDiagnostic] Ribledger: Enabled={Ribledger.Enabled}, Visible={Ribledger.Visible}, Text='{Ribledger.Text}', Items.Count={Ribledger.Items.Count}, Items=[{ledgerCaptions}]");
            }
            catch (Exception ex)
            {
                GlobalsEx.Context?.Logger?.LogException(ex, "[RibbonDiagnostic] Ribledger dump failed");
            }

            try
            {
                string segCaptions = "";
                foreach (var item in RibSegS.Items)
                {
                    var captionProp = item.GetType().GetProperty("Caption");
                    string caption = captionProp?.GetValue(item)?.ToString() ?? "?";
                    segCaptions += (segCaptions.Length > 0 ? ", " : "") + caption;
                }
                GlobalsEx.Context?.Logger?.LogWarn($"[RibbonDiagnostic] RibSegS: Enabled={RibSegS.Enabled}, Visible={RibSegS.Visible}, Text='{RibSegS.Text}', Items.Count={RibSegS.Items.Count}, Items=[{segCaptions}]");
            }
            catch (Exception ex)
            {
                GlobalsEx.Context?.Logger?.LogException(ex, "[RibbonDiagnostic] RibSegS dump failed");
            }
        }

        private void RibVersionCheck_OnClick(object sender, IRibbonControl control, bool pressed)
        {
            GlobalsEx.Context?.Logger?.LogDebug($"RibVersionCheck_OnClick fired (pressed={pressed})");
            // Old monolith: AppState.Instance.VersionCheck = pressed - same reasoning as
            // RibDebug_OnClick above.
            GlobalsEx.Addin?.OnRibbonAction("VersionCheckToggled", pressed);
        }

        private void RibAbout_OnClick(object sender, IRibbonControl control, bool pressed)
        {
            GlobalsEx.Context?.Logger?.LogDebug($"RibAbout_OnClick fired (pressed={pressed})");
            // Old monolith: opened GLAbout modally via SafeInvokeWpf. GLAbout now lives in
            // Addin.Core - see AddinEntry's "ShowAbout" case (Group I).
            _ribbonController?.ExecuteAction("ShowAbout");
        }

        private void RibHelp_OnClick(object sender, IRibbonControl control, bool pressed)
        {
            GlobalsEx.Context?.Logger?.LogDebug($"RibHelp_OnClick fired (pressed={pressed})");
            // Old monolith: built the help URL from AppState.Instance.LoginUrl/LoginToken
            // (both live in Addin.Core's AppDomain) and called Process.Start directly.
            // Process.Start has no Excel-COM/UI-thread affinity requirement, so the whole
            // thing (including the "please log in" guard) moved into Addin.Core - see
            // AddinEntry.ShowHelp() (Group I). ExecuteAction(buttonId) takes no payload,
            // and this needs none, so it goes through the same path as RibDDConfiguration/
            // RibDrillJobs above.
            _ribbonController?.ExecuteAction("ShowHelp");
        }

        // ------------------------------------------------------------------
        // Excel Application events (ADXExcelAppEvents component, wired in the
        // designer). This component - and these handlers - live in AddinModule,
        // which is NEVER unloaded/reloaded. That's deliberate: a COM event sink is a
        // long-lived delegate Excel holds onto directly. If it pointed into the
        // hot-reloadable child AppDomain (GLSense.Addin.Core) and that domain got
        // unloaded, Excel would be left holding a dead reference. So: subscribe once,
        // here, and just forward through the mutable GlobalsEx.Addin pointer - it gets
        // re-pointed after every reload, so Excel never needs to know a reload happened.
        //
        // IMPORTANT: only primitive/serializable values (string, bool, ...) are passed
        // to OnExcelEvent - never a live Excel COM object or an AddinExpress event-args
        // object. Those aren't [Serializable] and would throw a SerializationException
        // when the call is marshaled across the AppDomain boundary. Extract whatever you
        // need (name, address, ...) into a plain string/bool here first.
        // ------------------------------------------------------------------

        private void adxExcelAppEvents1_SheetActivate(object sender, object hostObj)
        {
            GlobalsEx.Context?.Logger?.LogDebug($"SheetActivate event fired (sheet={GetSheetName(hostObj)})");

            try
            {
                GlobalsEx.Addin?.OnExcelEvent("SheetActivate", new object[] { GetSheetName(hostObj) });
            }
            catch (Exception ex)
            {
                GlobalsEx.Context?.Logger?.LogException(ex, "adxExcelAppEvents1_SheetActivate");
            }

            // Old monolith's Helpers\RibbonStateHelper.ApplySheetActiveState/ProcessActiveSheet
            // - contextually enables/disables the Drilldowns ribbon buttons based on the
            // now-active sheet (a live Excel.Worksheet inspection, so - same reasoning as
            // DrillDownsStart/IsValidDrilldownSheet above - this stays entirely host-side;
            // nothing here needs to cross into Addin.Core). Previously dropped when this
            // handler was first ported (logging-only stub); restored as part of the
            // end-to-end completeness pass. Reuses IsBalancesDrilldown/IsJournalsDrilldown/
            // GetA1Text/GetDrilldownSheetMarkerValue/IsValidDrilldownSheet, all already
            // ported for the double-click dispatch above.
            try
            {
                ApplySheetActiveState();
            }
            catch (Exception ex)
            {
                GlobalsEx.Context?.Logger?.LogException(ex, "adxExcelAppEvents1_SheetActivate(RibbonState)");
            }
        }

        private void ApplySheetActiveState()
        {
            // Old monolith's RibbonStateHelper.ShouldProcessActiveState: ExcelApp != null &&
            // AppState.Instance.IsLoginCompleted. AppState lives in Addin.Core now, so the
            // login half of this guard uses RibbonController.IsLoggedIn instead (tracks the
            // last SetState(...) call - see RibbonController.cs) rather than crossing the
            // AppDomain boundary just to ask.
            if (ExcelApp == null || _ribbonController == null || !_ribbonController.IsLoggedIn)
                return;

            if (ExcelApp.ActiveSheet is not Excel.Worksheet activeSheet)
                return;

            ApplyDrilldownState(IsValidDrilldownSheet(activeSheet));

            if (CheckForBalanceFormulas(activeSheet))
            {
                EnableBalanceDrilldownControls();
            }
            else
            {
                DisableAllDrilldownControls();
                EnableDrilldownBasedOnSheetType(activeSheet);
            }
        }

        private void ApplyDrilldownState(bool isDrilldown)
        {
            try
            {
                // RibUserConfig deliberately left OUT of this list. Ported fix from
                // FinalWorkingCode\GLSense\Helpers\RibbonStateHelper.cs: it's a login/cube-
                // session-level settings window, not something tied to whichever sheet happens
                // to be active, so it must not be toggled by per-sheet state changes at all.
                // Previously it was disabled here whenever the active sheet was a drilldown
                // result sheet (isDrilldown=true), so opening a drilldown result tab would grey
                // out User Preferences even while fully logged in with a cube selected. Its
                // enabled/disabled state is owned entirely by RibbonControlIds.cs's login-state
                // lists (already includes RibUserConfig in LoggedInEnabledControls and the
                // various logged-out/disabled lists) - enabled whenever logged in, disabled when
                // logged out, irrespective of sheet.
                string[] controls =
                {
                    "RibDBL1", "RibGetCube", "Ribledger", "RibAccount", "RibRollerGroup", "RibLOVs", "RibFSG", "RibHideRows", "RibUnHideRows",
                    "RibLiveCalc", "RibSegS", "RibSegmentDiscover", "RigSegDiscover", "RibSegProperty", "RibSegmentExpand",
                    "RibSegmentExplode", "RibExpodeAll", "RibbonExplode1Level", "RibDiscoverPeriod", "RibAsFormula", "RibRefreshRange", "RibRefreshAll", "RibRefreshBook",
                    "RibClearSheet", "RibClear", "RibHighlight", "RibCellHighlight", "RibSnapShot", "RibSnapWorksheet", "RibSnapWorkbook", "RibSnapSubmit",
                    "RibFunctionsMenu", "RibSegmentEnabledFlag", "RibSummaryFlag", "RibSegment", "RibNextSegment",
                    "RibPreviousSegment", "RibSegmentDFF", "RibPeriod", "RibPeriodbyDate", "RibPeriodbyYear", "RibPeriodNum", "RibPeriodQtr", "RibPeriodYear",
                    "RibPeriodStart", "RibPeriodEnd", "RibDailyRate", "RibVersionCheck", "RibHelp"
                };

                if (isDrilldown)
                {
                    _ribbonController?.DisableControls(controls);
                }
                else
                {
                    // Old monolith's own "else" list is identical to the disable list above
                    // plus RibDDConfiguration/RibDrillJobs (Customization/Monitoring buttons
                    // that a drilldown-result sheet doesn't disable, but that DO get
                    // positively re-enabled here) - ported verbatim, asymmetry included.
                    _ribbonController?.EnableControls(controls);
                    _ribbonController?.EnableControls(new[] { "RibDDConfiguration", "RibDrillJobs" });
                }
            }
            catch (Exception ex)
            {
                GlobalsEx.Context?.Logger?.LogException(ex, "ApplyDrilldownState");
            }
        }

        private static bool CheckForBalanceFormulas(Excel.Worksheet activeSheet)
        {
            long formulaCount = CountFormulaCells(activeSheet);
            if (formulaCount <= 0)
                return false;

            return HasBalanceFormulaAnywhere(activeSheet);
        }

        private static long CountFormulaCells(Excel.Worksheet sheet)
        {
            try
            {
                return ((Excel.Range)sheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeFormulas)).Count;
            }
            catch (COMException ex) when ((uint)ex.ErrorCode == 0x800A03EC)
            {
                // Excel's SpecialCells throws this specific COMException ("No cells were
                // found") whenever the sheet has zero cells of the requested type - e.g. a
                // blank sheet or one with no formulas anywhere. That's a normal, expected
                // outcome (not a bug), so it's swallowed silently here instead of being
                // logged as an exception on every formula-free sheet.
                return 0;
            }
            catch (Exception ex)
            {
                GlobalsEx.Context?.Logger?.LogException(ex, "AddinModule.CountFormulaCells");
                return 0;
            }
        }

        private static bool HasBalanceFormulaAnywhere(Excel.Worksheet sheet)
        {
            try
            {
                Excel.Range foundRange = sheet.Cells.Find(
                    AppConstants_GlBal,
                    Type.Missing,
                    Excel.XlFindLookIn.xlFormulas,
                    Excel.XlLookAt.xlPart,
                    Excel.XlSearchOrder.xlByRows,
                    Excel.XlSearchDirection.xlNext,
                    false);

                return foundRange != null;
            }
            catch (Exception ex)
            {
                GlobalsEx.Context?.Logger?.LogException(ex, "AddinModule.HasBalanceFormulaAnywhere");
                return false;
            }
        }

        private void EnableBalanceDrilldownControls()
        {
            RibBalanceDD.Enabled = true;
            RibBalanceJournalDD.Enabled = true;
            RibBalanceSubLedgerDD.Enabled = true;
            RibJournalDD.Enabled = false;
            RibSubledgerDD.Enabled = false;
            // Ported gating from FinalWorkingCode\GLSense\Helpers\RibbonStateHelper.cs: Unified
            // Drilldown fails server-side for view-based/EBS cubes, so grey it out up front
            // instead of only failing after the user clicks it. _ribbonController.IsCubeViewBased
            // is pushed from Addin.Core's AppState.SelectedCube setter (see RibbonController.cs).
            RibTotaDD.Enabled = !(_ribbonController?.IsCubeViewBased ?? false);
            RibBalancesDDToSubLedger.Enabled = false;
            RibBalancesDDToUnified.Enabled = false;
        }

        private void DisableAllDrilldownControls()
        {
            RibBalanceDD.Enabled = false;
            RibBalanceJournalDD.Enabled = false;
            RibBalanceSubLedgerDD.Enabled = false;
            RibJournalDD.Enabled = false;
            RibSubledgerDD.Enabled = false;
            RibTotaDD.Enabled = false;
            RibBalancesDDToSubLedger.Enabled = false;
            RibBalancesDDToUnified.Enabled = false;
        }

        private void EnableDrilldownBasedOnSheetType(Excel.Worksheet sheet)
        {
            string a1Text = GetA1Text(sheet);
            string sheetName = sheet.Name;
            string markerSheetName = GetDrilldownSheetMarkerValue(sheet);

            bool isJournalDrilldown = IsBalancesDrilldown(a1Text, sheetName, markerSheetName);
            bool isSubledgerDrilldown = IsJournalsDrilldown(a1Text, sheetName, markerSheetName);

            // Same EBS/view-based gating as EnableBalanceDrilldownControls above - Balances
            // Drilldown to Unified fails server-side for such cubes too.
            RibBalancesDDToUnified.Enabled = isJournalDrilldown && !(_ribbonController?.IsCubeViewBased ?? false);
            RibBalancesDDToSubLedger.Enabled = isJournalDrilldown;
            RibJournalDD.Enabled = isJournalDrilldown;
            RibSubledgerDD.Enabled = isSubledgerDrilldown;
        }

        private void adxExcelAppEvents1_SheetBeforeDoubleClick(object sender, ADXExcelSheetBeforeEventArgs e)
        {
            GlobalsEx.Context?.Logger?.LogDebug($"SheetBeforeDoubleClick event fired (sheet={GetSheetName(e?.Sheet)}, range={GetRangeAddress(e?.Range)})");

            try
            {
                // NOTE: e.Sheet / e.Target / e.Cancel follow AddinExpress's documented naming
                // convention for this event (mirroring Excel's own
                // Worksheet_BeforeDoubleClick(Target, Cancel)) - double check against
                // IntelliSense if this doesn't compile as-is; adjust the property names below.
                string sheetName = GetSheetName(e?.Sheet);
                string rangeAddress = GetRangeAddress(e?.Range);

                bool proceed = GlobalsEx.Addin?.OnExcelEvent("SheetBeforeDoubleClick", new object[] { sheetName, rangeAddress }) ?? true;

                // Only touch e.Cancel if the add-in explicitly said not to proceed - leave
                // Excel's default double-click behavior alone otherwise.
                if (!proceed && e != null)
                    e.Cancel = true;

                // Old monolith's DrillDownsStart/ResolveDrilldownType/RunDrilldown/
                // RunBalancePrecedentDrilldown - the classification itself has to stay
                // host-side: it inspects the live Excel.Range/Worksheet the event handed
                // us (formula text, sheet name, A1 marker cell), and a live COM object
                // can't be handed to Addin.Core the way a WPF FrameworkElement can't -
                // simplest and safest to just never let a live Range/Worksheet cross the
                // AppDomain boundary at all (see PORTING_GUIDE.md). Once classified, only
                // primitive strings (ddType + a fully-qualified external address built
                // HERE, from e.Target specifically - NOT re-derived from Selection in
                // Addin.Core, since Excel hasn't necessarily moved the selection to the
                // double-clicked cell yet at this point in the event) cross via
                // GlobalsEx.Addin?.OnRibbonAction("RunDrilldownExternal", "ddType|external")
                // - see AddinEntry.RunDrilldownByExternalAddress for the Addin.Core side.
                if (e?.Range is not Excel.Range currRange || e?.Sheet is not Excel.Worksheet currSheet)
                    return;

                bool actionTaken = false;
                DrillDownsStart(currRange, currSheet, ref actionTaken);

                // Combine with (rather than overwrite) whatever the generic OnExcelEvent
                // forward above already decided, so a future OnExcelEvent-side cancel
                // reason is never silently undone here.
                if (e != null) e.Cancel = e.Cancel || actionTaken;
            }
            catch (Exception ex)
            {
                GlobalsEx.Context?.Logger?.LogException(ex, "adxExcelAppEvents1_SheetBeforeDoubleClick");
            }
        }

        /// <summary>
        /// Old monolith's AddinModule.DrillDownsStart - host-side classification only
        /// (live Range/Worksheet inspection); dispatches the actual drilldown work into
        /// Addin.Core via a plain ddType+external string pair once classified. See the
        /// double-click TODO this resolves in AddinEntry.cs's OnExcelEvent
        /// "SheetBeforeDoubleClick" case (that generic forward above is left in place for
        /// logging - this is the actual dispatch, restored from the old monolith).
        /// </summary>
        private static void DrillDownsStart(Excel.Range cellRange, Excel.Worksheet currentSheet, ref bool noAction)
        {
            try
            {
                if (HasNoBalanceFormula(cellRange))
                {
                    DispatchDrilldown("EP", cellRange);
                    noAction = true;
                    return;
                }

                string ddType = ResolveDrilldownType(cellRange, currentSheet);
                if (string.IsNullOrWhiteSpace(ddType))
                    return;

                DispatchDrilldown(ddType, cellRange);
                noAction = true;
            }
            catch (Exception ex)
            {
                GlobalsEx.Context?.Logger?.LogException(ex, "DrillDownsStart");
            }
            finally
            {
                try
                {
                    cellRange?.Worksheet?.ClearArrows();
                }
                catch (Exception ex)
                {
                    GlobalsEx.Context?.Logger?.LogException(ex, "DrillDownsStart: ClearArrows");
                }
            }
        }

        private static void DispatchDrilldown(string ddType, Excel.Range cellRange)
        {
            string external = BuildExternalAddress(cellRange);
            if (string.IsNullOrEmpty(external)) return;

            GlobalsEx.Addin?.OnRibbonAction("RunDrilldownExternal", $"{ddType}|{external}");
        }

        private static bool HasNoBalanceFormula(Excel.Range cellRange)
        {
            return TryGetSingleCellFormula(cellRange, out string formulaString) &&
                   formulaString.IndexOf(AppConstants_GlBal, StringComparison.OrdinalIgnoreCase) < 0;
        }

        private static bool HasBalanceFormula(Excel.Range cellRange)
        {
            return TryGetSingleCellFormula(cellRange, out string formulaString) &&
                   formulaString.IndexOf(AppConstants_GlBal, StringComparison.OrdinalIgnoreCase) >= 0;
        }

        private static string ResolveDrilldownType(Excel.Range cellRange, Excel.Worksheet currentSheet)
        {
            if (HasBalanceFormula(cellRange))
                return "BL";

            string a1Text = GetA1Text(currentSheet);
            if (string.IsNullOrWhiteSpace(a1Text))
                return string.Empty;

            string sheetName = currentSheet.Name;
            string markerSheetName = GetDrilldownSheetMarkerValue(currentSheet);

            if (IsBalancesDrilldown(a1Text, sheetName, markerSheetName))
                return "JL";

            if (IsJournalsDrilldown(a1Text, sheetName, markerSheetName))
                return "SL";

            return string.Empty;
        }

        private static string GetA1Text(Excel.Worksheet currentSheet)
        {
            var a1Val = currentSheet?.Range["A1"]?.Value2;
            return a1Val?.ToString().Trim().ToUpperInvariant() ?? string.Empty;
        }

        // Mirrors GLSense.Addin.Core.AppConstants.DrilldownSheetMarkerCellAddress ("XEZ5") -
        // duplicated host-side for the same reason as AppConstants_GlBal above.
        private const string DrilldownSheetMarkerCellAddress = "XEZ5";

        private static string GetDrilldownSheetMarkerValue(Excel.Worksheet sheet)
        {
            try
            {
                var markerCell = (Excel.Range)sheet.Range[DrilldownSheetMarkerCellAddress];
                var value = markerCell?.Value2 ?? markerCell?.Value;
                return value?.ToString() ?? string.Empty;
            }
            catch (Exception ex)
            {
                GlobalsEx.Context?.Logger?.LogException(ex, "AddinModule.GetDrilldownSheetMarkerValue");
                return string.Empty;
            }
        }

        private static bool IsBalancesDrilldown(string a1Text, string sheetName, string markerSheetName)
        {
            if (a1Text.Equals("BALANCES DRILLDOWN", StringComparison.OrdinalIgnoreCase))
                return true;

            return MatchesBalancesPattern(sheetName) || MatchesBalancesPattern(markerSheetName);
        }

        private static bool IsJournalsDrilldown(string a1Text, string sheetName, string markerSheetName)
        {
            if (a1Text.Equals("JOURNALS DRILLDOWN", StringComparison.OrdinalIgnoreCase))
                return true;

            return MatchesJournalsPattern(sheetName) || MatchesJournalsPattern(markerSheetName);
        }

        private static bool MatchesBalancesPattern(string sheetName)
        {
            return !string.IsNullOrWhiteSpace(sheetName) &&
                   sheetName.IndexOf("_BL_", StringComparison.OrdinalIgnoreCase) >= 0 &&
                   sheetName.IndexOf("_JL_", StringComparison.OrdinalIgnoreCase) < 0 &&
                   sheetName.IndexOf("_SL_", StringComparison.OrdinalIgnoreCase) < 0 &&
                   sheetName.IndexOf("_CM_", StringComparison.OrdinalIgnoreCase) >= 0;
        }

        private static bool MatchesJournalsPattern(string sheetName)
        {
            return !string.IsNullOrWhiteSpace(sheetName) &&
                   sheetName.IndexOf("_BL_", StringComparison.OrdinalIgnoreCase) >= 0 &&
                   sheetName.IndexOf("_JL_", StringComparison.OrdinalIgnoreCase) >= 0 &&
                   sheetName.IndexOf("_SL_", StringComparison.OrdinalIgnoreCase) < 0;
        }

        /// <summary>
        /// Old monolith's Helpers\ExcelExternalRef.BuildExternalAddress, duplicated
        /// host-side (like TryGetSingleCellFormula above) since the host must never take
        /// a compile-time dependency on Addin.Core, where the real (fuller) copy already
        /// lives and is used for every ribbon-triggered drilldown (see DD_BL.cs/
        /// AddinEntry.RunBalanceDrilldown).
        /// </summary>
        private static string BuildExternalAddress(Excel.Range rng)
        {
            if (rng == null) return null;
            return rng.Address[
                RowAbsolute: true,
                ColumnAbsolute: true,
                ReferenceStyle: Excel.XlReferenceStyle.xlA1,
                External: true
            ];
        }

        private void adxExcelAppEvents1_SheetChange(object sender, object sheet, object range)
        {
            GlobalsEx.Context?.Logger?.LogDebug($"SheetChange event fired (sheet={GetSheetName(sheet)}, range={GetRangeAddress(range)})");

            try
            {
                GlobalsEx.Addin?.OnExcelEvent("SheetChange", new object[] { GetSheetName(sheet), GetRangeAddress(range) });
            }
            catch (Exception ex)
            {
                GlobalsEx.Context?.Logger?.LogException(ex, "adxExcelAppEvents1_SheetChange");
            }

            // Old monolith's adxExcelAppEvents1_SheetChange: if the cell that just changed
            // now contains a GLSense_GetBalance formula, re-apply the "LoggedIn" ribbon
            // state (this is what re-enables the Drilldowns/Balance-formula-dependent
            // buttons the moment a user types/pastes a balance formula into a fresh sheet).
            // Pure Excel-COM + host-only ribbon-state concern, stays host-side entirely -
            // previously dropped (logging-only stub); restored as part of the end-to-end
            // completeness pass.
            try
            {
                if (ExcelApp == null || _ribbonController == null || !_ribbonController.IsLoggedIn)
                    return;

                if (range is not Excel.Range rng)
                    return;

                if (TryGetSingleCellFormula(rng, out string formula) &&
                    formula.IndexOf(AppConstants_GlBal, StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    _ribbonController.SetState("LoggedIn");
                }
            }
            catch (Exception ex)
            {
                GlobalsEx.Context?.Logger?.LogException(ex, "adxExcelAppEvents1_SheetChange(RibbonState)");
            }
        }

        private void adxExcelAppEvents1_SheetFollowHyperlink(object sender, object sheet, object hyperlink)
        {
            GlobalsEx.Context?.Logger?.LogDebug($"SheetFollowHyperlink event fired (sheet={GetSheetName(sheet)})");

            try
            {
                GlobalsEx.Addin?.OnExcelEvent("SheetFollowHyperlink", new object[] { GetSheetName(sheet), SafeToString(hyperlink) });
            }
            catch (Exception ex)
            {
                GlobalsEx.Context?.Logger?.LogException(ex, "adxExcelAppEvents1_SheetFollowHyperlink");
            }

            // Old monolith's adxExcelAppEvents1_SheetFollowHyperlink dispatch (custom
            // drilldown vs. journal-attachment link). IsValidDrilldownSheet/
            // IsCustomDrilldownHyperlink/header-label lookup are pure Excel-COM
            // inspection of the live sheet/hyperlink/range the event handed us, so - same
            // reasoning as DrillDownsStart above - this classification stays host-side;
            // only plain strings cross into Addin.Core's CustomDrilldown.RunCustomDrilldown/
            // JournalAttachments.RunJournalAttachmentFlow (both already ported, see their
            // file headers for the exact preconditions/parameters they expect - matched
            // exactly here).
            try
            {
                if (sheet is not Excel.Worksheet sht || !IsValidDrilldownSheet(sht))
                    return;

                if (hyperlink is not Excel.Hyperlink hyprLink)
                    return;

                if (IsCustomDrilldownHyperlink(hyprLink))
                {
                    DispatchCustomDrilldownHyperlink(sht, hyprLink);
                }
                else if (hyprLink.Parent is Excel.Range hyperlinkRange)
                {
                    DispatchJournalAttachmentHyperlink(hyperlinkRange);
                }
            }
            catch (Exception ex)
            {
                GlobalsEx.Context?.Logger?.LogException(ex, "adxExcelAppEvents1_SheetFollowHyperlink(Drilldown)");
            }
        }

        private static bool IsValidDrilldownSheet(Excel.Worksheet sht)
        {
            return sht?.ListObjects.Count > 0 && sht.ListObjects[1].Name.StartsWith("ORB_DD_");
        }

        private static bool IsCustomDrilldownHyperlink(Excel.Hyperlink hyprLink)
        {
            return !string.IsNullOrEmpty(hyprLink?.ScreenTip) &&
                   hyprLink.ScreenTip.IndexOf("CUSTOM DRILLDOWN", StringComparison.OrdinalIgnoreCase) >= 0;
        }

        private static void DispatchCustomDrilldownHyperlink(Excel.Worksheet sht, Excel.Hyperlink hyprLink)
        {
            try
            {
                Excel.Range rng = sht.Range[hyprLink.SubAddress];
                Excel.Range rngNew = sht.Cells[5, rng.Column] as Excel.Range;
                object cellValue = rngNew?.Value2;

                if (cellValue == null) return;

                string headerLabel = cellValue.ToString() ?? string.Empty;
                if (string.IsNullOrWhiteSpace(headerLabel))
                {
                    GlobalsEx.Context?.Logger?.LogWarn("Exiting the hyperlink sub since the label header is null or empty");
                    return;
                }

                string external = BuildExternalAddress(rng);
                if (string.IsNullOrEmpty(external)) return;

                GlobalsEx.Addin?.OnRibbonAction("RunCustomDrilldown", new string[] { sht.Name, external, headerLabel });
            }
            catch (Exception ex)
            {
                GlobalsEx.Context?.Logger?.LogException(ex, "DispatchCustomDrilldownHyperlink: Unable to get the header label.");
            }
        }

        private static void DispatchJournalAttachmentHyperlink(Excel.Range hyperlinkRange)
        {
            try
            {
                long journalHeaderId = (long)Math.Truncate(Convert.ToDouble(hyperlinkRange.Value2, System.Globalization.CultureInfo.InvariantCulture));
                GlobalsEx.Addin?.OnRibbonAction("RunJournalAttachmentFlow", journalHeaderId.ToString(System.Globalization.CultureInfo.InvariantCulture));
            }
            catch (Exception ex)
            {
                GlobalsEx.Context?.Logger?.LogException(ex, "DispatchJournalAttachmentHyperlink");
            }
        }

        private void adxExcelAppEvents1_SheetSelectionChange(object sender, object sheet, object range)
        {
            GlobalsEx.Context?.Logger?.LogDebug($"SheetSelectionChange event fired (sheet={GetSheetName(sheet)}, range={GetRangeAddress(range)})");

            try
            {
                GlobalsEx.Addin?.OnExcelEvent("SheetSelectionChange", new object[] { GetSheetName(sheet), GetRangeAddress(range) });
            }
            catch (Exception ex)
            {
                GlobalsEx.Context?.Logger?.LogException(ex, "adxExcelAppEvents1_SheetSelectionChange");
            }

            // Old monolith's second SheetSelectionChange handler: when the Balance
            // Configurator pane is visible, keep it synced to the active cell - relaunch
            // (reload from the cell's GLSense_GetBalance formula) if the new selection has
            // one, otherwise reset its cell reference. Purely host-side (pane visibility is
            // a host WinForms concern); RelaunchPane/ResetPaneReference are the only calls
            // that actually cross into Addin.Core (via the HWND-reparenting bridge). The old
            // `!AppState.Instance.IsLoginCompleted` early-out is dropped here - it was a
            // perf guard only, and blpane.Visible already can't be true unless the user
            // logged in and opened the pane via RibFSG_OnClick.
            try
            {
                if (range is not Excel.Range rng) return;
                if (rng.Rows.Count != 1 || rng.Columns.Count != 1) return;

                var blpane = GetPaneInstance();
                if (blpane != null && blpane.Visible)
                {
                    if (TryGetSingleCellFormula(rng, out string formula) &&
                        formula.IndexOf(AppConstants_GlBal, StringComparison.OrdinalIgnoreCase) >= 0)
                    {
                        _ = blpane.RelaunchPane();
                    }
                    else
                    {
                        _ = blpane.ResetPaneReference();
                    }

                    // Regression fix: GLConfiguratorPane.EmbedContent AttachThreadInput's this
                    // (Excel's own main UI) thread to Addin.Core's dedicated WPF thread so
                    // keyboard/Tab navigation flows INTO the reparented pane content - but
                    // attaching thread input only merges which window is considered
                    // "focused"; it does not make Win32 automatically hand focus back to
                    // Excel just because the user clicked a worksheet cell. Once any pane
                    // control has taken keyboard focus, it stays there across attached
                    // threads until something explicitly moves it - so every subsequent
                    // keystroke (even after clicking a different cell) kept going to the
                    // task pane instead of Excel's own cell editor. Since this handler is
                    // running on the exact same thread GLConfiguratorPane attached
                    // (Excel's own main UI thread - ADX event handlers and WinForms task
                    // panes both run there), explicitly reclaiming focus for Excel's own
                    // main window here - on every selection change while the pane is
                    // visible, regardless of whether this specific cell has a balance
                    // formula - reliably returns typing to the worksheet.
                    IntPtr excelHwnd = GlobalsEx.Context?.ExcelHandle ?? IntPtr.Zero;
                    if (excelHwnd != IntPtr.Zero)
                    {
                        SetFocus(excelHwnd);
                    }
                }
            }
            catch (Exception ex)
            {
                GlobalsEx.Context?.Logger?.LogException(ex, "adxExcelAppEvents1_SheetSelectionChange(Configurator)");
            }
        }

        // Mirrors GLSense.Addin.Core.AppConstants.glBal ("GLSense_GetBalance"). Duplicated
        // as a host-side constant (rather than referencing Addin.Core's AppConstants
        // directly) because the host project must never take a compile-time dependency on
        // the hot-reloadable Addin.Core assembly - see PORTING_GUIDE.md's AppDomain
        // boundary rules.
        private const string AppConstants_GlBal = "GLSense_GetBalance";

        /// <summary>
        /// Old monolith's AddinModule.TryGetSingleCellFormula - pure Excel COM/host-side
        /// helper, no Addin.Core involvement, ported verbatim.
        /// </summary>
        private static bool TryGetSingleCellFormula(Excel.Range cellRange, out string formulaString)
        {
            formulaString = string.Empty;

            try
            {
                if (cellRange == null)
                    return false;

                if (cellRange.Areas.Count != 1 || cellRange.Rows.Count != 1 || cellRange.Columns.Count != 1)
                    return false;

                if (!(cellRange.HasFormula is bool hf) || !hf)
                    return false;

                formulaString = cellRange.Formula?.ToString() ?? string.Empty;
                return true;
            }
            catch (Exception ex)
            {
                GlobalsEx.Context?.Logger?.LogException(ex, "AddinModule.TryGetSingleCellFormula");
                return false;
            }
        }

        private void adxExcelAppEvents1_WorkbookActivate(object sender, object hostObj)
        {
            GlobalsEx.Context?.Logger?.LogDebug($"WorkbookActivate event fired (workbook={GetWorkbookName(hostObj)})");

            try
            {
                GlobalsEx.Addin?.OnExcelEvent("WorkbookActivate", new object[] { GetWorkbookName(hostObj) });
            }
            catch (Exception ex)
            {
                GlobalsEx.Context?.Logger?.LogException(ex, "adxExcelAppEvents1_WorkbookActivate");
            }
        }

        private void adxExcelAppEvents1_WorkbookBeforeSave(object sender, ADXHostBeforeSaveEventArgs e)
        {
            GlobalsEx.Context?.Logger?.LogDebug($"WorkbookBeforeSave event fired (saveAsUI={e?.SaveAsUI})");

            try
            {
                // NOTE: e.SaveAsUI / e.Cancel follow AddinExpress's documented naming
                // convention for this event (mirroring Excel's own
                // Workbook_BeforeSave(SaveAsUI, Cancel)) - double check against IntelliSense
                // if this doesn't compile as-is.
                string workbookName = GetWorkbookName(ExcelApp?.ActiveWorkbook);

                bool proceed = GlobalsEx.Addin?.OnExcelEvent("WorkbookBeforeSave", new object[] { workbookName, e?.SaveAsUI }) ?? true;

                if (e != null)
                    e.Cancel = !proceed;
            }
            catch (Exception ex)
            {
                GlobalsEx.Context?.Logger?.LogException(ex, "adxExcelAppEvents1_WorkbookBeforeSave");
                // Deliberately do not set e.Cancel here - a bug in our own event handling
                // should never block the user from saving their workbook.
            }
        }

        // ---- Small, defensive helpers for extracting safe-to-marshal data from the
        // ---- late-bound "object" parameters ADX/Excel hand us. Every one of these
        // ---- swallows COM exceptions (the underlying object can already be gone by
        // ---- the time we look at it) and returns an empty string rather than throwing.

        private static string GetSheetName(object sheetOrChart)
        {
            try
            {
                if (sheetOrChart is Excel.Worksheet ws) return ws.Name;
                if (sheetOrChart is Excel.Chart ch) return ch.Name;
            }
            catch (COMException ex)
            {
                // The sheet can already be gone (e.g. deleted mid-event) - not an error we act on.
                GlobalsEx.Context?.Logger?.LogDebug($"AddinModule.GetSheetName: sheet/chart unavailable ({ex.Message})");
            }
            return string.Empty;
        }

        private static string GetRangeAddress(object rangeObj)
        {
            try
            {
                if (rangeObj is Excel.Range rng) return rng.Address[false, false];
            }
            catch (COMException ex)
            {
                // Same as above - the range can already be invalid by the time we read it.
                GlobalsEx.Context?.Logger?.LogDebug($"AddinModule.GetRangeAddress: range unavailable ({ex.Message})");
            }
            return string.Empty;
        }

        private static string GetWorkbookName(object workbookObj)
        {
            try
            {
                if (workbookObj is Excel.Workbook wb) return wb.Name;
            }
            catch (COMException ex)
            {
                // ignore - workbook already closed/gone
                GlobalsEx.Context?.Logger?.LogDebug($"AddinModule.GetWorkbookName: workbook unavailable ({ex.Message})");
            }
            return string.Empty;
        }

        private static string SafeToString(object obj)
        {
            try { return obj?.ToString() ?? string.Empty; }
            catch (COMException ex) { GlobalsEx.Context?.Logger?.LogDebug($"AddinModule.SafeToString: ToString failed ({ex.Message})"); return string.Empty; }
        }
    }
}

