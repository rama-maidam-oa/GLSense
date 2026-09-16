// IGLSenseAddin.cs in GLSense.Contracts
namespace GLSense.Contracts
{
    public interface IGLSenseAddin
    {
        void Initialize(IGLSenseContext context);
        void OnRibbonAction(string action, object parameter);

        /// <summary>
        /// Generic forwarder for Excel Application-level events (SheetChange, SheetActivate,
        /// SheetSelectionChange, WorkbookActivate, WorkbookBeforeSave, SheetBeforeDoubleClick,
        /// SheetFollowHyperlink, etc.). Subscriptions live in the never-reloaded host
        /// AddinModule (AddinExpress ADXExcelAppEvents component); this call crosses the
        /// AppDomain boundary into whichever GLSense.Addin.Core build is currently loaded.
        ///
        /// IMPORTANT for callers (AddinModule.cs): only pass primitive/serializable values in
        /// <paramref name="args"/> (string, bool, double, etc.) - never a live Excel COM range/
        /// sheet reference or an AddinExpress event-args object (ADXExcelSheetBeforeEventArgs,
        /// ADXHostBeforeSaveEventArgs, ...). Those aren't marked [Serializable] and will throw
        /// a SerializationException at the AppDomain boundary. Extract the sheet name / range
        /// address / etc. into a string in the host BEFORE calling this method.
        ///
        /// Return value: true = no objection / proceed as normal; false = the caller should
        /// cancel the underlying operation where that's meaningful (currently only
        /// WorkbookBeforeSave's Cancel flag). Events that have no cancel concept ignore the
        /// return value.
        /// </summary>
        bool OnExcelEvent(string eventName, object[] args);

        /// <summary>
        /// Dispatches one of the 16 GLSense_* custom Excel functions (GLSenseExcelFunctions.cs,
        /// XLLContainer) by name. The host's thin UDF wrapper methods are the only callers:
        /// each wrapper first runs the ADX/reflection-dependent bits that must stay host-side
        /// (ValidateInputs - inspects the calling wrapper method's own ExcelParamAttribute
        /// metadata via a stack-trace lookup, so it only works when called from the actual
        /// wrapper method; GetCellCallerAddress - uses the ADX-only
        /// Module.CallWorksheetFunction(ADXExcelWorksheetFunction.Caller) API, only meaningful
        /// for GLSense_GetBalance's R1C1 self-reference), then calls this method with the
        /// validated/defaulted parameter array (see UdfSentinels.cs remarks + inline comments
        /// in GLSenseExcelFunctions.cs for the exact per-function arg-array shape) to run the
        /// actual business logic (formula-result caching, period/segment lookups, REST calls).
        ///
        /// <paramref name="args"/> must only contain primitives (string/double/bool/null) -
        /// never a live Excel COM object or an AddinExpress-specific type
        /// (ADXExcelAsyncCallObject, ADXExcelError, ADXExcelRef, System.Reflection.Missing).
        /// The 3 async UDFs (GLSense_GetSegmentDFF, GLSense_GetDailyRate, GLSense_GetBalance)
        /// are called from inside a host-side Task.Run/BackgroundWorker (this call blocks
        /// until the underlying REST call finishes), and their result is then handed to
        /// asyncCallObject.ReturnResult(...) by the host.
        ///
        /// Wherever the old monolith returned AddinExpress.MSO.ADXExcelError directly, this
        /// method returns one of the plain-string sentinels in UdfSentinels instead (Addin.Core
        /// has no AddinExpress.MSO reference); the host translates the sentinel back to the
        /// real ADXExcelError enum value before returning it to Excel.
        /// </summary>
        object ExecuteUdf(string functionName, object[] args);
        void Shutdown();

        /// <summary>
        /// Group H (Balance Configurator pane) - HWND-reparenting bridge. The host's
        /// GLConfiguratorPane (an ADXExcelTaskPane, WinForms, host-only) cannot directly
        /// host a WPF UserControl created in this AppDomain the way the old monolith's
        /// GLConfiguratorPane did (new GLBalanceConfigurator(this) + ElementHost) - a WPF
        /// FrameworkElement isn't MarshalByRefObject and can't cross an AppDomain boundary
        /// by reference. Instead, Addin.Core creates a real, self-contained, borderless
        /// top-level Window (same "own WPF thread via WpfAppManager" trick every other
        /// dialog in this migration already uses) whose Content is GLBalanceConfigurator,
        /// and hands back only its native window handle (a blittable IntPtr, safe to cross
        /// via .NET Remoting since both AppDomains share one process's address space). The
        /// host then Win32 SetParent's that handle into its own WinForms panel and resizes
        /// it with MoveWindow - see GLSense\GLConfiguratorPane.cs. Idempotent: returns the
        /// existing handle if the content was already created.
        /// </summary>
        System.IntPtr CreateConfiguratorPaneContent();

        /// <summary>
        /// Group H - old monolith's GLConfiguratorPane.RelaunchPane() (re-reads the active
        /// cell's balance formula and reloads the configurator's fields). Called by the
        /// host from RibFSG_OnClick (pane just shown) and from SheetSelectionChange (pane
        /// already visible, active cell changed) - both host-side visibility checks, since
        /// ADXExcelTaskPane.Visible is host-only WinForms state.
        /// </summary>
        void RelaunchConfiguratorPane();

        /// <summary>
        /// Group H - old monolith's GLConfiguratorPane.ResetPaneReference() (just updates
        /// the displayed active-cell address without a full reload, used by
        /// SheetSelectionChange when the new active cell has no balance formula).
        /// </summary>
        void ResetConfiguratorPaneReference();

        /// <summary>
        /// Group H - tears down the hosted configurator Window (used on Shutdown/logoff
        /// and before a hot-reload swap, so the old AppDomain's Window doesn't linger).
        /// </summary>
        void CloseConfiguratorPaneContent();

        /// <summary>
        /// Returns the current login state (LoginUrl/LoginToken/IsLoggedIn) so the
        /// host's RibReload picker (GLReloadSourcePicker, host-side, no dependency on
        /// this project) can build the Online "check for update" request without
        /// AppState (which lives entirely in this project) crossing the AppDomain
        /// boundary directly.
        ///
        /// Added after the very first shipped version of this interface - any caller
        /// MUST wrap this call in try/catch (or otherwise tolerate a
        /// MissingMethodException/RemotingException), because the Release History
        /// browser can reload an OLDER historical build of GLSense.Addin.Core whose own
        /// compiled copy of this interface predates this member. Treat any failure
        /// exactly like "not logged in". See
        /// docs/superpowers/specs/2026-08-30-hotreload-release-history-design.md
        /// section 9.
        /// </summary>
        LoginInfo GetLoginInfo();
    }
}
