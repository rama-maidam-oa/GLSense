// GLSense.Addin.Core/AppState.cs
// Session-scoped mutable state. Kept intentionally minimal per PORTING_GUIDE.md #5 -
// only add a property here when a ported feature actually reads/writes it.
// LoginUrl/LoginToken/LoginUserName/IsLoggedIn added for the Login/Logout ribbon group
// (GLLogin.xaml.cs, ApiHelper, EdgeAddinHelper). SelectedCube/SelectedLedger added for
// Group B (Cube/Ledger selection - GLCubeDetails.xaml.cs is what actually sets them, and
// is also where IsLoginCompleted first becomes true, once both are chosen).
using GLSense.Addin.Core.Infrastructure;
using GLSense.Addin.Core.Models;
using System;
using System.Collections.Generic;
using System.Data;

namespace GLSense.Addin.Core
{
    public sealed class AppState
    {
        // Singleton
        private static readonly Lazy<AppState> _instance = new(() => new AppState());
        public static AppState Instance => _instance.Value;
        private AppState() { } // Private constructor

        // Login & authentication
        public string LoginUrl { get; set; }
        public string LoginToken { get; set; }
        public string LoginUserName { get; set; }

        private bool _isLoggedIn;
        /// <summary>Not necessarily the same as IsLoginCompleted - see that property's
        /// summary for the difference between a bare successful sign-in and a fully
        /// completed login (cube+ledger chosen).</summary>
        public bool IsLoggedIn
        {
            get => _isLoggedIn;
            set
            {
                if (_isLoggedIn == value) return;
                _isLoggedIn = value;
                ServiceLocator.Logger?.LogDebug($"AppState.IsLoggedIn changed to {value}");
            }
        }

        private bool _isLoginCompleted;
        public bool IsLoginCompleted
        {
            get => _isLoginCompleted;
            set
            {
                if (_isLoginCompleted == value) return;
                _isLoginCompleted = value;
                ServiceLocator.Logger?.LogDebug($"AppState.IsLoginCompleted changed to {value}");
            }
        }

        /// <summary>
        /// Group F (Refresh/Clear/Highlight/Hide-Rows) - cache of the balance-formula
        /// calculation results returned by the last bulk refresh (written by
        /// BulkRefreshProcess.BalancestoDTAsync via DataTableBuilder.ToDataTable, read by
        /// BalanceHighlighter/the hide-rows logic to find zero-value cached balances, and
        /// marked cache=false per-sheet by AddinEntry.BalancesReset on Clear/ClearSheet).
        /// Columns: excelSheet, excelCell, inputFormula, formulaKey, balanceValue, cache.
        /// </summary>
        public DataTable CalculatedBalances { get; set; }

        /// <summary>
        /// Group F - set true for the duration of RibRefreshRange's single-selection
        /// refresh (RangeRefresher), mirroring the old monolith's
        /// AppState.Instance.SingleRefresh flag.
        /// </summary>
        public bool SingleRefresh { get; set; }

        /// <summary>
        /// Group F - set true for the duration of RibClearSheet/RibClear's
        /// EnableCalculation-toggle reset (AddinEntry.ResetBalances), mirroring the old
        /// monolith's AppState.Instance.ResetFormulas flag.
        /// </summary>
        public bool ResetFormulas { get; set; }

        /// <summary>
        /// Group F (transitive - BulkRefreshProcess) - result flag of the last
        /// bulk-refresh/snapshot run, mirroring the old monolith's
        /// AppState.Instance.SnapshotSuccess.
        /// </summary>
        public bool SnapshotSuccess { get; set; }

        /// <summary>
        /// Group F (transitive - BulkRefreshProcess.PrepareBatchCache) - precomputed
        /// formulaKey -> balanceValue lookup built alongside CalculatedBalances for fast
        /// UDF cache hits, mirroring the old monolith's AppState.Instance.PreComputedBalances.
        /// </summary>
        public Dictionary<string, object> PreComputedBalances { get; set; }

        /// <summary>
        /// Group F (transitive - BatchCalcScope) - true while a bulk refresh/recalculation
        /// batch is in progress, mirroring the old monolith's AppState.Instance.StartBatchCalc.
        /// </summary>
        public bool StartBatchCalc { get; set; }

        /// <summary>
        /// Group F (transitive - BalanceRefresh) - whether balance formulas should be
        /// checked for version compatibility (and offered an update) before a refresh,
        /// mirroring the old monolith's AppState.Instance.VersionCheck.
        /// </summary>
        public bool VersionCheck { get; set; }

        private CubeRecord _selectedCube;
        /// <summary>Cube selected by the user in GLCubeDetails (Group B).</summary>
        public CubeRecord SelectedCube
        {
            get => _selectedCube;
            set
            {
                _selectedCube = value;
                ServiceLocator.Logger?.LogDebug($"AppState.SelectedCube changed to '{value?.CubeName ?? "<null>"}' (CubeId={value?.CubeId.ToString() ?? "<null>"})");

                // Ported from FinalWorkingCode\GLSense\Helpers\RibbonStateHelper.cs's
                // IsViewBasedCube() check (also duplicated in Views\GLUserConfig.xaml.cs): the
                // host's AddinModule.cs (GLSense, not Addin.Core) needs to know whether the
                // selected cube is view-based/EBS so it can grey out the Unified Drilldown /
                // Balances-to-Unified ribbon buttons for such cubes - but the host must never
                // take a compile-time dependency on GLSense.Addin.Core.AppState (would break the
                // AppDomain hot-reload isolation this project is built around). So instead of the
                // host reading this directly, push the flag through IRibbonController.
                // SetCubeViewBased whenever it changes - mirrors how RibbonController.IsLoggedIn
                // already lets the host ask "am I logged in" without reaching into Addin.Core.
                try
                {
                    bool isViewBased = (value?.ViewBased ?? false)
                        || string.Equals(value?.ErpType, "EBS", StringComparison.OrdinalIgnoreCase);
                    ServiceLocator.RibbonController?.SetCubeViewBased(isViewBased);
                }
                catch (Exception ex)
                {
                    ServiceLocator.Logger?.LogException(ex, "AppState.SelectedCube setter: SetCubeViewBased push failed");
                }
            }
        }

        private LedgerModel _selectedLedger;
        /// <summary>Ledger selected by the user in GLCubeDetails (Group B).</summary>
        public LedgerModel SelectedLedger
        {
            get => _selectedLedger;
            set
            {
                _selectedLedger = value;
                ServiceLocator.Logger?.LogDebug($"AppState.SelectedLedger changed to '{value?.LedgerName ?? "<null>"}' (LedgerId={value?.LedgerId.ToString() ?? "<null>"}, CoaId={value?.CoaId.ToString() ?? "<null>"})");
            }
        }

        /// <summary>
        /// Segment currently picked from the RibSegS ribbon combo (Group D - Segment/
        /// Period discoverers). Set by AddinEntry.SegmentChanged (RibSegS_OnChange) and
        /// read by SegmentDiscoverer/GLSegmentDiscovery to know which segment's values
        /// apply to the active cell. These were deliberately deferred from Group C/B -
        /// see the (now resolved) TODO that used to live here.
        /// </summary>
        public string DefaultSegment { get; set; }

        /// <summary>Index of DefaultSegment within the RibSegS combo (Group D). -1 = none picked.</summary>
        public int SegmentPickedIndex { get; set; }

        /// <summary>
        /// Group G (Snapshot/Job submission) - true when the RibSnapSubmit ribbon checkbox
        /// is pressed, meaning RibSnapShot should submit an async background snapshot job
        /// (BalanceRefresh.SubmitSnapAsync) instead of an inline synchronous snapshot
        /// (BalanceRefresh.RefreshingBalancesAsync("Snapshot", mode)). Mirrors the old
        /// monolith's AppState.Instance.SnapshotJob flag.
        /// </summary>
        public bool SnapshotJob { get; set; }

        /// <summary>
        /// Group I (Config/Debug/About/Help) - true while the RibDebug ribbon toggle is
        /// pressed, mirroring the old monolith's AppState.Instance.DebugLogs. Set by
        /// AddinEntry.OnRibbonAction's "DebugLogsToggled" case, which also starts/flushes
        /// the debug trace buffer via ServiceLocator.Logger.LogDebug/FlushDebugLogs.
        /// </summary>
        public bool DebugLogs { get; set; }

        /// <summary>
        /// Double-click/hyperlink drilldown pass (journal attachment flow) -
        /// FILE_ID -&gt; FILE_NAME lookup populated by JournalAttachments.
        /// PopulateJournalDictionary from the server's journal-attachment-files response,
        /// and read by Views.AttachmentsDialog to list the checkbox items, mirroring the
        /// old monolith's AppState.Instance.JournalDictionary.
        /// </summary>
        public Dictionary<string, string> JournalDictionary { get; set; } = new();

        /// <summary>
        /// Double-click/hyperlink drilldown pass (journal attachment flow) -
        /// comma-separated FILE_ID list the user checked in Views.AttachmentsDialog,
        /// consumed by JournalAttachments.DownloadSelectedAttachments, mirroring the old
        /// monolith's AppState.Instance.AttachIDs.
        /// </summary>
        public string AttachIDs { get; set; }

        /// <summary>
        /// Resets session state on logout. Does not touch anything Excel-COM related -
        /// this project gets Excel access from ServiceLocator.ExcelApp (supplied by the
        /// host), not from a field on AppState, so there is nothing to re-fetch here
        /// (unlike the old monolith's AppState.Reset(), which re-pulled
        /// AddinModule.CurrentInstance.HostApplication).
        /// </summary>
        public void Reset()
        {
            ServiceLocator.Logger?.LogDebug("AppState.Reset: clearing session state (logout).");
            LoginUrl = null;
            LoginToken = null;
            LoginUserName = null;
            IsLoggedIn = false;
            IsLoginCompleted = false;
            SelectedCube = null;
            SelectedLedger = null;
            CalculatedBalances = null;
            PreComputedBalances = null;
            SingleRefresh = false;
            ResetFormulas = false;
            SnapshotSuccess = false;
            StartBatchCalc = false;
            VersionCheck = false;
            DebugLogs = false;
            DefaultSegment = null;
            SegmentPickedIndex = AppConstants.DefaultSegmentPickedIndex;
            SnapshotJob = false;
            JournalDictionary.Clear();
            AttachIDs = null;
        }
    }
}
