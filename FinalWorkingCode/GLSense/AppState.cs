using GLSense.Interfaces;
using GLSense.Models;
using GLSense.Utilities;
using GLSense.Views;
using System;
using System.Collections.Generic;
using System.Diagnostics.Contracts;
using System.Globalization;
using Excel = Microsoft.Office.Interop.Excel;

namespace GLSense
{
    public sealed class AppState
    {
        // Singleton
        private static readonly Lazy<AppState> _instance = new(() => new AppState());
        public static AppState Instance => _instance.Value;

        private AppState() { } // Private constructor

        // Excel & Add-in
        // Null-conditional on CurrentInstance: this initializer runs the first time
        // AppState.Instance is touched anywhere (e.g. now also from
        // ExcelRefEditControl's constructor, which calls LogUtility.LogDebug ->
        // AppState.Instance.DebugLogs). In the VS XAML designer (and any other
        // context before OnConnection has run), AddinModule.CurrentInstance is null,
        // so the old unguarded ".HostApplication" threw a NullReferenceException here
        // - which the designer surfaced as an XDG0003 "Object reference not set" on
        // every ExcelRefEditControl instance on the page. Real runtime behavior is
        // unchanged since CurrentInstance is always set by the time a window opens.
        public Excel.Application ExcelApp { get; set; } = (Excel.Application)AddinModule.CurrentInstance?.HostApplication;

        // Current selections
        private CubeRecord _selectedCube;
        public CubeRecord SelectedCube
        {
            get => _selectedCube;
            set
            {
                _selectedCube = value;
                LogUtility.LogDebug($"AppState.SelectedCube changed. CubeName={value?.CubeName}, CubeId={value?.CubeId}");
            }
        }

        private LedgerModel _selectedLedger;
        public LedgerModel SelectedLedger
        {
            get => _selectedLedger;
            set
            {
                _selectedLedger = value;
                LogUtility.LogDebug($"AppState.SelectedLedger changed. LedgerName={value?.LedgerName}, LedgerId={value?.LedgerId}");
            }
        }

        // Login & authentication
        public string LoginUrl { get; set; }
        public string LoginToken { get; set; }
        public string LoginUserName { get; set; }

        private bool _isLoggedIn;
        public bool IsLoggedIn
        {
            get => _isLoggedIn;
            set
            {
                _isLoggedIn = value;
                LogUtility.LogDebug($"AppState.IsLoggedIn changed to {value}.");
            }
        }

        private bool _isLoginCompleted;
        public bool IsLoginCompleted
        {
            get => _isLoginCompleted;
            set
            {
                _isLoginCompleted = value;
                LogUtility.LogDebug($"AppState.IsLoginCompleted changed to {value}.");
            }
        }

        // Feature & job state
        public bool SnapshotSuccess { get; set; }
        public bool DebugLogs { get; set; }
        public bool VersionCheck { get; set; }
        public bool ResetFormulas { get; set; }
        public bool StartBatchCalc { get; set; }
        public bool SnapshotJob { get; set; }
        public bool SingleRefresh { get; set; }

        // Set for the duration of any CommonMethods.DisableExcelSettings/
        // EnableExcelSettings-bracketed bulk operation (mirrors the VB.NET sibling's
        // FSGExecute flag). Application.EnableEvents=false, also toggled by that same
        // bracket, already stops Excel from raising SheetSelectionChange/SheetActivate/
        // WorkbookActivate at all during the bracket - this flag is deliberately set
        // slightly wider (before EnableEvents goes false, cleared after it's restored) as
        // defense-in-depth against a trailing write landing right at that boundary, and is
        // checked explicitly by handlers that also need to guard against non-bulk
        // selection-changing operations like cut/copy/paste (see
        // adxExcelAppEvents1_SheetSelectionChange's separate CutCopyMode check).
        public bool IsBulkExcelOperationRunning { get; set; }

        // UI & configuration
        public string DefaultSegment { get; set; }
        public int SegmentPickedIndex { get; set; } = AppConstants.DefaultSegmentPickedIndex;

        //Balance Configurator Pane
        public GLConfiguratorPane BalancePane { get; set; }
        public bool displayConfigurator { get; set; } = false;

        //Balance Configurator floating window (Option A prototype - task pane above is untouched)
        //  WinForms Form (not a pure WPF Window) - see GLBalanceConfiguratorForm's own
        //  doc comment for why: a plain WPF Window repeatedly lost real Win32 keyboard
        //  focus back to Excel's grid in this in-process, same-thread hosting context.
        public GLBalanceConfiguratorForm BalanceWindow { get; set; }

        // Mirrors the VB.NET sibling's "Show Always" ribbon checkbox (FSGShow) - when
        // true, the window stays open and live-refreshes on every cell/sheet selection
        // instead of auto-hiding when the selection leaves a GLSense_GetBalance cell. See
        // RibShowAlways_OnClick and AddinModule.ApplyBalanceWindowVisibility.
        public bool BalanceWindowShowAlways { get; set; }

        // Data caches
        public System.Data.DataTable CalculatedBalances { get; set; } = new System.Data.DataTable();
        public Dictionary<string,object> PreComputedBalances { get; set; } = new Dictionary<string, object>();

        public Dictionary<string, string> JournalDictionary { get; set; } = new Dictionary<string, string>();
        public string AttachIDs { get; set; }

        // XLEdge COM add-in lookup (~1.4s via COMAddIns enumeration) - found once at ribbon
        // load (AddinModule_OnRibbonLoaded) and cached here so every other call site
        // (login, logout, XLEdge-permission check) reads the cached result instead of
        // re-searching. EdgeAddinSearchCompleted is tracked separately from
        // EdgeAddinInstance so a "not found" result is also cached (a null
        // EdgeAddinInstance alone can't distinguish "never searched" from "searched, XLEdge
        // isn't installed"). Excluded from Reset(): the XLEdge add-in's own COM
        // registration doesn't change across a GLSense login/logout, so there's no reason
        // to pay the search cost again on next login.
        public object EdgeAddinInstance { get; set; }
        public bool EdgeAddinSearchCompleted { get; set; }

        public void Reset()
        {
            LogUtility.LogDebug("AppState.Reset invoked. Resetting all writable properties to their defaults.");
            Type type = GetType();
            foreach (var prop in type.GetProperties())
            {
                if (prop.CanWrite && prop.Name != nameof(EdgeAddinInstance) && prop.Name != nameof(EdgeAddinSearchCompleted))
                {
                    object defaultValue = GetDefault(prop.PropertyType);
                    prop.SetValue(this, defaultValue);
                }
            }

            AppState.Instance.ExcelApp = (Excel.Application)AddinModule.CurrentInstance?.HostApplication;
        }

        private static object GetDefault(Type type)
        {
            if (type.IsValueType)
            {
                return Activator.CreateInstance(type);
            }
            return null;
        }
    }

    
}
