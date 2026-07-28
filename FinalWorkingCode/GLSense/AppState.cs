using GLSense.Interfaces;
using GLSense.Models;
using GLSense.Utilities;
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

        // UI & configuration
        public string DefaultSegment { get; set; }
        public int SegmentPickedIndex { get; set; } = AppConstants.DefaultSegmentPickedIndex;

        //Balance Configurator Pane
        public GLConfiguratorPane BalancePane { get; set; }
        public bool displayConfigurator { get; set; } = false;

        // Data caches
        public System.Data.DataTable CalculatedBalances { get; set; } = new System.Data.DataTable();
        public Dictionary<string,object> PreComputedBalances { get; set; } = new Dictionary<string, object>();

        public Dictionary<string, string> JournalDictionary { get; set; } = new Dictionary<string, string>();
        public string AttachIDs { get; set; }

        public void Reset()
        {
            LogUtility.LogDebug("AppState.Reset invoked. Resetting all writable properties to their defaults.");
            Type type = GetType();
            foreach (var prop in type.GetProperties())
            {
                if (prop.CanWrite)
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
