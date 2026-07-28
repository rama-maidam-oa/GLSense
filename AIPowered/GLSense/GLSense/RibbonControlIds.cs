//RibbonControlIds.cs in GLSense
using System;
using System.Collections.ObjectModel;

namespace GLSense
{
    internal static class RibbonControlIds
    {
        // Individual control IDs
        public const string RibDBL1 = "RibDBL1";
        public const string RibGetCube = "RibGetCube";
        public const string RibLedger = "Ribledger";
        public const string RibAccount = "RibAccount";
        public const string RibRollerGroup = "RibRollerGroup";
        public const string RibLOVs = "RibLOVs";
        public const string RibFSG = "RibFSG";
        public const string RibHideRows = "RibHideRows";
        public const string RibUnHideRows = "RibUnHideRows";
        public const string RibLiveCalc = "RibLiveCalc";
        public const string RibSegS = "RibSegS";
        public const string RibSegmentDiscover = "RibSegmentDiscover";
        public const string RigSegDiscover = "RigSegDiscover";
        public const string RibSegProperty = "RibSegProperty";
        // RibSegmentExpand: was a menu hosting RibExpandAll/RibbonExpand1Level (both
        // removed - GLExpandOptions.xaml now chooses level + fill direction).
        public const string RibSegmentExpand = "RibSegmentExpand";
        public const string RibSegmentExplode = "RibSegmentExplode";
        public const string RibExpodeAll = "RibExpodeAll";
        public const string RibbonExplode1Level = "RibbonExplode1Level";
        public const string RibDiscoverPeriod = "RibDiscoverPeriod";
        public const string RibAsFormula = "RibAsFormula";
        public const string RibRefreshRange = "RibRefreshRange";
        public const string RibRefreshAll = "RibRefreshAll";
        public const string RibRefreshBook = "RibRefreshBook";
        public const string RibClearSheet = "RibClearSheet";
        public const string RibClear = "RibClear";
        public const string RibHighlight = "RibHighlight";
        public const string RibCellHighlight = "RibCellHighlight";
        public const string RibSnapShot = "RibSnapShot";
        public const string RibSnapWorksheet = "RibSnapWorksheet";
        public const string RibSnapWorkbook = "RibSnapWorkbook";
        public const string RibSnapSubmit = "RibSnapSubmit";
        public const string RibDrilldownMenu = "RibDrilldownMenu";
        public const string RibBalanceDD = "RibBalanceDD";
        public const string RibBalanceJournalDD = "RibBalanceJournalDD";
        public const string RibBalanceSubLedgerDD = "RibBalanceSubLedgerDD";
        public const string RibJournalDD = "RibJournalDD";
        public const string RibSubledgerDD = "RibSubledgerDD";
        public const string RibTotaDD = "RibTotaDD";
        public const string RibDDConfiguration = "RibDDConfiguration";
        public const string RibDrillJobs = "RibDrillJobs";
        public const string RibFunctionsMenu = "RibFunctionsMenu";
        public const string RibSegmentEnabledFlag = "RibSegmentEnabledFlag";
        public const string RibSummaryFlag = "RibSummaryFlag";
        public const string RibSegment = "RibSegment";
        public const string RibNextSegment = "RibNextSegment";
        public const string RibPreviousSegment = "RibPreviousSegment";
        public const string RibSegmentDFF = "RibSegmentDFF";
        public const string RibPeriod = "RibPeriod";
        public const string RibPeriodbyDate = "RibPeriodbyDate";
        public const string RibPeriodbyYear = "RibPeriodbyYear";
        public const string RibPeriodNum = "RibPeriodNum";
        public const string RibPeriodQtr = "RibPeriodQtr";
        public const string RibPeriodYear = "RibPeriodYear";
        public const string RibPeriodStart = "RibPeriodStart";
        public const string RibPeriodEnd = "RibPeriodEnd";
        public const string RibDailyRate = "RibDailyRate";
        public const string RibVersionCheck = "RibVersionCheck";
        public const string RibUserConfig = "RibUserConfig";
        public const string RibHelp = "RibHelp";

        public const string RibLogin = "RibLogin";
        public const string RibLogout = "RibLogout";
        public const string RibUrl = "Riburl";
        public const string RibDebug = "RibDebug";
        public const string RibAbout = "RibAbout";

        // Shared reusable groups
        public static readonly ReadOnlyCollection<string> CommonDisabledControls =
            Array.AsReadOnly(new[]
            {
                RibDBL1, RibGetCube, RibLedger, RibAccount, RibRollerGroup, RibLOVs, RibFSG,
                RibHideRows, RibUnHideRows, RibLiveCalc, RibSegS, RibSegmentDiscover, RigSegDiscover,
                RibSegProperty, RibSegmentExpand, RibSegmentExplode,
                RibExpodeAll, RibbonExplode1Level, RibDiscoverPeriod, RibAsFormula, RibRefreshRange,
                RibRefreshAll, RibRefreshBook, RibClearSheet, RibClear, RibHighlight, RibCellHighlight,
                RibSnapShot, RibSnapWorksheet, RibSnapWorkbook, RibSnapSubmit, RibDrilldownMenu,
                RibBalanceDD, RibBalanceJournalDD, RibBalanceSubLedgerDD, RibJournalDD, RibSubledgerDD,
                RibTotaDD, RibDDConfiguration, RibDrillJobs, RibFunctionsMenu, RibSegmentEnabledFlag,
                RibSummaryFlag, RibSegment, RibNextSegment, RibPreviousSegment, RibSegmentDFF,
                RibPeriod, RibPeriodbyDate, RibPeriodbyYear, RibPeriodNum, RibPeriodQtr, RibPeriodYear,
                RibPeriodStart, RibPeriodEnd, RibDailyRate, RibVersionCheck, RibUserConfig, RibHelp
            });

        public static readonly ReadOnlyCollection<string> DefaultEnabledControls =
            Array.AsReadOnly(new[]
            {
                RibLogin, RibUrl, RibDebug, RibAbout
            });

        public static readonly ReadOnlyCollection<string> DefaultUnpressedControls =
            Array.AsReadOnly(new[]
            {
                RibDebug, RibLiveCalc, RibAsFormula, RibSnapWorksheet,
                RibSnapWorkbook, RibSnapSubmit, RibVersionCheck
            });

        public static readonly ReadOnlyCollection<string> PartialLoginDisabledControls =
            Array.AsReadOnly(new[]
            {
                RibLedger, RibAccount, RibRollerGroup, RibLOVs, RibFSG, RibHideRows, RibUnHideRows,
                RibLiveCalc, RibSegS, RibSegmentDiscover, RigSegDiscover, RibSegProperty,
                RibSegmentExpand, RibSegmentExplode, RibExpodeAll,
                RibbonExplode1Level, RibDiscoverPeriod, RibAsFormula, RibRefreshRange, RibRefreshAll,
                RibRefreshBook, RibClearSheet, RibClear, RibHighlight, RibCellHighlight, RibSnapShot,
                RibSnapWorksheet, RibSnapWorkbook, RibSnapSubmit, RibDrilldownMenu, RibBalanceDD,
                RibBalanceJournalDD, RibBalanceSubLedgerDD, RibJournalDD, RibSubledgerDD, RibTotaDD,
                RibDDConfiguration, RibDrillJobs, RibFunctionsMenu, RibSegmentEnabledFlag, RibSummaryFlag,
                RibSegment, RibNextSegment, RibPreviousSegment, RibSegmentDFF, RibPeriod,
                RibPeriodbyDate, RibPeriodbyYear, RibPeriodNum, RibPeriodQtr, RibPeriodYear,
                RibPeriodStart, RibPeriodEnd, RibDailyRate, RibVersionCheck, RibUserConfig
            });

        public static readonly ReadOnlyCollection<string> PartialLoginEnabledControls =
            Array.AsReadOnly(new[]
            {
                RibLogout, RibDBL1, RibGetCube, RibUrl, RibDebug, RibAbout, RibHelp
            });

        public static readonly ReadOnlyCollection<string> LoggedInEnabledControls =
            Array.AsReadOnly(new[]
            {
                RibDBL1, RibGetCube, RibLedger, RibAccount, RibRollerGroup, RibLOVs, RibFSG,
                RibHideRows, RibUnHideRows, RibLiveCalc, RibSegS, RibSegmentDiscover, RigSegDiscover,
                RibSegProperty, RibSegmentExpand, RibSegmentExplode,
                RibExpodeAll, RibbonExplode1Level, RibDiscoverPeriod, RibAsFormula, RibRefreshRange,
                RibRefreshAll, RibRefreshBook, RibClearSheet, RibClear, RibHighlight, RibCellHighlight,
                RibSnapShot, RibSnapWorksheet, RibSnapWorkbook, RibSnapSubmit, RibDrilldownMenu,
                RibBalanceDD, RibBalanceJournalDD, RibBalanceSubLedgerDD, RibJournalDD, RibSubledgerDD,
                RibTotaDD, RibDDConfiguration, RibDrillJobs, RibFunctionsMenu, RibSegmentEnabledFlag,
                RibSummaryFlag, RibSegment, RibNextSegment, RibPreviousSegment, RibSegmentDFF,
                RibPeriod, RibPeriodbyDate, RibPeriodbyYear, RibPeriodNum, RibPeriodQtr, RibPeriodYear,
                RibPeriodStart, RibPeriodEnd, RibDailyRate, RibVersionCheck, RibHelp, RibUserConfig
            });

        public static readonly ReadOnlyCollection<string> DrilldownSheetDisabledControls =
            Array.AsReadOnly(new[]
            {
                RibDBL1, RibGetCube, RibLedger, RibAccount, RibRollerGroup, RibLOVs, RibFSG,
                RibHideRows, RibUnHideRows, RibLiveCalc, RibSegS, RibSegmentDiscover, RigSegDiscover,
                RibSegProperty, RibSegmentExpand, RibSegmentExplode,
                RibExpodeAll, RibbonExplode1Level, RibDiscoverPeriod, RibAsFormula, RibRefreshRange,
                RibRefreshAll, RibRefreshBook, RibClearSheet, RibClear, RibHighlight, RibCellHighlight,
                RibSnapShot, RibSnapWorksheet, RibSnapWorkbook, RibSnapSubmit, RibFunctionsMenu,
                RibSegmentEnabledFlag, RibSummaryFlag, RibSegment, RibNextSegment, RibPreviousSegment,
                RibSegmentDFF, RibPeriod, RibPeriodbyDate, RibPeriodbyYear, RibPeriodNum, RibPeriodQtr,
                RibPeriodYear, RibPeriodStart, RibPeriodEnd, RibDailyRate, RibVersionCheck,
                RibHelp, RibUserConfig
            });

        public static readonly ReadOnlyCollection<string> DrilldownSheetEnabledControls =
            Array.AsReadOnly(new[]
            {
                RibBalanceDD, RibBalanceJournalDD, RibBalanceSubLedgerDD,
                RibJournalDD, RibSubledgerDD, RibTotaDD, RibDDConfiguration, RibDrillJobs
            });

        public static readonly ReadOnlyCollection<string> DisableRibbonDisabledControls =
            Array.AsReadOnly(new[]
            {
                RibDBL1, RibGetCube, RibLedger, RibAccount, RibRollerGroup, RibLOVs, RibFSG,
                RibHideRows, RibUnHideRows, RibLiveCalc, RibSegS, RibSegmentDiscover, RigSegDiscover,
                RibSegProperty, RibSegmentExpand, RibSegmentExplode,
                RibExpodeAll, RibbonExplode1Level, RibDiscoverPeriod, RibAsFormula, RibRefreshRange,
                RibRefreshAll, RibRefreshBook, RibClearSheet, RibClear, RibHighlight, RibCellHighlight,
                RibSnapShot, RibSnapWorksheet, RibSnapWorkbook, RibSnapSubmit, RibDrilldownMenu,
                RibBalanceDD, RibBalanceJournalDD, RibBalanceSubLedgerDD, RibJournalDD, RibSubledgerDD,
                RibTotaDD, RibDDConfiguration, RibDrillJobs, RibFunctionsMenu, RibSegmentEnabledFlag,
                RibSummaryFlag, RibSegment, RibNextSegment, RibPreviousSegment, RibSegmentDFF,
                RibPeriod, RibPeriodbyDate, RibPeriodbyYear, RibPeriodNum, RibPeriodQtr, RibPeriodYear,
                RibPeriodStart, RibPeriodEnd, RibDailyRate, RibVersionCheck, RibUserConfig, RibHelp,
                RibLogout, RibUrl, RibDebug, RibAbout
            });

        public static readonly ReadOnlyCollection<string> NoCubesDisabledControls =
            Array.AsReadOnly(new[]
            {
                RibLogin, RibLogout, RibUrl, RibDebug, RibAbout,
                RibDBL1, RibGetCube, RibLedger, RibAccount, RibRollerGroup, RibLOVs, RibFSG,
                RibHideRows, RibUnHideRows, RibLiveCalc, RibSegS, RibSegmentDiscover, RigSegDiscover,
                RibSegProperty, RibSegmentExpand, RibSegmentExplode,
                RibExpodeAll, RibbonExplode1Level, RibDiscoverPeriod, RibAsFormula, RibRefreshRange,
                RibRefreshAll, RibRefreshBook, RibClearSheet, RibClear, RibHighlight, RibCellHighlight,
                RibSnapShot, RibSnapWorksheet, RibSnapWorkbook, RibSnapSubmit, RibDrilldownMenu,
                RibBalanceDD, RibBalanceJournalDD, RibBalanceSubLedgerDD, RibJournalDD, RibSubledgerDD,
                RibTotaDD, RibDDConfiguration, RibDrillJobs, RibFunctionsMenu, RibSegmentEnabledFlag,
                RibSummaryFlag, RibSegment, RibNextSegment, RibPreviousSegment, RibSegmentDFF,
                RibPeriod, RibPeriodbyDate, RibPeriodbyYear, RibPeriodNum, RibPeriodQtr, RibPeriodYear,
                RibPeriodStart, RibPeriodEnd, RibDailyRate, RibVersionCheck, RibUserConfig, RibHelp
            });

        public static readonly ReadOnlyCollection<string> ProcessingDisabledControls =
            Array.AsReadOnly(new[]
            {
                RibDBL1, RibGetCube, RibLedger, RibAccount, RibRollerGroup, RibLOVs, RibFSG,
                RibHideRows, RibUnHideRows, RibLiveCalc, RibSegS, RibSegmentDiscover, RigSegDiscover,
                RibSegProperty, RibSegmentExpand, RibSegmentExplode,
                RibExpodeAll, RibbonExplode1Level, RibDiscoverPeriod, RibAsFormula, RibRefreshRange,
                RibRefreshAll, RibRefreshBook, RibClearSheet, RibClear, RibHighlight, RibCellHighlight,
                RibSnapShot, RibSnapWorksheet, RibSnapWorkbook, RibSnapSubmit, RibDrilldownMenu,
                RibBalanceDD, RibBalanceJournalDD, RibBalanceSubLedgerDD, RibJournalDD, RibSubledgerDD,
                RibTotaDD, RibDDConfiguration, RibDrillJobs, RibFunctionsMenu, RibSegmentEnabledFlag,
                RibSummaryFlag, RibSegment, RibNextSegment, RibPreviousSegment, RibSegmentDFF,
                RibPeriod, RibPeriodbyDate, RibPeriodbyYear, RibPeriodNum, RibPeriodQtr, RibPeriodYear,
                RibPeriodStart, RibPeriodEnd, RibDailyRate, RibVersionCheck, RibUserConfig, RibHelp,
                RibLogin
            });

        public static readonly ReadOnlyCollection<string> LoggedInPressedControls =
            Array.AsReadOnly(new[]
            {
                RibAsFormula, RibSnapWorksheet
            });
    }
}