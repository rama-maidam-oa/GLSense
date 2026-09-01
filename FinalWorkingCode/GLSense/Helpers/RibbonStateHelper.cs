using GLSense.Utilities;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using Excel = Microsoft.Office.Interop.Excel;

namespace GLSense.Helpers
{
    /// <summary>
    /// Central helper for enabling, disabling and toggling ribbon control visibility.
    /// Works directly with AddinModule.CurrentInstance.
    /// </summary>
    public class RibbonStateHelper
    {
        private readonly AddinModule _addinModule;
        private readonly AddinExpress.MSO.IRibbonUI _ribbon;
        private readonly Dictionary<string, bool> _enabledStates;

        public RibbonStateHelper(AddinModule addinModule, AddinExpress.MSO.IRibbonUI ribbon)
        {
            _addinModule = addinModule;
            _ribbon = ribbon;
            _enabledStates = new Dictionary<string, bool>(StringComparer.OrdinalIgnoreCase);
        }

        // ---------- Generic helpers ----------
        public void SetControlPressed(string controlName, bool pressed)
        {
            try
            {
                object ctrl = GetRibbonControl(controlName);
                if (ctrl == null)
                    return;

                // Try setting Pressed (for Add-in Express toggle controls)
                PropertyInfo propPressed = ctrl.GetType().GetProperty("Pressed");
                if (propPressed != null)
                {
                    propPressed.SetValue(ctrl, pressed, null);
                    return;
                }

                // Fallback: some controls expose 'Checked' instead
                PropertyInfo propChecked = ctrl.GetType().GetProperty("Checked");
                if (propChecked != null)
                {
                    propChecked.SetValue(ctrl, pressed, null);
                    return;
                }

                LogUtility.LogError("[RibbonStateHelper] No Pressed/Checked property on control " + controlName);
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex, $"[RibbonStateHelper] SetControlPressed: control='{controlName}', pressed={pressed}");
            }
        }
        public void SetControlEnabled(string controlName, bool enabled)
        {
            try
            {
                object ctrl = GetRibbonControl(controlName);
                if (ctrl == null)
                    return;

                PropertyInfo prop = ctrl.GetType().GetProperty("Enabled");
                if (prop != null)
                {
                    prop.SetValue(ctrl, enabled, null);
                    _enabledStates[controlName] = enabled;
                }
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex, $"[RibbonStateHelper] SetControlEnabled: control='{controlName}', enabled={enabled}");
            }
        }

        public void SetControlVisible(string controlName, bool visible)
        {
            try
            {
                object ctrl = GetRibbonControl(controlName);
                if (ctrl == null)
                    return;

                PropertyInfo prop = ctrl.GetType().GetProperty("Visible");
                prop?.SetValue(ctrl, visible, null);
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex, $"[RibbonStateHelper] SetControlVisible: control='{controlName}', visible={visible}");
            }
        }

        public void EnableControls(IEnumerable<string> controlNames)
        {
            foreach (string name in controlNames)
                SetControlEnabled(name, true);

        }

        public void DisableControls(IEnumerable<string> controlNames)
        {
            foreach (string name in controlNames)
                SetControlEnabled(name, false);
        }

        public void RestorePreviousState()
        {
            foreach (KeyValuePair<string, bool> kvp in _enabledStates)
                SetControlEnabled(kvp.Key, kvp.Value);
        }

        // ---------- Predefined logical states ----------
        private void RefreshRibbon()
        {
            try
            {
                _ribbon.Invalidate();
            }
            catch (Exception ex)
            {
                LogUtility.LogError("[RibbonStateHelper] RefreshRibbon: " + ex.Message);
            }
        }
        public void ApplyState(string stateName)
        {
            LogUtility.LogDebug($"[RibbonStateHelper] ApplyState: {stateName}");
            switch (stateName)
            {
                case "LoggedOut":
                    ApplyLoggedOutState();
                    break;

                case "PartialLoggedIn":
                    ApplyPartialLoginState();
                    break;

                case "LoggedIn":
                    ApplyLoggedInState();
                    break;

                case "ApplySheetActiveState":
                    ApplySheetActiveState();
                    break;
                case "NoCubes":
                    ApplyNoCubesState();
                    break;

                default:
                    LogUtility.LogWarn("[RibbonStateHelper] Unknown state: " + stateName);
                    break;
            }
            RefreshRibbon();
        }
        private void ApplyLoggedOutState()
        {
            try
            {
                DisableControls(
                [
                   "RibDBL1","RibGetCube","Ribledger","RibAccount","RibRollerGroup","RibLOVs","RibFSG","RibHideRows","RibUnHideRows",
                   "RibLiveCalc","RibSegS","RibSegmentDiscover","RigSegDiscover","RibSegProperty","RibSegmentExpand",
                   "RibSegmentExplode","RibExpodeAll","RibbonExplode1Level","RibDiscoverPeriod","RibDiscoverPeriodByDate","RibAsFormula","RibRefreshRange","RibRefreshAll","RibRefreshBook",
                   "RibClearSheet","RibClear","RibHighlight","RibCellHighlight","RibSnapShot","RibSnapWorksheet","RibSnapWorkbook","RibSnapSubmit",
                   "RibDrilldownMenu","RibBalanceDD","RibBalanceJournalDD","RibBalanceSubLedgerDD","RibJournalDD","RibSubledgerDD","RibTotaDD","RibBalancesDDToSubLedger","RibBalancesDDToUnified",
                   "RibDDConfiguration","RibDDDeleteConfiguration","RibDrillJobs","RibFunctionsMenu","RibSegmentEnabledFlag","RibSummaryFlag","RibSegment","RibNextSegment","RibSegmentAccountType",
                   "RibPreviousSegment","RibSegmentDFF","RibPeriod","RibPeriodbyDate","RibPeriodbyYear","RibPeriodNum","RibPeriodQtr","RibPeriodYear",
                   "RibPeriodStart","RibPeriodEnd","RibDailyRate","RibVersionCheck","RibUserConfig","RibHelp"
                 ]);

                EnableControls(["RibLogin", "Riburl", "RibDebug", "RibAbout"]);

                SetControlVisible("RibLogin", true);
                SetControlVisible("RibLogout", false);
                SetControlPressed("RibDebug", false);
                SetControlPressed("RibLiveCalc", false);
                SetControlPressed("RibAsFormula", false);
                SetControlPressed("RibSnapWorksheet", false);
                SetControlPressed("RibSnapWorkbook", false);
                SetControlPressed("RibSnapSubmit", false);
                SetControlPressed("RibVersionCheck", false);
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex, "[RibbonStateHelper] ApplyLoggedOutState");
            }
        }
        private void ApplyNoCubesState()
        {
            try
            {
                DisableControls(
                [
                   "RibLogin", "RibLogout", "Riburl", "RibDebug", "RibAbout",
                   "RibDBL1","RibGetCube","Ribledger","RibAccount","RibRollerGroup","RibLOVs","RibFSG","RibHideRows","RibUnHideRows",
                   "RibLiveCalc","RibSegS","RibSegmentDiscover","RigSegDiscover","RibSegProperty","RibSegmentExpand",
                   "RibSegmentExplode","RibExpodeAll","RibbonExplode1Level","RibDiscoverPeriod","RibDiscoverPeriodByDate","RibAsFormula","RibRefreshRange","RibRefreshAll","RibRefreshBook",
                   "RibClearSheet","RibClear","RibHighlight","RibCellHighlight","RibSnapShot","RibSnapWorksheet","RibSnapWorkbook","RibSnapSubmit",
                   "RibDrilldownMenu","RibBalanceDD","RibBalanceJournalDD","RibBalanceSubLedgerDD","RibJournalDD","RibSubledgerDD","RibTotaDD","RibBalancesDDToSubLedger","RibBalancesDDToUnified",
                   "RibDDConfiguration","RibDDDeleteConfiguration","RibDrillJobs","RibFunctionsMenu","RibSegmentEnabledFlag","RibSummaryFlag","RibSegment","RibNextSegment","RibSegmentAccountType",
                   "RibPreviousSegment","RibSegmentDFF","RibPeriod","RibPeriodbyDate","RibPeriodbyYear","RibPeriodNum","RibPeriodQtr","RibPeriodYear",
                   "RibPeriodStart","RibPeriodEnd","RibDailyRate","RibVersionCheck","RibUserConfig","RibHelp"
                 ]);
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex, "[RibbonStateHelper] ApplyNoCubesState");
            }
        }
        private void ApplyPartialLoginState()
        {
            try
            {
                DisableControls(
                [
                   "Ribledger","RibAccount","RibRollerGroup","RibLOVs","RibFSG","RibHideRows","RibUnHideRows",
                   "RibLiveCalc","RibSegS","RibSegmentDiscover","RigSegDiscover","RibSegProperty","RibSegmentExpand",
                   "RibSegmentExplode","RibExpodeAll","RibbonExplode1Level","RibDiscoverPeriod","RibDiscoverPeriodByDate","RibAsFormula","RibRefreshRange","RibRefreshAll","RibRefreshBook",
                   "RibClearSheet","RibClear","RibHighlight","RibCellHighlight","RibSnapShot","RibSnapWorksheet","RibSnapWorkbook","RibSnapSubmit",
                   "RibDrilldownMenu","RibBalanceDD","RibBalanceJournalDD","RibBalanceSubLedgerDD","RibJournalDD","RibSubledgerDD","RibTotaDD","RibBalancesDDToSubLedger","RibBalancesDDToUnified",
                   "RibDDConfiguration","RibDDDeleteConfiguration","RibDrillJobs","RibFunctionsMenu","RibSegmentEnabledFlag","RibSummaryFlag","RibSegment","RibNextSegment","RibSegmentAccountType",
                   "RibPreviousSegment","RibSegmentDFF","RibPeriod","RibPeriodbyDate","RibPeriodbyYear","RibPeriodNum","RibPeriodQtr","RibPeriodYear",
                   "RibPeriodStart","RibPeriodEnd","RibDailyRate","RibVersionCheck","RibUserConfig"
                 ]);

                EnableControls(["RibLogout", "RibDBL1", "RibGetCube", "Riburl", "RibDebug", "RibAbout", "RibHelp"]);

                SetControlVisible("RibLogin", false);
                SetControlVisible("RibLogout", true);
                SetControlPressed("RibAsFormula", false);
                SetControlPressed("RibSnapWorksheet", false);
                SetControlPressed("RibSnapWorkbook", false);
                SetControlPressed("RibSnapSubmit", false);
                SetControlPressed("RibVersionCheck", false);
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex, "[RibbonStateHelper] ApplyPartialLoginState");
            }
        }
        private void ApplyLoggedInState()
        {
            try
            {
                EnableControls(
                [
                   "RibDBL1","RibGetCube","Ribledger","RibAccount","RibRollerGroup","RibLOVs","RibFSG","RibHideRows","RibUnHideRows",
                   "RibLiveCalc","RibSegS","RibSegmentDiscover","RigSegDiscover","RibSegProperty","RibSegmentExpand",
                   "RibSegmentExplode","RibExpodeAll","RibbonExplode1Level","RibDiscoverPeriod","RibDiscoverPeriodByDate","RibAsFormula","RibRefreshRange","RibRefreshAll","RibRefreshBook",
                   "RibClearSheet","RibClear","RibHighlight","RibCellHighlight","RibSnapShot","RibSnapWorksheet","RibSnapWorkbook","RibSnapSubmit",
                   "RibDrilldownMenu","RibBalanceDD","RibBalanceJournalDD","RibBalanceSubLedgerDD","RibJournalDD","RibSubledgerDD","RibTotaDD","RibBalancesDDToSubLedger","RibBalancesDDToUnified",
                   "RibDDConfiguration","RibDDDeleteConfiguration","RibDrillJobs","RibFunctionsMenu","RibSegmentEnabledFlag","RibSummaryFlag","RibSegment","RibNextSegment","RibSegmentAccountType",
                   "RibPreviousSegment","RibSegmentDFF","RibPeriod","RibPeriodbyDate","RibPeriodbyYear","RibPeriodNum","RibPeriodQtr","RibPeriodYear",
                   "RibPeriodStart","RibPeriodEnd","RibDailyRate","RibVersionCheck","RibHelp","RibUserConfig"
                 ]);

                SetControlVisible("RibLogin", false);
                SetControlVisible("RibLogout", true);
                SetControlPressed("RibAsFormula", true);
                SetControlPressed("RibSnapWorksheet", true);
                SetControlPressed("RibSnapWorkbook", false);
                SetControlPressed("RibSnapSubmit", false);
                SetControlPressed("RibVersionCheck", false);

                ApplySheetActiveState();
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex, "[RibbonStateHelper] ApplyLoggedInState");
            }
        }
        private void ApplyDrilldownState(bool isDrilldown)
        {
            try
            {
                // RibUserConfig (User Preferences window) is deliberately excluded from both
                // lists below - it's a login/cube-session-level settings window, not something
                // tied to whichever sheet happens to be active, so it must not be toggled by
                // per-sheet state changes at all. Previously it was disabled here whenever the
                // active sheet was a drilldown result sheet (isDrilldown=true), so opening a
                // drilldown result tab would grey out User Preferences even while fully logged
                // in with a cube selected. Its enabled/disabled state is owned entirely by the
                // login-state methods above (ApplyLoggedOutState/ApplyNoCubesState/
                // ApplyPartialLoginState/ApplyLoggedInState) - enabled whenever logged in
                // (with a cube selected), disabled when logged out, irrespective of sheet.
                if (isDrilldown)
                {
                    DisableControls(
                     [
                       "RibDBL1","RibGetCube","Ribledger","RibAccount","RibRollerGroup","RibLOVs","RibFSG","RibHideRows","RibUnHideRows",
                       "RibLiveCalc","RibSegS","RibSegmentDiscover","RigSegDiscover","RibSegProperty","RibSegmentExpand",
                       "RibSegmentExplode","RibExpodeAll","RibbonExplode1Level","RibDiscoverPeriod","RibDiscoverPeriodByDate","RibAsFormula","RibRefreshRange","RibRefreshAll","RibRefreshBook",
                       "RibClearSheet","RibClear","RibHighlight","RibCellHighlight","RibSnapShot","RibSnapWorksheet","RibSnapWorkbook","RibSnapSubmit",
                       "RibFunctionsMenu","RibSegmentEnabledFlag","RibSummaryFlag","RibSegment","RibNextSegment","RibSegmentAccountType",
                       "RibPreviousSegment","RibSegmentDFF","RibPeriod","RibPeriodbyDate","RibPeriodbyYear","RibPeriodNum","RibPeriodQtr","RibPeriodYear",
                       "RibPeriodStart","RibPeriodEnd","RibDailyRate","RibVersionCheck","RibHelp"
                     ]);
                }
                else
                {
                    EnableControls(
                     [
                       "RibDBL1","RibGetCube","Ribledger","RibAccount","RibRollerGroup","RibLOVs","RibFSG","RibHideRows","RibUnHideRows",
                       "RibLiveCalc","RibSegS","RibSegmentDiscover","RigSegDiscover","RibSegProperty","RibSegmentExpand",
                       "RibSegmentExplode","RibExpodeAll","RibbonExplode1Level","RibDiscoverPeriod","RibDiscoverPeriodByDate","RibAsFormula","RibRefreshRange","RibRefreshAll","RibRefreshBook",
                       "RibClearSheet","RibClear","RibHighlight","RibCellHighlight","RibSnapShot","RibSnapWorksheet","RibSnapWorkbook","RibSnapSubmit",
                       "RibDDConfiguration","RibDDDeleteConfiguration","RibDrillJobs","RibFunctionsMenu","RibSegmentEnabledFlag","RibSummaryFlag","RibSegment","RibNextSegment","RibSegmentAccountType",
                       "RibPreviousSegment","RibSegmentDFF","RibPeriod","RibPeriodbyDate","RibPeriodbyYear","RibPeriodNum","RibPeriodQtr","RibPeriodYear",
                       "RibPeriodStart","RibPeriodEnd","RibDailyRate","RibVersionCheck","RibHelp"
                     ]);
                }
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex, $"[RibbonStateHelper] ApplyDrilldownState(isDrilldown={isDrilldown})");
            }
        }
        private void ApplySheetActiveState()
        {
            if (!ShouldProcessActiveState())
                return;

            try
            {
                ProcessActiveSheet();
            }
            catch (Exception ex)
            {
                LogUtility.LogError("[RibbonStateHelper] ApplySheetActiveState: " + ex.Message);
            }
        }
        private bool ShouldProcessActiveState()
        {
            return _addinModule.ExcelApp != null && AppState.Instance.IsLoginCompleted;
        }

        private void ProcessActiveSheet()
        {
            var activeSheet = GetActiveWorksheet();
            if (activeSheet == null)
                return;

            ApplyDrilldownState(IsValidDrilldownSheet(activeSheet));

            bool balanceExists = CheckForBalanceFormulas(activeSheet);
            LogUtility.LogDebug($"[RibbonStateHelper] ProcessActiveSheet: sheet='{activeSheet.Name}', balanceExists={balanceExists}");

            if (balanceExists)
            {
                EnableBalanceDrilldownControls();
            }
            else
            {
                DisableAllDrilldownControls();
                EnableDrilldownBasedOnSheetType(activeSheet);
            }
        }
        private static bool IsValidDrilldownSheet(Excel.Worksheet sht)
        {
            return sht?.ListObjects.Count > 0 && sht.ListObjects[1].Name.StartsWith("ORB_DD_");
        }
        private Excel.Worksheet GetActiveWorksheet()
        {
            return _addinModule.ExcelApp.ActiveSheet as Excel.Worksheet;
        }

        private static bool CheckForBalanceFormulas(Excel.Worksheet activeSheet)
        {
            var formulaCount = CountFormulaCells(activeSheet);
            if (formulaCount <= 0)
                return false;

            return HasBalanceFormula(activeSheet);
        }

        private static long CountFormulaCells(Excel.Worksheet sheet)
        {
            try
            {
                return sheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeFormulas).Count;
            }
            catch (Exception ex)
            {
                LogUtility.LogDebug($"[RibbonStateHelper] CountFormulaCells: no formula cells or SpecialCells failed on sheet '{sheet?.Name}' ({ex.Message}).");
                return 0;
            }
        }

        private static bool HasBalanceFormula(Excel.Worksheet sheet)
        {
            const string balanceFunctionName = AppConstants.glBal;

            try
            {
                var foundRange = sheet.Cells.Find(
                    balanceFunctionName,
                    Type.Missing,
                    Excel.XlFindLookIn.xlFormulas,
                    Excel.XlLookAt.xlPart,
                    Excel.XlSearchOrder.xlByRows,
                    Excel.XlSearchDirection.xlNext,
                    false
                );

                return foundRange != null;
            }
            catch (Exception ex)
            {
                LogUtility.LogDebug($"[RibbonStateHelper] HasBalanceFormula: search failed on sheet '{sheet?.Name}' ({ex.Message}).");
                return false;
            }
        }

        private void EnableBalanceDrilldownControls()
        {
            var ribbon = _addinModule;
            ribbon.RibBalanceDD.Enabled = true;
            ribbon.RibBalanceJournalDD.Enabled = true;
            ribbon.RibBalanceSubLedgerDD.Enabled = true;
            ribbon.RibJournalDD.Enabled = false;
            ribbon.RibSubledgerDD.Enabled = false;
            ribbon.RibTotaDD.Enabled = !IsViewBasedCube();
            ribbon.RibBalancesDDToSubLedger.Enabled = false;
            ribbon.RibBalancesDDToUnified.Enabled = false;
        }

        // Unified Drilldown (RibTotaDD, ddType "UF") and Balances Drilldown to Unified Drilldown
        // (RibBalancesDDToUnified, ddType "BLDD_UF") both fail server-side for view-based/EBS
        // cubes - the same restriction GLUserConfig.xaml.cs::IsViewBasedCube() already encodes
        // for the "Run as Job" checkbox and for hiding the Unified Drilldown row in the
        // preferences window ("view based cubes have limitations that prevent running unified
        // drilldown as a job"). Named/implemented to match that same convention, reused here so
        // these two ribbon buttons grey out up front for such cubes instead of only failing
        // after the user clicks them.
        private static bool IsViewBasedCube()
        {
            return (AppState.Instance.SelectedCube?.ViewBased ?? false)
                || string.Equals(AppState.Instance.SelectedCube?.ErpType, "EBS", StringComparison.OrdinalIgnoreCase);
        }

        private void DisableAllDrilldownControls()
        {
            var ribbon = _addinModule;
            ribbon.RibBalanceDD.Enabled = false;
            ribbon.RibBalanceJournalDD.Enabled = false;
            ribbon.RibBalanceSubLedgerDD.Enabled = false;
            ribbon.RibJournalDD.Enabled = false;
            ribbon.RibSubledgerDD.Enabled = false;
            ribbon.RibTotaDD.Enabled = false;
            ribbon.RibBalancesDDToSubLedger.Enabled = false;
            ribbon.RibBalancesDDToUnified.Enabled = false;
        }

        private void EnableDrilldownBasedOnSheetType(Excel.Worksheet sheet)
        {
            var drilldownType = GetDrilldownTypeFromCellA1(sheet);
            var sheetName = sheet.Name;
            var markerSheetName = GetSheetMarkerValue(sheet);

            bool isJournalDrilldown = IsJournalDrilldownSheet(drilldownType, sheetName, markerSheetName);
            bool isSubledgerDrilldown = IsSubledgerDrilldownSheet(drilldownType, sheetName, markerSheetName);

            _addinModule.RibBalancesDDToUnified.Enabled = isJournalDrilldown && !IsViewBasedCube();
            _addinModule.RibBalancesDDToSubLedger.Enabled = isJournalDrilldown;
            _addinModule.RibJournalDD.Enabled = isJournalDrilldown;
            _addinModule.RibSubledgerDD.Enabled = isSubledgerDrilldown;
        }

        private static string GetDrilldownTypeFromCellA1(Excel.Worksheet sheet)
        {
            var value = sheet.Range["A1"].Value;
            return value?.ToString().Trim().ToUpper() ?? string.Empty;
        }

        private static string GetSheetMarkerValue(Excel.Worksheet sheet)
        {
            try
            {
                var markerCell = sheet.Range[AppConstants.DrilldownSheetMarkerCellAddress];
                var value = markerCell?.Value2 ?? markerCell?.Value;
                return value?.ToString().Trim().ToUpper() ?? string.Empty;
            }
            catch (Exception ex)
            {
                LogUtility.LogDebug($"[RibbonStateHelper] GetSheetMarkerValue: failed reading marker cell on sheet '{sheet?.Name}' ({ex.Message}).");
                return string.Empty;
            }
        }

        private static bool IsJournalDrilldownSheet(string drilldownType, string sheetName, string markerSheetName)
        {
            return drilldownType.Equals("BALANCES DRILLDOWN", StringComparison.OrdinalIgnoreCase) ||
                   (MatchesBalancesPattern(sheetName) || MatchesBalancesPattern(markerSheetName));
        }

        private static bool IsSubledgerDrilldownSheet(string drilldownType, string sheetName, string markerSheetName)
        {
            return drilldownType.Equals("JOURNALS DRILLDOWN", StringComparison.OrdinalIgnoreCase) ||
                   (MatchesJournalsPattern(sheetName) || MatchesJournalsPattern(markerSheetName));
        }

        private static bool MatchesBalancesPattern(string sheetName)
        {
            return !string.IsNullOrWhiteSpace(sheetName) &&
                   StringContains(sheetName, "_BL_") &&
                   !StringContains(sheetName, "_JL_") &&
                   !StringContains(sheetName, "_SL_") &&
                   StringContains(sheetName, "_CM_");
        }

        private static bool MatchesJournalsPattern(string sheetName)
        {
            return !string.IsNullOrWhiteSpace(sheetName) &&
                   StringContains(sheetName, "_BL_") &&
                   StringContains(sheetName, "_JL_") &&
                   !StringContains(sheetName, "_SL_");
        }

        private static bool StringContains(string source, string value)
        {
            return source.IndexOf(value, StringComparison.OrdinalIgnoreCase) >= 0;
        }
        // ---------- Internal ----------

        private object GetRibbonControl(string controlName)
        {
            try
            {
                return RibbonReflectionHelper.GetRibbonControl(_addinModule, controlName);
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex, "[RibbonStateHelper] GetRibbonControl: " + controlName);
                return null;
            }

        }
    }
}
