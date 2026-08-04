// IRibbonController.cs in GLSense.Contracts
using System;
using System.Collections.Generic;

namespace GLSense.Contracts
{
    public interface IRibbonController
    {
        // EXISTING: Action execution - UNCHANGED
        void ExecuteAction(string buttonId);
        void RegisterAction(string buttonId, Action action);

        // NEW: State management
        void SetState(string stateName);

        // NEW: Individual control methods
        void SetControlEnabled(string controlName, bool enabled);
        void SetControlVisible(string controlName, bool visible);
        void SetControlPressed(string controlName, bool pressed);

        // Group D (Segment/Period discoverers) - SegmentDiscoverer/PeriodsDiscoverer/
        // GLSegmentDiscovery all need to read the "As Formula" ribbon toggle
        // (RibAsFormula.Pressed in the old monolith) from Addin.Core, which cannot
        // reference AddinExpress.MSO/ADXRibbonCheckBox directly. No existing Group A/B/C
        // mechanism reads a control's *current* boolean state back (SetControlPressed is
        // write-only), so this mirrors GetComboText's read-back pattern instead.
        bool GetControlPressed(string controlName);
        void SetControlLabel(string controlName, string label);
        void EnableControls(IEnumerable<string> controlNames);
        void DisableControls(IEnumerable<string> controlNames);

        // Ported from FinalWorkingCode's RibbonStateHelper.IsViewBasedCube() gating: the host
        // (AddinModule.cs) needs to know whether the currently selected cube is view-based/EBS
        // so it can grey out the Unified Drilldown / Balances-to-Unified ribbon buttons, but it
        // must never take a compile-time dependency on GLSense.Addin.Core.AppState. Addin.Core
        // pushes the flag through this whenever AppState.Instance.SelectedCube changes (see
        // AppState.cs's SelectedCube setter), mirroring how RibbonController.IsLoggedIn already
        // lets the host ask "am I logged in" without reaching into Addin.Core.
        void SetCubeViewBased(bool isViewBased);

        // NEW: Combo/dropdown ribbon controls (e.g. Ribledger, RibSegS) - these hold a
        // list of ADXRibbonItem entries plus free-text. Addin.Core can't reference
        // AddinExpress.MSO/ADXRibbonItem directly, so the host implements this over
        // reflection the same way it already does for SetControlEnabled/SetControlLabel.
        void SetComboItems(string controlName, IEnumerable<string> items);
        void ClearComboItems(string controlName);
        void SetComboText(string controlName, string text);
        string GetComboText(string controlName);

        // UI update methods
        void Invalidate(string controlId);
        void InvalidateAll();
        void ToggleTaskPane();

        // Group H (Balance Configurator pane) - GLConfiguratorPane/its ADX
        // TaskPaneInstances collection are host-only WinForms/ADX constructs (Addin.Core
        // can't reach them directly), so Addin.Core-initiated flows that need to affect
        // the pane's visibility go through these instead of a new dedicated interface:
        //   - HideAllTaskPanes: old monolith's AddinModule.HideTaskPanes(), called from
        //     AddinEntry.Logout() (was a TODO left by Group A/B, now resolved here).
        //   - HideConfiguratorPane: old monolith's GLConfiguratorPane's
        //     `_wpfControl.OnCloseRequested += () => this.Visible = false;` - the ported
        //     GLBalanceConfigurator's Cancel button raises the same event, but it now
        //     lives in Addin.Core and needs to reach back into the host to hide the pane.
        //   - RelaunchConfiguratorPaneIfVisible: old monolith's ledger-change flow relaunch
        //     (`if (BalancePane != null && BalancePane.Visible) _ = BalancePane.RelaunchPane();`),
        //     called from AddinEntry.LedgerChanged after a successful ledger switch. The
        //     host does the Visible check (host-only WinForms state) before deciding
        //     whether to call back into RelaunchConfiguratorPane.
        void HideAllTaskPanes();
        void HideConfiguratorPane();
        void RelaunchConfiguratorPaneIfVisible();
    }
}
