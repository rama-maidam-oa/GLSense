// RibbonController.cs in GLSense
using GLSense.Contracts;
using System;
using System.Collections.Generic;
using System.Linq;
using AddinExpress.MSO;

namespace GLSense
{
    public class RibbonController : MarshalByRefObject, IRibbonController
    {
        private const string StateLoggedOut = "LoggedOut";
        private const string StatePartialLoggedIn = "PartialLoggedIn";
        private const string StateLoggedIn = "LoggedIn";

        // Not a real login-state transition - re-evaluates the active sheet's
        // drilldown/balance-formula ribbon state without touching IsLoggedIn. Mirrors
        // FinalWorkingCode's Helpers\RibbonStateHelper.ApplyState("ApplySheetActiveState").
        private const string StateApplySheetActiveState = "ApplySheetActiveState";

        private readonly AddinModule _addinModule;
        private readonly IRibbonUI _ribbon;
        private readonly ILogger _logger;
        private readonly Dictionary<string, object> _controlCache;
        private readonly Dictionary<string, bool> _enabledStates;

        /// <summary>
        /// Tracks whether the last-applied ribbon state represents a completed login
        /// (StateLoggedIn). Added so host-side Excel Application-event handlers
        /// (AddinModule.adxExcelAppEvents1_SheetActivate/SheetChange/WorkbookActivate) can
        /// replicate the old monolith's `AppState.Instance.IsLoginCompleted` guard without
        /// the host taking a direct dependency on GLSense.Addin.Core.AppState (which would
        /// break the AppDomain hot-reload isolation this whole project is built around).
        /// </summary>
        public bool IsLoggedIn { get; private set; }

        /// <summary>
        /// Cached mirror of whether the currently selected cube is view-based/EBS (pushed from
        /// Addin.Core's AppState.SelectedCube setter via SetCubeViewBased below), so host-side
        /// drilldown ribbon-state methods (AddinModule.EnableBalanceDrilldownControls/
        /// EnableDrilldownBasedOnSheetType) can gate Unified Drilldown / Balances-to-Unified
        /// without taking a dependency on GLSense.Addin.Core.AppState - same reasoning as
        /// IsLoggedIn above.
        /// </summary>
        public bool IsCubeViewBased { get; private set; }

        // EXISTING: Action registration system - KEEP AS IS
        private readonly Dictionary<string, Action> _actionMap;

        public RibbonController(AddinModule addinModule, IRibbonUI ribbon, ILogger logger)
        {
            _addinModule = addinModule;
            _ribbon = ribbon;
            _logger = logger;
            _controlCache = new Dictionary<string, object>(StringComparer.OrdinalIgnoreCase);
            _enabledStates = new Dictionary<string, bool>(StringComparer.OrdinalIgnoreCase);

            // EXISTING: Initialize action map - KEEP AS IS
            _actionMap = new Dictionary<string, Action>(StringComparer.OrdinalIgnoreCase);
            InitializeActions();
        }

        public override object InitializeLifetimeService() => null;

        #region EXISTING: Action Registration System (UNCHANGED)

        private void InitializeActions()
        {
            RegisterAction("Login", () => GlobalsEx.Addin?.OnRibbonAction("Login", null));
            RegisterAction("ShowMessage", () => GlobalsEx.Addin?.OnRibbonAction("ShowMessage", null));
            RegisterAction("ShowMessage1", () => GlobalsEx.Addin?.OnRibbonAction("ShowMessage1", null));
            RegisterAction("Refresh", () => GlobalsEx.Addin?.OnRibbonAction("Refresh", null));
            RegisterAction("CreateTable", () => GlobalsEx.Addin?.OnRibbonAction("CreateTable", null));
            RegisterAction("Export", () => GlobalsEx.Addin?.OnRibbonAction("Export", null));
            RegisterAction("Logout", () => GlobalsEx.Addin?.OnRibbonAction("Logout", null));
        }

        public void RegisterAction(string buttonId, Action action)
        {
            if (!_actionMap.ContainsKey(buttonId))
            {
                _actionMap[buttonId] = action;
            }
        }

        public void ExecuteAction(string buttonId)
        {
            try
            {
                if (_actionMap.TryGetValue(buttonId, out var action))
                {
                    _logger?.LogDebug($"Executing action: {buttonId}");
                    action();
                }
                else
                {
                    _logger?.LogDebug($"Action '{buttonId}' not found, delegating to Addin");
                    GlobalsEx.Addin?.OnRibbonAction(buttonId, null);
                }
            }
            catch (Exception ex)
            {
                _logger?.LogError($"Error executing action {buttonId}", ex);
            }
        }

        #endregion

        #region NEW: State Management System

        public void SetState(string stateName)
        {
            _logger?.LogDebug($"Applying ribbon state: {stateName}");

            // Handled before the IsLoggedIn assignment below on purpose: this isn't a
            // login-state transition, just an on-demand re-evaluation of the active
            // sheet's drilldown/balance-formula ribbon state (see Drilldowns\
            // DDDatatoWorksheet.cs / DD_EP.cs), so it must never stomp the last real
            // login state IsLoggedIn is tracking.
            if (stateName == StateApplySheetActiveState)
            {
                try
                {
                    _addinModule?.ApplySheetActiveState();
                }
                catch (Exception ex)
                {
                    _logger?.LogError("ApplySheetActiveState (via SetState) failed", ex);
                }

                InvalidateAll();
                return;
            }

            IsLoggedIn = stateName == StateLoggedIn;

            switch (stateName)
            {
                case StateLoggedOut:
                    ApplyLoggedOutState();
                    break;

                case StatePartialLoggedIn:
                    ApplyPartialLoginState();
                    break;

                case StateLoggedIn:
                    ApplyLoggedInState();
                    break;

                default:
                    _logger?.LogWarn($"Unknown state: {stateName}");
                    break;
            }

            InvalidateAll();
        }

        #endregion

        #region NEW: State Implementations

        private void ApplyLoggedOutState()
        {
            try
            {
                DisableControls(RibbonControlIds.CommonDisabledControls);
                EnableControls(RibbonControlIds.DefaultEnabledControls);

                SetControlVisible(RibbonControlIds.RibLogin, true);
                SetControlVisible(RibbonControlIds.RibLogout, false);

                ResetPressedControls(RibbonControlIds.DefaultUnpressedControls);
            }
            catch (Exception ex)
            {
                _logger?.LogError("ApplyLoggedOutState failed", ex);
            }
        }

        private void ApplyPartialLoginState()
        {
            try
            {
                DisableControls(RibbonControlIds.PartialLoginDisabledControls);
                EnableControls(RibbonControlIds.PartialLoginEnabledControls);

                SetControlVisible(RibbonControlIds.RibLogin, false);
                SetControlVisible(RibbonControlIds.RibLogout, true);

                SetControlPressed(RibbonControlIds.RibAsFormula, false);
                SetControlPressed(RibbonControlIds.RibSnapWorksheet, false);
                SetControlPressed(RibbonControlIds.RibSnapWorkbook, false);
                SetControlPressed(RibbonControlIds.RibSnapSubmit, false);
                SetControlPressed(RibbonControlIds.RibVersionCheck, false);
            }
            catch (Exception ex)
            {
                _logger?.LogError("ApplyPartialLoginState failed", ex);
            }
        }

        private void ApplyLoggedInState()
        {
            try
            {
                EnableControls(RibbonControlIds.LoggedInEnabledControls);

                SetControlVisible(RibbonControlIds.RibLogin, false);
                SetControlVisible(RibbonControlIds.RibLogout, true);

                SetControlPressed(RibbonControlIds.RibAsFormula, true);
                SetControlPressed(RibbonControlIds.RibSnapWorksheet, true);
                SetControlPressed(RibbonControlIds.RibSnapWorkbook, false);
                SetControlPressed(RibbonControlIds.RibSnapSubmit, false);
                SetControlPressed(RibbonControlIds.RibVersionCheck, false);

                // FinalWorkingCode's RibbonStateHelper.ApplyLoggedInState re-runs the
                // sheet-scoped drilldown/balance-formula narrowing as its own last step,
                // since the 6 balance/journal/subledger/unified drilldown buttons are
                // deliberately excluded from LoggedInEnabledControls above - they must
                // only ever be enabled by what's actually on the active sheet, never by
                // login alone. Mirrored here for parity.
                _addinModule?.ApplySheetActiveState();
            }
            catch (Exception ex)
            {
                _logger?.LogError("ApplyLoggedInState failed", ex);
            }
        }

        #endregion

        #region NEW: Core Control Methods

        public void SetControlEnabled(string controlName, bool enabled)
        {
            try
            {
                _logger?.LogDebug($"RibbonController.SetControlEnabled: '{controlName}' -> {enabled}");
                object ctrl = GetRibbonControl(controlName);
                if (ctrl == null) return;

                var prop = ctrl.GetType().GetProperty("Enabled");
                if (prop != null && prop.CanWrite)
                {
                    prop.SetValue(ctrl, enabled, null);
                    _enabledStates[controlName] = enabled;
                }
                else
                {
                    _logger?.LogWarn($"SetControlEnabled: '{controlName}' has no writable 'Enabled' property - request to set it to {enabled} was silently ignored.");
                }
            }
            catch (Exception ex)
            {
                _logger?.LogError($"SetControlEnabled failed for '{controlName}'", ex);
            }
        }

        public void SetControlVisible(string controlName, bool visible)
        {
            try
            {
                _logger?.LogDebug($"RibbonController.SetControlVisible: '{controlName}' -> {visible}");
                object ctrl = GetRibbonControl(controlName);
                if (ctrl == null) return;

                var prop = ctrl.GetType().GetProperty("Visible");
                if (prop != null && prop.CanWrite)
                {
                    prop.SetValue(ctrl, visible, null);
                }
                else
                {
                    _logger?.LogWarn($"SetControlVisible: '{controlName}' has no writable 'Visible' property - request to set it to {visible} was silently ignored.");
                }
            }
            catch (Exception ex)
            {
                _logger?.LogError($"SetControlVisible failed for '{controlName}'", ex);
            }
        }

        public void SetControlPressed(string controlName, bool pressed)
        {
            try
            {
                _logger?.LogDebug($"RibbonController.SetControlPressed: '{controlName}' -> {pressed}");
                object ctrl = GetRibbonControl(controlName);
                if (ctrl == null) return;

                var propPressed = ctrl.GetType().GetProperty("Pressed");
                if (propPressed != null && propPressed.CanWrite)
                {
                    propPressed.SetValue(ctrl, pressed, null);
                    return;
                }

                var propChecked = ctrl.GetType().GetProperty("Checked");
                if (propChecked != null && propChecked.CanWrite)
                {
                    propChecked.SetValue(ctrl, pressed, null);
                }
                else
                {
                    _logger?.LogWarn($"SetControlPressed: '{controlName}' has no writable 'Pressed' or 'Checked' property - request to set it to {pressed} was silently ignored.");
                }
            }
            catch (Exception ex)
            {
                _logger?.LogError($"SetControlPressed failed for '{controlName}'", ex);
            }
        }

        public void SetControlLabel(string controlName, string label)
        {
            try
            {
                object ctrl = GetRibbonControl(controlName);
                if (ctrl == null) return;

                // ADXRibbonButton's visible display text property is "Caption", not
                // "Label" - this was reflecting on a property that either doesn't
                // exist or has no effect on what's actually rendered, so every call
                // (e.g. GLCubeDetails/AddinEntry updating RibGetCube with the selected
                // cube's name) was silently doing nothing: no exception, no log, just
                // no visible change. Confirmed against SetComboItems' own working
                // reflection in this same class (uses "Caption" on ADXRibbonItem) and
                // against FinalWorkingCode's proven "RibGetCube.Caption = ..." pattern.
                var prop = ctrl.GetType().GetProperty("Caption");
                if (prop != null && prop.CanWrite)
                {
                    prop.SetValue(ctrl, label, null);
                }
                else
                {
                    _logger?.LogWarn($"SetControlLabel: '{controlName}' has no writable 'Caption' property - request to set it to '{label}' was silently ignored.");
                }
            }
            catch (Exception ex)
            {
                _logger?.LogError($"SetControlLabel failed for '{controlName}'", ex);
            }
        }

        public void SetCubeViewBased(bool isViewBased)
        {
            _logger?.LogDebug($"RibbonController.SetCubeViewBased: {isViewBased}");
            IsCubeViewBased = isViewBased;
        }

        public void EnableControls(IEnumerable<string> controlNames)
        {
            foreach (string name in controlNames)
            {
                SetControlEnabled(name, true);
            }
        }

        public void DisableControls(IEnumerable<string> controlNames)
        {
            foreach (string name in controlNames)
            {
                SetControlEnabled(name, false);
            }
        }

        /// <summary>
        /// Replaces the item list of a combo/dropdown ribbon control (e.g. Ribledger,
        /// RibSegS). Each control's "Items" collection holds ADXRibbonItem instances -
        /// since this project doesn't reference AddinExpress.MSO types by name here,
        /// the ADXRibbonItem type and its "Caption" property are both discovered via
        /// reflection off the collection's own Add() method, the same "never hardcode a
        /// designer type" approach GetRibbonControlInternal already uses below.
        /// </summary>
        public void SetComboItems(string controlName, IEnumerable<string> items)
        {
            try
            {
                _logger?.LogDebug($"RibbonController.SetComboItems: '{controlName}' <- {items?.Count() ?? 0} item(s)");
                object ctrl = GetRibbonControl(controlName);
                if (ctrl == null) return;

                var itemsProp = ctrl.GetType().GetProperty("Items");
                var itemsCollection = itemsProp?.GetValue(ctrl, null);
                if (itemsCollection == null)
                {
                    _logger?.LogWarn($"SetComboItems: '{controlName}' has no readable 'Items' property - no items were set.");
                    return;
                }

                var itemsType = itemsCollection.GetType();
                itemsType.GetMethod("Clear", System.Reflection.BindingFlags.Public | System.Reflection.BindingFlags.Instance, null, Type.EmptyTypes, null)?.Invoke(itemsCollection, null);

                // Both previous approaches to resolving the right "Add" overload were
                // reflection guesswork and both broke: "prefer whichever Add isn't
                // Add(object)" picked an Add(string) convenience overload instead of the
                // real Add(ADXRibbonItem) one; "prefer whichever Add's parameter type has
                // a writable Caption property" found NOTHING at all (confirmed by
                // debugging - every candidate's GetProperty("Caption") came back null,
                // meaning the real Add(ADXRibbonItem)-shaped overload's parameter type
                // isn't exposing "Caption" the way GetProperty("Caption") expects - likely
                // because the Add overload actually takes a base/interface type that
                // doesn't declare Caption itself, only ADXRibbonItem does).
                // This file already has "using AddinExpress.MSO;" at the top, so there is
                // no real need to reflect for the ITEM type at all - construct the actual
                // ADXRibbonItem directly (strongly typed, Caption is a real compile-time
                // property, no ambiguity possible), and only use reflection to find which
                // "Add" overload can actually ACCEPT one, via IsAssignableFrom - the
                // correct way to test "can I pass an ADXRibbonItem here", which works
                // whether the parameter type is ADXRibbonItem itself or some base/
                // interface type of it.
                System.Reflection.MethodInfo addMethod = null;
                foreach (var m in itemsType.GetMethods(System.Reflection.BindingFlags.Public | System.Reflection.BindingFlags.Instance))
                {
                    if (m.Name != "Add") continue;
                    var parameters = m.GetParameters();
                    if (parameters.Length != 1) continue;

                    if (parameters[0].ParameterType.IsAssignableFrom(typeof(ADXRibbonItem)))
                    {
                        addMethod = m;
                        break;
                    }
                }

                if (addMethod == null)
                {
                    _logger?.LogError($"SetComboItems: could not find an Add(x) overload that accepts an ADXRibbonItem for '{controlName}' - no items were set.");
                    return;
                }

                int expectedCount = items?.Count() ?? 0;
                foreach (var text in items ?? Array.Empty<string>())
                {
                    var item = new ADXRibbonItem { Caption = text };
                    addMethod.Invoke(itemsCollection, new object[] { item });
                }

                // Self-verifying check, not just a status log: this is exactly the class
                // of bug this whole session was spent chasing - a method that logs "N
                // items <- passed in" and looks fine, while the live collection ends up
                // with a different count (0, in our case) because something upstream
                // silently no-op'd. Comparing expected vs. actual and using LogWarn (not
                // LogDebug) means a future recurrence is visible in the log even without
                // Debug mode enabled - the exact gap that made this bug take this long to
                // pin down originally.
                int actualCount = (int)(itemsCollection.GetType().GetProperty("Count")?.GetValue(itemsCollection, null) ?? 0);
                if (actualCount != expectedCount)
                {
                    _logger?.LogWarn($"SetComboItems: '{controlName}' expected {expectedCount} item(s) but the live collection now has {actualCount} - population did not fully succeed.");
                }
                else
                {
                    _logger?.LogDebug($"RibbonController.SetComboItems: '{controlName}' actually populated with {actualCount} item(s) in the live collection.");
                }
            }
            catch (Exception ex)
            {
                _logger?.LogError($"SetComboItems failed for '{controlName}'", ex);
            }
        }

        public void ClearComboItems(string controlName)
        {
            try
            {
                object ctrl = GetRibbonControl(controlName);
                if (ctrl == null) return;

                var itemsProp = ctrl.GetType().GetProperty("Items");
                var itemsCollection = itemsProp?.GetValue(ctrl, null);
                if (itemsCollection == null)
                {
                    _logger?.LogWarn($"ClearComboItems: '{controlName}' has no readable 'Items' property - nothing was cleared.");
                    return;
                }
                itemsCollection.GetType().GetMethod("Clear", System.Reflection.BindingFlags.Public | System.Reflection.BindingFlags.Instance, null, Type.EmptyTypes, null)?.Invoke(itemsCollection, null);
            }
            catch (Exception ex)
            {
                _logger?.LogError($"ClearComboItems failed for '{controlName}'", ex);
            }
        }

        public void SetComboText(string controlName, string text)
        {
            try
            {
                object ctrl = GetRibbonControl(controlName);
                if (ctrl == null) return;

                var prop = ctrl.GetType().GetProperty("Text");
                if (prop == null || !prop.CanWrite)
                {
                    _logger?.LogWarn($"SetComboText: '{controlName}' has no writable 'Text' property - request to set it to '{text}' was silently ignored.");
                    return;
                }

                prop.SetValue(ctrl, text, null);

                // Self-verifying, same reasoning as SetComboItems' expected-vs-actual
                // check above: read the value back immediately and flag any mismatch via
                // LogWarn so it's visible without Debug mode - catches the control
                // silently rejecting/altering the value (e.g. text not matching any real
                // item in a strict-selection combo) right where it happens instead of
                // surfacing later as "the ribbon just doesn't show the right thing."
                var actual = prop.GetValue(ctrl, null) as string;
                if (!string.Equals(actual, text, StringComparison.Ordinal))
                {
                    _logger?.LogWarn($"SetComboText: '{controlName}' was set to '{text}' but reading it back immediately shows '{actual}' - mismatch.");
                }
            }
            catch (Exception ex)
            {
                _logger?.LogError($"SetComboText failed for '{controlName}'", ex);
            }
        }

        public string GetComboText(string controlName)
        {
            try
            {
                object ctrl = GetRibbonControl(controlName);
                if (ctrl == null) return string.Empty;

                var prop = ctrl.GetType().GetProperty("Text");
                return prop?.GetValue(ctrl, null) as string ?? string.Empty;
            }
            catch (Exception ex)
            {
                _logger?.LogError($"GetComboText failed for '{controlName}'", ex);
                return string.Empty;
            }
        }

        /// <summary>
        /// Read-back counterpart to SetControlPressed (Group D) - checks the "Pressed"
        /// property first (ADXRibbonToggleButton/ADXRibbonCheckBox) and falls back to
        /// "Checked", mirroring SetControlPressed's own property lookup order.
        /// </summary>
        public bool GetControlPressed(string controlName)
        {
            try
            {
                object ctrl = GetRibbonControl(controlName);
                if (ctrl == null) return false;

                var propPressed = ctrl.GetType().GetProperty("Pressed");
                if (propPressed != null && propPressed.CanRead)
                {
                    return propPressed.GetValue(ctrl, null) is bool pressed && pressed;
                }

                var propChecked = ctrl.GetType().GetProperty("Checked");
                if (propChecked != null && propChecked.CanRead)
                {
                    return propChecked.GetValue(ctrl, null) is bool @checked && @checked;
                }

                return false;
            }
            catch (Exception ex)
            {
                _logger?.LogError($"GetControlPressed failed for '{controlName}'", ex);
                return false;
            }
        }

        private void ResetPressedControls(IEnumerable<string> controlNames)
        {
            foreach (string name in controlNames)
            {
                SetControlPressed(name, false);
            }
        }

        #endregion

        #region NEW: IRibbonController Implementation

        public void UpdateButton(string controlId, bool enabled, bool visible)
        {
            SetControlEnabled(controlId, enabled);
            SetControlVisible(controlId, visible);
        }
        public void Invalidate(string controlId)
        {
            try
            {
                // Was previously reflecting for an "Invalidate" method on the control
                // object itself (ctrl.GetType().GetMethod("Invalidate")) - that's
                // unreliable: AddinExpress ribbon controls can expose more than one
                // method/inherited member matching that name, and reflection with no
                // parameter-type filter picks whichever GetMethod() happens to resolve
                // first, which may not be the one that actually tells Excel's Ribbon
                // engine to re-pull this control's cached getItemCount/getItemLabel/
                // getText state. _ribbon (AddinExpress.MSO.IRibbonUI, the same object
                // InvalidateAll() calls .Invalidate() on) has a real, documented
                // InvalidateControl(string) member for exactly this - invalidate one
                // control's Ribbon-side cache without touching the rest of the ribbon.
                _ribbon?.InvalidateControl(controlId);
            }
            catch (Exception ex)
            {
                _logger?.LogError($"Invalidate failed for '{controlId}'", ex);
            }
        }

        public void InvalidateAll()
        {
            try
            {
                _ribbon?.Invalidate();
            }
            catch (Exception ex)
            {
                _logger?.LogError("InvalidateAll failed", ex);
            }
        }

        public void ToggleTaskPane()
        {
            // Implement if needed
        }

        /// <summary>
        /// Group H - old monolith's AddinModule.HideTaskPanes() (iterates every live
        /// GLConfiguratorPane instance and hides it). Called from AddinEntry.Logout() via
        /// this interface since ADXExcelTaskPanesCollectionItem.TaskPaneInstances is a
        /// host-only ADX construct.
        /// </summary>
        public void HideAllTaskPanes()
        {
            try
            {
                var item = _addinModule?.adxExcelTaskPanesCollectionItem1;
                if (item?.TaskPaneInstances == null) return;

                foreach (GLConfiguratorPane pane in item.TaskPaneInstances)
                {
                    pane.Visible = false;
                }
            }
            catch (Exception ex)
            {
                _logger?.LogError("HideAllTaskPanes failed", ex);
            }
        }

        /// <summary>
        /// Group H - hides just the currently-instantiated configurator pane (used by
        /// GLBalanceConfigurator's Cancel button, ported from the old monolith's
        /// `_wpfControl.OnCloseRequested += () => this.Visible = false;`).
        /// </summary>
        public void HideConfiguratorPane()
        {
            try
            {
                var pane = _addinModule?.GetPaneInstance();
                if (pane != null)
                {
                    pane.Visible = false;
                }
            }
            catch (Exception ex)
            {
                _logger?.LogError("HideConfiguratorPane failed", ex);
            }
        }

        /// <summary>
        /// Group H - old monolith's ledger-change relaunch
        /// (`if (BalancePane != null && BalancePane.Visible) _ = BalancePane.RelaunchPane();`).
        /// Called from AddinEntry.LedgerChanged after a successful ledger switch - the
        /// Visible check has to happen here since it's host-only WinForms state.
        /// </summary>
        public void RelaunchConfiguratorPaneIfVisible()
        {
            try
            {
                var pane = _addinModule?.GetPaneInstance();
                if (pane != null && pane.Visible)
                {
                    _ = pane.RelaunchPane();
                }
            }
            catch (Exception ex)
            {
                _logger?.LogError("RelaunchConfiguratorPaneIfVisible failed", ex);
            }
        }

        #endregion

        #region Control Lookup

        private object GetRibbonControl(string controlName)
        {
            if (string.IsNullOrWhiteSpace(controlName))
                return null;

            if (_controlCache.TryGetValue(controlName, out object cachedControl))
                return cachedControl;

            try
            {
                object control = GetRibbonControlInternal(controlName);
                if (control != null)
                {
                    _controlCache[controlName] = control;
                }
                else
                {
                    // This is the single choke point every SetComboItems/SetComboText/
                    // SetControlEnabled/SetControlVisible/SetControlPressed/
                    // SetControlLabel/ClearComboItems/GetComboText call goes through -
                    // every one of those methods used to just silently "if (ctrl == null)
                    // return;" with zero trace, so a renamed/missing control (typo,
                    // AddinExpress designer regeneration, etc.) would fail completely
                    // silently, exactly the kind of bug this whole session was spent
                    // chasing blind through logs that looked fine. LogWarn here (not
                    // LogDebug) so this is ALWAYS visible in the log even without Debug
                    // mode enabled - this is meant to be diagnosable at a client site with
                    // no source access and no debugger, not just in dev.
                    _logger?.LogWarn($"RibbonController.GetRibbonControl: control '{controlName}' could not be resolved (not found as a public property or field on AddinModule) - any operation on it will silently no-op.");
                }

                return control;
            }
            catch (Exception ex)
            {
                _logger?.LogError($"GetRibbonControl failed for '{controlName}'", ex);
                return null;
            }
        }

        private object GetRibbonControlInternal(string controlName)
        {
            if (_addinModule == null)
                return null;

            var t = _addinModule.GetType();

            try
            {
                var prop = t.GetProperty(controlName, System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.Public);
                if (prop != null && prop.CanRead)
                {
                    try { return prop.GetValue(_addinModule, null); }
                    catch (Exception ex) { _logger?.LogException(ex, $"RibbonController.GetRibbonControlInternal: public property read failed for '{controlName}'"); }
                }
            }
            catch (Exception ex) { _logger?.LogException(ex, $"RibbonController.GetRibbonControlInternal: GetProperty lookup failed for '{controlName}'"); }

            try
            {
                var field = t.GetField(controlName, System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.Public);
                if (field != null)
                {
                    try { return field.GetValue(_addinModule); }
                    catch (Exception ex) { _logger?.LogException(ex, $"RibbonController.GetRibbonControlInternal: public field read failed for '{controlName}'"); }
                }
            }
            catch (Exception ex) { _logger?.LogException(ex, $"RibbonController.GetRibbonControlInternal: GetField lookup failed for '{controlName}'"); }

            try
            {
                var nonPublicField = t.GetField(controlName, System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.NonPublic);
                if (nonPublicField != null)
                {
                    try
                    {
                        var val = nonPublicField.GetValue(_addinModule);
                        if (val != null)
                            return val;
                    }
                    catch (Exception ex) { _logger?.LogException(ex, $"RibbonController.GetRibbonControlInternal: non-public field read failed for '{controlName}'"); }
                }
            }
            catch (Exception ex) { _logger?.LogException(ex, $"RibbonController.GetRibbonControlInternal: GetField (non-public) lookup failed for '{controlName}'"); }

            return null;
        }

        #endregion
    }
}