// GLBalanceConfigurator.xaml.cs in GLSense.Addin.Core
// Group H (Balance Configurator) - port of GLSense\Views\GLBalanceConfigurator.xaml.cs
// (FinalWorkingCode). The WPF content hosted (via Views\ConfiguratorPaneHost.cs, already
// written - see that file's header for the full HWND-reparenting rationale) inside the
// host's GLConfiguratorPane task pane.
//
// Architectural changes vs. the original (see task brief for the full rationale):
//   1. Constructor is now parameterless: "public GLBalanceConfigurator()" - the old
//      "GLBalanceConfigurator(GLConfiguratorPane parentPane = null)" and its "_parentPane"
//      field are both gone. GLConfiguratorPane is a host-only WinForms type
//      (AddinExpress.XL.ADXExcelTaskPane) that cannot be referenced from this project, and
//      in this architecture GLBalanceConfigurator is no longer embedded via an ElementHost
//      directly inside that pane - ConfiguratorPaneHost.cs instead puts it inside its own
//      self-contained WPF Window, whose native HWND the host then Win32-reparents into its
//      task pane panel (SetParent + style-bit rewrite + MoveWindow-on-resize).
//   2. "_parentPane.Resize" -> this control's own SizeChanged event. The host's
//      GLConfiguratorPane.ResizeContent() (GLSense\GLConfiguratorPane.cs) already calls
//      MoveWindow on the reparented HWND to match its own ClientSize, so this control's
//      own bounds already reflect what the host pane is doing - no parent-pane resize
//      event needs to be listened to separately. See OnSizeChanged below.
//   3. The old EnsureMinimumWidth() additionally pushed a minimum width up into the parent
//      pane ("_parentPane.Width = (int)MinimumConfiguratorWidth"). Dropped: the host's own
//      GLConfiguratorPane.SetBoundsCore/WndProc overrides already enforce a DPI-aware
//      minimum size (600x300 dip) independently of anything this control does, so pushing
//      a width up into a parent (which no longer exists as a reference here anyway) is
//      both impossible and unnecessary.
//
// OnCloseRequested / ReLoadConfigurator / ResetCellReference (relied on by
// Views\ConfiguratorPaneHost.cs, already written against this exact contract - see that
// file's header for how it calls into these three members):
//   - These three members ALREADY EXISTED with these exact names/signatures in the old
//     GLBalanceConfigurator.xaml.cs (this is not new design - the old file already had
//     "public event Action OnCloseRequested;", "public async Task ReLoadConfigurator()",
//     and "public static void ResetCellReference()", all driven by the old
//     GLConfiguratorPane.cs's RelaunchPane()/ResetPaneReference() calling into them and by
//     BtnCancelBottom_Click raising OnCloseRequested). Porting them is therefore a direct,
//     unchanged copy of this control's own pre-existing API - only the ExcelApp/logging
//     source changed (ServiceLocator instead of AppState.Instance.ExcelApp/LogUtility).
//   - OnCloseRequested is raised from both the header's inline close (X) button and the
//     bottom Cancel button (BtnCancelBottom_Click), exactly as in the old file - both were
//     already wired to the same handler there.
//
// Other re-pointing vs. the original (no logic changes):
//   - GLSense.Utilities.LogUtility.* (static) -> Infrastructure.ServiceLocator.Logger?.*.
//   - AppState.Instance.ExcelApp.Application -> Infrastructure.ServiceLocator.ExcelApp
//     (this project's AppState has no ExcelApp property - same fix already applied by
//     every other ported window/ViewModel that touches Excel COM).
//   - GLSense.Repositories.DataRepository -> GLSense.Addin.Core.Repositories.DataRepository
//     (GetTableItemsCount, static, already ported, identical signature).
//   - GLSense.Utilities.CommonFunctions -> GLSense.Addin.Core.Utilities.CommonFunctions
//     (FillResponsibilitiesAsync/RemoveInDirect/FormulaParameters/FormulaValues, already
//     ported, identical signatures).
//   - GLSense.Models.LedgerRecord -> GLSense.Addin.Core.Models.LedgerRecord (already
//     ported, identical shape).
//   - AppState/AppConstants/GlobalStateViewModel resolve without an explicit "using" -
//     same namespace-nesting resolution already relied on throughout this project (see
//     ViewModels\GLConfiguratorViewModel.cs's header for the detailed explanation).
// No functional changes to the cell-parsing/formula-detection/reload flow vs. the
// original.
using GLSense.Addin.Core.Extensions;
using GLSense.Addin.Core.Helpers;
using GLSense.Addin.Core.Infrastructure;
using GLSense.Addin.Core.Interfaces;
using GLSense.Addin.Core.Models;
using GLSense.Addin.Core.Repositories;
using GLSense.Addin.Core.Utilities;
using GLSense.Addin.Core.ViewModels;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Threading;
using Excel = Microsoft.Office.Interop.Excel;

namespace GLSense.Addin.Core.Views
{
    /// <summary>
    /// Interaction logic for GLBalanceConfigurator.xaml
    /// </summary>
    public partial class GLBalanceConfigurator : UserControl, IWarningHost
    {
        private const double MinimumConfiguratorWidth = 600;

        private sealed class CellData
        {
            public string Address { get; set; }
            public string LedgerName { get; set; }
            public List<string> LedgerNames { get; set; }
            public List<string> FuncArgs { get; set; }
            public List<string> FuncValues { get; set; }
            public bool ZeroesChecked { get; set; }
        }

        private sealed class FormulaInfo
        {
            public List<string> FuncArgs { get; set; }
            public List<string> FuncValues { get; set; }
        }

        private readonly GLConfiguratorViewModel vm;

        /// <summary>
        /// Raised when the user asks to close this pane (header X or bottom Cancel
        /// button). ConfiguratorPaneHost.cs subscribes to this
        /// (OnContentCloseRequested -> ServiceLocator.RibbonController?.HideConfiguratorPane())
        /// so the host can hide its task pane in response. This is a real, actively-raised
        /// event (not a no-op hook) - both close affordances in the XAML wire to it, same
        /// as the old monolith's GLConfiguratorPane which set
        /// "_wpfControl.OnCloseRequested += () => this.Visible = false;".
        /// </summary>
        public event Action OnCloseRequested;

        public GLBalanceConfigurator()
        {
            InitializeComponent();
            ServiceLocator.Logger?.LogDebug("GLBalanceConfigurator constructor invoked");

            MinWidth = MinimumConfiguratorWidth;
            MainScrollViewer.MinWidth = MinimumConfiguratorWidth;

            // OISR: scroll wheel/touchpad gestures did nothing here (only click-and-drag
            // of the scrollbar thumb worked). This control's own host Window (created by
            // ConfiguratorPaneHost.CreateContent) is a plain System.Windows.Window, not a
            // BaseWindow, so it doesn't get BaseWindow's blanket fix - and even if it did,
            // HWND-reparenting into the host's task pane never gives it Win32 keyboard
            // focus just from mouse hover, which is what WM_MOUSEWHEEL routing requires.
            // See MouseWheelFocusHelper for the full explanation.
            MouseWheelFocusHelper.EnableHoverToScroll(this);

            // Follow-up OISR finding: once the pane has focus, wheel behavior still
            // varied wildly depending on which control was under the cursor - the top
            // field rows (wrapped in their own nested ScrollViewer) silently swallowed
            // it, ExcelRefEditControl's own TextBox scrolled its text instead of the
            // page, and the RichTextBox (Balance Parameters) let it fall through to
            // Excel's worksheet. Intercepting at the tunneling stage on the root, before
            // any of those controls get a chance to touch the event, makes every one of
            // them behave the same way. See OnPreviewMouseWheel below.
            this.PreviewMouseWheel += OnPreviewMouseWheel;

            vm = new GLConfiguratorViewModel(Dispatcher)
            {
                ExcelApp = ServiceLocator.ExcelApp,
                ShowWarningAction = (msg) => Dispatcher.Invoke(() => AppOverlayControl.ShowWarning(msg)),
                ShowInfoAction = (msg) => Dispatcher.Invoke(() => AppOverlayControl.ShowInfo(msg)),
                ShowBusyAction = async (txt, cancel) =>
                        await Dispatcher.InvokeAsync(async () =>
                            await AppOverlayControl.ShowBusyasynTask(txt, cancel)),
                HideBusyAsyncAction = async () => await Dispatcher.InvokeAsync(async () => await AppOverlayControl.HideBusyAsync())
            };

            DataContext = vm;

            // Subscribe to events
            this.Loaded += OnLoaded;
            this.SizeChanged += OnSizeChanged;
            this.IsVisibleChanged += OnIsVisibleChanged;

            // GLSegmentRef (the "Account Assignment Configurator" popup) is now ported -
            // wire AcctsRef's Edit button to open it, mirroring the old monolith's
            // GLAccountsRef.xaml.cs BtnEdit_Click (see AcctsRef_EditRequested below).
            AcctsRef.EditRequested += AcctsRef_EditRequested;
        }

        // Centralizes ALL wheel-scrolling through MainScrollViewer, regardless of which
        // child control is under the cursor - see the OISR follow-up comment in the
        // constructor. Handling this during the tunneling (Preview) phase on the root
        // means it runs before any descendant control (nested ScrollViewer, TextBox,
        // RichTextBox) gets a chance to swallow, redirect, or leak the event out to
        // Excel's own worksheet.
        private void OnPreviewMouseWheel(object sender, MouseWheelEventArgs e)
        {
            try
            {
                if (MainScrollViewer == null) return;

                int notches = e.Delta / 120;
                if (notches == 0)
                    notches = e.Delta > 0 ? 1 : -1;

                int lines = Math.Abs(notches) * 3;
                for (int i = 0; i < lines; i++)
                {
                    if (notches > 0) MainScrollViewer.LineUp();
                    else MainScrollViewer.LineDown();
                }

                e.Handled = true;
            }
            catch (Exception ex)
            {
                ServiceLocator.Logger?.LogException(ex, "GLBalanceConfigurator.OnPreviewMouseWheel");
            }
        }

        /// <summary>
        /// Opens the "Account Assignment Configurator" picker and copies the user's
        /// selection back into AcctsRef.Text on OK - same flow the old monolith's
        /// GLAccountsRef.xaml.cs BtnEdit_Click had inline, now split across AcctsRef's
        /// EditRequested event (this control doesn't reference the picker window itself)
        /// and this handler (which does, since both live in this project's Views
        /// namespace already).
        ///
        /// TRIAL: currently opens GLSegmentManager (the master-detail redesign - segment
        /// list on the left, Value/Reference/Hierarchy/Search/dual-grid detail panel on
        /// the right that rebinds to whichever segment is selected) instead of the
        /// original GLSegmentRef, so it gets exercised through real usage. GLSegmentRef is
        /// untouched and still fully compiled into the project - if GLSegmentManager turns
        /// up problems, roll back by changing the one line below from
        /// "new GLSegmentManager(AcctsRef.Text)" to "new GLSegmentRef(AcctsRef.Text)".
        /// Both share the exact same constructor signature and GLSegments_SelectedValue
        /// contract, so nothing else here needs to change either way.
        /// </summary>
        private void AcctsRef_EditRequested(object sender, EventArgs e)
        {
            ServiceLocator.Logger?.LogDebug("GLBalanceConfigurator.AcctsRef_EditRequested invoked");
            try
            {
                var dlg = new GLSegmentManager(AcctsRef.Text)
                {
                    EnableExcelCentering = false,
                    WindowStartupLocation = WindowStartupLocation.CenterOwner
                };

                if (dlg.ShowDialog() == true)
                {
                    AcctsRef.Text = dlg.GLSegments_SelectedValue;
                    ServiceLocator.Logger?.LogDebug($"GLBalanceConfigurator.AcctsRef_EditRequested: user confirmed selection - {dlg.GLSegments_SelectedValue}");
                }
            }
            catch (Exception ex)
            {
                ServiceLocator.Logger?.LogException(ex, "GLBalanceConfigurator.AcctsRef_EditRequested");
            }
        }

        public void CellSelectionWarning(string message)
        {
            try
            {
                AppOverlayControl.ShowWarning(message);
            }
            catch (Exception ex)
            {
                ServiceLocator.Logger?.LogException(ex, "GLBalanceConfigurator.CellSelectionWarning");
            }
        }

        private void OnLoaded(object sender, RoutedEventArgs e)
        {
            ServiceLocator.Logger?.LogDebug("GLBalanceConfigurator.OnLoaded invoked");
            string Address = GetActiveCellInfo();
            GlobalStateViewModel.Instance.ReferenceText = Address;

            EnsureMinimumWidth();
            this.UpdateLayout();
            MainScrollViewer?.UpdateLayout();

            // Initialize DatePicker - creates tooltip automatically
            dtpStartDate.SetupTooltip(
                title: "Select Start Date",        // Appears in tooltip header
                dispatcher: this.Dispatcher,      // For UI thread safety
                dateFormat: "yyyy-MM-dd",         // Date format
                instructionText: "Click calendar icon to select/change date",  // Footer text
                onDateChangedAction: (dp) =>      // Custom action when date changes
                {
                    vm?.OnDateChanged();
                }
            );

            dtpEndDate.SetupTooltip(
                    title: "Select End Date",        // Appears in tooltip header
                    dispatcher: this.Dispatcher,      // For UI thread safety
                    dateFormat: "yyyy-MM-dd",         // Date format
                    instructionText: "Click calendar icon to select/change date",  // Footer text
                    onDateChangedAction: (dp) =>      // Custom action when date changes
                    {
                        vm?.OnDateChanged();
                    }
                );

            _ = ReLoadConfigurator();
        }

        private void DatePicker_CalendarOpenedEx(object sender, RoutedEventArgs e)
        {
            try
            {
                if (vm == null) return;

                // Determine allowed range from Periods collection. Take the actual min
                // StartDate / max EndDate across the whole collection rather than assuming
                // Periods[0]/Periods[last] are the earliest/latest - custom fiscal calendars
                // (e.g. a "GOV Calendar" period set starting in July instead of January, or
                // one a user has shifted to start in a different quarter) are not guaranteed
                // to come back from the repository in chronological order, so relying on
                // list position silently truncated the selectable range to whatever period
                // happened to be first/last in the query result instead of the true bounds.
                // Ported from FinalWorkingCode's identical fix.
                if (vm.Periods == null || vm.Periods.Count == 0)
                    return;

                DateTime minDate = vm.Periods.Min(p => p.StartDate).Date;
                DateTime maxDate = vm.Periods.Max(p => p.EndDate).Date;

                if (sender is DatePicker dp)
                {
                    // Set the calendar display bounds
                    dp.DisplayDateStart = minDate;
                    dp.DisplayDateEnd = maxDate;

                    // Clear any previous blackout ranges and add blackout for outside range
                    dp.BlackoutDates.Clear();
                    if (minDate > DateTime.MinValue)
                    {
                        dp.BlackoutDates.Add(new CalendarDateRange(DateTime.MinValue, minDate.AddDays(-1)));
                    }
                    if (maxDate < DateTime.MaxValue)
                    {
                        dp.BlackoutDates.Add(new CalendarDateRange(maxDate.AddDays(1), DateTime.MaxValue));
                    }

                    if (dp.Name != null && dp.Name == "dtpStartDate")
                    {
                        dp.DisplayDate = minDate;
                    }
                    else if (dp.Name != null && dp.Name == "dtpEndDate")
                    {
                        dp.DisplayDate = maxDate;
                    }
                }
            }
            catch (Exception ex)
            {
                ServiceLocator.Logger?.LogException(ex, "DatePicker_CalendarOpenedEx");
            }
        }

        private void OnIsVisibleChanged(object sender, DependencyPropertyChangedEventArgs e)
        {
            if (this.IsVisible)
            {
                Dispatcher.BeginInvoke(new Action(() =>
                {
                    EnsureMinimumWidth();
                    this.UpdateLayout();
                    MainScrollViewer?.UpdateLayout();
                }), DispatcherPriority.Loaded);
            }
        }

        private void OnSizeChanged(object sender, SizeChangedEventArgs e)
        {
            // Replaces the old "_parentPane.Resize" subscription - see this file's header.
            // The host's GLConfiguratorPane.ResizeContent() already keeps this control's
            // reparented HWND in sync with its own ClientSize via MoveWindow, so this
            // control's own SizeChanged carries the same signal the old ParentPaneResize
            // handler used to react to.
            Dispatcher.BeginInvoke(new Action(() =>
            {
                EnsureMinimumWidth();
                this.UpdateLayout();
                MainScrollViewer?.UpdateLayout();
            }), DispatcherPriority.Loaded);
        }

        private void EnsureMinimumWidth()
        {
            this.MinWidth = MinimumConfiguratorWidth;
            if (MainScrollViewer != null)
            {
                MainScrollViewer.MinWidth = MinimumConfiguratorWidth;
            }

            // Old code also pushed this minimum width up into the parent ADX task pane
            // ("_parentPane.Width = ..."). Dropped - see this file's header: the host's
            // GLConfiguratorPane already enforces a DPI-aware 600x300 dip minimum size on
            // its own, independent of this control.
        }

        private static string GetActiveCellInfo()
        {
            var rng = ServiceLocator.ExcelApp.ActiveCell;
            string sheetName = ((Excel.Worksheet)rng.Parent).Name;
            string cellAddress = rng.Address[true, true, Excel.XlReferenceStyle.xlA1, false];
            return $"'{sheetName}'!{cellAddress}";
        }

        private async Task ShowBusyOverlayAsync(CancellationHelper helper, string message)
        {
            void ShowOverlay()
            {
                AppOverlayControl.ShowBusyasyn(
                    message: message + " (Click cancel to stop)",
                    cancelAction: async () =>
                    {
                        if (!helper.IsCancellationRequested)
                        {
                            helper.Cancel();
                            ServiceLocator.Logger?.LogWarn($"Operation cancelled by user: {message}");
                        }
                        await Task.Delay(80); // small feedback delay
                                              // HideBusyAsync() is already called in the overlay's handler
                    }
                );
            }

            // Callers of ExecuteWithBusyOverlay (e.g. ReLoadConfigurator, invoked via
            // WpfAppManager.InvokeOnWpfThread) are already marshaled onto this
            // control's Dispatcher thread before this method runs. Routing through an
            // awaited Dispatcher.InvokeAsync(..., Background) here was an unnecessary
            // extra hop, and that hop's DispatcherOperation await was observed to
            // sometimes resume ExecuteWithBusyOverlay's subsequent "await
            // action(helper)" continuation on a thread other than this Dispatcher's
            // own thread - causing "The calling thread cannot access this object
            // because a different thread owns it" as soon as the caller's action
            // touched a UI element (e.g. ReLoadConfigurator's
            // BalanceParametersExpander.IsExpanded = false). Calling directly when
            // already on the right thread removes that risk; the InvokeAsync fallback
            // (raised to Send priority) keeps this safe for any future caller that
            // genuinely is on a background thread.
            if (Dispatcher.CheckAccess())
            {
                ShowOverlay();
            }
            else
            {
                await Dispatcher.InvokeAsync(ShowOverlay, DispatcherPriority.Send);
            }
        }

        public async Task ExecuteWithBusyOverlay(
                string message,
                Func<CancellationHelper, Task> action)
        {
            var helper = new CancellationHelper();

            try
            {
                await ShowBusyOverlayAsync(helper, message);
                await action(helper);
            }
            finally
            {
                await AppOverlayControl.HideBusyAsync();
            }
        }

        /// <summary>
        /// Re-reads the active cell's balance formula (if any) and repopulates the
        /// ViewModel's fields from it - old GLConfiguratorPane.RelaunchPane() dispatched
        /// into this exact method (via the WPF thread) whenever the host's task pane was
        /// shown/re-shown; ConfiguratorPaneHost.Relaunch() does the same here.
        /// </summary>
        public async Task ReLoadConfigurator()
        {
            ServiceLocator.Logger?.LogDebug("GLBalanceConfigurator.ReLoadConfigurator invoked");
            try
            {
                await ExecuteWithBusyOverlay("Reloading Configurator", async helper =>
                {
                    // Defensive: guard against this running off the UI thread even if
                    // some other future path into this lambda misses the Dispatcher
                    // marshal that ShowBusyOverlayAsync now guarantees (see its
                    // comment) - this is the exact statement that previously threw
                    // "The calling thread cannot access this object because a
                    // different thread owns it".
                    if (BalanceParametersExpander.Dispatcher.CheckAccess())
                    {
                        BalanceParametersExpander.IsExpanded = false;
                    }
                    else
                    {
                        BalanceParametersExpander.Dispatcher.Invoke(() => BalanceParametersExpander.IsExpanded = false);
                    }

                    if (!HasValidCubeAndLedger())
                    {
                        ServiceLocator.Logger?.LogDebug("GLBalanceConfigurator.ReLoadConfigurator: no cube/ledger selected, aborting reload");
                        return;
                    }

                    var cellData = await ExtractCellDataAsync();
                    var config = ProcessBalanceFormula(cellData);
                    await LoadConfiguratorDataAsync(config);
                    ServiceLocator.Logger?.LogDebug("GLBalanceConfigurator.ReLoadConfigurator: reload completed successfully");
                });
            }
            catch (Exception ex)
            {
                ServiceLocator.Logger?.LogException(ex, "GLBalanceConfigurator.ReLoadConfigurator");
            }
        }

        /// <summary>
        /// Re-reads the active cell's address into GlobalStateViewModel.ReferenceText -
        /// old GLConfiguratorPane.ResetPaneReference() dispatched into this exact method
        /// (via the WPF thread) on SheetSelectionChange when the new active cell has no
        /// balance formula; ConfiguratorPaneHost.ResetReference() does the same here.
        /// </summary>
        public static void ResetCellReference()
        {
            string Address = GetActiveCellInfo();
            ServiceLocator.Logger?.LogDebug($"GLBalanceConfigurator.ResetCellReference invoked: address={Address}");
            GlobalStateViewModel.Instance.ReferenceText = Address;
        }

        private CellData ProcessBalanceFormula(CellData cellData)
        {
            // If no balance formula detected, use default values
            if (!IsBalanceFormulaDetected(cellData))
            {
                return new CellData
                {
                    LedgerName = AppState.Instance.SelectedLedger.LedgerName,
                    LedgerNames = new List<string> { AppState.Instance.SelectedLedger.LedgerName },
                    FuncArgs = null,
                    FuncValues = null,
                    ZeroesChecked = true,
                    Address = cellData.Address
                };
            }

            var ledgerNames = cellData.LedgerNames ?? GetLedgerNamesFromFormula(cellData.FuncValues);
            var primaryLedgerName = ledgerNames.FirstOrDefault(name => !string.IsNullOrWhiteSpace(name)) ?? AppState.Instance.SelectedLedger.LedgerName;

            return new CellData
            {
                LedgerName = primaryLedgerName,
                LedgerNames = ledgerNames,
                FuncArgs = cellData.FuncArgs,
                FuncValues = cellData.FuncValues,
                ZeroesChecked = cellData.ZeroesChecked,
                Address = cellData.Address
            };
        }

        private static bool IsBalanceFormulaDetected(CellData cellData)
        {
            return !string.IsNullOrWhiteSpace(cellData.LedgerName);
        }

        private async Task<CellData> ExtractCellDataAsync()
        {
            return await Dispatcher.InvokeAsync(() =>
            {
                var rng = ServiceLocator.ExcelApp.ActiveCell;
                var cellData = ParseCellData(rng);
                GlobalStateViewModel.Instance.ReferenceText = cellData.Address;
                return cellData;
            });
        }

        private CellData ParseCellData(Excel.Range rng)
        {
            var sheetName = ((Excel.Worksheet)rng.Parent).Name;
            var cellAddress = rng.Address[true, true, Excel.XlReferenceStyle.xlA1, false];
            var address = $"'{sheetName}'!{cellAddress}";

            var ledgerName = string.Empty;
            var funcArgs = (List<string>)null;
            var funcValues = (List<string>)null;
            var zeroesChecked = true;

            if ((bool)rng.HasFormula && IsBalanceFormula(rng))
            {
                var formulaInfo = ParseBalanceFormula(rng);
                var ledgerNames = GetLedgerNamesFromFormula(formulaInfo.FuncValues);
                ledgerName = ledgerNames.FirstOrDefault() ?? string.Empty;
                funcArgs = formulaInfo.FuncArgs;
                funcValues = formulaInfo.FuncValues;
                zeroesChecked = ParseZeroesSetting(rng.NumberFormat.ToString());

                return new CellData
                {
                    Address = address,
                    LedgerName = ledgerName,
                    LedgerNames = ledgerNames,
                    FuncArgs = funcArgs,
                    FuncValues = funcValues,
                    ZeroesChecked = zeroesChecked
                };
            }

            return new CellData
            {
                Address = address,
                LedgerName = ledgerName,
                LedgerNames = ledgerName == string.Empty ? new List<string>() : new List<string> { ledgerName },
                FuncArgs = funcArgs,
                FuncValues = funcValues,
                ZeroesChecked = zeroesChecked
            };
        }

        private static bool IsBalanceFormula(Excel.Range rng)
        {
            return rng.Formula.ToString().IndexOf(AppConstants.glBal, StringComparison.OrdinalIgnoreCase) >= 0;
        }

        private FormulaInfo ParseBalanceFormula(Excel.Range rng)
        {
            string funcName = rng.Formula.ToString();
            return new FormulaInfo
            {
                FuncArgs = CommonFunctions.FormulaParameters(funcName),
                FuncValues = CommonFunctions.FormulaValues(funcName)
            };
        }

        private static List<string> GetLedgerNamesFromFormula(List<string> funcValues)
        {
            if (funcValues?.Count > 1)
            {
                var ledgerName = funcValues[1].Replace("\"", "");
                return ledgerName
                    .Split(new[] { ';' }, StringSplitOptions.None)
                    .Select(name => name.Trim())
                    .ToList();
            }

            return new List<string>();
        }

        private static bool ParseZeroesSetting(string cellFormatStr)
        {
            try
            {
                if (!cellFormatStr.Contains(";"))
                    return true;

                var parts = cellFormatStr.Split(';');
                return parts.Length < 3 || !string.IsNullOrEmpty(parts[2]);
            }
            catch (Exception ex)
            {
                ServiceLocator.Logger?.LogWarn($"GLBalanceConfigurator.ParseZeroesSetting: failed to parse cell format '{cellFormatStr}' - {ex.Message}");
                return true;
            }
        }

        private async Task LoadConfiguratorDataAsync(CellData cellData)
        {
            var formulaLedgerNames = (cellData.LedgerNames ?? new List<string>())
                .Where(name => !string.IsNullOrWhiteSpace(name))
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .ToList();

            string ledgerName = formulaLedgerNames.FirstOrDefault(name => !string.IsNullOrWhiteSpace(name)) ?? cellData.LedgerName;
            var funcArgs = cellData.FuncArgs;
            var funcValues = cellData.FuncValues;

            if (formulaLedgerNames.Count > 0)
            {
                await EnsureFormulaLedgersLoadedAsync(formulaLedgerNames);
            }
            else
            {
                ledgerName = AppState.Instance.SelectedLedger.LedgerName;
            }

            var ledger = AppState.Instance.SelectedCube.GetLedgerByName(ledgerName);
            if (ledger == null && AppState.Instance.SelectedLedger != null)
            {
                ledger = new LedgerRecord
                {
                    LedgerId = AppState.Instance.SelectedLedger.LedgerId,
                    LedgerName = AppState.Instance.SelectedLedger.LedgerName,
                    Coaid = AppState.Instance.SelectedLedger.CoaId,
                    PeriodSetName = AppState.Instance.SelectedLedger.PeriodSetName,
                    CurrencyCode = AppState.Instance.SelectedLedger.CurrencyCode,
                    PeriodType = string.Empty,
                    LedgerData = string.Empty
                };
            }

            if (ledger == null)
            {
                return;
            }

            await vm.LoadConfiguratorAsync(cellData.ZeroesChecked, ledger, funcArgs, funcValues);
            await Dispatcher.InvokeAsync(() => CmbLedgers.Text = formulaLedgerNames.Count > 0 ? string.Join(";", formulaLedgerNames) : ledgerName);

            // Update datepicker tooltips after periods load so the tooltip shows the
            // dynamic available date range (start..end).
            try
            {
                await Dispatcher.InvokeAsync(() =>
                {
                    try { dtpStartDate.UpdateTooltip(); } catch { ServiceLocator.Logger?.LogWarn("Exception in setting start date tooltip"); }
                    try { dtpEndDate.UpdateTooltip(); } catch { ServiceLocator.Logger?.LogWarn("Exception in setting end date tooltip"); }
                }, System.Windows.Threading.DispatcherPriority.Background);
            }
            catch (Exception ex)
            {
                ServiceLocator.Logger?.LogWarn($"Exception in updating datepicker tooltips: {ex.Message}");
            }
        }

        private static async Task EnsureFormulaLedgersLoadedAsync(IEnumerable<string> formulaLedgerNames)
        {
            var cube = AppState.Instance.SelectedCube;
            if (cube == null)
                return;

            foreach (var ledgerName in formulaLedgerNames)
            {
                var ledgerRecord = cube.GetLedgerByName(ledgerName);
                if (ledgerRecord == null)
                    continue;

                var segmentCount = DataRepository.GetTableItemsCount(cube.CubeId, ledgerRecord.LedgerId, "SEGMENTS");
                if (segmentCount == 0)
                {
                    await CommonFunctions.FillResponsibilitiesAsync(ledgerRecord.LedgerId, cube.CubeId, CancellationToken.None);
                }
            }
        }

        private static bool HasValidCubeAndLedger()
        {
            return AppState.Instance.SelectedCube != null && AppState.Instance.SelectedLedger != null;
        }

        private void BtnOKBottom_Click(object sender, RoutedEventArgs e)
        {
            ServiceLocator.Logger?.LogDebug("GLBalanceConfigurator.BtnOKBottom_Click invoked");
            if (string.IsNullOrWhiteSpace(CellReference.Text))
            {
                ServiceLocator.Logger?.LogDebug("GLBalanceConfigurator.BtnOKBottom_Click: validation failed - cell reference is empty");
                AppOverlayControl.ShowWarning("Please select a cell reference for entering formula.");
            }
            else
            {
                try
                {
                    if (!string.IsNullOrWhiteSpace(CellReference.Text))
                    {
                        Excel.Range rng = CommonFunctions.RemoveInDirect(CellReference.Text);
                        if (rng != null)
                        {
                            CellFormat(rng);
                            ServiceLocator.Logger?.LogDebug("GLBalanceConfigurator.BtnOKBottom_Click: formula written to cell successfully");
                        }
                        else
                        {
                            ServiceLocator.Logger?.LogWarn($"GLBalanceConfigurator.BtnOKBottom_Click: cell reference '{CellReference.Text}' did not resolve to a valid range");
                            AppOverlayControl.ShowWarning("The specified cell reference does not refer to a valid cell in the current workbook.");
                        }
                    }
                    else
                    {
                        AppOverlayControl.ShowWarning("Cell reference for get balance cannot be blank.Try providing a cell reference for generating balance formula.");
                    }
                }
                catch (Exception ex)
                {
                    ServiceLocator.Logger?.LogException(ex, "GLBalanceConfigurator.BtnOKBottom_Click");
                }
            }
        }

        private void CellFormat(Excel.Range rng)
        {
            if (vm.IsZeroesChecked)
            {
                rng.NumberFormat = "#,##0.00_);[Red](#,##0.00)";
            }
            else
            {
                rng.NumberFormat = "#,##0.00_);[Red](#,##0.00);;@";
            }
            vm.WriteFormulaToCell(rng);
        }

        private void BtnCancelBottom_Click(object sender, RoutedEventArgs e)
        {
            ServiceLocator.Logger?.LogDebug("GLBalanceConfigurator.BtnCancelBottom_Click invoked - raising OnCloseRequested");
            OnCloseRequested?.Invoke();
        }
    }
}
