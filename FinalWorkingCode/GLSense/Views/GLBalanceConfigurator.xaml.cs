using GLSense.Extensions;
using GLSense.Helpers;
using GLSense.Interfaces;
using GLSense.Models;
using GLSense.Repositories;
using GLSense.Utilities;
using GLSense.ViewModels;
using MahApps.Metro.IconPacks;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Threading;
using Excel = Microsoft.Office.Interop.Excel;

namespace GLSense.Views
{
    /// <summary>
    /// Interaction logic for GLBalanceConfigurator.xaml
    /// </summary>
    public partial class GLBalanceConfigurator : System.Windows.Controls.UserControl, IWarningHost
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
        private GLConfiguratorPane _parentPane;
        public event Action OnCloseRequested;

        public GLBalanceConfigurator(GLConfiguratorPane parentPane = null)
        {
            LogUtility.LogDebug($"GLBalanceConfigurator.ctor invoked - parentPane={(parentPane != null)}");
            InitializeComponent();

            MinWidth = MinimumConfiguratorWidth;
            MainScrollViewer.MinWidth = MinimumConfiguratorWidth;

            // OISR: scroll wheel/touchpad gestures did nothing here (only click-and-drag
            // of the scrollbar thumb worked) because this control is embedded into the
            // Excel task pane via ElementHost, which is never given Win32 keyboard focus
            // just from mouse hover - and WM_MOUSEWHEEL is routed to whichever window
            // currently has focus, not whichever is under the cursor. See
            // MouseWheelFocusHelper for the full explanation.
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
                ExcelApp = AppState.Instance.ExcelApp.Application, // Pass the Excel application instance to the ViewModel
                ShowWarningAction = (msg) => Dispatcher.Invoke(() => AppOverlayControl.ShowWarning(msg)),
                ShowInfoAction = (msg) => Dispatcher.Invoke(() => AppOverlayControl.ShowInfo(msg)),
                ShowBusyAction = async (txt, cancel) =>
                        await Dispatcher.InvokeAsync(async () =>
                            await AppOverlayControl.ShowBusyasynTask(txt, cancel)),
                HideBusyAsyncAction = async () => await Dispatcher.InvokeAsync(async () => await AppOverlayControl.HideBusyAsync())
            };

            _parentPane = parentPane;
            DataContext = vm;

            // Subscribe to events
            this.Loaded += OnLoaded;
            this.SizeChanged += OnSizeChanged;
            this.IsVisibleChanged += OnIsVisibleChanged;

            if (_parentPane != null)
            {
                _parentPane.Resize += OnParentPaneResize;
            }
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
                LogUtility.LogException(ex, "GLBalanceConfigurator.OnPreviewMouseWheel");
            }
        }

        public void CellSelectionWarning(string message)
        {
            LogUtility.LogDebug($"GLBalanceConfigurator.CellSelectionWarning invoked - message={message}");
            try
            {
                AppOverlayControl.ShowWarning(message);
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex, "GLBalanceConfigurator.CellSelectionWarning");
            }
        }
        private void OnLoaded(object sender, RoutedEventArgs e)
        {
            LogUtility.LogDebug("GLBalanceConfigurator.OnLoaded invoked");
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

                // Determine allowed range from Periods collection
                if (vm.Periods == null || vm.Periods.Count == 0)
                    return;

                DateTime minDate = vm.Periods[0].StartDate.Date;
                DateTime maxDate = vm.Periods[vm.Periods.Count - 1].EndDate.Date;

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

                    //// If selected date is null or outside range, and today is outside range,
                    //// select the last available date in the range (maxDate)
                    //DateTime? selected = dp.SelectedDate;
                    //DateTime today = DateTime.Today;
                    //bool todayInRange = today >= minDate && today <= maxDate;

                    //if ((selected == null || selected < minDate || selected > maxDate) && !todayInRange)
                    //{
                    //    dp.SelectedDate = maxDate;
                    //}
                }
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex, "DatePicker_CalendarOpenedEx");
            }
        }
        private void OnIsVisibleChanged(object sender, DependencyPropertyChangedEventArgs e)
        {
            LogUtility.LogDebug($"GLBalanceConfigurator.OnIsVisibleChanged invoked - IsVisible={this.IsVisible}");
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
            EnsureMinimumWidth();
        }
        private void OnParentPaneResize(object sender, EventArgs e)
        {
            LogUtility.LogDebug("GLBalanceConfigurator.OnParentPaneResize invoked");
            // Update control when task pane resizes
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

            if (_parentPane != null && _parentPane.Width < MinimumConfiguratorWidth)
            {
                _parentPane.Width = (int)MinimumConfiguratorWidth;
            }
        }
        private static string GetActiveCellInfo()
        {
            var rng = AppState.Instance.ExcelApp.ActiveCell;
            string sheetName = ((Excel.Worksheet)rng.Parent).Name;
            string cellAddress = rng.Address[true, true, Excel.XlReferenceStyle.xlA1, false];
            return $"'{sheetName}'!{cellAddress}";
        }
        private async Task ShowBusyOverlayAsync(CancellationHelper helper, string message)
        {
            await Dispatcher.InvokeAsync(() =>
            {
                AppOverlayControl.ShowBusyasyn(
                    message: message + " (Click cancel to stop)",
                    cancelAction: async () =>
                    {
                        if (!helper.IsCancellationRequested)
                        {
                            helper.Cancel();
                            LogUtility.LogWarn($"Operation cancelled by user: {message}");
                        }
                        await Task.Delay(80); // small feedback delay
                                              // HideBusyAsync() is already called in the overlay's handler
                    }
                );
            }, DispatcherPriority.Background);
        }
        public async Task ExecuteWithBusyOverlay(
                string message,
                Func<CancellationHelper, Task> action)
        {
            LogUtility.LogDebug($"GLBalanceConfigurator.ExecuteWithBusyOverlay invoked - message={message}");
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
        public async Task ReLoadConfigurator()
        {
            LogUtility.LogDebug("GLBalanceConfigurator.ReLoadConfigurator invoked");
            await ExecuteWithBusyOverlay("Reloading Configurator", async helper =>
            {
                BalanceParametersExpander.IsExpanded = false;

                if (!HasValidCubeAndLedger())
                {
                    LogUtility.LogDebug("GLBalanceConfigurator.ReLoadConfigurator: no valid cube/ledger selected, aborting reload");
                    return;
                }

                var cellData = await ExtractCellDataAsync();
                var config = ProcessBalanceFormula(cellData);
                await LoadConfiguratorDataAsync(config);
            });
        }
        public static void ResetCellReference()
        {
            LogUtility.LogDebug("GLBalanceConfigurator.ResetCellReference invoked");
            string Address = GetActiveCellInfo();
            GlobalStateViewModel.Instance.ReferenceText = Address;
        }
        private CellData ProcessBalanceFormula(CellData cellData)
        {
            // If no balance formula detected, use default values
            if (!IsBalanceFormulaDetected(cellData))
            {
                LogUtility.LogDebug($"GLBalanceConfigurator.ProcessBalanceFormula: no balance formula detected at {cellData.Address}, using default ledger values");
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
                var rng = AppState.Instance.ExcelApp.ActiveCell;
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
                LogUtility.LogDebug($"GLBalanceConfigurator.ParseCellData: balance formula detected at {address}");
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
                LogUtility.LogException(ex, "GLBalanceConfigurator.ParseZeroesSetting");
                return true;
            }
        }
        private async Task LoadConfiguratorDataAsync(CellData cellData)
        {
            LogUtility.LogDebug($"GLBalanceConfigurator.LoadConfiguratorDataAsync invoked - address={cellData?.Address}, ledgerName={cellData?.LedgerName}");
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
                LogUtility.LogDebug($"GLBalanceConfigurator.LoadConfiguratorDataAsync: no ledger found for ledgerName={ledgerName}, aborting load");
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
                    try { dtpStartDate.UpdateTooltip(); } catch { LogUtility.LogWarn($"Exception in setting start date tooltip"); }
                    try { dtpEndDate.UpdateTooltip(); } catch { LogUtility.LogWarn($"Exception in setting end date tooltip"); }
                }, System.Windows.Threading.DispatcherPriority.Background);
            }
            catch (Exception ex)
            {
                LogUtility.LogWarn($"Exception in updating datepicker tooltips: {ex.Message}");
            }
        }

        private static async Task EnsureFormulaLedgersLoadedAsync(IEnumerable<string> formulaLedgerNames)
        {
            LogUtility.LogDebug($"GLBalanceConfigurator.EnsureFormulaLedgersLoadedAsync invoked - ledgerNames={string.Join(",", formulaLedgerNames ?? new List<string>())}");
            var cube = AppState.Instance.SelectedCube;
            if (cube == null)
            {
                LogUtility.LogDebug("GLBalanceConfigurator.EnsureFormulaLedgersLoadedAsync: no selected cube, aborting");
                return;
            }

            foreach (var ledgerName in formulaLedgerNames)
            {
                var ledgerRecord = cube.GetLedgerByName(ledgerName);
                if (ledgerRecord == null)
                    continue;

                var segmentCount = DataRepository.GetTableItemsCount(cube.CubeId, ledgerRecord.LedgerId, "SEGMENTS");
                if (segmentCount == 0)
                {
                    LogUtility.LogDebug($"GLBalanceConfigurator.EnsureFormulaLedgersLoadedAsync: segments not loaded for ledger={ledgerName}, fetching responsibilities");
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
            LogUtility.LogDebug($"GLBalanceConfigurator.BtnOKBottom_Click invoked - CellReference={CellReference.Text}");
            if (string.IsNullOrWhiteSpace(CellReference.Text))
            {
                LogUtility.LogDebug("GLBalanceConfigurator.BtnOKBottom_Click: validation failed - cell reference is blank");
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
                        }
                        else
                        {
                            LogUtility.LogDebug($"GLBalanceConfigurator.BtnOKBottom_Click: validation failed - cell reference '{CellReference.Text}' does not resolve to a valid range");
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
                    LogUtility.LogException(ex, "GLBalanceConfigurator.BtnOKBottom_Click");
                }
            }
        }
        private void CellFormat(Excel.Range rng)
        {
            LogUtility.LogDebug($"GLBalanceConfigurator.CellFormat invoked - IsZeroesChecked={vm.IsZeroesChecked}");
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
            LogUtility.LogDebug("GLBalanceConfigurator.BtnCancelBottom_Click invoked");
            OnCloseRequested?.Invoke();
        }
        
    }
}
