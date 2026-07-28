// GLLOVs.xaml.cs in GLSense.Addin.Core
// Port of GLSense\Views\GLLOVs.xaml.cs (FinalWorkingCode) - opened by ribbon button
// RibLOVs (Group H - Balance Configurator pane + LOVs/Roller/Account dialogs, last
// remaining piece). Follows the exact same pattern already established by GLDailyRates.
// xaml.cs/GLSegmentFunctions.xaml.cs (BaseWindow instead of DpiAwareWindow,
// TitleBar_MouseLeftButtonDown instead of EnhancedDragDropHelper.EnableWindowDrag,
// ServiceLocator.ExcelApp instead of AppState.Instance.ExcelApp.Application,
// ServiceLocator.Logger?.* instead of LogUtility.*). Re-pointed additionally: GLSense.
// Repositories.DataRepository -> GLSense.Addin.Core.Repositories.DataRepository;
// GLSense.Service.ServiceLocator.PeriodDataService.GetPeriodsForLedger ->
// GLSense.Addin.Core.Services.DataServiceLocator.PeriodDataService.GetPeriodsForLedger
// (see Services\DataServiceLocator.cs's header for why the rename). No logic changes vs.
// the original.
using GLSense.Addin.Core.Helpers;
using GLSense.Addin.Core.Infrastructure;
using GLSense.Addin.Core.Interfaces;
using GLSense.Addin.Core.Repositories;
using GLSense.Addin.Core.Services;
using GLSense.Addin.Core.Utilities;
using GLSense.Addin.Core.ViewModels;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Input;
using System.Windows.Threading;
using Excel = Microsoft.Office.Interop.Excel;

namespace GLSense.Addin.Core.Views
{
    /// <summary>
    /// Interaction logic for GLLOVs.xaml
    /// </summary>
    public partial class GLLOVs : BaseWindow, IWarningHost
    {
        private readonly GLLovViewModel vm;
        public GLLOVs()
        {
            InitializeComponent();
            ServiceLocator.Logger?.LogDebug("GLLOVs constructor invoked");

            // "Available LOVs" (index 0, previously the highest-weighted 2* column) fills any
            // left-over width instead of leaving a blank gap now that every column is
            // Width="Auto" (see DataGridColumnFillHelper for why the star-width columns were
            // removed).
            DataGridColumnFillHelper.EnableFillColumn(dgLovs, dgLovs.Columns[0]);

            // Add any initialization after the InitializeComponent() call.
            vm = new GLLovViewModel(this.Dispatcher)
            {
                ExcelApp = ServiceLocator.ExcelApp,
                ShowWarningAction = (msg) => Dispatcher.Invoke(() => AppOverlayControl.ShowWarning(msg)),
                ShowBusyAction = async (txt, cancel) =>
                        await Dispatcher.InvokeAsync(async () =>
                            await AppOverlayControl.ShowBusyasynTask(txt, cancel)),
                HideBusyAsyncAction = async () => await Dispatcher.InvokeAsync(() => AppOverlayControl.HideBusyAsync()),
                // LOVRows gets populated fire-and-forget (LOV_SelectedLedger's setter ->
                // LoadLovRows() -> Task.Run(LoadLovRowsAsync)), detached from
                // Window_Loaded's own await chain, so BaseWindow.OnLoaded's SizeToContent
                // resettle always ran against an empty dgLovs. Resettle again once real
                // rows are actually in place. See CLAUDE.md section 1.4b.
                DataLoadedAction = () =>
                {
                    ForceSizeToContentResettle();
                    PumpDispatcherFrame();
                }
            };
            this.DataContext = vm;
        }

        private void TitleBar_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            try
            {
                if (e.LeftButton == MouseButtonState.Pressed)
                {
                    this.DragMove();
                }
            }
            catch (Exception ex)
            {
                ServiceLocator.Logger?.LogException(ex, "TitleBar_MouseLeftButtonDown error");
            }
        }

        private async void Window_Loaded(object sender, RoutedEventArgs e)
        {
            ServiceLocator.Logger?.LogDebug("GLLOVs.Window_Loaded invoked");
            try
            {
                Excel.Range rng = ServiceLocator.ExcelApp.ActiveCell;
                string sheetName = ((Excel.Worksheet)rng.Parent).Name;
                string cellAddress = rng.Address[true, true, Excel.XlReferenceStyle.xlA1, false];
                string addr = $"'{sheetName}'!{cellAddress}";

                GlobalStateViewModel.Instance.ReferenceText = addr;

                if (AppState.Instance.SelectedCube != null && AppState.Instance.SelectedLedger != null)
                {
                    ServiceLocator.Logger?.LogDebug($"GLLOVs.Window_Loaded: loading data for cubeId={AppState.Instance.SelectedCube.CubeId}, ledgerId={AppState.Instance.SelectedLedger.LedgerId}");
                    await vm.LoadDataAsync(AppState.Instance.SelectedCube.CubeId, AppState.Instance.SelectedLedger.LedgerId);

                    await Dispatcher.InvokeAsync(() =>
                    {
                        cmbLedgers.Text = vm.LOV_SelectedLedger.LedgerName;
                    });
                }
            }
            catch (Exception ex)
            {
                ServiceLocator.Logger?.LogException(ex, "GLLOVs.Window_Loaded");
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
                ServiceLocator.Logger?.LogException(ex, "GLLOVs.CellSelectionWarning");
            }
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
                            ServiceLocator.Logger?.LogWarn($"Operation cancelled by user: {message}");
                        }
                        await Task.Delay(80); // small feedback delay
                                              // HideBusyAsync() is already called in the overlay's handler
                    }
                );
            }, DispatcherPriority.Background);
        }
        private async void CmdSubmit_Click(object sender, RoutedEventArgs e)
        {
            ServiceLocator.Logger?.LogDebug("GLLOVs.CmdSubmit_Click invoked");
            CancellationHelper ctsHelper = new();
            var SelLov = vm.SelectedLov;

            if (SelLov == null)
            {
                ServiceLocator.Logger?.LogDebug("GLLOVs.CmdSubmit_Click: validation failed - no LOV selected");
                await AppOverlayControl.ShowWarningAsync("Please select a LOV to proceed.");
                return;
            }
            if (excelRefEdit.Text == null)
            {
                ServiceLocator.Logger?.LogDebug("GLLOVs.CmdSubmit_Click: validation failed - no range selected");
                await AppOverlayControl.ShowWarningAsync("Please select a range to copy lov.");
                return;
            }
            if (SelLov.ItemsCount == 0)
            {
                ServiceLocator.Logger?.LogDebug("GLLOVs.CmdSubmit_Click: validation failed - selected LOV has no items");
                await AppOverlayControl.ShowWarningAsync("The selected LOV has no items to copy.");
                return;
            }

            try
            {
                await ShowBusyOverlayAsync(ctsHelper, "Please wait while we set the excel dependencies...");

                await CreateLOVSheetAsync();

                var comments = TxtComments.Text.Trim();

                if (string.IsNullOrEmpty(comments))
                {
                    comments = vm.LOV_SelectedLedger.LedgerName;
                }

                if (SelLov.Category == "Segment")
                {
                    comments = $"Segments( {SelLov.Name} ): {comments}";
                }

                if (comments.Length > 255)
                {
                    comments = comments.Substring(0, 255);
                }

                var dvTitle = $"{SelLov.Name}_{vm.LOV_SelectedLedger.LedgerId}";

                string cleanedName = await CleanUpNamedRangeAsync(dvTitle.Trim());

                bool nameExists = NameRangeExists(cleanedName);

                if (!string.IsNullOrEmpty(cleanedName) && nameExists)
                {
                    Excel.Range rng = ServiceLocator.ExcelApp.Range[excelRefEdit.Text];
                    rng.Clear();
                    rng.NumberFormat = "@";

                    var validation = rng.Validation;
                    validation.Delete();
                    validation.Add(
                        Excel.XlDVType.xlValidateList,
                        Excel.XlDVAlertStyle.xlValidAlertStop,
                        Excel.XlFormatConditionOperator.xlBetween,
                        $"={cleanedName}"
                    );
                    validation.IgnoreBlank = true;
                    validation.InCellDropdown = true;
                    validation.ShowInput = true;
                    validation.ErrorTitle = dvTitle;
                    validation.ErrorMessage = "Values should be from the list.";
                    validation.ShowError = true;

                    if (!string.IsNullOrEmpty(comments))
                    {
                        rng.AddComment(comments);
                    }

                    await AppOverlayControl.HideBusyAsync();
                    await AppOverlayControl.ShowInfoAsync("LOV copied successfully.");
                    ServiceLocator.Logger?.LogDebug("GLLOVs.CmdSubmit_Click: LOV copied successfully");
                }
                else
                {
                    ServiceLocator.Logger?.LogWarn($"GLLOVs.CmdSubmit_Click: named range '{cleanedName}' could not be created/found, skipping validation setup");
                }
            }
            catch (OperationCanceledException)
            {
                ServiceLocator.Logger?.LogDebug("GLLOVs.CmdSubmit_Click: operation cancelled by user");
                await AppOverlayControl.HideBusyAsync();
                await AppOverlayControl.ShowWarningAsync("Operation cancelled by user.");
            }
            catch (Exception ex)
            {
                ServiceLocator.Logger?.LogException(ex, "GLLOVs.CmdSubmit_Click");
            }
            finally
            {
                try
                {
                    if (!ctsHelper.IsCancellationRequested)
                        ctsHelper.Cancel();

                    ctsHelper.Dispose();
                }
                catch
                {
                    //ignore cancellation helper dispose exceptions
                }
                await AppOverlayControl.HideBusyAsync();
            }
        }
        private async Task CreateLOVSheetAsync()
        {
            long CubeId = AppState.Instance.SelectedCube.CubeId;
            string CubeName = AppState.Instance.SelectedCube.CubeName;
            long LedgerIdlng = vm.LOV_SelectedLedger.LedgerId;
            string LedgerIdStr = vm.LOV_SelectedLedger.LedgerId.ToString();
            string LedgerName = vm.LOV_SelectedLedger.LedgerName;
            try
            {
                string sheetName;
                int maxLength = 24;

                sheetName = $"{(LedgerIdStr.Length > maxLength ? LedgerIdStr.Substring(0, maxLength) : LedgerIdStr)}_LOV";

                Excel.Worksheet WrkSheet;

                if (!await CommonFunctions.SheetExistsAsync(sheetName))
                {
                    WrkSheet = (Excel.Worksheet)ServiceLocator.ExcelApp.Worksheets.Add();
                    WrkSheet.Name = sheetName;
                }
                else
                {
                    WrkSheet = (Excel.Worksheet)ServiceLocator.ExcelApp.ActiveWorkbook.Worksheets[sheetName];
                }

                WrkSheet.Visible = Excel.XlSheetVisibility.xlSheetHidden;
                WrkSheet.Cells.Clear();

                var ItemsDict = new Dictionary<string, List<string>>();

                var repository = new DataRepository();

                //All segments
                var segments = DataRepository.GetAllSegmentValues(CubeId, LedgerIdlng);

                if (segments != null && segments.Any())
                {
                    // Group segments by SegmentName and process each group
                    var segmentGroups = segments.GroupBy(s => s.SegmentName);

                    foreach (var segmentGroup in segmentGroups)
                    {
                        string segmentName = segmentGroup.Key;
                        List<string> segmentValues = segmentGroup.Select(s => s.SegmentValue).ToList();
                        ItemsDict[segmentName] = segmentValues;
                    }
                }

                //Activity
                var activities = repository.GetActivities(CubeId, LedgerIdlng);
                List<string> activityValues = activities.Select(a => a.ShortName).ToList();
                ItemsDict["Activity"] = activityValues;

                //Balance Activity
                List<string> balanceTypes = new List<string> { "PTD", "YTD", "QTD", "PJTD", "CTD", "JED", "JEDP", "JEDU" };
                ItemsDict["BalanceType"] = balanceTypes;

                //Periods
                var periods = DataServiceLocator.PeriodDataService.GetPeriodsForLedger(LedgerName);
                List<string> periodValues = periods.Select(p => p.PeriodName).ToList();
                ItemsDict["Periods"] = periodValues;

                //Currencies
                var currencies = repository.GetCurrencies(CubeId, LedgerIdlng);
                List<string> currencyValues = currencies.Select(c => c.CurrencyCode).ToList();
                ItemsDict["Currencies"] = currencyValues;

                //Currency Types
                List<string> currencyTypes = new List<string> { "Total", "E", "T", "C" };
                ItemsDict["CurrencyTypes"] = currencyTypes;

                //Actual Flags
                List<string> actualFlags = new List<string> { "A", "B", "E", "A+E" };
                ItemsDict["ActualFlags"] = actualFlags;

                //Budgets
                var budgets = repository.GetBudgets(CubeId, LedgerIdlng);
                List<string> budgetValues = budgets.Select(b => b.BudgetName).ToList();
                ItemsDict["Budgets"] = budgetValues;

                //Encumbrances
                var encumbrances = repository.GetEncumbrances(CubeId, LedgerIdlng);
                List<string> encumbranceValues = encumbrances.Select(e => e.EncumbranceType).ToList();
                ItemsDict["Encumbrances"] = encumbranceValues;

                //Journal Sources
                var journalSources = repository.GetJournalSources(CubeId, LedgerIdlng);
                List<string> sourceValues = journalSources.Select(js => js.SourceName).ToList();
                ItemsDict["JournalSources"] = sourceValues;

                //Journal Categories
                var journalCategories = repository.GetJournalCategories(CubeId, LedgerIdlng);
                List<string> categoryValues = journalCategories.Select(jc => jc.CategoryName).ToList();
                ItemsDict["JournalCategories"] = categoryValues;

                // Write all items to worksheet
                int colIndex = 1;
                foreach (var kvp in ItemsDict)
                {
                    string columnTitle = kvp.Key;
                    List<string> values = kvp.Value;
                    object[,] arr = ToColumnArray(values);
                    await WriteListToColumnAsync(WrkSheet, LedgerIdStr, colIndex++, columnTitle, arr);
                }
            }
            catch (Exception ex)
            {
                await LogErrorAsync(ex, $"Exception occurred while creating LOV sheet for Cube : {CubeName} and ledger : {LedgerName}");
            }
        }


        private static object[,] ToColumnArray<T>(List<T> values)
        {
            if (values == null || values.Count == 0)
                return new object[0, 0];

            var result = new object[values.Count, 1];
            for (int i = 0; i < values.Count; i++)
            {
                result[i, 0] = values[i];
            }
            return result;
        }
        private static async Task WriteListToColumnAsync(Excel.Worksheet ws, string LedgerIDSelected, int colIndex, string columnTitle, object[,] arr)
        {
            string ColTitle = System.Globalization.CultureInfo.CurrentCulture.TextInfo.ToTitleCase(columnTitle.ToLower());

            ((Excel.Range)ws.Cells[1, colIndex]).Value = ColTitle;

            if (arr != null && arr.Length > 0)
            {
                try
                {
                    // Cast ws.Cells[2, colIndex] to Excel.Range before calling Resize
                    Excel.Range startCell = (Excel.Range)ws.Cells[2, colIndex];
                    Excel.Range resizedRange = startCell.Resize[arr.GetLength(0), 1];
                    if (resizedRange != null)
                    {
                        resizedRange.NumberFormat = "@";
                        resizedRange.Value = arr;

                        string columnAddress = resizedRange.Address[true, true, Excel.XlReferenceStyle.xlA1, Type.Missing];
                        string cellAddress = $"'{ws.Name}'!{columnAddress}";
                        string formulaStr = $"=INDIRECT(\"{cellAddress}\")";
                        string rangeName = $"{columnTitle}_{LedgerIDSelected}";

                        await CreateNamedRangeAsync(rangeName, formulaStr);
                    }
                }
                catch (Exception ex)
                {
                    await LogErrorAsync(ex, $"Exception occurred while writing array to excel range in WriteListToColumn");
                }
            }
        }

        private static async Task<string> CleanUpNamedRangeAsync(string nmName)
        {
            string PatternStr = "[^a-zA-Z0-9_]";
            string SubstitutionStr = "";

            try
            {
                return await Task.Run(() =>
                {
                    System.Text.RegularExpressions.Regex regex = new(PatternStr);
                    string TestStr = regex.Replace(nmName, SubstitutionStr);

                    return !string.IsNullOrEmpty(TestStr) ? TestStr : string.Empty;
                });
            }
            catch (Exception ex)
            {
                await LogErrorAsync(ex, $"Exception in cleaning named range : {nmName}");
                return string.Empty;
            }
        }

        private static async Task DeleteNamedRangeAsync(string rngnm)
        {
            try
            {
                if (ServiceLocator.ExcelApp.ActiveWorkbook.Names.Count > 0)
                {
                    foreach (Excel.Name NM in ServiceLocator.ExcelApp.ActiveWorkbook.Names)
                    {
                        if (NM.Name.ToUpper().Trim() == rngnm.ToUpper().Trim())
                        {
                            NM.Delete();
                            return;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                await LogErrorAsync(ex, $"Exception in deleting named range : {rngnm}");
            }
        }

        private static async Task CreateNamedRangeAsync(string nm, string formulaStr)
        {
            try
            {
                string cleanedName = await CleanUpNamedRangeAsync(nm.Trim());

                if (NameRangeExists(cleanedName))
                {
                    await DeleteNamedRangeAsync(cleanedName);
                }

                ServiceLocator.ExcelApp.ActiveWorkbook.Names.Add(cleanedName, RefersToR1C1: formulaStr);
            }
            catch (Exception ex)
            {
                await LogErrorAsync(ex, $"Exception in creating named range : {nm} and formula : {formulaStr}");
            }
        }

        private static bool NameRangeExists(string rangeName)
        {
            if (ServiceLocator.ExcelApp.ActiveWorkbook.Names.Count > 0)
            {
                foreach (Excel.Name name in ServiceLocator.ExcelApp.ActiveWorkbook.Names)
                {
                    if (name.Name.Equals(rangeName, StringComparison.OrdinalIgnoreCase))
                    {
                        return true;
                    }
                }
            }
            return false;
        }

        private static async Task LogErrorAsync(Exception ex, string message = null)
        {
            await Task.Run(() =>
            {
                if (message != null)
                {
                    ServiceLocator.Logger?.LogError($"{message}: {ex.Message}");
                }
                else
                {
                    ServiceLocator.Logger?.LogException(ex);
                }
            });
        }

        private void BtnClose_Click(object sender, RoutedEventArgs e)
        {
            ServiceLocator.Logger?.LogDebug("GLLOVs.BtnClose_Click invoked - closing window");
            Close();
        }
    }
}
