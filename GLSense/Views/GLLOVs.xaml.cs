using GLSense.Helpers;
using GLSense.Interfaces;
using GLSense.Repositories;
using GLSense.Service;
using GLSense.Utilities;
using GLSense.ViewModels;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Threading;
using Excel = Microsoft.Office.Interop.Excel;

namespace GLSense.Views
{
    /// <summary>
    /// Interaction logic for GLLOVs.xaml
    /// </summary>
    public partial class GLLOVs : DpiAwareWindow, IWarningHost
    {
        private readonly GLLovViewModel vm;
        public GLLOVs()
        {
            LogUtility.LogDebug("GLLOVs.ctor invoked");
            InitializeComponent();
            EnhancedDragDropHelper.EnableWindowDrag(this);

            // Add any initialization after the InitializeComponent() call.
            vm = new GLLovViewModel(this.Dispatcher)
            {
                ExcelApp = AppState.Instance.ExcelApp.Application,  // Pass the Excel application instance to the ViewModel
                ShowWarningAction = (msg) => Dispatcher.Invoke(() => AppOverlayControl.ShowWarning(msg)),
                ShowBusyAction = async (txt, cancel) =>
                        await Dispatcher.InvokeAsync(async () =>
                            await AppOverlayControl.ShowBusyasynTask(txt, cancel)),
                HideBusyAsyncAction = async () => await Dispatcher.InvokeAsync(() => AppOverlayControl.HideBusyAsync())
            };
            this.DataContext = vm;
        }
        private async void Window_Loaded(object sender, RoutedEventArgs e)
        {
            LogUtility.LogDebug("GLLOVs.Window_Loaded invoked");
            try
            {
                Excel.Range rng = AppState.Instance.ExcelApp.ActiveCell;
                string sheetName = ((Excel.Worksheet)rng.Parent).Name;
                string cellAddress = rng.Address[true, true, Excel.XlReferenceStyle.xlA1, false];
                string addr = $"'{sheetName}'!{cellAddress}";

                excelRefEdit.Text = addr;
                LogUtility.LogDebug($"GLLOVs.Window_Loaded: active cell reference={addr}");

                if (AppState.Instance.SelectedCube != null && AppState.Instance.SelectedLedger != null)
                {
                    await vm.LoadDataAsync(AppState.Instance.SelectedCube.CubeId, AppState.Instance.SelectedLedger.LedgerId);

                    await Dispatcher.InvokeAsync(() =>
                    {
                        cmbLedgers.Text = vm.LOV_SelectedLedger.LedgerName;
                    });
                    LogUtility.LogDebug("GLLOVs.Window_Loaded: LOV data loaded successfully");
                }
                else
                {
                    LogUtility.LogDebug("GLLOVs.Window_Loaded: validation failed - no cube/ledger selected, skipping load");
                }
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex, "GLLOVs.Window_Loaded");
            }
        }
        public void CellSelectionWarning(string message)
        {
            LogUtility.LogDebug($"GLLOVs.CellSelectionWarning invoked - message={message}");
            try
            {
                AppOverlayControl.ShowWarning(message);
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex, "GLLOVs.CellSelectionWarning");
            }
        }
        private async Task ShowBusyOverlayAsync(CancellationHelper helper, string message)
        {
            LogUtility.LogDebug($"GLLOVs.ShowBusyOverlayAsync invoked - message={message}");
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
        private async void CmdSubmit_Click(object sender, RoutedEventArgs e)
        {
            LogUtility.LogDebug("GLLOVs.CmdSubmit_Click invoked");
            CancellationHelper ctsHelper = new();
            var SelLov = vm.SelectedLov;

            if (SelLov == null)
            {
                LogUtility.LogDebug("GLLOVs.CmdSubmit_Click: validation failed - no LOV selected");
                await AppOverlayControl.ShowWarningAsync("Please select a LOV to proceed.");
                return;
            }
            if (excelRefEdit.Text == null)
            {
                LogUtility.LogDebug("GLLOVs.CmdSubmit_Click: validation failed - no range selected");
                await AppOverlayControl.ShowWarningAsync("Please select a range to copy lov.");
                return;
            }
            if (SelLov.ItemsCount == 0)
            {
                LogUtility.LogDebug($"GLLOVs.CmdSubmit_Click: validation failed - selected LOV '{SelLov.Name}' has no items");
                await AppOverlayControl.ShowWarningAsync("The selected LOV has no items to copy.");
                return;
            }

            try
            {

                await ShowBusyOverlayAsync(ctsHelper, "Please wait while we set the excel dependencies...");

                LogUtility.LogDebug("GLLOVs.CmdSubmit_Click: calling CreateLOVSheetAsync");
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

                LogUtility.LogDebug($"GLLOVs.CmdSubmit_Click: cleanedName={cleanedName}, nameExists={nameExists}");
                if (!string.IsNullOrEmpty(cleanedName) && nameExists)
                {
                    Excel.Range rng = AppState.Instance.ExcelApp.Range[excelRefEdit.Text];
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
                    LogUtility.LogDebug("GLLOVs.CmdSubmit_Click: LOV copied successfully");
                }
                else
                {
                    LogUtility.LogDebug("GLLOVs.CmdSubmit_Click: named range not created or does not exist, skipping validation setup");
                }
            }
            catch (OperationCanceledException)
            {
                LogUtility.LogWarn("GLLOVs.CmdSubmit_Click: operation cancelled by user");
                await AppOverlayControl.HideBusyAsync();
                await AppOverlayControl.ShowWarningAsync("Operation cancelled by user.");
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex, "GLLOVs.CmdSubmit_Click");
            }
            finally
            {
                try
                {
                    if (!ctsHelper.IsCancellationRequested)
                        ctsHelper.Cancel();

                    ctsHelper.Dispose();
                }
                catch (Exception ex)
                {
                    // Swallow dispose exceptions (Excel COM weirdness) but still log for diagnostics.
                    LogUtility.LogWarn($"GLLOVs.CmdSubmit_Click: exception disposing CancellationHelper (ignored): {ex.Message}");
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
            LogUtility.LogDebug($"GLLOVs.CreateLOVSheetAsync invoked - CubeId={CubeId}, CubeName={CubeName}, LedgerId={LedgerIdlng}, LedgerName={LedgerName}");
            try

            {
                string sheetName = string.Empty;
                int maxLength = 24;

                sheetName = $"{(LedgerIdStr.Length > maxLength ? LedgerIdStr.Substring(0, maxLength) : LedgerIdStr)}_LOV";

                Excel.Worksheet WrkSheet;

                if (!await CommonFunctions.SheetExistsAsync(sheetName))
                {
                    WrkSheet = (Excel.Worksheet)AppState.Instance.ExcelApp.Worksheets.Add();
                    WrkSheet.Name = sheetName;
                }
                else
                {
                    WrkSheet = (Excel.Worksheet)AppState.Instance.ExcelApp.ActiveWorkbook.Worksheets[sheetName];
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
                var periods = ServiceLocator.PeriodDataService.GetPeriodsForLedger(LedgerName);
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
                LogUtility.LogDebug($"GLLOVs.CreateLOVSheetAsync: completed successfully for sheet={sheetName}, columns written={ItemsDict.Count}");

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
            LogUtility.LogDebug($"GLLOVs.WriteListToColumnAsync invoked - columnTitle={columnTitle}, colIndex={colIndex}");
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
            LogUtility.LogDebug($"GLLOVs.CleanUpNamedRangeAsync invoked - nmName={nmName}");
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
            LogUtility.LogDebug($"GLLOVs.DeleteNamedRangeAsync invoked - rngnm={rngnm}");
            try
            {
                if (AppState.Instance.ExcelApp.ActiveWorkbook.Names.Count > 0)
                {
                    foreach (Excel.Name NM in AppState.Instance.ExcelApp.ActiveWorkbook.Names)
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
            LogUtility.LogDebug($"GLLOVs.CreateNamedRangeAsync invoked - nm={nm}");
            try
            {
                string cleanedName = await CleanUpNamedRangeAsync(nm.Trim());

                if (NameRangeExists(cleanedName))
                {
                    await DeleteNamedRangeAsync(cleanedName);
                }

                AppState.Instance.ExcelApp.ActiveWorkbook.Names.Add(cleanedName, RefersToR1C1: formulaStr);
                LogUtility.LogDebug($"GLLOVs.CreateNamedRangeAsync: named range '{cleanedName}' created successfully");
            }
            catch (Exception ex)
            {
                await LogErrorAsync(ex, $"Exception in creating named range : {nm} and formula : {formulaStr}");
            }
        }

        // You'll also need this method if it doesn't exist
        private static bool NameRangeExists(string rangeName)
        {
            if (AppState.Instance.ExcelApp.ActiveWorkbook.Names.Count > 0)
            {
                foreach (Excel.Name name in AppState.Instance.ExcelApp.ActiveWorkbook.Names)
                {
                    if (name.Name.Equals(rangeName, StringComparison.OrdinalIgnoreCase))
                    {
                        return true;
                    }
                }
            }
            return false;
        }

        // Async version of LogError (you'll need to implement this based on your logging framework)
        private static async Task LogErrorAsync(Exception ex, string message = null)
        {
            await Task.Run(() =>
            {
                // Your logging implementation here
                // For example:
                if (message != null)
                {
                    LogUtility.LogError($"{message}: {ex.Message}");
                }
                else
                {
                    LogUtility.LogException(ex);
                }
            });
        }

        private void BtnClose_Click(object sender, RoutedEventArgs e)
        {
            LogUtility.LogDebug("GLLOVs.BtnClose_Click invoked");
            Close();
        }
    }
}

