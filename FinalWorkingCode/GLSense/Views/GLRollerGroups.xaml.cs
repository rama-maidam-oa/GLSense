using GLSense.Helpers;
using GLSense.Interfaces;
using GLSense.Models;
using GLSense.Utilities;
using GLSense.ViewModels;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Threading;

namespace GLSense.Views
{
    /// <summary>
    /// Interaction logic for GLRollerGroups.xaml
    /// </summary>
    public partial class GLRollerGroups : DpiAwareWindow, IWarningHost
    {
        private readonly SimpleSegmentViewModel vm;
        public GLRollerGroups()
        {
            InitializeComponent();

            LogUtility.LogDebug("GLRollerGroups constructor invoked");

            EnhancedDragDropHelper.EnableWindowDrag(this);

            vm = new SimpleSegmentViewModel(Dispatcher)
            {
                ExcelApp = AppState.Instance.ExcelApp.Application, // Pass the Excel application instance to the ViewModel
                ShowWarningAction = (msg) => Dispatcher.Invoke(() => AppOverlayControl.ShowWarning(msg)),
                ShowBusyAction = async (txt, cancel) =>
                        await Dispatcher.InvokeAsync(async () =>
                            await AppOverlayControl.ShowBusyasynTask(txt, cancel)),
                HideBusyAsyncAction = async () => await Dispatcher.InvokeAsync(async () => await AppOverlayControl.HideBusyAsync())
            };
            DataContext = vm;

            // Subscribe to scroll messages
            vm.ScrollToTopRequested += OnScrollToTopRequested;
            vm.PropertyChanged += Vm_PropertyChanged;
        }

        // Keeps the Overwrite/Insert radio buttons (bound to IsMultipleRowsEnabled for their
        // IsEnabled state) from getting stuck on a stale "Insert" selection once they're
        // disabled - e.g. the user picks Insert while a single segment is selected, then
        // selects items from a second segment, which disables the whole row-mode section.
        // Forcing rbOverwrite back on here (rather than just letting both radios grey out)
        // guarantees the write path this window actually takes matches what's visibly shown -
        // an unchecked, disabled "Insert" would otherwise still read as IsChecked=true.
        private void Vm_PropertyChanged(object sender, PropertyChangedEventArgs e)
        {
            if (e.PropertyName == nameof(SimpleSegmentViewModel.IsMultipleRowsEnabled) && !vm.IsMultipleRowsEnabled)
            {
                rbOverwrite.IsChecked = true;
                rbByRows.IsChecked = true;
            }
        }

        private void OnScrollToTopRequested(ScrollToTopMessage message)
        {

            if (message.ScrollLeft && dgLeft.Items.Count > 0)
                dgLeft.ScrollIntoView(dgLeft.Items[0]);

            if (message.ScrollRight && dgRight.Items.Count > 0)
                dgRight.ScrollIntoView(dgRight.Items[0]);
        }
        public void CellSelectionWarning(string message)
        {
            AppOverlayControl.ShowWarning(message);
        }
        private async void Window_Loaded(object sender, RoutedEventArgs e)
        {
            LogUtility.LogDebug("GLRollerGroups.Window_Loaded invoked");

            Microsoft.Office.Interop.Excel.Range rng = AppState.Instance.ExcelApp.ActiveCell;
            string sheetName = ((Microsoft.Office.Interop.Excel.Worksheet)rng.Parent).Name;
            string cellAddress = rng.Address[true, true, Microsoft.Office.Interop.Excel.XlReferenceStyle.xlA1, false];
            string addr = $"'{sheetName}'!{cellAddress}";

            excelRefEdit.Text = addr;
            LogUtility.LogDebug($"GLRollerGroups.Window_Loaded: active cell address={addr}");

            if (AppState.Instance.SelectedCube != null && AppState.Instance.SelectedLedger != null)
            {
                LogUtility.LogDebug($"GLRollerGroups.Window_Loaded: loading segments for CubeId={AppState.Instance.SelectedLedger.CubeId}, LedgerId={AppState.Instance.SelectedLedger.LedgerId}");
                await vm.LoadSegmentsAsync(AppState.Instance.SelectedLedger.CubeId, AppState.Instance.SelectedLedger.LedgerId);
            }
            else
            {
                LogUtility.LogDebug("GLRollerGroups.Window_Loaded: SelectedCube or SelectedLedger is null - skipping LoadSegmentsAsync");
            }
            await Dispatcher.InvokeAsync(() =>
            {
                cmbSegments.Text = vm.SelectedSegment.SegmentName;
                cmbSearchType.Text = vm.SelectedSearchType.DisplayName;
            });
        }
        private void BtnClose_Click(object sender, RoutedEventArgs e)
        {
            LogUtility.LogDebug("GLRollerGroups.BtnClose_Click invoked");
            vm.ScrollToTopRequested -= OnScrollToTopRequested;
            vm.PropertyChanged -= Vm_PropertyChanged;
            Close();
        }
        private void dgLeft_SelectionChanged(object sender, System.Windows.Controls.SelectionChangedEventArgs e)
        {
            var dataGrid = sender as DataGrid;
            if (dataGrid == null) return;

            // Remove any TitleRow objects from selection
            var titleRows = dataGrid.SelectedItems
                .OfType<TitleRow>()
                .ToList();

            foreach (var titleRow in titleRows)
            {
                dataGrid.SelectedItems.Remove(titleRow);
            }
        }
        private void BtnAdd_Click(object sender, RoutedEventArgs e)
        {
            var viewModel = (SimpleSegmentViewModel)DataContext;
            var selectedItems = dgLeft.SelectedItems.Cast<SegmentDataRow>().ToList();
            LogUtility.LogDebug($"GLRollerGroups.BtnAdd_Click invoked - selectedCount={selectedItems.Count}");
            viewModel.AddSelection(selectedItems);

            ResetGridSelections();
        }
        private void DgLeft_MouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
            var viewModel = (SimpleSegmentViewModel)DataContext;
            var selected = dgLeft.SelectedItems.Cast<SegmentDataRow>().ToList();
            LogUtility.LogDebug($"GLRollerGroups.DgLeft_MouseDoubleClick invoked - selectedCount={selected.Count}");
            viewModel.AddSelection(selected);

            ResetGridSelections();
        }
        private void ResetGridSelections()
        {
            // Reset left grid selection
            dgLeft.UnselectAll();
            dgLeft.SelectedItems.Clear();

            // Optional: Scroll to top to show the titles clearly
            if (dgLeft.Items.Count > 0)
            {
                dgLeft.ScrollIntoView(dgLeft.Items[0]);
            }
        }
        private void BtnRemove_Click(object sender, RoutedEventArgs e)
        {
            var viewModel = (SimpleSegmentViewModel)DataContext;
            var selectedItems = dgRight.SelectedItems.Cast<SegmentSelectionModel>().ToList();
            LogUtility.LogDebug($"GLRollerGroups.BtnRemove_Click invoked - selectedCount={selectedItems.Count}");
            viewModel.RemoveSelection(selectedItems);
        }
        private void DgRight_MouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
            var viewModel = (SimpleSegmentViewModel)DataContext;
            var selected = dgRight.SelectedItems.Cast<SegmentSelectionModel>().ToList();
            LogUtility.LogDebug($"GLRollerGroups.DgRight_MouseDoubleClick invoked - selectedCount={selected.Count}");
            viewModel.RemoveSelection(selected);
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
        private async void BtnOK_Click(object sender, RoutedEventArgs e)
        {
            LogUtility.LogDebug("GLRollerGroups.BtnOK_Click invoked");

            if (!ValidateInputs())
                return;

            string startAddress = excelRefEdit.Text.Trim();
            var rng = AppState.Instance.ExcelApp.Range[startAddress];
            var grouped = GroupSelectedItems();

            try
            {
                using var cancellationHelper = new CancellationHelper();

                await ShowBusyOverlayAsync(cancellationHelper, "Writing data to Excel...");

                LogUtility.LogDebug($"GLRollerGroups.BtnOK_Click: writing to Excel - startAddress={startAddress}, multipleRows={vm.IsMultipleRowsChecked}, groupCount={grouped.Count}, insertMode={rbInsert.IsChecked == true}");

                if (vm.IsMultipleRowsChecked)
                {
                    var flatCount = grouped.Values.Sum(v => v.Count);
                    // Orientation only applies to this multi-cell path - the single-cell path
                    // below always writes one concatenated string into one cell, so there is
                    // nothing for "columns" to mean there (rbByColumns is disabled/reset to
                    // rows whenever chkMultipleRows itself is disabled - see Vm_PropertyChanged).
                    bool columnWise = rbByColumns.IsChecked == true;
                    PerformInsertIfNeeded(rng, flatCount, columnWise);
                    // Re-resolve the reference after a possible Insert: Excel's Range object
                    // for `rng` was bound to the cell(s) at startAddress BEFORE the insert, and
                    // Insert(xlShiftDown/xlShiftToRight) carries that same object along with the
                    // original content it shifts - so writing into the (stale) `rng` reference
                    // here would land on the shifted-down/shifted-right old content instead of
                    // the newly-opened blank cells. A fresh address lookup always resolves to
                    // whatever now occupies that address, which is exactly the blank space
                    // Insert just created. This is a no-op re-fetch (same cells) when Overwrite
                    // is selected.
                    rng = AppState.Instance.ExcelApp.Range[startAddress];
                    WriteMultipleRows(rng, grouped, columnWise);
                }
                else
                {
                    PerformInsertIfNeeded(rng, 1, columnWise: false);
                    rng = AppState.Instance.ExcelApp.Range[startAddress];
                    WriteSingleRow(rng, grouped);
                }

                if (cancellationHelper.IsCancellationRequested)
                {
                    LogUtility.LogWarn("GLRollerGroups.BtnOK_Click: Excel write operation was cancelled by user");
                    vm.ShowWarningAction?.Invoke("Operation was cancelled by user.");
                    return;
                }

                LogUtility.LogDebug("GLRollerGroups.BtnOK_Click: write to Excel completed successfully");
                DialogResult = true;
                Close();
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex, "GLRollerGroups.BtnOK_Click: Failed to write to Excel");
                vm.ShowWarningAction?.Invoke("Failed to write to Excel: " + ex.Message);
            }
            finally
            {
                // Hide busy overlay
                await AppOverlayControl.HideBusyAsync();
            }
        }
        private bool ValidateInputs()
        {
            if (vm.SelectedItemsRight == null || vm.SelectedItemsRight.Count == 0)
            {
                LogUtility.LogDebug("GLRollerGroups.ValidateInputs: validation failed - no items added");
                vm.ShowWarningAction?.Invoke("No items added. Please add values before clicking Insert.");
                return false;
            }

            if (string.IsNullOrWhiteSpace(excelRefEdit.Text))
            {
                LogUtility.LogDebug("GLRollerGroups.ValidateInputs: validation failed - no cell reference selected");
                vm.ShowWarningAction?.Invoke("Please select a cell reference.");
                return false;
            }

            if (vm.IsMultipleRowsChecked && vm.SelectedItemsRight.Select(x => x.Segment).Distinct().Count() > 1)
            {
                LogUtility.LogDebug("GLRollerGroups.ValidateInputs: validation failed - multiple rows mode requires a single segment");
                vm.ShowWarningAction?.Invoke("Multiple rows mode is only allowed if all values belong to the same segment.");
                return false;
            }

            return true;
        }
        private Dictionary<string, List<string>> GroupSelectedItems()
        {
            return vm.SelectedItemsRight
                .GroupBy(x => x.Segment)
                .ToDictionary(
                    g => g.Key,
                    g => g.Select(x => string.IsNullOrEmpty(x.Value2) ? x.Value1 : x.Value1 + "|" + x.Value2)
                         .ToList()
                );
        }

        // When Insert is selected, opens space for the values about to be written by
        // selecting a range the size of rowCount at the reference cell and shifting
        // existing content down - mirrors GLSegmentDiscovery.PerformInsertIfNeeded's
        // "Down" direction case. When Overwrite is selected (the default, matching this
        // window's original behavior), this is a no-op and the write methods below write
        // directly into the existing cells, same as before this feature existed.
        private void PerformInsertIfNeeded(Microsoft.Office.Interop.Excel.Range rng, int count, bool columnWise)
        {
            if (rbInsert.IsChecked != true || count <= 0)
                return;

            LogUtility.LogDebug($"GLRollerGroups.PerformInsertIfNeeded: inserting {count} {(columnWise ? "column(s)" : "row(s)")} at {rng.Address}");
            if (columnWise)
            {
                rng.Resize[1, count].Select();
                var selectedRange = AppState.Instance.ExcelApp.Selection as Microsoft.Office.Interop.Excel.Range;
                selectedRange.Insert(Microsoft.Office.Interop.Excel.XlInsertShiftDirection.xlShiftToRight);
            }
            else
            {
                rng.Resize[count, 1].Select();
                var selectedRange = AppState.Instance.ExcelApp.Selection as Microsoft.Office.Interop.Excel.Range;
                selectedRange.Insert(Microsoft.Office.Interop.Excel.XlInsertShiftDirection.xlShiftDown);
            }
        }

        private static void WriteMultipleRows(Microsoft.Office.Interop.Excel.Range rng, Dictionary<string, List<string>> grouped, bool columnWise)
        {
            var flat = grouped.Values.SelectMany(x => x).ToList();
            if (columnWise)
                WriteValuesHorizontallyToExcel(rng, flat);
            else
                WriteValuesVerticallyToExcel(rng, flat);
        }
        private void WriteSingleRow(Microsoft.Office.Interop.Excel.Range rng, Dictionary<string, List<string>> grouped)
        {
            var allSegments = vm.Segments.Select(s => s.SegmentName).ToList();
            var parts = allSegments.Select(segName =>
                grouped.TryGetValue(segName, out var values) && values.Count > 0
                    ? string.Join(",", values)
                    : ""
            ).ToList();

            string resultString = string.Join(";", parts);
            WriteStringToExcel(rng, resultString);
        }

        // Reuse your existing Excel helper methods or include these (if you don't have them already).
        private static void WriteValuesVerticallyToExcel(Microsoft.Office.Interop.Excel.Range rng, System.Collections.Generic.List<string> values)
        {
            for (int i = 0; i < values.Count; i++)
            {
                rng.Offset[i].NumberFormat = "@";
                rng.Offset[i].Value = values[i];
            }
        }

        // Write as Multiple Columns: same shape as WriteValuesVerticallyToExcel, offsetting
        // across columns (row offset 0, increasing column offset) instead of down rows.
        private static void WriteValuesHorizontallyToExcel(Microsoft.Office.Interop.Excel.Range rng, System.Collections.Generic.List<string> values)
        {
            for (int i = 0; i < values.Count; i++)
            {
                rng.Offset[0, i].NumberFormat = "@";
                rng.Offset[0, i].Value = values[i];
            }
        }
        private static void WriteStringToExcel(Microsoft.Office.Interop.Excel.Range rng, string value)
        {
            rng.Value = value;
        }
    }
}

