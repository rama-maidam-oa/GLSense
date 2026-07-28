using GLSense.Helpers;
using GLSense.Interfaces;
using GLSense.Models;
using GLSense.Utilities;
using GLSense.ViewModels;
using System;
using System.Collections.Generic;
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
                ShowWarningAction = (msg) => Dispatcher.Invoke(() => AppOverlayControl.ShowWarning(msg))
            };
            DataContext = vm;

            // Subscribe to scroll messages
            vm.ScrollToTopRequested += OnScrollToTopRequested;
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

                LogUtility.LogDebug($"GLRollerGroups.BtnOK_Click: writing to Excel - startAddress={startAddress}, multipleRows={vm.IsMultipleRowsChecked}, groupCount={grouped.Count}");

                if (vm.IsMultipleRowsChecked)
                {
                    WriteMultipleRows(rng, grouped);
                }
                else
                {
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

        private static void WriteMultipleRows(Microsoft.Office.Interop.Excel.Range rng, Dictionary<string, List<string>> grouped)
        {
            var flat = grouped.Values.SelectMany(x => x).ToList();
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
        private static void WriteStringToExcel(Microsoft.Office.Interop.Excel.Range rng, string value)
        {
            rng.Value = value;
        }
    }
}

