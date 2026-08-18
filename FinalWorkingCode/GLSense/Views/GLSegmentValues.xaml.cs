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
using System.Windows.Input;
using System.Windows.Threading;
using Excel = Microsoft.Office.Interop.Excel;

namespace GLSense.Views
{
    /// <summary>
    /// Interaction logic for GLSegmentValues.xaml
    /// </summary>
    public partial class GLSegmentValues : DpiAwareWindow, IWarningHost
    {
        private readonly SegmentSelectorViewModel vm;
        public GLSegmentValues()
        {
            LogUtility.LogDebug("GLSegmentValues.ctor invoked");
            InitializeComponent();
            EnhancedDragDropHelper.EnableWindowDrag(this);
            vm = new SegmentSelectorViewModel(Dispatcher, "val", string.Empty)
            {
                ExcelApp = AppState.Instance.ExcelApp.Application, // Pass the Excel application instance to the ViewModel
                ShowWarningAction = (msg) => Dispatcher.Invoke(() => AppOverlayControl.ShowWarning(msg)),
                ShowInfoAction = (msg) => Dispatcher.Invoke(() => AppOverlayControl.ShowInfo(msg)),
                ShowBusyAction = async (txt, cancel) =>
                        await Dispatcher.InvokeAsync(async () =>
                            await AppOverlayControl.ShowBusyasynTask(txt, cancel)),
                HideBusyAsyncAction = async () => await Dispatcher.InvokeAsync(async () => await AppOverlayControl.HideBusyAsync())
            };
            DataContext = vm;

            // Wire up custom ComboBox events
            cmbHierarchy.SelectionCommitted += async (obj) => await CmbHierarchy_SelectionCommitted(obj);
            cmbHierarchy.InvalidSelection += (invalidText) =>
            {
                Dispatcher.Invoke(() =>
                {
                    AppOverlayControl.ShowWarning($"Invalid hierarchy: '{invalidText}'. Please select a valid one.");
                });
            };

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
            if (e.PropertyName == nameof(SegmentSelectorViewModel.IsMultipleRowsEnabled) && !vm.IsMultipleRowsEnabled)
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
        protected override void OnClosed(EventArgs e)
        {
            LogUtility.LogDebug("GLSegmentValues.OnClosed invoked");
            vm.ScrollToTopRequested -= OnScrollToTopRequested;
            vm.PropertyChanged -= Vm_PropertyChanged;
            base.OnClosed(e);
        }
        private async void Window_Loaded(object sender, RoutedEventArgs e)
        {
            LogUtility.LogDebug("GLSegmentValues.Window_Loaded invoked");
            Excel.Range rng = AppState.Instance.ExcelApp.ActiveCell;
            string sheetName = ((Excel.Worksheet)rng.Parent).Name;
            string cellAddress = rng.Address[true, true, Excel.XlReferenceStyle.xlA1, false];
            string addr = $"'{sheetName}'!{cellAddress}";
            LogUtility.LogDebug($"GLSegmentValues.Window_Loaded: active cell address={addr}");

            if (AppState.Instance.SelectedCube != null && AppState.Instance.SelectedLedger != null)
            {
                LogUtility.LogDebug($"GLSegmentValues.Window_Loaded: loading segments - cubeId={AppState.Instance.SelectedLedger.CubeId}, ledgerId={AppState.Instance.SelectedLedger.LedgerId}");
                await vm.LoadSegmentsAsync(AppState.Instance.SelectedLedger.CubeId, AppState.Instance.SelectedLedger.LedgerId);
            }
            else
            {
                LogUtility.LogDebug("GLSegmentValues.Window_Loaded: validation failed - no cube/ledger selected, skipping segment load");
            }

            await Dispatcher.InvokeAsync(() =>
            {
                excelRefEdit.Text = addr;
                cmbSegments.Text = vm.SelectedSegment.SegmentName;
                cmbSearchType.Text = vm.SelectedSearchType.DisplayName;
            });

            await Dispatcher.InvokeAsync(() => RefreshWindowLayout(), DispatcherPriority.Render);
        }

        private async Task CmbHierarchy_SelectionCommitted(object obj)
        {
            LogUtility.LogDebug("GLSegmentValues.CmbHierarchy_SelectionCommitted invoked");
            if (obj is SegmentValueModel selectedHierarchy)
            {
                LogUtility.LogDebug($"GLSegmentValues.CmbHierarchy_SelectionCommitted: hierarchy selected - {selectedHierarchy}");
                await vm.LoadSegmentValuesAsync(null, selectedHierarchy, true);
            }
            else
            {
                LogUtility.LogDebug("GLSegmentValues.CmbHierarchy_SelectionCommitted: no valid hierarchy selection, falling back to regular segment values");
                await vm.LoadSegmentValuesAsync(); // fallback to regular segment values
            }
        }
        private async void CmbHierarchy_LostFocus(object sender, RoutedEventArgs e)
        {
            var typedText = cmbHierarchy.SelectedItem;
            if (typedText is not SegmentValueModel)
            {
                LogUtility.LogDebug("GLSegmentValues.CmbHierarchy_LostFocus: typed text is not a valid hierarchy selection, reloading segment values");
                await vm.LoadSegmentValuesAsync();
            }
        }

        public void CellSelectionWarning(string message)
        {
            LogUtility.LogDebug($"GLSegmentValues.CellSelectionWarning invoked - message={message}");
            AppOverlayControl.ShowWarning(message);
        }
        private void BtnAdd_Click(object sender, RoutedEventArgs e)
        {
            LogUtility.LogDebug("GLSegmentValues.BtnAdd_Click invoked");
            var viewModel = (SegmentSelectorViewModel)DataContext;
            var selected = dgLeft.SelectedItems.Cast<SegmentValueModel>().ToList();
            LogUtility.LogDebug($"GLSegmentValues.BtnAdd_Click: adding {selected.Count} selected item(s)");
            viewModel.AddSelection(selected);
        }
        private void BtnFirst_Click(object sender, RoutedEventArgs e)
        {
            LogUtility.LogDebug("GLSegmentValues.BtnFirst_Click invoked");
            vm.GoFirstPage();
        }

        private void BtnPrev_Click(object sender, RoutedEventArgs e)
        {
            LogUtility.LogDebug("GLSegmentValues.BtnPrev_Click invoked");
            vm.GoPreviousPage();
        }

        private void BtnNext_Click(object sender, RoutedEventArgs e)
        {
            LogUtility.LogDebug("GLSegmentValues.BtnNext_Click invoked");
            vm.GoNextPage();
        }

        private void BtnLast_Click(object sender, RoutedEventArgs e)
        {
            LogUtility.LogDebug("GLSegmentValues.BtnLast_Click invoked");
            vm.GoLastPage();
        }

        private void BtnRemove_Click(object sender, RoutedEventArgs e)
        {
            LogUtility.LogDebug("GLSegmentValues.BtnRemove_Click invoked");
            var viewModel = (SegmentSelectorViewModel)DataContext;
            var selected = dgRight.SelectedItems.Cast<SegmentSelectionModel>().ToList();
            LogUtility.LogDebug($"GLSegmentValues.BtnRemove_Click: removing {selected.Count} selected item(s)");
            viewModel.RemoveSelection(selected);
        }

        private void BtnBetween_Click(object sender, RoutedEventArgs e)
        {
            LogUtility.LogDebug("GLSegmentValues.BtnBetween_Click invoked");
            var viewModel = (SegmentSelectorViewModel)DataContext;
            viewModel.AddBetweenSelection(dgLeft.SelectedItems.Cast<SegmentValueModel>().ToList(), false);
        }

        private void BtnNotBetween_Click(object sender, RoutedEventArgs e)
        {
            LogUtility.LogDebug("GLSegmentValues.BtnNotBetween_Click invoked");
            var viewModel = (SegmentSelectorViewModel)DataContext;
            viewModel.AddBetweenSelection(dgLeft.SelectedItems.Cast<SegmentValueModel>().ToList(), true);
        }

        private void BtnExclude_Click(object sender, RoutedEventArgs e)
        {
            LogUtility.LogDebug("GLSegmentValues.BtnExclude_Click invoked");
            var viewModel = (SegmentSelectorViewModel)DataContext;
            viewModel.AddExcludeSelection(dgLeft.SelectedItems.Cast<SegmentValueModel>().ToList());
        }

        private void DgRight_MouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
            LogUtility.LogDebug("GLSegmentValues.DgRight_MouseDoubleClick invoked");
            var viewModel = (SegmentSelectorViewModel)DataContext;
            var selected = dgRight.SelectedItems.Cast<SegmentSelectionModel>().ToList();
            viewModel.RemoveSelection(selected);
        }

        private void DgLeft_MouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
            LogUtility.LogDebug("GLSegmentValues.DgLeft_MouseDoubleClick invoked");
            var viewModel = (SegmentSelectorViewModel)DataContext;
            var selected = dgLeft.SelectedItems.Cast<SegmentValueModel>().ToList();
            viewModel.AddSelection(selected);
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
            LogUtility.LogDebug("GLSegmentValues.BtnOK_Click invoked");
            if (!ValidateInputs())
            {
                LogUtility.LogDebug("GLSegmentValues.BtnOK_Click: validation failed, aborting");
                return;
            }

            string startAddress = excelRefEdit.Text?.Trim() ?? string.Empty;
            Excel.Range rng = AppState.Instance.ExcelApp.Range[startAddress];

            var allSegments = vm.Segments.Select(s => s.SegmentName).ToList();
            var grouped = GroupSelectedItems();
            LogUtility.LogDebug($"GLSegmentValues.BtnOK_Click: writing to Excel - startAddress={startAddress}, multipleRows={vm.IsMultipleRowsChecked}, groupedSegments={grouped.Count}, insertMode={rbInsert.IsChecked == true}");

            try
            {
                using var cancellationHelper = new CancellationHelper();

                await ShowBusyOverlayAsync(cancellationHelper, "Writing data to Excel...");

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
                    WriteSingleRow(rng, allSegments, grouped);
                }

                if (cancellationHelper.IsCancellationRequested)
                {
                    LogUtility.LogWarn("GLSegmentValues.BtnOK_Click: operation was cancelled by user");
                    vm.ShowWarningAction?.Invoke("Operation was cancelled by user.");
                    return;
                }

                LogUtility.LogDebug("GLSegmentValues.BtnOK_Click: write to Excel completed successfully, closing dialog");
                DialogResult = true;
                Close();
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex, "GLSegmentValues.BtnOK_Click: Failed to write to Excel");
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
                LogUtility.LogDebug("GLSegmentValues.ValidateInputs: validation failed - no items selected");
                vm.ShowWarningAction?.Invoke("No items added. Please add values before clicking Insert.");
                return false;
            }

            if (string.IsNullOrWhiteSpace(excelRefEdit.Text))
            {
                LogUtility.LogDebug("GLSegmentValues.ValidateInputs: validation failed - no cell reference selected");
                vm.ShowWarningAction?.Invoke("Please select a cell reference.");
                return false;
            }

            if (vm.IsMultipleRowsChecked)
            {
                var segments = vm.SelectedItemsRight.Select(x => x.Segment).Distinct().Count();
                if (segments > 1)
                {
                    LogUtility.LogDebug($"GLSegmentValues.ValidateInputs: validation failed - multiple rows mode requires a single segment, found {segments}");
                    vm.ShowWarningAction?.Invoke("Multiple rows mode is only allowed if all values belong to the same segment.");
                    return false;
                }
            }

            return true;
        }

        private Dictionary<string, List<string>> GroupSelectedItems()
        {
            return vm.SelectedItemsRight
                .GroupBy(x => x.Segment)
                .ToDictionary(
                    g => g.Key,
                    g => g.Select(x => string.IsNullOrEmpty(x.Value2) ? x.Value1 : x.Value1 + "|" + x.Value2).ToList()
                );
        }

        // When Insert is selected, opens space for the values about to be written by
        // selecting a range the size of rowCount at the reference cell and shifting
        // existing content down - mirrors GLSegmentDiscovery.PerformInsertIfNeeded's
        // "Down" direction case. When Overwrite is selected (the default, matching this
        // window's original behavior), this is a no-op and the write methods below write
        // directly into the existing cells, same as before this feature existed.
        private void PerformInsertIfNeeded(Excel.Range rng, int count, bool columnWise)
        {
            if (rbInsert.IsChecked != true || count <= 0)
                return;

            LogUtility.LogDebug($"GLSegmentValues.PerformInsertIfNeeded: inserting {count} {(columnWise ? "column(s)" : "row(s)")} at {rng.Address}");
            if (columnWise)
            {
                rng.Resize[1, count].Select();
                var selectedRange = AppState.Instance.ExcelApp.Selection as Excel.Range;
                selectedRange.Insert(Excel.XlInsertShiftDirection.xlShiftToRight);
            }
            else
            {
                rng.Resize[count, 1].Select();
                var selectedRange = AppState.Instance.ExcelApp.Selection as Excel.Range;
                selectedRange.Insert(Excel.XlInsertShiftDirection.xlShiftDown);
            }
        }

        private static void WriteMultipleRows(Excel.Range rng, Dictionary<string, List<string>> grouped, bool columnWise)
        {
            var flat = grouped.Values.SelectMany(x => x).ToList();
            if (columnWise)
                WriteValuesHorizontallyToExcel(rng, flat);
            else
                WriteValuesVerticallyToExcel(rng, flat);
        }

        private static void WriteSingleRow(Excel.Range rng, List<string> allSegments, Dictionary<string, List<string>> grouped)
        {
            var parts = allSegments.Select(segName =>
            {
                if (grouped.TryGetValue(segName, out var values) && values.Count > 0)
                {
                    return string.Join(",", values);
                }
                return string.Empty;
            }).ToList();

            string resultString = string.Join(";", parts);
            WriteStringToExcel(rng, resultString);
        }


        // Reuse your existing Excel helper methods or include these (if you don't have them already).
        private static void WriteValuesVerticallyToExcel(Excel.Range rng, System.Collections.Generic.List<string> values)
        {
            for (int i = 0; i < values.Count; i++)
            {
                rng.Offset[i].NumberFormat = "@";
                rng.Offset[i].Value = values[i];
            }
        }

        // Write as Multiple Columns: same shape as WriteValuesVerticallyToExcel, offsetting
        // across columns (row offset 0, increasing column offset) instead of down rows.
        private static void WriteValuesHorizontallyToExcel(Excel.Range rng, System.Collections.Generic.List<string> values)
        {
            for (int i = 0; i < values.Count; i++)
            {
                rng.Offset[0, i].NumberFormat = "@";
                rng.Offset[0, i].Value = values[i];
            }
        }
        private static void WriteStringToExcel(Excel.Range rng, string value)
        {
            rng.Value = value;
        }

        private void BtnClose_Click(object sender, RoutedEventArgs e)
        {
            LogUtility.LogDebug("GLSegmentValues.BtnClose_Click invoked");
            Close();
        }
    }
}
