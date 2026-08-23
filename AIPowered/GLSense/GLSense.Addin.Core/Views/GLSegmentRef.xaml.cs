// GLSegmentRef.xaml.cs in GLSense.Addin.Core
// Port of GLSense\Views\GLSegmentRef.xaml.cs (FinalWorkingCode) - "Account Assignment
// Configurator" popup opened from GLAccountsRef.xaml.cs's Edit button (used by
// GLBalanceConfigurator's "Account Assignment(s)" row). Previously deferred (see the old
// header comments in GLAccountsRef.xaml.cs and SegmentSelectorViewModel.cs) because this
// window's only dependency, SegmentSelectorViewModel, was ported ahead of time in full
// (including its "Ref"-mode branches) specifically so this window could be added later
// without touching the ViewModel again - this file is that follow-up.
//
// Follows the exact same pattern already established by GLSegmentValues.xaml.cs (shares
// SegmentSelectorViewModel, just constructed with windowName="Ref" instead of "val"):
//   - Base class DpiAwareWindow -> DpiAwareWindow. EnhancedDragDropHelper.EnableWindowDrag(this)
//     -> TitleBar_MouseLeftButtonDown (same handler every other window here uses).
//   - AppState.Instance.ExcelApp.Application -> ServiceLocator.ExcelApp.
//   - LogUtility.* -> ServiceLocator.Logger?.*.
//   - dlg.ShowDialogWithOwner((IntPtr)AppState.Instance.ExcelApp.Hwnd) (caller-side, in
//     GLAccountsRef/GLBalanceConfigurator) -> plain ShowDialog(); DpiAwareWindow already sets
//     the Excel owner automatically via ServiceLocator.ExcelHandle/ModalToExcel.
//   - The DataGrid columns here (Value/Description/Is-Summary on the left,
//     Value1/Value2/Segment on the right) are Width="Auto" with DataGridColumnFillHelper
//     filling the Description/Segment columns, exactly like GLSegmentValues - the original
//     used Width="*" columns, which is the same star-width-under-SizeToContent bug fixed
//     project-wide earlier in this migration (see DataGridColumnFillHelper's own header).
//   - The old RefreshWindowLayout() call (DpiAwareWindow-only, never existed on DpiAwareWindow)
//     is dropped - DpiAwareWindow's own Loaded handler already re-applies DPI/work-area layout.
// No other logic changes vs. the original.
using GLSense.Addin.Core.Helpers;
using GLSense.Addin.Core.Infrastructure;
using GLSense.Addin.Core.Interfaces;
using GLSense.Addin.Core.Models;
using GLSense.Addin.Core.Utilities;
using GLSense.Addin.Core.ViewModels;
using System;
using System.Linq;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Threading;

namespace GLSense.Addin.Core.Views
{
    /// <summary>
    /// Interaction logic for GLSegmentRef.xaml
    /// </summary>
    public partial class GLSegmentRef : DpiAwareWindow, IWarningHost
    {
        public string GLSegments_SelectedValue { get; set; }
        private readonly SegmentSelectorViewModel vm;

        // Guards DataLoadedAction below so the window only resettles its SizeToContent size
        // ONCE, the first time real data lands after the initial async load (its original
        // purpose - see CLAUDE.md section 1.4b/1.4c). DataLoadedAction is invoked from
        // SegmentSelectorViewModel.UpdatePagingAndGrid(), which is ALSO the choke point
        // SelectedItemsRight's setter funnels through - so without this guard, every add/
        // remove of a value in the right-hand grid during normal interactive use re-triggers
        // a full window resettle. See GLSegmentManager.xaml.cs's identical fix/comment for
        // the full writeup (same shared ViewModel, same bug).
        private bool _hasResettledAfterInitialLoad;

        public GLSegmentRef(string selectedSegValues)
        {
            InitializeComponent();
            ServiceLocator.Logger?.LogDebug($"GLSegmentRef constructor invoked - selectedSegValues={selectedSegValues}");
            EnableExcelCentering = false;

            // "Description"/"Segment" fill any left-over width in their respective grids
            // instead of leaving a blank gap now that every column is Width="Auto" (see
            // DataGridColumnFillHelper for why the star-width columns were removed).
            DataGridColumnFillHelper.EnableFillColumn(dgLeft, dgLeft.Columns[1]);
            DataGridColumnFillHelper.EnableFillColumn(dgRight, dgRight.Columns[2]);

            vm = new SegmentSelectorViewModel(Dispatcher, "Ref", selectedSegValues)
            {
                ExcelApp = ServiceLocator.ExcelApp,
                ShowWarningAction = (msg) => Dispatcher.Invoke(() => AppOverlayControl.ShowWarning(msg)),
                ShowInfoAction = (msg) => Dispatcher.Invoke(() => AppOverlayControl.ShowInfo(msg)),
                ShowBusyAction = async (txt, cancel) =>
                        await Dispatcher.InvokeAsync(async () =>
                            await AppOverlayControl.ShowBusyasynTask(txt, cancel)),
                HideBusyAsyncAction = async () => await Dispatcher.InvokeAsync(async () => await AppOverlayControl.HideBusyAsync()),
                // See CLAUDE.md section 1.4b - dgLeft/dgRight populate fire-and-forget,
                // detached from Window_Loaded's own await chain. Only resettle the FIRST
                // time this fires (initial load) - see _hasResettledAfterInitialLoad's
                // comment above for why later invocations (every right-grid add/remove)
                // must not re-trigger this.
                DataLoadedAction = () =>
                {
                    if (_hasResettledAfterInitialLoad) return;
                    _hasResettledAfterInitialLoad = true;

                    ForceSizeToContentResettle();
                    PumpDispatcherFrame();
                }
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

        private void OnScrollToTopRequested(ScrollToTopMessage message)
        {
            if (message.ScrollLeft && dgLeft.Items.Count > 0)
                dgLeft.ScrollIntoView(dgLeft.Items[0]);

            if (message.ScrollRight && dgRight.Items.Count > 0)
                dgRight.ScrollIntoView(dgRight.Items[0]);
        }

        protected override void OnClosed(EventArgs e)
        {
            ServiceLocator.Logger?.LogDebug("GLSegmentRef.OnClosed invoked");
            vm.ScrollToTopRequested -= OnScrollToTopRequested;
            base.OnClosed(e);
        }

        private async void Window_Loaded(object sender, RoutedEventArgs e)
        {
            ServiceLocator.Logger?.LogDebug("GLSegmentRef.Window_Loaded invoked");
            try
            {
                if (AppState.Instance.SelectedCube != null && AppState.Instance.SelectedLedger != null)
                {
                    ServiceLocator.Logger?.LogDebug($"GLSegmentRef.Window_Loaded: loading segments - cubeId={AppState.Instance.SelectedLedger.CubeId}, ledgerId={AppState.Instance.SelectedLedger.LedgerId}");
                    await vm.LoadSegmentsAsync(AppState.Instance.SelectedLedger.CubeId, AppState.Instance.SelectedLedger.LedgerId);
                    await Dispatcher.InvokeAsync(() => FocusFirstSegmentTextBox(), DispatcherPriority.Background);
                }
                else
                {
                    ServiceLocator.Logger?.LogDebug("GLSegmentRef.Window_Loaded: validation failed - no cube/ledger selected, skipping segment load");
                }
            }
            catch (Exception ex)
            {
                ServiceLocator.Logger?.LogException(ex, "GLSegmentRef.Window_Loaded");
            }
        }

        private async Task CmbHierarchy_SelectionCommitted(object obj)
        {
            ServiceLocator.Logger?.LogDebug("GLSegmentRef.CmbHierarchy_SelectionCommitted invoked");
            try
            {
                if (obj is SegmentValueModel selectedHierarchy)
                {
                    ServiceLocator.Logger?.LogDebug($"GLSegmentRef.CmbHierarchy_SelectionCommitted: hierarchy selected - {selectedHierarchy}");
                    await vm.LoadSegmentValuesAsync(null, selectedHierarchy, true);
                }
                else
                {
                    ServiceLocator.Logger?.LogDebug("GLSegmentRef.CmbHierarchy_SelectionCommitted: no valid hierarchy selection, falling back to regular segment values");
                    await vm.LoadSegmentValuesAsync(); // fallback to regular segment values
                }
            }
            catch (Exception ex)
            {
                ServiceLocator.Logger?.LogException(ex, "GLSegmentRef.CmbHierarchy_SelectionCommitted");
            }
        }

        private async void CmbHierarchy_LostFocus(object sender, RoutedEventArgs e)
        {
            try
            {
                var typedText = cmbHierarchy.SelectedItem;
                if (typedText is not SegmentValueModel)
                {
                    ServiceLocator.Logger?.LogDebug("GLSegmentRef.CmbHierarchy_LostFocus: typed text is not a valid hierarchy selection, reloading segment values");
                    await vm.LoadSegmentValuesAsync();
                }
            }
            catch (Exception ex)
            {
                ServiceLocator.Logger?.LogException(ex, "GLSegmentRef.CmbHierarchy_LostFocus");
            }
        }

        private void FocusFirstSegmentTextBox()
        {
            if (SegmentsItemsControl.Items.Count == 0) return;

            var container = SegmentsItemsControl.ItemContainerGenerator.ContainerFromIndex(0);
            if (container == null) return; // Container not yet generated

            var textBox = FindVisualChild<TextBox>(container);
            if (textBox != null)
            {
                textBox.Focus();
                textBox.SelectAll(); // optional: select all text to ease editing
            }
        }

        public void CellSelectionWarning(string message)
        {
            ServiceLocator.Logger?.LogDebug($"GLSegmentRef.CellSelectionWarning invoked - message={message}");
            AppOverlayControl.ShowWarning(message);
        }

        private async void SegmentControl_GotFocus(object sender, RoutedEventArgs e)
        {
            if (!TryGetElements(sender, out var fe, out var seg, out var vm1))
            {
                ServiceLocator.Logger?.LogDebug("GLSegmentRef.SegmentControl_GotFocus: validation failed - unable to resolve elements from sender");
                return;
            }

            vm1.SelectedSegment = seg;

            // Re-enable or disable left grid depending on this segment's state
            await UpdateGridsAsync(seg, vm);

            // Get the ItemsControl (parent to all rows); adjust as needed
            var itemsControl = SegmentsItemsControl;

            ResetAllRowStyles(itemsControl);
            HighlightFocusedRow(fe);
        }

        private void ResetAllRowStyles(ItemsControl itemsControl)
        {
            for (int i = 0; i < itemsControl.Items.Count; i++)
            {
                var container = itemsControl.ItemContainerGenerator.ContainerFromIndex(i) as ContentPresenter;
                var grid = FindVisualChild<Grid>(container);
                if (grid != null)
                {
                    foreach (var child in grid.Children)
                    {
                        if (child is Label label)
                        {
                            label.FontWeight = FontWeights.Normal;
                            label.Foreground = Brushes.Black;
                        }
                    }
                }
            }
        }

        private static void HighlightFocusedRow(FrameworkElement fe)
        {
            if (VisualTreeHelper.GetParent(fe) is Grid parentGrid)
            {
                foreach (var child in parentGrid.Children)
                {
                    if (child is Label label)
                    {
                        label.FontWeight = FontWeights.Bold;
                        label.Foreground = Brushes.Red;
                    }
                }
            }
        }

        private async Task UpdateGridsAsync(SegmentModel seg, SegmentSelectorViewModel vm)
        {
            if (string.IsNullOrWhiteSpace(seg.Reference))
            {
                vm.IsLeftGridEnabled = true;
                await vm.LoadSegmentValuesAsync(seg);
                ScrollDataGridToTop(dgLeft);
            }
            else
            {
                vm.IsLeftGridEnabled = false;
                vm.SelectedItemsRight.Clear();
            }
        }

        private bool TryGetElements(object sender, out FrameworkElement fe, out SegmentModel seg, out SegmentSelectorViewModel vm)
        {
            fe = sender as FrameworkElement;
            if (fe == null)
            {
                seg = null;
                vm = null;
                return false;
            }

            seg = fe.DataContext as SegmentModel;
            if (seg == null)
            {
                vm = null;
                return false;
            }

            vm = this.DataContext as SegmentSelectorViewModel;
            return vm != null;
        }

        private void ScrollDataGridToTop(DataGrid dg)
        {
            if (dg == null) return;

            var sv = FindVisualChild<ScrollViewer>(dg);
            sv?.ScrollToTop();
        }

        private T FindVisualChild<T>(DependencyObject parent) where T : DependencyObject
        {
            if (parent == null) return null;

            for (int i = 0; i < VisualTreeHelper.GetChildrenCount(parent); i++)
            {
                var child = VisualTreeHelper.GetChild(parent, i);
                if (child is T result)
                {
                    return result;
                }
                else
                {
                    var foundChild = FindVisualChild<T>(child);
                    if (foundChild != null) return foundChild;
                }
            }
            return null;
        }

        private void BtnAdd_Click(object sender, RoutedEventArgs e)
        {
            ServiceLocator.Logger?.LogDebug("GLSegmentRef.BtnAdd_Click invoked");
            var selected = dgLeft.SelectedItems.Cast<SegmentValueModel>().ToList();
            ServiceLocator.Logger?.LogDebug($"GLSegmentRef.BtnAdd_Click: adding {selected.Count} selected item(s)");
            vm.AddSelection(selected);
        }

        private void BtnFirst_Click(object sender, RoutedEventArgs e)
        {
            ServiceLocator.Logger?.LogDebug("GLSegmentRef.BtnFirst_Click invoked");
            vm.GoFirstPage();
        }

        private void BtnPrev_Click(object sender, RoutedEventArgs e)
        {
            ServiceLocator.Logger?.LogDebug("GLSegmentRef.BtnPrev_Click invoked");
            vm.GoPreviousPage();
        }

        private void BtnNext_Click(object sender, RoutedEventArgs e)
        {
            ServiceLocator.Logger?.LogDebug("GLSegmentRef.BtnNext_Click invoked");
            vm.GoNextPage();
        }

        private void BtnLast_Click(object sender, RoutedEventArgs e)
        {
            ServiceLocator.Logger?.LogDebug("GLSegmentRef.BtnLast_Click invoked");
            vm.GoLastPage();
        }

        private void BtnRemove_Click(object sender, RoutedEventArgs e)
        {
            ServiceLocator.Logger?.LogDebug("GLSegmentRef.BtnRemove_Click invoked");
            var selected = dgRight.SelectedItems.Cast<SegmentSelectionModel>().ToList();
            ServiceLocator.Logger?.LogDebug($"GLSegmentRef.BtnRemove_Click: removing {selected.Count} selected item(s)");
            vm.RemoveSelection(selected);
        }

        private void BtnBetween_Click(object sender, RoutedEventArgs e)
        {
            ServiceLocator.Logger?.LogDebug("GLSegmentRef.BtnBetween_Click invoked");
            vm.AddBetweenSelection(dgLeft.SelectedItems.Cast<SegmentValueModel>().ToList(), false);
        }

        private void BtnNotBetween_Click(object sender, RoutedEventArgs e)
        {
            ServiceLocator.Logger?.LogDebug("GLSegmentRef.BtnNotBetween_Click invoked");
            vm.AddBetweenSelection(dgLeft.SelectedItems.Cast<SegmentValueModel>().ToList(), true);
        }

        private void BtnExclude_Click(object sender, RoutedEventArgs e)
        {
            ServiceLocator.Logger?.LogDebug("GLSegmentRef.BtnExclude_Click invoked");
            vm.AddExcludeSelection(dgLeft.SelectedItems.Cast<SegmentValueModel>().ToList());
        }

        private void DgRight_MouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
            ServiceLocator.Logger?.LogDebug("GLSegmentRef.DgRight_MouseDoubleClick invoked");
            var selected = dgRight.SelectedItems.Cast<SegmentSelectionModel>().ToList();
            vm.RemoveSelection(selected);
        }

        private void DgLeft_MouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
            ServiceLocator.Logger?.LogDebug("GLSegmentRef.DgLeft_MouseDoubleClick invoked");
            var selected = dgLeft.SelectedItems.Cast<SegmentValueModel>().ToList();
            vm.AddSelection(selected);
        }

        private void BtnClose_Click(object sender, RoutedEventArgs e)
        {
            ServiceLocator.Logger?.LogDebug("GLSegmentRef.BtnClose_Click invoked");
            Close();
        }

        private void BtnClearDefaults_Click(object sender, RoutedEventArgs e)
        {
            ServiceLocator.Logger?.LogDebug("GLSegmentRef.BtnClearDefaults_Click invoked");
            vm.ClearDefaults();
        }

        private void BtnOK_Click(object sender, RoutedEventArgs e)
        {
            ServiceLocator.Logger?.LogDebug("GLSegmentRef.BtnOK_Click invoked");
            var output = vm.GetAllSegmentValues();
            GLSegments_SelectedValue = output.Replace("\"", "");
            ServiceLocator.Logger?.LogDebug($"GLSegmentRef.BtnOK_Click: selected value set, closing dialog with result true - {GLSegments_SelectedValue}");
            this.DialogResult = true;
            this.Close();
        }
    }
}
