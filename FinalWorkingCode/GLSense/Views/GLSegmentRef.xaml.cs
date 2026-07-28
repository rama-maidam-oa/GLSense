using GLSense.Helpers;
using GLSense.Interfaces;
using GLSense.Models;
using GLSense.Utilities;
using GLSense.ViewModels;
using System;
using System.Linq;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Threading;

namespace GLSense.Views
{
    /// <summary>
    /// Interaction logic for GLSegmentRef.xaml
    /// </summary>
    public partial class GLSegmentRef : DpiAwareWindow, IWarningHost
    {
        public string GLSegments_SelectedValue { get; set; }
        private readonly SegmentSelectorViewModel vm;
        public GLSegmentRef(string SelectedSegValues)
        {
            LogUtility.LogDebug($"GLSegmentRef.ctor invoked - SelectedSegValues={SelectedSegValues}");
            InitializeComponent();
            EnhancedDragDropHelper.EnableWindowDrag(this);

            vm = new SegmentSelectorViewModel(Dispatcher, "Ref", SelectedSegValues)
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
            LogUtility.LogDebug("GLSegmentRef.OnClosed invoked");
            vm.ScrollToTopRequested -= OnScrollToTopRequested;
            base.OnClosed(e);
        }
        private async void Window_Loaded(object sender, RoutedEventArgs e)
        {
            LogUtility.LogDebug("GLSegmentRef.Window_Loaded invoked");
            if (AppState.Instance.SelectedCube != null && AppState.Instance.SelectedLedger != null)
            {
                LogUtility.LogDebug($"GLSegmentRef.Window_Loaded: loading segments - cubeId={AppState.Instance.SelectedLedger.CubeId}, ledgerId={AppState.Instance.SelectedLedger.LedgerId}");
                await vm.LoadSegmentsAsync(AppState.Instance.SelectedLedger.CubeId, AppState.Instance.SelectedLedger.LedgerId);
                await Dispatcher.InvokeAsync(() => FocusFirstSegmentTextBox(), DispatcherPriority.Background);
            }
            else
            {
                LogUtility.LogDebug("GLSegmentRef.Window_Loaded: validation failed - no cube/ledger selected, skipping segment load");
            }

            await Dispatcher.InvokeAsync(() => RefreshWindowLayout(), DispatcherPriority.Render);
        }
        private async Task CmbHierarchy_SelectionCommitted(object obj)
        {
            LogUtility.LogDebug("GLSegmentRef.CmbHierarchy_SelectionCommitted invoked");
            if (obj is SegmentValueModel selectedHierarchy)
            {
                LogUtility.LogDebug($"GLSegmentRef.CmbHierarchy_SelectionCommitted: hierarchy selected - {selectedHierarchy}");
                await vm.LoadSegmentValuesAsync(null, selectedHierarchy, true);
            }
            else
            {
                LogUtility.LogDebug("GLSegmentRef.CmbHierarchy_SelectionCommitted: no valid hierarchy selection, falling back to regular segment values");
                await vm.LoadSegmentValuesAsync(); // fallback to regular segment values
            }
        }
        private async void CmbHierarchy_LostFocus(object sender, RoutedEventArgs e)
        {
            var typedText = cmbHierarchy.SelectedItem;
            if (typedText is not SegmentValueModel)
            {
                LogUtility.LogDebug("GLSegmentRef.CmbHierarchy_LostFocus: typed text is not a valid hierarchy selection, reloading segment values");
                await vm.LoadSegmentValuesAsync();
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
            LogUtility.LogDebug($"GLSegmentRef.CellSelectionWarning invoked - message={message}");
            AppOverlayControl.ShowWarning(message);
        }
        private async void SegmentControl_GotFocus(object sender, RoutedEventArgs e)
        {
            if (!TryGetElements(sender, out var fe, out var seg, out var vm1))
            {
                LogUtility.LogDebug("GLSegmentRef.SegmentControl_GotFocus: validation failed - unable to resolve elements from sender");
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
                var grid = FindVisualChild<Grid>(container);  // Now compiles
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
            LogUtility.LogDebug("GLSegmentRef.BtnAdd_Click invoked");
            var viewModel = (SegmentSelectorViewModel)DataContext;
            var selected = dgLeft.SelectedItems.Cast<SegmentValueModel>().ToList();
            LogUtility.LogDebug($"GLSegmentRef.BtnAdd_Click: adding {selected.Count} selected item(s)");
            viewModel.AddSelection(selected);
        }
        private void BtnFirst_Click(object sender, RoutedEventArgs e)
        {
            LogUtility.LogDebug("GLSegmentRef.BtnFirst_Click invoked");
            vm.GoFirstPage();
        }

        private void BtnPrev_Click(object sender, RoutedEventArgs e)
        {
            LogUtility.LogDebug("GLSegmentRef.BtnPrev_Click invoked");
            vm.GoPreviousPage();
        }

        private void BtnNext_Click(object sender, RoutedEventArgs e)
        {
            LogUtility.LogDebug("GLSegmentRef.BtnNext_Click invoked");
            vm.GoNextPage();
        }

        private void BtnLast_Click(object sender, RoutedEventArgs e)
        {
            LogUtility.LogDebug("GLSegmentRef.BtnLast_Click invoked");
            vm.GoLastPage();
        }

        private void BtnRemove_Click(object sender, RoutedEventArgs e)
        {
            LogUtility.LogDebug("GLSegmentRef.BtnRemove_Click invoked");
            var viewModel = (SegmentSelectorViewModel)DataContext;
            var selected = dgRight.SelectedItems.Cast<SegmentSelectionModel>().ToList();
            LogUtility.LogDebug($"GLSegmentRef.BtnRemove_Click: removing {selected.Count} selected item(s)");
            viewModel.RemoveSelection(selected);
        }

        private void BtnBetween_Click(object sender, RoutedEventArgs e)
        {
            LogUtility.LogDebug("GLSegmentRef.BtnBetween_Click invoked");
            var viewModel = (SegmentSelectorViewModel)DataContext;
            viewModel.AddBetweenSelection(dgLeft.SelectedItems.Cast<SegmentValueModel>().ToList(), false);
        }

        private void BtnNotBetween_Click(object sender, RoutedEventArgs e)
        {
            LogUtility.LogDebug("GLSegmentRef.BtnNotBetween_Click invoked");
            var viewModel = (SegmentSelectorViewModel)DataContext;
            viewModel.AddBetweenSelection(dgLeft.SelectedItems.Cast<SegmentValueModel>().ToList(), true);
        }

        private void BtnExclude_Click(object sender, RoutedEventArgs e)
        {
            LogUtility.LogDebug("GLSegmentRef.BtnExclude_Click invoked");
            var viewModel = (SegmentSelectorViewModel)DataContext;
            viewModel.AddExcludeSelection(dgLeft.SelectedItems.Cast<SegmentValueModel>().ToList());
        }

        private void DgRight_MouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
            LogUtility.LogDebug("GLSegmentRef.DgRight_MouseDoubleClick invoked");
            var viewModel = (SegmentSelectorViewModel)DataContext;
            var selected = dgRight.SelectedItems.Cast<SegmentSelectionModel>().ToList();
            viewModel.RemoveSelection(selected);
        }

        private void DgLeft_MouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
            LogUtility.LogDebug("GLSegmentRef.DgLeft_MouseDoubleClick invoked");
            var viewModel = (SegmentSelectorViewModel)DataContext;
            var selected = dgLeft.SelectedItems.Cast<SegmentValueModel>().ToList();
            viewModel.AddSelection(selected);
        }

        private void BtnClose_Click(object sender, RoutedEventArgs e)
        {
            LogUtility.LogDebug("GLSegmentRef.BtnClose_Click invoked");
            Close();
        }
        private void BtnClearDefaults_Click(object sender, RoutedEventArgs e)
        {
            LogUtility.LogDebug("GLSegmentRef.BtnClearDefaults_Click invoked");
            vm.ClearDefaults();
        }
        private void BtnOK_Click(object sender, RoutedEventArgs e)
        {
            LogUtility.LogDebug("GLSegmentRef.BtnOK_Click invoked");
            var vm1 = this.DataContext as SegmentSelectorViewModel;
            var output = vm1.GetAllSegmentValues();
            GLSegments_SelectedValue = output.Replace("\"", "");
            LogUtility.LogDebug($"GLSegmentRef.BtnOK_Click: selected value set, closing dialog with result true");
            this.DialogResult = true;
            this.Close();
        }
    }
}

