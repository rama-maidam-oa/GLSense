// GLSegmentManager.xaml.cs in GLSense.Addin.Core
// Master-detail redesign of GLSegmentRef.xaml.cs, trialed side by side with it (both
// files exist and compile - nothing about GLSegmentRef changed). Same
// SegmentSelectorViewModel, constructed the same way with windowName="Ref", same
// GLSegments_SelectedValue contract on the public constructor/property, so callers can
// swap between the two with a one-line change (see GLBalanceConfigurator.xaml.cs's
// AcctsRef_EditRequested, which currently opens this window - change "new
// GLSegmentManager(...)" back to "new GLSegmentRef(...)" there to roll back).
//
// What's different from GLSegmentRef, and why this file is noticeably shorter:
//   - GLSegmentRef laid every segment out as its own row (Value textbox + Reference
//     box), and used whichever row last had keyboard focus (SegmentControl_GotFocus) to
//     decide which segment the Hierarchy/Search/dual-grid below was operating on -
//     tracked manually via ResetAllRowStyles/HighlightFocusedRow/FindVisualChild walking
//     the visual tree. This window instead puts one segment per row in a plain ListBox
//     (lstSegments) whose SelectedItem is bound TwoWay straight to
//     SegmentSelectorViewModel.SelectedSegment - so a segment is explicitly highlighted,
//     and no focus-tracking/visual-tree code is needed at all.
//   - SegmentSelectorViewModel.SelectedSegment's setter already does everything that used
//     to be manual: it saves the outgoing segment's SelectedItemsRight into its
//     SelectedValues, restores the incoming segment's saved selections, reloads
//     PagedSegmentValues for it, and (via UpdatePagingAndGrid -> ScrollDataGridsToTop)
//     scrolls both grids back to top - all before this file even gets involved. So there
//     is no SegmentControl_GotFocus/UpdateGridsAsync/TryGetElements equivalent here; the
//     ListBox binding is the entire wiring.
//   - Value/Reference are no longer per-row inline controls; they're two controls in the
//     detail panel bound through the selection: "SelectedSegment.Value" / ".Reference" /
//     ".IsTextEnabled" / ".IsRefEditEnabled" (all already on SegmentModel - unchanged).
//   - The left list's subtitle (SegmentSummaryConverter, in Converters.cs) is the one
//     genuinely new piece - it doesn't exist on GLSegmentRef because GLSegmentRef shows
//     every segment's Value/Reference textbox at once, so nothing needed summarizing.
using GLSense.Addin.Core.Helpers;
using GLSense.Addin.Core.Infrastructure;
using GLSense.Addin.Core.Interfaces;
using GLSense.Addin.Core.Models;
using GLSense.Addin.Core.ViewModels;
using System;
using System.Linq;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Input;
using System.Windows.Threading;

namespace GLSense.Addin.Core.Views
{
    /// <summary>
    /// Interaction logic for GLSegmentManager.xaml
    /// </summary>
    public partial class GLSegmentManager : BaseWindow, IWarningHost
    {
        public string GLSegments_SelectedValue { get; set; }
        private readonly SegmentSelectorViewModel vm;

        public GLSegmentManager(string selectedSegValues)
        {
            InitializeComponent();
            ServiceLocator.Logger?.LogDebug($"GLSegmentManager constructor invoked - selectedSegValues={selectedSegValues}");
            EnableExcelCentering = false;

            // No DataGridColumnFillHelper wiring here anymore (CLAUDE.md 26.8): this window
            // is now SizeToContent="Manual" with a fixed Width/Height (see GLSegmentManager.xaml's
            // header comment), so dgLeft's "Description" and dgRight's "Segment" columns are
            // plain, native Width="*" - they size correctly on their own, the same way
            // GLSegmentValues.xaml's identical dual-grid layout always has. That also means
            // none of BaseWindow's SizeToContent-only resettle/pump-dispatcher-frame logic
            // runs for this window (it's gated on SizeToContent != Manual), so there's no
            // "only resettle once" bookkeeping needed here either - this window's size never
            // resettles itself at all now, matching GLSegmentValues.

            vm = new SegmentSelectorViewModel(Dispatcher, "Ref", selectedSegValues)
            {
                ExcelApp = ServiceLocator.ExcelApp,
                ShowWarningAction = (msg) => Dispatcher.Invoke(() => AppOverlayControl.ShowWarning(msg)),
                ShowInfoAction = (msg) => Dispatcher.Invoke(() => AppOverlayControl.ShowInfo(msg)),
                ShowBusyAction = async (txt, cancel) =>
                        await Dispatcher.InvokeAsync(async () =>
                            await AppOverlayControl.ShowBusyasynTask(txt, cancel)),
                HideBusyAsyncAction = async () => await Dispatcher.InvokeAsync(async () => await AppOverlayControl.HideBusyAsync()),
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
            ServiceLocator.Logger?.LogDebug("GLSegmentManager.OnClosed invoked");
            vm.ScrollToTopRequested -= OnScrollToTopRequested;
            base.OnClosed(e);
        }

        private async void Window_Loaded(object sender, RoutedEventArgs e)
        {
            ServiceLocator.Logger?.LogDebug("GLSegmentManager.Window_Loaded invoked");
            try
            {
                if (AppState.Instance.SelectedCube != null && AppState.Instance.SelectedLedger != null)
                {
                    ServiceLocator.Logger?.LogDebug($"GLSegmentManager.Window_Loaded: loading segments - cubeId={AppState.Instance.SelectedLedger.CubeId}, ledgerId={AppState.Instance.SelectedLedger.LedgerId}");
                    await vm.LoadSegmentsAsync(AppState.Instance.SelectedLedger.CubeId, AppState.Instance.SelectedLedger.LedgerId);
                    // SegmentSelectorViewModel.ProcessSegments already auto-selects an
                    // initial segment (SelectInitialSegment), so the list arrives with a
                    // row already highlighted - just move keyboard focus there.
                    await Dispatcher.InvokeAsync(() => lstSegments.Focus(), DispatcherPriority.Background);
                }
                else
                {
                    ServiceLocator.Logger?.LogDebug("GLSegmentManager.Window_Loaded: validation failed - no cube/ledger selected, skipping segment load");
                }
                // No post-load column-fill/resettle step needed anymore (CLAUDE.md 26.8) -
                // dgLeft/dgRight's Description/Segment columns are plain Width="*" now, and
                // this window is a fixed-size SizeToContent="Manual" window, so there's no
                // "grid got a too-small ActualWidth before the window settled" case to work
                // around; the columns are correct from the very first layout pass.
            }
            catch (Exception ex)
            {
                ServiceLocator.Logger?.LogException(ex, "GLSegmentManager.Window_Loaded");
            }
        }

        private async Task CmbHierarchy_SelectionCommitted(object obj)
        {
            ServiceLocator.Logger?.LogDebug("GLSegmentManager.CmbHierarchy_SelectionCommitted invoked");
            try
            {
                if (obj is SegmentValueModel selectedHierarchy)
                {
                    ServiceLocator.Logger?.LogDebug($"GLSegmentManager.CmbHierarchy_SelectionCommitted: hierarchy selected - {selectedHierarchy}");
                    await vm.LoadSegmentValuesAsync(null, selectedHierarchy, true);
                }
                else
                {
                    ServiceLocator.Logger?.LogDebug("GLSegmentManager.CmbHierarchy_SelectionCommitted: no valid hierarchy selection, falling back to regular segment values");
                    await vm.LoadSegmentValuesAsync(); // fallback to regular segment values
                }
            }
            catch (Exception ex)
            {
                ServiceLocator.Logger?.LogException(ex, "GLSegmentManager.CmbHierarchy_SelectionCommitted");
            }
        }

        private async void CmbHierarchy_LostFocus(object sender, RoutedEventArgs e)
        {
            try
            {
                var typedText = cmbHierarchy.SelectedItem;
                if (typedText is not SegmentValueModel)
                {
                    ServiceLocator.Logger?.LogDebug("GLSegmentManager.CmbHierarchy_LostFocus: typed text is not a valid hierarchy selection, reloading segment values");
                    await vm.LoadSegmentValuesAsync();
                }
            }
            catch (Exception ex)
            {
                ServiceLocator.Logger?.LogException(ex, "GLSegmentManager.CmbHierarchy_LostFocus");
            }
        }

        public void CellSelectionWarning(string message)
        {
            ServiceLocator.Logger?.LogDebug($"GLSegmentManager.CellSelectionWarning invoked - message={message}");
            AppOverlayControl.ShowWarning(message);
        }

        private void BtnAdd_Click(object sender, RoutedEventArgs e)
        {
            ServiceLocator.Logger?.LogDebug("GLSegmentManager.BtnAdd_Click invoked");
            var selected = dgLeft.SelectedItems.Cast<SegmentValueModel>().ToList();
            ServiceLocator.Logger?.LogDebug($"GLSegmentManager.BtnAdd_Click: adding {selected.Count} selected item(s)");
            vm.AddSelection(selected);
        }

        private void BtnFirst_Click(object sender, RoutedEventArgs e)
        {
            ServiceLocator.Logger?.LogDebug("GLSegmentManager.BtnFirst_Click invoked");
            vm.GoFirstPage();
        }

        private void BtnPrev_Click(object sender, RoutedEventArgs e)
        {
            ServiceLocator.Logger?.LogDebug("GLSegmentManager.BtnPrev_Click invoked");
            vm.GoPreviousPage();
        }

        private void BtnNext_Click(object sender, RoutedEventArgs e)
        {
            ServiceLocator.Logger?.LogDebug("GLSegmentManager.BtnNext_Click invoked");
            vm.GoNextPage();
        }

        private void BtnLast_Click(object sender, RoutedEventArgs e)
        {
            ServiceLocator.Logger?.LogDebug("GLSegmentManager.BtnLast_Click invoked");
            vm.GoLastPage();
        }

        private void BtnRemove_Click(object sender, RoutedEventArgs e)
        {
            ServiceLocator.Logger?.LogDebug("GLSegmentManager.BtnRemove_Click invoked");
            var selected = dgRight.SelectedItems.Cast<SegmentSelectionModel>().ToList();
            ServiceLocator.Logger?.LogDebug($"GLSegmentManager.BtnRemove_Click: removing {selected.Count} selected item(s)");
            vm.RemoveSelection(selected);
        }

        private void BtnBetween_Click(object sender, RoutedEventArgs e)
        {
            ServiceLocator.Logger?.LogDebug("GLSegmentManager.BtnBetween_Click invoked");
            vm.AddBetweenSelection(dgLeft.SelectedItems.Cast<SegmentValueModel>().ToList(), false);
        }

        private void BtnNotBetween_Click(object sender, RoutedEventArgs e)
        {
            ServiceLocator.Logger?.LogDebug("GLSegmentManager.BtnNotBetween_Click invoked");
            vm.AddBetweenSelection(dgLeft.SelectedItems.Cast<SegmentValueModel>().ToList(), true);
        }

        private void BtnExclude_Click(object sender, RoutedEventArgs e)
        {
            ServiceLocator.Logger?.LogDebug("GLSegmentManager.BtnExclude_Click invoked");
            vm.AddExcludeSelection(dgLeft.SelectedItems.Cast<SegmentValueModel>().ToList());
        }

        private void DgRight_MouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
            ServiceLocator.Logger?.LogDebug("GLSegmentManager.DgRight_MouseDoubleClick invoked");
            var selected = dgRight.SelectedItems.Cast<SegmentSelectionModel>().ToList();
            vm.RemoveSelection(selected);
        }

        private void DgLeft_MouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
            ServiceLocator.Logger?.LogDebug("GLSegmentManager.DgLeft_MouseDoubleClick invoked");
            var selected = dgLeft.SelectedItems.Cast<SegmentValueModel>().ToList();
            vm.AddSelection(selected);
        }

        private void BtnClose_Click(object sender, RoutedEventArgs e)
        {
            ServiceLocator.Logger?.LogDebug("GLSegmentManager.BtnClose_Click invoked");
            Close();
        }

        private void BtnClearDefaults_Click(object sender, RoutedEventArgs e)
        {
            ServiceLocator.Logger?.LogDebug("GLSegmentManager.BtnClearDefaults_Click invoked");
            vm.ClearDefaults();
        }

        private void BtnOK_Click(object sender, RoutedEventArgs e)
        {
            ServiceLocator.Logger?.LogDebug("GLSegmentManager.BtnOK_Click invoked");
            var output = vm.GetAllSegmentValues();
            GLSegments_SelectedValue = output.Replace("\"", "");
            ServiceLocator.Logger?.LogDebug($"GLSegmentManager.BtnOK_Click: selected value set, closing dialog with result true - {GLSegments_SelectedValue}");
            this.DialogResult = true;
            this.Close();
        }
    }
}
