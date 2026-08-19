// GLJobsMonitor.xaml.cs in GLSense.Addin.Core
// Port of GLSense\Views\GLJobsMonitor.xaml.cs (FinalWorkingCode) - the "Processed Jobs"
// window opened by the RibDrillJobs ribbon button, backed by GLSubmittedJobsViewModel
// (ported this pass, ViewModels\GLSubmittedJobsViewModel.cs).
//
// Adjustments made when porting into this project's architecture (mirrors GLCubeDetails.
// xaml.cs - the other BaseWindow that owns a DataGrid+AppOverlay - and GLDrilldownCustomization.
// xaml.cs, both already ported this pass/prior passes; see those files' own header comments
// for the general rules referenced below):
//   - Base class DpiAwareWindow -> BaseWindow. EnhancedDragDropHelper.EnableWindowDrag(this)
//     + the three AddHandler(Preview*Event, ...) toast-dismiss subscriptions -> the
//     dedicated TitleBar_MouseLeftButtonDown handler (drag) already used everywhere else in
//     this project. The three PreviewMouseDown/PreviewKeyDown/PreviewTextInput toast-dismiss
//     handlers were NOT carried forward: grepping every other already-ported BaseWindow in
//     this project (GLCubeDetails, GLDrilldownCustomization, GLLogin, GLSegmentFunctions,
//     GLGetPeriod*) confirms none of them wire up window-level toast-dismiss-on-any-input -
//     AppOverlay's own DismissToast() is only ever called from inside AppOverlay.xaml.cs
//     itself in this project. Reproducing the old window-level hook here would be a
//     one-off inconsistency vs. every sibling window instead of a straight port of an
//     established pattern - dropped rather than guessed at.
//   - LogUtility.* (static) -> N/A here (no direct logger calls in the original file).
//   - AppState.Instance.ExcelApp.Application -> ServiceLocator.ExcelApp (this project's
//     AppState has no ExcelApp field - same gap DD_BL.cs/GLSubmittedJobsViewModel.cs
//     document).
//   - Window_MouseDown (manual DragMove on left-button-down, unused/dead in the original -
//     no XAML MouseDown handler wires it up) - dropped as dead code, consistent with this
//     project already dropping other confirmed-dead old handlers.
// No functional changes to the job-monitor behavior itself (refresh/download/delete flows
// all live in the ViewModel, unchanged).
using GLSense.Addin.Core.Infrastructure;
using GLSense.Addin.Core.Models;
using GLSense.Addin.Core.ViewModels;
using System;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;
using GLSense.Addin.Core.Utilities;

namespace GLSense.Addin.Core.Views
{
    /// <summary>
    /// Interaction logic for GLJobsMonitor.xaml
    /// </summary>
    public partial class GLJobsMonitor : BaseWindow
    {
        private readonly GLSubmittedJobsViewModel vm;

        public GLJobsMonitor()
        {
            InitializeComponent();
            ServiceLocator.Logger?.LogDebug("GLJobsMonitor constructor invoked");

            // "Name" (index 1, previously the highest-weighted 3* column) fills any left-over
            // width instead of leaving a blank gap now that every column is Width="Auto" (see
            // DataGridColumnFillHelper for why the star-width columns were removed).
            DataGridColumnFillHelper.EnableFillColumn(dgJobs, dgJobs.Columns[1]);

            vm = new GLSubmittedJobsViewModel
            {
                ExcelApp = ServiceLocator.ExcelApp,
                ShowWarningAction = (msg) => Dispatcher.Invoke(() => AppOverlayControl.ShowWarning(msg)),
                ShowInfoAction = (msg) => Dispatcher.Invoke(() => AppOverlayControl.ShowInfo(msg)),
                ShowInfoAsyncAction = async (msg) => await Dispatcher.InvokeAsync(async () => await AppOverlayControl.ShowInfoAsync(msg)),
                ShowWarningAsyncAction = async (msg) => await Dispatcher.InvokeAsync(async () => await AppOverlayControl.ShowWarningAsync(msg)),
                ShowStatusAsyncAction = async (msg) => await Dispatcher.InvokeAsync(async () => await AppOverlayControl.ShowStatusAsync(msg)),
                ShowConfirmAction = (msg) => Dispatcher.Invoke(() => AppOverlayControl.ShowConfirmAsync(msg)),
                ShowBusyAction = async (txt, cancel) =>
                        await Dispatcher.InvokeAsync(async () =>
                            await AppOverlayControl.ShowBusyasynTask(txt, cancel)),
                HideBusyAsyncAction = async () => await Dispatcher.InvokeAsync(async () => await AppOverlayControl.HideBusyAsync())
            };

            DataContext = vm;
        }

        // ---------- Title bar (drag / close) ----------

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
            ServiceLocator.Logger?.LogDebug("GLJobsMonitor.Window_Loaded invoked - loading jobs");
            try
            {
                await vm.LoadJobsAsync();
                ServiceLocator.Logger?.LogDebug("GLJobsMonitor.Window_Loaded: jobs loaded successfully");
            }
            catch (Exception ex)
            {
                ServiceLocator.Logger?.LogException(ex, "GLJobsMonitor.Window_Loaded");
            }
        }

        private void BtnClose_Click(object sender, RoutedEventArgs e)
        {
            ServiceLocator.Logger?.LogDebug("GLJobsMonitor.BtnClose_Click invoked - closing window");
            Close();
        }

        private async void BtnRefresh_Click(object sender, RoutedEventArgs e)
        {
            ServiceLocator.Logger?.LogDebug("GLJobsMonitor.BtnRefresh_Click invoked");
            try
            {
                await vm.RefreshJobsAsync();
                ServiceLocator.Logger?.LogDebug("GLJobsMonitor.BtnRefresh_Click: jobs refreshed successfully");
            }
            catch (Exception ex)
            {
                ServiceLocator.Logger?.LogException(ex, "GLJobsMonitor.BtnRefresh_Click");
            }
        }

        private async void BtnDownloadLogs_Click(object sender, RoutedEventArgs e)
        {
            ServiceLocator.Logger?.LogDebug("GLJobsMonitor.BtnDownloadLogs_Click invoked");
            try
            {
                await vm.DownloadLogsAsync();
                ServiceLocator.Logger?.LogDebug("GLJobsMonitor.BtnDownloadLogs_Click: download completed");
            }
            catch (Exception ex)
            {
                ServiceLocator.Logger?.LogException(ex, "GLJobsMonitor.BtnDownloadLogs_Click");
            }
        }

        private async void BtnDownloadOutputs_Click(object sender, RoutedEventArgs e)
        {
            ServiceLocator.Logger?.LogDebug("GLJobsMonitor.BtnDownloadOutputs_Click invoked");
            try
            {
                await vm.DownloadOutputsAsync();
                ServiceLocator.Logger?.LogDebug("GLJobsMonitor.BtnDownloadOutputs_Click: download completed");
            }
            catch (Exception ex)
            {
                ServiceLocator.Logger?.LogException(ex, "GLJobsMonitor.BtnDownloadOutputs_Click");
            }
        }

        private async void BtnDelete_Click(object sender, RoutedEventArgs e)
        {
            ServiceLocator.Logger?.LogDebug("GLJobsMonitor.BtnDelete_Click invoked");
            try
            {
                await vm.DeleteSelectedAsync();
                ServiceLocator.Logger?.LogDebug("GLJobsMonitor.BtnDelete_Click: selected jobs deleted");
            }
            catch (Exception ex)
            {
                ServiceLocator.Logger?.LogException(ex, "GLJobsMonitor.BtnDelete_Click");
            }
        }

        private async void BtnDeleteAll_Click(object sender, RoutedEventArgs e)
        {
            ServiceLocator.Logger?.LogDebug("GLJobsMonitor.BtnDeleteAll_Click invoked");
            try
            {
                await vm.DeleteAllAsync();
                ServiceLocator.Logger?.LogDebug("GLJobsMonitor.BtnDeleteAll_Click: all jobs deleted");
            }
            catch (Exception ex)
            {
                ServiceLocator.Logger?.LogException(ex, "GLJobsMonitor.BtnDeleteAll_Click");
            }
        }

        private void ChkSelectAll_Checked(object sender, RoutedEventArgs e)
        {
            if (DataContext is GLSubmittedJobsViewModel vm1 && vm1.Jobs != null)
            {
                foreach (var job in vm1.Jobs)
                {
                    job.IsSelected = true;
                }
            }
        }

        private void ChkSelectAll_Unchecked(object sender, RoutedEventArgs e)
        {
            if (DataContext is GLSubmittedJobsViewModel vm1 && vm1.Jobs != null)
            {
                foreach (var job in vm.Jobs)
                {
                    job.IsSelected = false;
                }
            }
        }

        private void DgJobs_Loaded(object sender, RoutedEventArgs e)
        {
            // Subscribe to property changes on jobs
            if (DataContext is GLSubmittedJobsViewModel vm1)
            {
                vm1.Jobs.CollectionChanged += (s, args) =>
                {
                    UpdateHeaderCheckbox();
                };

                foreach (var job in vm1.Jobs)
                {
                    job.PropertyChanged += (s, args) =>
                    {
                        if (args.PropertyName == nameof(GLJobModel.IsSelected))
                        {
                            UpdateHeaderCheckbox();
                        }
                    };
                }
            }
        }

        private void UpdateHeaderCheckbox()
        {
            if (DataContext is GLSubmittedJobsViewModel vm1 && vm1.Jobs != null && vm1.Jobs.Count > 0)
            {
                var allSelected = vm1.Jobs.All(j => j.IsSelected);
                var anySelected = vm1.Jobs.Any(j => j.IsSelected);

                // Update header checkbox state
                if (allSelected)
                {
                    chkSelectAll.IsChecked = true;
                }
                else if (anySelected)
                {
                    chkSelectAll.IsChecked = null; // Indeterminate state
                }
                else
                {
                    chkSelectAll.IsChecked = false;
                }
            }
            else
            {
                chkSelectAll.IsChecked = false;
            }
        }

        private void DgJobsRow_PreviewMouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            if (e.OriginalSource is DependencyObject source)
            {
                if (FindAncestor<CheckBox>(source) != null)
                    return;

                var row = FindAncestor<DataGridRow>(source);
                if (row?.DataContext is GLJobModel job)
                {
                    job.IsSelected = !job.IsSelected;
                    row.IsSelected = job.IsSelected;
                    UpdateHeaderCheckbox();
                    e.Handled = true;
                }
            }
        }

        private static T FindAncestor<T>(DependencyObject current) where T : DependencyObject
        {
            while (current != null)
            {
                if (current is T typed)
                    return typed;

                current = VisualTreeHelper.GetParent(current);
            }

            return null;
        }
    }
}
