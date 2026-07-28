using GLSense.Helpers;
using GLSense.Models;
using GLSense.Utilities;
using GLSense.ViewModels;
using System;
using System.Linq;
using System.Windows.Controls;
using System.Windows;
using System.Windows.Input;
using System.Windows.Media;

namespace GLSense.Views
{
    /// <summary>
    /// Interaction logic for GLJobsMonitor.xaml
    /// </summary>
    public partial class GLJobsMonitor : DpiAwareWindow
    {
        private readonly GLSubmittedJobsViewModel vm;
        public GLJobsMonitor()
        {
            LogUtility.LogDebug("GLJobsMonitor.ctor invoked");
            InitializeComponent();

            EnhancedDragDropHelper.EnableWindowDrag(this);

            AddHandler(UIElement.PreviewMouseDownEvent, new MouseButtonEventHandler(Window_PreviewMouseDown), true);
            AddHandler(UIElement.PreviewKeyDownEvent, new KeyEventHandler(Window_PreviewKeyDown), true);
            AddHandler(UIElement.PreviewTextInputEvent, new TextCompositionEventHandler(Window_PreviewTextInput), true);

            vm = new GLSubmittedJobsViewModel
            {
                ExcelApp = AppState.Instance.ExcelApp.Application,
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
        private async void Window_Loaded(object sender, RoutedEventArgs e)
        {
            LogUtility.LogDebug("GLJobsMonitor.Window_Loaded invoked");

            await vm.LoadJobsAsync();
            LogUtility.LogDebug("GLJobsMonitor.Window_Loaded: LoadJobsAsync completed");
        }

        private void BtnClose_Click(object sender, RoutedEventArgs e)
        {
            LogUtility.LogDebug("GLJobsMonitor.BtnClose_Click invoked");
            Close();
        }

        private async void BtnRefresh_Click(object sender, RoutedEventArgs e)
        {
            LogUtility.LogDebug("GLJobsMonitor.BtnRefresh_Click invoked");
            await vm.RefreshJobsAsync();
            LogUtility.LogDebug("GLJobsMonitor.BtnRefresh_Click: RefreshJobsAsync completed");
        }

        private async void BtnDownloadLogs_Click(object sender, RoutedEventArgs e)
        {
            LogUtility.LogDebug("GLJobsMonitor.BtnDownloadLogs_Click invoked");
            await vm.DownloadLogsAsync();
            LogUtility.LogDebug("GLJobsMonitor.BtnDownloadLogs_Click: DownloadLogsAsync completed");
        }

        private async void BtnDownloadOutputs_Click(object sender, RoutedEventArgs e)
        {
            LogUtility.LogDebug("GLJobsMonitor.BtnDownloadOutputs_Click invoked");
            await vm.DownloadOutputsAsync();
            LogUtility.LogDebug("GLJobsMonitor.BtnDownloadOutputs_Click: DownloadOutputsAsync completed");
        }

        private async void BtnDelete_Click(object sender, RoutedEventArgs e)
        {
            LogUtility.LogDebug("GLJobsMonitor.BtnDelete_Click invoked");
            await vm.DeleteSelectedAsync();
            LogUtility.LogDebug("GLJobsMonitor.BtnDelete_Click: DeleteSelectedAsync completed");
        }

        private async void BtnDeleteAll_Click(object sender, RoutedEventArgs e)
        {
            LogUtility.LogDebug("GLJobsMonitor.BtnDeleteAll_Click invoked");
            await vm.DeleteAllAsync();
            LogUtility.LogDebug("GLJobsMonitor.BtnDeleteAll_Click: DeleteAllAsync completed");
        }

        private void Window_PreviewMouseDown(object sender, MouseButtonEventArgs e)
        {
            AppOverlayControl.DismissToast();
        }

        private void Window_PreviewKeyDown(object sender, KeyEventArgs e)
        {
            AppOverlayControl.DismissToast();
        }

        private void Window_PreviewTextInput(object sender, TextCompositionEventArgs e)
        {
            AppOverlayControl.DismissToast();
        }

        // Enable window dragging
        private void Window_MouseDown(object sender, MouseButtonEventArgs e)
        {
            if (e.ChangedButton == MouseButton.Left)
                DragMove();
        }
        private void ChkSelectAll_Checked(object sender, RoutedEventArgs e)
        {
            LogUtility.LogDebug("GLJobsMonitor.ChkSelectAll_Checked invoked");
            if (DataContext is GLSubmittedJobsViewModel vm1 && vm1.Jobs != null)
            {
                foreach (var job in vm1.Jobs)
                {
                    job.IsSelected = true;
                }
                LogUtility.LogDebug($"GLJobsMonitor.ChkSelectAll_Checked: selected all {vm1.Jobs.Count} job(s)");
            }
        }

        private void ChkSelectAll_Unchecked(object sender, RoutedEventArgs e)
        {
            LogUtility.LogDebug("GLJobsMonitor.ChkSelectAll_Unchecked invoked");
            if (DataContext is GLSubmittedJobsViewModel vm1 && vm1.Jobs != null)
            {
                foreach (var job in vm.Jobs)
                {
                    job.IsSelected = false;
                }
                LogUtility.LogDebug($"GLJobsMonitor.ChkSelectAll_Unchecked: deselected all {vm1.Jobs.Count} job(s)");
            }
        }
        private void DgJobs_Loaded(object sender, RoutedEventArgs e)
        {
            LogUtility.LogDebug("GLJobsMonitor.DgJobs_Loaded invoked");
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
                    LogUtility.LogDebug($"GLJobsMonitor.DgJobsRow_PreviewMouseLeftButtonDown: job row toggled, IsSelected={job.IsSelected}");
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

