// GLSubmittedJobsViewModel.cs in GLSense.Addin.Core
// Port of GLSense\ViewModels\GLSubmittedJobsViewModel.cs (FinalWorkingCode) for Group E
// (Drilldowns) - backs the GLJobsMonitor window (RibDrillJobs ribbon button), a
// background-jobs monitor that lists submitted drilldown/snapshot jobs (matched against
// "GLSense_DD*" Excel named ranges), polls the finance drilldown-processes API, and
// supports refresh/download-logs/download-outputs/delete/delete-all actions.
//
// Re-pointed vs. the original (business logic/REST URLs/payload shapes unchanged),
// following the same conventions already established by Group C's ViewModels (see
// GLSegmentFuncsViewModel.cs header) and Group E's DD_BL.cs/DDDatatoWorksheet.cs:
//   - GLSense.Drilldowns.DDDatatoWorksheet -> GLSense.Addin.Core.Drilldowns.DDDatatoWorksheet
//     (already ported this pass - constructor/DD_DatetoWorksheet() signature unchanged).
//   - GLSense.Helpers.* (ApiHelper/ApiResponseHelper/JsonGlobals/StrictCertificateValidator)
//     -> GLSense.Addin.Core.Helpers.* (all already ported, signatures unchanged).
//   - GLSense.Models.GLJobModel -> GLSense.Addin.Core.Models.GLJobModel (ported this pass,
//     Models\GLJobModel.cs).
//   - GLSense.Service.SearchTypeService/SearchTypeModel -> GLSense.Addin.Core.Models.
//     SearchTypeService/SearchTypeModel (ported this pass into Models\PeriodModels.cs,
//     see that file's header for why - AttributeTypeService already lived there).
//   - LogUtility.* (static) -> ServiceLocator.Logger?.* (Infrastructure.ServiceLocator).
//   - AppState.Instance.LoginUrl/LoginToken -> unchanged (this project's AppState still
//     has both). AppState.Instance.ExcelApp does NOT exist on this project's AppState
//     (same gap DD_BL.cs's header documents) - the old GLJobsMonitor.xaml.cs set
//     vm.ExcelApp = AppState.Instance.ExcelApp.Application in its own constructor; this
//     port's GLJobsMonitor.xaml.cs instead sets vm.ExcelApp = ServiceLocator.ExcelApp,
//     same fix already applied by every other ported window that needed Excel COM access.
//   - System.Windows.Forms: the original had a stray "using System.Windows.Forms;" that
//     is unused by any code in this class (grepped - no WinForms type referenced anywhere
//     in the method bodies) - dropped rather than carried forward, since this project has
//     no WinForms reference at all (see PORTING_GUIDE.md's MessageBoxIcon/MessageBoxButtons
//     ban).
//   - AppState/AppConstants resolve without an explicit "using" here because this
//     namespace (GLSense.Addin.Core.ViewModels) nests directly under GLSense.Addin.Core,
//     where both live (same resolution already relied on throughout DD_BL.cs).
// No logic changes vs. the original.
using GLSense.Addin.Core.Drilldowns;
using GLSense.Addin.Core.Helpers;
using GLSense.Addin.Core.Infrastructure;
using GLSense.Addin.Core.Models;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Net.Http;
using System.Reflection;
using System.Text;
using System.Text.Json;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Data;
using System.Windows.Threading;
using Excel = Microsoft.Office.Interop.Excel;

namespace GLSense.Addin.Core.ViewModels
{
    public class GLSubmittedJobsViewModel : INotifyPropertyChanged
    {
        private readonly Dispatcher _dispatcher;

        // Actions for window overlay controls
        public Action<string> ShowWarningAction { get; set; }
        public Action<string> ShowInfoAction { get; set; }
        public Func<string, Task> ShowInfoAsyncAction { get; set; }
        public Func<string, Task> ShowWarningAsyncAction { get; set; }
        // Lightweight, non-blurring notification for benign/expected "nothing to show" states
        // (e.g. no drilldown jobs exist yet) - see ShowStatusMessageAsync below (OISR-21811).
        public Func<string, Task> ShowStatusAsyncAction { get; set; }
        public Func<string, Func<Task>, Task> ShowBusyAction { get; set; }
        public Func<string, Task<bool?>> ShowConfirmAction { get; set; }
        public Func<Task> HideBusyAsyncAction { get; set; }

        // Data collections
        private ObservableCollection<GLJobModel> _jobs;
        public ObservableCollection<GLJobModel> Jobs
        {
            get => _jobs;
            set
            {
                if (SetProperty(ref _jobs, value))
                {
                    ConfigureJobsView();
                    RefreshFilteredJobs();
                }
            }
        }

        private ICollectionView _jobsView;
        public ICollectionView JobsView
        {
            get => _jobsView;
            private set => SetProperty(ref _jobsView, value);
        }

        // Search properties
        private string _searchText;
        public string SearchText
        {
            get => _searchText;
            set
            {
                if (SetProperty(ref _searchText, value))
                {
                    RefreshFilteredJobs();
                }
            }
        }

        private SearchTypeModel _selectedSearchType;
        public SearchTypeModel SelectedSearchType
        {
            get => _selectedSearchType;
            set
            {
                if (SetProperty(ref _selectedSearchType, value))
                {
                    RefreshFilteredJobs();
                }
            }
        }

        // Static search types
        public static ObservableCollection<SearchTypeModel> SearchTypes => SearchTypeService.GetSearchTypes();
        private System.Collections.ArrayList _drillJobsList;

        public GLSubmittedJobsViewModel()
        {
            _dispatcher = System.Windows.Application.Current.Dispatcher;
            _jobs = new ObservableCollection<GLJobModel>();
            _selectedSearchType = SearchTypeService.GetDefaultSearchType();
            ConfigureJobsView();
        }

        private void ConfigureJobsView()
        {
            var view = CollectionViewSource.GetDefaultView(Jobs);
            if (view != null)
            {
                view.Filter = FilterJob;
            }

            JobsView = view;
        }

        private void RefreshFilteredJobs()
        {
            JobsView?.Refresh();
        }

        private bool FilterJob(object item)
        {
            if (!(item is GLJobModel job))
                return false;

            var searchText = SearchText?.Trim();
            if (string.IsNullOrWhiteSpace(searchText))
                return true;

            var criteria = SelectedSearchType?.Value ?? "Contains";
            var searchTarget = string.Join(" ", new[]
            {
                job.ProcessId,
                job.Name,
                job.JobType,
                job.Phase,
                job.Status,
                job.DisplayDate?.ToString()
            }.Where(value => !string.IsNullOrWhiteSpace(value)));

            return MatchesSearch(searchTarget, searchText, criteria);
        }

        private static bool MatchesSearch(string value, string searchText, string criteria)
        {
            if (string.IsNullOrWhiteSpace(value))
                return false;

            switch (criteria)
            {
                case "StartsWith":
                    return value.StartsWith(searchText, StringComparison.OrdinalIgnoreCase);
                case "DoesNotStartWith":
                    return !value.StartsWith(searchText, StringComparison.OrdinalIgnoreCase);
                case "EndsWith":
                    return value.EndsWith(searchText, StringComparison.OrdinalIgnoreCase);
                case "DoesNotEndWith":
                    return !value.EndsWith(searchText, StringComparison.OrdinalIgnoreCase);
                case "Equals":
                    return string.Equals(value, searchText, StringComparison.OrdinalIgnoreCase);
                case "NotEquals":
                    return !string.Equals(value, searchText, StringComparison.OrdinalIgnoreCase);
                case "NotContains":
                    return value.IndexOf(searchText, StringComparison.OrdinalIgnoreCase) < 0;
                case "Contains":
                default:
                    return value.IndexOf(searchText, StringComparison.OrdinalIgnoreCase) >= 0;
            }
        }
        // Main load method
        public async Task LoadJobsAsync()
        {
            ServiceLocator.Logger?.LogDebug("GLSubmittedJobsViewModel.LoadJobsAsync started.");
            try
            {
                await ShowBusyAsync("Loading processed jobs...");

                // Check if there are drilldown jobs in Excel
                bool xlJobsExist = DrillDownJobsExists();
                ServiceLocator.Logger?.LogDebug($"GLSubmittedJobsViewModel.LoadJobsAsync: DrillDownJobsExists={xlJobsExist}");

                if (xlJobsExist)
                {
                    _drillJobsList = GetExcelDrilldownJobs();
                    ServiceLocator.Logger?.LogDebug($"GLSubmittedJobsViewModel.LoadJobsAsync: found {_drillJobsList?.Count ?? 0} drilldown/snapshot named range(s) in workbook.");
                }
                else
                {
                    ServiceLocator.Logger?.LogDebug("GLSubmittedJobsViewModel.LoadJobsAsync: no drilldown/snapshot jobs exist in the workbook.");
                    await HideBusyAsync();
                    // Benign/expected state (not an error) - use the non-blurring status
                    // notification so the window doesn't look like something went wrong
                    // every time this screen is opened with no jobs yet (OISR-21811).
                    await ShowStatusMessageAsync("No drilldown/snapshot jobs exist.");
                    return;
                }

                // Get processed jobs from API (from your VB.NET GetProcessRecords())
                var processedJobs = await GetProcessRecordsAsync();
                ServiceLocator.Logger?.LogDebug($"GLSubmittedJobsViewModel.LoadJobsAsync: GetProcessRecordsAsync returned {(string.IsNullOrWhiteSpace(processedJobs) ? "empty" : $"{processedJobs.Length} char(s)")} response.");

                if (!string.IsNullOrWhiteSpace(processedJobs))
                {
                    await ParseAndDisplayJobs(processedJobs);
                }

            }
            catch (Exception ex)
            {
                LogError(ex);
                await ShowWarningMessageAsync($"Error loading jobs: {ex.Message}");
            }
            finally
            {
                await HideBusyAsync();
            }
        }
        private bool DrillDownJobsExists()
        {
            try
            {
                if (_excelApp?.ActiveWorkbook?.Names == null)
                    return false;

                foreach (Excel.Name name in _excelApp.ActiveWorkbook.Names)
                {
                    if (name.Name.Contains("GLSense_DD"))
                        return true;
                }
                return false;
            }
            catch (Exception ex)
            {
                ServiceLocator.Logger?.LogException(ex, "GLSubmittedJobsViewModel.DrillDownJobsExists");
                return false;
            }
        }

        private System.Collections.ArrayList GetExcelDrilldownJobs()
        {
            var drillJobsList = new System.Collections.ArrayList();
            try
            {
                foreach (Excel.Name name in _excelApp.ActiveWorkbook.Names)
                {
                    if (name.Name.Contains("GLSense_DD"))
                    {
                        drillJobsList.Add(name.Value?.ToString()?.Replace("=", ""));
                    }
                }
            }
            catch (Exception ex)
            {
                LogError(ex);
            }
            return drillJobsList;
        }

        private static async Task<string> GetProcessRecordsAsync()
        {
            string apiUrl =
                $"{AppState.Instance.LoginUrl}/rest/secure/finance/drilldown-processes?limit=100&page=1";


            try
            {
                using var cts =
                    new CancellationTokenSource(TimeSpan.FromSeconds(300));

                string response =
                    await ApiHelper.ServerAPI(apiUrl, "Form", "", "GET", cts.Token);

                ValidateTransportResponse(response);

                var result =
                    ApiResponseHelper.Parse<JsonElement>(response, JsonGlobals.Options);

                if (!result.IsSuccess)
                {
                    ServiceLocator.Logger?.LogWarn(
                        $"{nameof(GetProcessRecordsAsync)}|API: {apiUrl}");
                    ServiceLocator.Logger?.LogWarn(
                        $"{nameof(GetProcessRecordsAsync)}|Response: {response}");

                    return string.Empty;
                }


                return response;
            }
            catch (OperationCanceledException)
            {
                ServiceLocator.Logger?.LogWarn(
                    "GetProcessRecordsAsync cancelled/timeout after 30s");
                return string.Empty;
            }
            catch (Exception ex)
            {
                ServiceLocator.Logger?.LogException(ex);
                return string.Empty;
            }
        }
        private static void ValidateTransportResponse(string response)
        {
            if (string.IsNullOrWhiteSpace(response))
                throw new InvalidOperationException("Empty API response.");

            if (response.IndexOf("(401)", StringComparison.OrdinalIgnoreCase) >= 0 ||
                response.IndexOf("401: Unauthorized", StringComparison.OrdinalIgnoreCase) >= 0)
                throw new UnauthorizedAccessException("Session expired.");

            if (response.StartsWith("Error", StringComparison.OrdinalIgnoreCase))
                throw new InvalidOperationException(response);

            if (response.StartsWith("ORA", StringComparison.OrdinalIgnoreCase))
                throw new InvalidOperationException(response);

            if (response.IndexOf("<!DOCTYPE", StringComparison.OrdinalIgnoreCase) >= 0)
                throw new InvalidOperationException("Unexpected HTML response.");
        }


        private async Task ParseAndDisplayJobs(string processedJobs)
        {
            string notificationMessage = null;
            // True only for benign/expected "nothing to show" outcomes (no records), as
            // opposed to genuine failures (bad status, parse errors) - see OISR-21811: the
            // background-dim toast should only be used for real errors, not for a routine
            // empty result.
            bool isBenignEmpty = false;

            await Task.Run(() =>
            {
                _dispatcher.Invoke(() =>
                {
                    Jobs.Clear();

                    try
                    {
                        var jobsData = JsonSerializer.Deserialize<DrilldownJobsResponse>(processedJobs, JsonGlobals.Options);

                        if (jobsData?.status != AppConstants.Success)
                        {
                            notificationMessage = jobsData?.msg ?? "Failed to load jobs.";
                            return;
                        }

                        if (jobsData.records == null || jobsData.records.Length == 0)
                        {
                            notificationMessage = jobsData.msg ?? "No jobs found.";
                            isBenignEmpty = true;
                            return;
                        }


                        var sorted = jobsData.records
                            .Where(r => ShouldIncludeJob(r, _drillJobsList))
                            .OrderByDescending(r => r.processId)
                            .Select(r => CreateJobModel(r))
                            .Where(j => j != null);

                        Jobs = new ObservableCollection<GLJobModel>(sorted);

                    }
                    catch (JsonException jsonEx)
                    {
                        LogError(jsonEx);
                        ServiceLocator.Logger?.LogRawJson("GLSubmittedJobsViewModel.ParseAndDisplayJobs", processedJobs);
                        notificationMessage = "Error parsing job data.";
                    }
                    catch (Exception ex)
                    {
                        LogError(ex);
                        ServiceLocator.Logger?.LogRawJson("GLSubmittedJobsViewModel.ParseAndDisplayJobs", processedJobs);
                        notificationMessage = "Error parsing job data.";
                    }
                });
            });

            if (!string.IsNullOrWhiteSpace(notificationMessage))
            {
                if (isBenignEmpty)
                    await ShowStatusMessageAsync(notificationMessage);
                else
                    await ShowWarningMessageAsync(notificationMessage);
            }
        }

        private static bool ShouldIncludeJob(JobRecord record, System.Collections.ArrayList drillJobsList)
        {

            // For time being ignoring all the cases for testing purpose

            if (drillJobsList == null)
                return false;

            // Check if job is in Excel named ranges
            if (!drillJobsList.Contains(record.processId.ToString()))
                return false;

            // Check for drilldown jobs (with * in name)
            if (record.description.Contains("*"))
            {
                var parts = record.description.Split('*');
                if (parts.Length >= 5)
                    return true;
            }

            // Check for snapshot jobs
            if (record.concurrentJobName == "FinanceSnapshotJob")
                return true;

            return false;
        }
        private GLJobModel CreateJobModel(JobRecord record)
        {
            var job = new GLJobModel
            {
                ProcessId = record.processId.ToString(),
                JobDescription = record.description,
                Phase = record.phase,
                Status = record.status,
                JobType = CleanConcurrentJobName(record.concurrentJobName),
                CreatedDate = ParseUnixTimestamp(record.createdDate),
                DrillType = GetDrillType(record),
                IsSelected = false
            };

            // Set job name
            if (record.description.Contains("*"))
            {
                var parts = record.description.Split('*');
                job.Name = parts.Length > 0 ? parts[0] : record.description;
                if (parts.Length >= 3)
                    job.DateInfo = parts[2];
            }
            else
            {
                job.Name = record.description ?? record.concurrentJobName;
            }

            return job;
        }
        private static string CleanConcurrentJobName(string concurrentJobName)
        {
            if (string.IsNullOrEmpty(concurrentJobName))
                return string.Empty;

            string[] wordsToRemove = { "Finance", "Job" };
            string result = concurrentJobName;

            // Remove each word (case insensitive)
            foreach (var word in wordsToRemove)
            {
                result = System.Text.RegularExpressions.Regex.Replace(
                    result,
                    word,
                    "",
                    System.Text.RegularExpressions.RegexOptions.IgnoreCase
                );
            }

            // Remove any double spaces that might result from removals
            result = System.Text.RegularExpressions.Regex.Replace(result, @"\s+", " ");

            // Trim any leading/trailing spaces
            result = result.Trim();

            // Convert to Proper Case (Title Case)
            if (!string.IsNullOrEmpty(result))
            {

                if (result == "BalanceTotalDrilldown")
                {
                    result = "UnifiedDrilldown";
                }

                // First, split by capital letters to handle camelCase
                var words = System.Text.RegularExpressions.Regex.Split(result, @"(?<!^)(?=[A-Z])");

                // Join with space and then apply Title Case
                result = string.Join(" ", words);

                System.Globalization.TextInfo textInfo = System.Globalization.CultureInfo.CurrentCulture.TextInfo;
                result = textInfo.ToTitleCase(result.ToLower());
            }

            return result;
        }

        private static string GetDrillType(JobRecord record)
        {
            if (record.concurrentJobName == "FinanceSnapshotJob")
                return "SS";

            if (record.description.Contains("*"))
            {
                var parts = record.description.Split('*');
                return parts.Length >= 4 ? parts[3] : string.Empty;
            }

            return string.Empty;
        }

        private static DateTime ParseUnixTimestamp(object timestamp)
        {
            try
            {
                if (timestamp == null)
                    return DateTime.MinValue;

                if (long.TryParse(timestamp.ToString(), out long milliseconds))
                {
                    var epoch = new DateTime(1970, 1, 1, 0, 0, 0, DateTimeKind.Utc);
                    return epoch.AddMilliseconds(milliseconds);
                }
                return DateTime.MinValue;
            }
            catch (Exception ex)
            {
                ServiceLocator.Logger?.LogException(ex, "GLSubmittedJobsViewModel.ParseUnixTimestamp");
                return DateTime.MinValue;
            }
        }
        // Command methods
        public async Task RefreshJobsAsync()
        {
            SearchText = "";
            await LoadJobsAsync();
        }

        public async Task DownloadLogsAsync()
        {
            ServiceLocator.Logger?.LogDebug("GLSubmittedJobsViewModel.DownloadLogsAsync started.");
            var selectedJobs = Jobs.Where(j => j.IsSelected).ToList();
            if (selectedJobs.Count == 0)
            {
                await ShowErrorMessageAsync("Please select jobs to download logs.");
                return;
            }

            try
            {
                await ShowBusyAsync("Downloading logs...");

                var downloadedFiles = new System.Text.StringBuilder();
                var failedJobs = new List<string>();
                foreach (var job in selectedJobs)
                {
                    try
                    {
                        var fileName = await DownloadJobLogsAsync(job.ProcessId);
                        if (!string.IsNullOrEmpty(fileName))
                        {
                            if (downloadedFiles.Length == 0)
                                downloadedFiles.Append($"\"{fileName}\"");
                            else
                                downloadedFiles.Append($", \"{fileName}\"");
                        }
                        else
                        {
                            failedJobs.Add(job.ProcessId);
                        }
                    }
                    catch (Exception ex)
                    {
                        ServiceLocator.Logger?.LogException(ex, $"GLSubmittedJobsViewModel.DownloadLogsAsync: job {job.ProcessId}");
                        failedJobs.Add(job.ProcessId);
                    }
                }

                // Hide busy overlay before showing toast to avoid race condition
                await HideBusyAsync();

                if (downloadedFiles.Length > 0)
                {
                    string strDestDir = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.UserProfile), "Downloads");
                    if (ShowInfoAsyncAction != null)
                    {
                        await ShowInfoAsyncAction($"Logs downloaded to folder {strDestDir}: {downloadedFiles}");
                    }
                }

                if (downloadedFiles.Length == 0 && failedJobs.Count > 0)
                {
                    await ShowWarningMessageAsync("Unable to download logs for the selected jobs.");
                }
            }
            catch (Exception ex)
            {
                LogError(ex);
                await ShowErrorMessageAsync($"Error downloading logs: {ex.Message}");
            }
        }

        public async Task DownloadOutputsAsync()
        {
            ServiceLocator.Logger?.LogDebug("GLSubmittedJobsViewModel.DownloadOutputsAsync started.");
            if (Jobs.Any(j => j.IsSelected && !IsCompletedSelectedJob(j)))
            {
                await ShowErrorMessage("Please select successful jobs to download outputs.");
                return;
            }

            var selectedJobs = GetCompletedSelectedJobs().ToList();
            if (!selectedJobs.Any())
            {
                await ShowErrorMessage("Please select successful jobs to download outputs.");
                return;
            }

            try
            {
                await ShowBusyAsync("Downloading outputs...");

                var downloadedFiles = await DownloadOutputsForJobsAsync(selectedJobs);

                // Hide busy overlay before showing toast to avoid race condition
                await HideBusyAsync();

                if (downloadedFiles.Length > 0)
                {
                    string strDestDir = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.UserProfile), "Downloads");
                    await ShowInfoMessageAsync($"Output(s) downloaded to folder {strDestDir}: {downloadedFiles}");
                }
            }
            catch (Exception ex)
            {
                LogError(ex);
                await ShowErrorMessage($"Error downloading outputs: {ex.Message}");
            }
        }
        private async Task<StringBuilder> DownloadOutputsForJobsAsync(IEnumerable<GLJobModel> jobs)
        {
            var downloadedFiles = new StringBuilder();

            foreach (var job in jobs)
            {
                await DownloadSingleJobOutputAsync(job, downloadedFiles);
            }

            return downloadedFiles;
        }

        private async Task DownloadSingleJobOutputAsync(GLJobModel job, StringBuilder downloadedFiles)
        {
            if (job.DrillType == "SS")
            {
                await DownloadSnapshotJobAsync(job, downloadedFiles);
                return;
            }

            await DownloadDrilldownAsync(job);
        }

        private async Task DownloadSnapshotJobAsync(GLJobModel job, StringBuilder downloadedFiles)
        {
            var fileName = await DownloadSnapshotAsync(job.ProcessId);
            if (string.IsNullOrEmpty(fileName))
                return;

            AppendFileName(downloadedFiles, fileName);
        }

        private static void AppendFileName(StringBuilder builder, string fileName)
        {
            if (builder.Length > 0)
                builder.Append(", ");

            builder.Append('"').Append(fileName).Append('"');
        }

        private IEnumerable<GLJobModel> GetCompletedSelectedJobs()
        {
            return Jobs.Where(IsCompletedSelectedJob);
        }
        private static bool IsCompletedSelectedJob(GLJobModel job)
        {
            if (!job.IsSelected) return false;

            var phase = job.Phase?.ToLowerInvariant() ?? string.Empty;
            var status = job.Status?.ToLowerInvariant() ?? string.Empty;

            var isComplete = phase == "complete" || phase == "completed";
            var isSuccess = status == AppConstants.Success;

            return isComplete && isSuccess;
        }


        public async Task DeleteSelectedAsync()
        {
            ServiceLocator.Logger?.LogDebug("GLSubmittedJobsViewModel.DeleteSelectedAsync started.");
            var selectedJobs = Jobs.Where(j => j.IsSelected).ToList();
            if (!selectedJobs.Any())
            {
                await ShowErrorMessage("Please select jobs to delete.");
                return;
            }

            if (ShowConfirmAction == null)
                return;

            var userConfirmed = await AskUserToConfirmDeletionAsync(selectedJobs.Count);
            if (!userConfirmed)
                return;

            await DeleteJobsWithBusyUiAsync(selectedJobs);
        }

        private async Task<bool> AskUserToConfirmDeletionAsync(int count)
        {
            var confirmTask = ShowConfirmAction?.Invoke($"Delete {count} selected job(s)?");
            if (confirmTask == null)
                return false;

            var userAction = await confirmTask;
            return userAction.HasValue && userAction.Value;
        }

        private async Task DeleteJobsWithBusyUiAsync(List<GLJobModel> selectedJobs)
        {
            try
            {
                await ShowBusyAsync("Deleting selected jobs...");

                DeleteJobsFromExcel(selectedJobs);
                RemoveJobsFromCollection(selectedJobs);

                await ShowInfoMessageAsync($"{selectedJobs.Count} job(s) deleted successfully.");

                await RefreshJobsAsync();
            }
            catch (Exception ex)
            {
                LogError(ex);
                await ShowErrorMessage($"Error deleting jobs: {ex.Message}");
            }
            finally
            {
                await HideBusyAsync();
            }
        }

        private void DeleteJobsFromExcel(IEnumerable<GLJobModel> jobs)
        {
            foreach (var job in jobs)
            {
                DeleteExcelNamedRange(job.ProcessId);
            }
        }

        private void RemoveJobsFromCollection(IEnumerable<GLJobModel> jobs)
        {
            foreach (var job in jobs.ToList())
            {
                Jobs.Remove(job);
            }
        }

        public async Task DeleteAllAsync()
        {
            ServiceLocator.Logger?.LogDebug($"GLSubmittedJobsViewModel.DeleteAllAsync started. Jobs.Count={Jobs.Count}");
            if (Jobs.Count == 0)
            {
                await ShowInfoMessageAsync("No jobs to delete.");
                return;
            }

            if (ShowConfirmAction != null)
            {
                var userActionTask = ShowConfirmAction?.Invoke($"Delete all {Jobs.Count} jobs?");
                if (userActionTask != null)
                {
                    var userAction = await userActionTask;
                    if (userAction.HasValue && userAction.Value)
                    {
                        try
                        {
                            await ShowBusyAsync("Deleting all jobs...");

                            // Delete all Excel named ranges
                            DeleteAllExcelNamedRanges();

                            // Clear collection
                            Jobs.Clear();

                            await ShowInfoMessageAsync("All jobs deleted successfully.");
                            await RefreshJobsAsync();
                        }
                        catch (Exception ex)
                        {
                            LogError(ex);
                            await ShowErrorMessage($"Error deleting jobs: {ex.Message}");
                        }
                        finally
                        {
                            await HideBusyAsync();
                        }
                    }
                }
            }
        }
        private async Task ShowErrorMessage(string message)
        {
            await HideBusyAsync();
            await ShowWarningMessageAsync(message);
        }

        private async Task ShowErrorMessageAsync(string message)
        {
            await HideBusyAsync();
            await ShowWarningMessageAsync(message);
        }

        private Task ShowInfoMessageAsync(string message)
        {
            if (ShowInfoAsyncAction != null)
                return ShowInfoAsyncAction(message);

            ShowInfoAction?.Invoke(message);
            return Task.CompletedTask;
        }

        private Task ShowWarningMessageAsync(string message)
        {
            if (ShowWarningAsyncAction != null)
                return ShowWarningAsyncAction(message);

            ShowWarningAction?.Invoke(message);
            return Task.CompletedTask;
        }

        // Non-blurring notification for benign/expected "nothing to show" states (e.g. no
        // jobs exist yet). Falls back to the regular warning toast if the window hasn't
        // wired up ShowStatusAsyncAction, so this degrades safely (OISR-21811).
        private Task ShowStatusMessageAsync(string message)
        {
            if (ShowStatusAsyncAction != null)
                return ShowStatusAsyncAction(message);

            return ShowWarningMessageAsync(message);
        }

        // Helper methods from VB.NET
        private async Task<string> DownloadJobLogsAsync(string processId)
        {
            // Implement from your VB.NET GLSense_DownloadProcessesc method
            var APIURL = $"{AppState.Instance.LoginUrl}{AppConstants.WebSecure}schedule/{processId}/log-zip";
            return await DownloadFileAsync(APIURL);
        }

        private async Task<string> DownloadSnapshotAsync(string processId)
        {
            var APIURL = $"{AppState.Instance.LoginUrl}/rest/secure/finance/snapshot-output?processId={processId}";
            return await DownloadFileAsync(APIURL);
        }
        private async Task<string> DownloadFileAsync(string url)
        {
            ServiceLocator.Logger?.LogDebug($"GLSubmittedJobsViewModel.DownloadFileAsync: downloading from {url}");
            try
            {
                var handler = new HttpClientHandler()
                {
                    SslProtocols = System.Security.Authentication.SslProtocols.Tls12 | System.Security.Authentication.SslProtocols.Tls13,
                    ServerCertificateCustomValidationCallback = StrictCertificateValidator.Validate
                };

                using var client = new HttpClient(handler);
                client.Timeout = Timeout.InfiniteTimeSpan;
                client.DefaultRequestHeaders.ExpectContinue = false;
                client.DefaultRequestHeaders.Authorization =
                    new System.Net.Http.Headers.AuthenticationHeaderValue("Bearer", AppState.Instance.LoginToken);

                var response = await client.GetAsync(url, HttpCompletionOption.ResponseHeadersRead);

                if (response.IsSuccessStatusCode)
                {
                    string fn = string.Empty;
                    string fname = string.Empty;

                    // Extract filename from headers if Content-Disposition exists
                    if (response.Content.Headers.ContentDisposition != null)
                    {
                        fn = response.Content.Headers.ContentDisposition.FileName;
                    }
                    else
                    {
                        // Default filename if header is missing
                        fn = $"DownloadedFile_{DateTime.Now.Ticks}.dat";
                    }

                    fname = fn.Replace("\"", ""); // Remove quotes if present
                    ServiceLocator.Logger?.LogDebug($"GLSubmittedJobsViewModel.DownloadFileAsync: download succeeded, saving as \"{fname}\"");

                    string strDestDir = Path.Combine(
                        Environment.GetFolderPath(Environment.SpecialFolder.UserProfile),
                        "Downloads",
                        fname);

                    // Read the response stream and save the file
                    using (Stream responseStream = await response.Content.ReadAsStreamAsync())
                    using (var fs = new FileStream(strDestDir, FileMode.Create, FileAccess.Write, FileShare.None, 8192, true))
                    {
                        await responseStream.CopyToAsync(fs);
                    }

                    return fname;
                }
                else
                {
                    // Log failure details
                    ServiceLocator.Logger?.LogWarn($"{MethodBase.GetCurrentMethod().Name}|HTTP Error: {response.StatusCode}");
                    ServiceLocator.Logger?.LogWarn($"{MethodBase.GetCurrentMethod().Name}|Error Description: {response.ReasonPhrase}");

                    return string.Empty;
                }

            }
            catch (Exception ex)
            {
                LogError(ex);
            }
            return string.Empty;
        }
        private async Task DownloadDrilldownAsync(GLJobModel job)
        {
            try
            {
                string apiUrl = $"{AppState.Instance.LoginUrl}/rest/secure/finance/drilldown-data?processId={job.ProcessId}";
                using var cts = new CancellationTokenSource(TimeSpan.FromSeconds(3600));
                string response = await ApiHelper.ServerAPI(apiUrl, "Form", "", "GET", cts.Token);

                ValidateTransportResponse(response);

                var result = ApiResponseHelper.Parse<JsonElement>(response, JsonGlobals.Options);

                if (!result.IsSuccess)
                {
                    ServiceLocator.Logger?.LogWarn(
                        $"{nameof(GetProcessRecordsAsync)}|API: {apiUrl}");
                    ServiceLocator.Logger?.LogWarn(
                        $"{nameof(GetProcessRecordsAsync)}|Response: {response}");

                    return;
                }

                var dataToSheet = new DDDatatoWorksheet(_excelApp, response, job.DrillType.Replace("_TL_", "_UF_"), job.JobDescription, cts.Token, null);
                await dataToSheet.DD_DatetoWorksheet();
            }
            catch (OperationCanceledException)
            {
                await ShowErrorMessage("Drilldown data download was canceled.");
                ServiceLocator.Logger?.LogWarn("Drilldown data download cancelled/timeout for 30s.");
            }
            catch (Exception ex)
            {
                await ShowErrorMessage($"Error downloading drilldown data: {ex.Message}");
            }
        }
        private void DeleteExcelNamedRange(string processId)
        {
            try
            {
                var rangeName = $"GLSense_DD_{processId}";
                foreach (Excel.Name name in _excelApp.ActiveWorkbook.Names)
                {
                    if (name.Name == rangeName)
                    {
                        name.Delete();
                        break;
                    }
                }
            }
            catch (Exception ex)
            {
                LogError(ex);
            }
        }

        private void DeleteAllExcelNamedRanges()
        {
            try
            {
                for (int i = _excelApp.ActiveWorkbook.Names.Count; i >= 1; i--)
                {
                    try
                    {
                        var name = _excelApp.ActiveWorkbook.Names.Item(i);
                        if (name.Name.Contains("GLSense_DD_"))
                        {
                            name.Delete();
                        }
                    }
                    catch (Exception ex)
                    {
                        ServiceLocator.Logger?.LogException(ex, "GLSubmittedJobsViewModel.DeleteAllExcelNamedRanges: deleting single named range (ignored, continuing)");
                        // Continue
                    }
                }
            }
            catch (Exception ex)
            {
                LogError(ex);
            }
        }

        // Busy overlay helpers
        private async Task ShowBusyAsync(string message)
        {
            if (ShowBusyAction != null)
            {
                await ShowBusyAction.Invoke(message, async () =>
                {
                    // Cancel action
                    await Task.CompletedTask;
                });
            }
        }

        private async Task HideBusyAsync()
        {
            if (HideBusyAsyncAction != null)
            {
                await HideBusyAsyncAction.Invoke();
            }
        }
        private static void LogError(Exception ex)
        {
            // Use existing project-wide logging
            ServiceLocator.Logger?.LogException(ex);
        }

        // INotifyPropertyChanged implementation
        public event PropertyChangedEventHandler PropertyChanged;

        protected virtual void OnPropertyChanged(string propertyName)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }

        protected bool SetProperty<T>(ref T field, T value, [System.Runtime.CompilerServices.CallerMemberName] string propertyName = null)
        {
            if (EqualityComparer<T>.Default.Equals(field, value)) return false;
            field = value;
            OnPropertyChanged(propertyName);
            return true;
        }

        // ---------------- Excel interop for named-range/COM access ----------------
        private Microsoft.Office.Interop.Excel.Application _excelApp;
        public Microsoft.Office.Interop.Excel.Application ExcelApp
        {
            get => _excelApp;
            set
            {
                _excelApp = value;
                OnPropertyChanged(nameof(ExcelApp));
            }
        }
    }
    // Supporting classes
    public class DrilldownJobsResponse
    {
        public string status { get; set; }
        public string msg { get; set; }
        public JobRecord[] records { get; set; }
    }

    public class JobRecord
    {
        public long processId { get; set; }
        public string description { get; set; }
        public string concurrentJobName { get; set; }
        public string phase { get; set; }
        public string status { get; set; }
        public object createdDate { get; set; }
    }
}
