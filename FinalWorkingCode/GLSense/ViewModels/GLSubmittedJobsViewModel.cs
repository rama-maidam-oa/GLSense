using GLSense.Drilldowns;
using GLSense.Helpers;
using GLSense.Models;
using GLSense.Service;
using GLSense.Utilities;
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
using System.Windows.Forms;
using System.Windows.Data;
using System.Windows.Threading;
using Excel = Microsoft.Office.Interop.Excel;

namespace GLSense.ViewModels
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
                LogUtility.LogDebug($"GLSubmittedJobsViewModel.Jobs (set): count={value?.Count ?? 0}");
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
                LogUtility.LogDebug($"GLSubmittedJobsViewModel.SearchText (set): value='{value}'");
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
                LogUtility.LogDebug($"GLSubmittedJobsViewModel.SelectedSearchType (set): value='{value?.Value}'");
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
            LogUtility.LogDebug("GLSubmittedJobsViewModel..ctor: entry");
            _dispatcher = System.Windows.Application.Current.Dispatcher;
            _jobs = new ObservableCollection<GLJobModel>();
            _selectedSearchType = SearchTypeService.GetDefaultSearchType();
            ConfigureJobsView();
        }

        private void ConfigureJobsView()
        {
            LogUtility.LogDebug("GLSubmittedJobsViewModel.ConfigureJobsView: entry");
            var view = CollectionViewSource.GetDefaultView(Jobs);
            if (view != null)
            {
                view.Filter = FilterJob;
            }

            JobsView = view;
        }

        private void RefreshFilteredJobs()
        {
            LogUtility.LogDebug("GLSubmittedJobsViewModel.RefreshFilteredJobs: refreshing JobsView");
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
                job.JobDescription,
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
            LogUtility.LogDebug("GLSubmittedJobsViewModel.LoadJobsAsync: entry");
            try
            {
                await ShowBusyAsync("Loading processed jobs...");

                // Check if there are drilldown jobs in Excel
                bool xlJobsExist = DrillDownJobsExists();
                LogUtility.LogDebug($"GLSubmittedJobsViewModel.LoadJobsAsync: xlJobsExist={xlJobsExist}");

                if (xlJobsExist)
                {
                    _drillJobsList = GetExcelDrilldownJobs();
                }
                else
                {
                    LogUtility.LogDebug("GLSubmittedJobsViewModel.LoadJobsAsync: no drilldown/snapshot named ranges found in workbook, aborting load.");
                    await HideBusyAsync();
                    // Benign/expected state (not an error) - use the non-blurring status
                    // notification so the window doesn't look like something went wrong
                    // every time this screen is opened with no jobs yet (OISR-21811).
                    await ShowStatusMessageAsync("No drilldown/snapshot jobs exist.");
                    return;
                }

                // Get processed jobs from API (from your VB.NET GetProcessRecords())
                LogUtility.LogDebug("GLSubmittedJobsViewModel.LoadJobsAsync: requesting processed jobs from API");
                var processedJobs = await GetProcessRecordsAsync();
                LogUtility.LogDebug($"GLSubmittedJobsViewModel.LoadJobsAsync: API response received, length={processedJobs?.Length ?? 0}");

                if (!string.IsNullOrWhiteSpace(processedJobs))
                {
                    await ParseAndDisplayJobs(processedJobs);
                }

            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex, "GLSubmittedJobsViewModel.LoadJobsAsync - failed to load processed jobs");
                await ShowWarningMessageAsync($"Error loading jobs: {ex.Message}");
            }
            finally
            {
                await HideBusyAsync();
                LogUtility.LogDebug("GLSubmittedJobsViewModel.LoadJobsAsync: exit");
            }
        }
        private bool DrillDownJobsExists()
        {
            LogUtility.LogDebug("GLSubmittedJobsViewModel.DrillDownJobsExists: entry");
            try
            {
                if (_excelApp?.ActiveWorkbook?.Names == null)
                {
                    LogUtility.LogDebug("GLSubmittedJobsViewModel.DrillDownJobsExists: ExcelApp/ActiveWorkbook/Names is null.");
                    return false;
                }

                foreach (Excel.Name name in _excelApp.ActiveWorkbook.Names)
                {
                    if (name.Name.Contains("GLSense_DD"))
                        return true;
                }
                return false;
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex, "GLSubmittedJobsViewModel.DrillDownJobsExists - failed to check workbook named ranges");
                return false;
            }
        }

        private System.Collections.ArrayList GetExcelDrilldownJobs()
        {
            LogUtility.LogDebug("GLSubmittedJobsViewModel.GetExcelDrilldownJobs: entry");
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
                LogUtility.LogDebug($"GLSubmittedJobsViewModel.GetExcelDrilldownJobs: found {drillJobsList.Count} drilldown/snapshot named range(s)");
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex, "GLSubmittedJobsViewModel.GetExcelDrilldownJobs - failed to enumerate workbook named ranges");
            }
            return drillJobsList;
        }

        private static async Task<string> GetProcessRecordsAsync()
        {
            string apiUrl =
                $"{AppState.Instance.LoginUrl}{AppConstants.RestSecure}drilldown-processes?limit=100&page=1";

            LogUtility.LogDebug($"GLSubmittedJobsViewModel.GetProcessRecordsAsync: entry, API={apiUrl}");
            try
            {
                using var cts =
                    new CancellationTokenSource(TimeSpan.FromSeconds(300));

                string response =
                    await ApiHelper.ServerAPI(apiUrl, "Form", "", "GET", cts.Token);

                LogUtility.LogDebug($"GLSubmittedJobsViewModel.GetProcessRecordsAsync: response received, length={response?.Length ?? 0}");

                ValidateTransportResponse(response);

                var result =
                    ApiResponseHelper.Parse<JsonElement>(response, JsonGlobals.Options);

                if (!result.IsSuccess)
                {
                    LogUtility.LogWarn(
                        $"{nameof(GetProcessRecordsAsync)}|API: {apiUrl}");
                    LogUtility.LogWarn(
                        $"{nameof(GetProcessRecordsAsync)}|Response: {response}");

                    return string.Empty;
                }


                return response;
            }
            catch (OperationCanceledException)
            {
                LogUtility.LogWarn(
                    "GetProcessRecordsAsync cancelled/timeout after 30s");
                return string.Empty;
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex, $"GLSubmittedJobsViewModel.GetProcessRecordsAsync - failed calling API {apiUrl}");
                return string.Empty;
            }
        }
        private static void ValidateTransportResponse(string response)
        {
            LogUtility.LogDebug("GLSubmittedJobsViewModel.ValidateTransportResponse: entry");
            if (string.IsNullOrWhiteSpace(response))
                throw new InvalidOperationException("Empty API response.");

            if (response.IndexOf("(401)", StringComparison.OrdinalIgnoreCase) >=0  ||
                response.IndexOf("401: Unauthorized", StringComparison.OrdinalIgnoreCase) >=0)
                throw new UnauthorizedAccessException("Session expired.");

            if (response.StartsWith("Error", StringComparison.OrdinalIgnoreCase))
                throw new InvalidOperationException(response);

            if (response.StartsWith("ORA", StringComparison.OrdinalIgnoreCase))
                throw new InvalidOperationException(response);

            if (response.IndexOf("<!DOCTYPE", StringComparison.OrdinalIgnoreCase) >=0)
                throw new InvalidOperationException("Unexpected HTML response.");
        }


        private async Task ParseAndDisplayJobs(string processedJobs)
        {
            LogUtility.LogDebug($"GLSubmittedJobsViewModel.ParseAndDisplayJobs: entry, processedJobs.Length={processedJobs?.Length ?? 0}");
            string notificationMessage = null;
            // True only for benign/expected "nothing to show" outcomes (no records), as
            // opposed to genuine failures (bad status, parse errors) - see OISR-21811: the
            // background-blur toast should only be used for real errors, not for a routine
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
                            LogUtility.LogDebug($"GLSubmittedJobsViewModel.ParseAndDisplayJobs: status='{jobsData?.status}' is not success.");
                            notificationMessage = jobsData?.msg ?? "Failed to load jobs.";
                            return;
                        }

                        if (jobsData.records == null || jobsData.records.Length == 0)
                        {
                            LogUtility.LogDebug("GLSubmittedJobsViewModel.ParseAndDisplayJobs: no records returned.");
                            notificationMessage = jobsData.msg ?? "No jobs found.";
                            isBenignEmpty = true;
                            return;
                        }


                        var sorted = jobsData.records
                            .Where(r => ShouldIncludeJob(r, _drillJobsList))
                            .Select(r => CreateJobModel(r))
                            .Where(j => j != null)
                            .OrderByDescending(j => j.ProcessId)
                            .ToList();

                        LogUtility.LogDebug($"GLSubmittedJobsViewModel.ParseAndDisplayJobs: {jobsData.records.Length} record(s) received, {sorted.Count} matched drilldown/snapshot criteria.");

                        Jobs = new ObservableCollection<GLJobModel>(sorted);

                    }
                    catch (JsonException jsonEx)
                    {
                        LogUtility.LogException(jsonEx, "GLSubmittedJobsViewModel.ParseAndDisplayJobs - failed to deserialize job data");
                        LogUtility.LogRawJson("GLSubmittedJobsViewModel.ParseAndDisplayJobs", processedJobs);
                        notificationMessage = "Error parsing job data.";
                    }
                    catch (Exception ex)
                    {
                        LogUtility.LogException(ex, "GLSubmittedJobsViewModel.ParseAndDisplayJobs - unexpected error parsing job data");
                        LogUtility.LogRawJson("GLSubmittedJobsViewModel.ParseAndDisplayJobs", processedJobs);
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
                LogUtility.LogException(ex, $"GLSubmittedJobsViewModel.ParseUnixTimestamp - failed to parse timestamp '{timestamp}'");
                return DateTime.MinValue;
            }
        }
        // Command methods
        public async Task RefreshJobsAsync()
        {
            LogUtility.LogDebug("GLSubmittedJobsViewModel.RefreshJobsAsync: entry");
            SearchText = "";
            await LoadJobsAsync();
        }

        public async Task DownloadLogsAsync()
        {
            var selectedJobs = Jobs.Where(j => j.IsSelected).ToList();
            LogUtility.LogDebug($"GLSubmittedJobsViewModel.DownloadLogsAsync: entry, selectedJobs.Count={selectedJobs.Count}");
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
                        LogUtility.LogDebug($"GLSubmittedJobsViewModel.DownloadLogsAsync: downloading logs for ProcessId={job.ProcessId}");
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
                            LogUtility.LogWarn($"GLSubmittedJobsViewModel.DownloadLogsAsync: no log file returned for ProcessId={job.ProcessId}");
                            failedJobs.Add(job.ProcessId);
                        }
                    }
                    catch (Exception ex)
                    {
                        LogUtility.LogException(ex, $"GLSubmittedJobsViewModel.DownloadLogsAsync - failed downloading logs for ProcessId={job.ProcessId}");
                        failedJobs.Add(job.ProcessId);
                    }
                }

                LogUtility.LogDebug($"GLSubmittedJobsViewModel.DownloadLogsAsync: downloaded={selectedJobs.Count - failedJobs.Count}, failed={failedJobs.Count}");

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
                LogUtility.LogException(ex, "GLSubmittedJobsViewModel.DownloadLogsAsync - unexpected error downloading logs");
                await ShowErrorMessageAsync($"Error downloading logs: {ex.Message}");
            }
        }

        public async Task DownloadOutputsAsync()
        {
            LogUtility.LogDebug("GLSubmittedJobsViewModel.DownloadOutputsAsync: entry");
            if (Jobs.Any(j => j.IsSelected && !IsCompletedSelectedJob(j)))
            {
                LogUtility.LogDebug("GLSubmittedJobsViewModel.DownloadOutputsAsync: one or more selected jobs are not completed/successful.");
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
                LogUtility.LogDebug($"GLSubmittedJobsViewModel.DownloadOutputsAsync: downloadedFiles='{downloadedFiles}'");

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
                LogUtility.LogException(ex, "GLSubmittedJobsViewModel.DownloadOutputsAsync - unexpected error downloading outputs");
                await ShowErrorMessage($"Error downloading outputs: {ex.Message}");
            }
        }
        private async Task<StringBuilder> DownloadOutputsForJobsAsync(IEnumerable<GLJobModel> jobs)
        {
            var jobList = jobs.ToList();
            LogUtility.LogDebug($"GLSubmittedJobsViewModel.DownloadOutputsForJobsAsync: entry, jobs.Count={jobList.Count}");
            var downloadedFiles = new StringBuilder();

            foreach (var job in jobList)
            {
                await DownloadSingleJobOutputAsync(job, downloadedFiles);
            }

            return downloadedFiles;
        }

        private async Task DownloadSingleJobOutputAsync(GLJobModel job, StringBuilder downloadedFiles)
        {
            LogUtility.LogDebug($"GLSubmittedJobsViewModel.DownloadSingleJobOutputAsync: ProcessId={job?.ProcessId}, DrillType={job?.DrillType}");
            if (job.DrillType == "SS")
            {
                await DownloadSnapshotJobAsync(job, downloadedFiles);
                return;
            }

            await DownloadDrilldownAsync(job);
        }

        private async Task DownloadSnapshotJobAsync(GLJobModel job, StringBuilder downloadedFiles)
        {
            LogUtility.LogDebug($"GLSubmittedJobsViewModel.DownloadSnapshotJobAsync: ProcessId={job?.ProcessId}");
            var fileName = await DownloadSnapshotAsync(job.ProcessId);
            if (string.IsNullOrEmpty(fileName))
            {
                LogUtility.LogWarn($"GLSubmittedJobsViewModel.DownloadSnapshotJobAsync: no file returned for ProcessId={job.ProcessId}");
                return;
            }

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
            var selectedJobs = Jobs.Where(j => j.IsSelected).ToList();
            LogUtility.LogDebug($"GLSubmittedJobsViewModel.DeleteSelectedAsync: entry, selectedJobs.Count={selectedJobs.Count}");
            if (!selectedJobs.Any())
            {
                await ShowErrorMessage("Please select jobs to delete.");
                return;
            }

            if (ShowConfirmAction == null)
            {
                LogUtility.LogWarn("GLSubmittedJobsViewModel.DeleteSelectedAsync: ShowConfirmAction is not wired up, aborting delete.");
                return;
            }

            var userConfirmed = await AskUserToConfirmDeletionAsync(selectedJobs.Count);
            LogUtility.LogDebug($"GLSubmittedJobsViewModel.DeleteSelectedAsync: userConfirmed={userConfirmed}");
            if (!userConfirmed)
                return;

            await DeleteJobsWithBusyUiAsync(selectedJobs);
        }

        private async Task<bool> AskUserToConfirmDeletionAsync(int count)
        {
            LogUtility.LogDebug($"GLSubmittedJobsViewModel.AskUserToConfirmDeletionAsync: entry, count={count}");
            var confirmTask = ShowConfirmAction?.Invoke($"Delete {count} selected job(s)?");
            if (confirmTask == null)
                return false;

            var userAction = await confirmTask;
            return userAction.HasValue && userAction.Value;
        }

        private async Task DeleteJobsWithBusyUiAsync(List<GLJobModel> selectedJobs)
        {
            LogUtility.LogDebug($"GLSubmittedJobsViewModel.DeleteJobsWithBusyUiAsync: entry, selectedJobs.Count={selectedJobs?.Count ?? 0}");
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
                LogUtility.LogException(ex, "GLSubmittedJobsViewModel.DeleteJobsWithBusyUiAsync - failed deleting selected jobs");
                await ShowErrorMessage($"Error deleting jobs: {ex.Message}");
            }
            finally
            {
                await HideBusyAsync();
            }
        }

        private void DeleteJobsFromExcel(IEnumerable<GLJobModel> jobs)
        {
            LogUtility.LogDebug("GLSubmittedJobsViewModel.DeleteJobsFromExcel: entry");
            foreach (var job in jobs)
            {
                DeleteExcelNamedRange(job.ProcessId);
            }
        }

        private void RemoveJobsFromCollection(IEnumerable<GLJobModel> jobs)
        {
            var jobList = jobs.ToList();
            LogUtility.LogDebug($"GLSubmittedJobsViewModel.RemoveJobsFromCollection: removing {jobList.Count} job(s) from collection");
            foreach (var job in jobList)
            {
                Jobs.Remove(job);
            }
        }

        public async Task DeleteAllAsync()
        {
            LogUtility.LogDebug($"GLSubmittedJobsViewModel.DeleteAllAsync: entry, Jobs.Count={Jobs.Count}");
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
                    LogUtility.LogDebug($"GLSubmittedJobsViewModel.DeleteAllAsync: userAction={userAction}");
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
                            LogUtility.LogException(ex, "GLSubmittedJobsViewModel.DeleteAllAsync - failed deleting all jobs");
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
            LogUtility.LogDebug($"GLSubmittedJobsViewModel.ShowErrorMessage: message='{message}'");
            await HideBusyAsync();
            await ShowWarningMessageAsync(message);
        }

        private async Task ShowErrorMessageAsync(string message)
        {
            LogUtility.LogDebug($"GLSubmittedJobsViewModel.ShowErrorMessageAsync: message='{message}'");
            await HideBusyAsync();
            await ShowWarningMessageAsync(message);
        }

        private Task ShowInfoMessageAsync(string message)
        {
            LogUtility.LogDebug($"GLSubmittedJobsViewModel.ShowInfoMessageAsync: message='{message}'");
            if (ShowInfoAsyncAction != null)
                return ShowInfoAsyncAction(message);

            ShowInfoAction?.Invoke(message);
            return Task.CompletedTask;
        }

        private Task ShowWarningMessageAsync(string message)
        {
            LogUtility.LogDebug($"GLSubmittedJobsViewModel.ShowWarningMessageAsync: message='{message}'");
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
            LogUtility.LogDebug($"GLSubmittedJobsViewModel.ShowStatusMessageAsync: message='{message}'");
            if (ShowStatusAsyncAction != null)
                return ShowStatusAsyncAction(message);

            return ShowWarningMessageAsync(message);
        }

        // Helper methods from VB.NET
        private async Task<string> DownloadJobLogsAsync(string processId)
        {
            // Implement from your VB.NET GLSense_DownloadProcessesc method
            var APIURL = $"{AppState.Instance.LoginUrl}{AppConstants.WebSecure}schedule/{processId}/log-zip";
            LogUtility.LogDebug($"GLSubmittedJobsViewModel.DownloadJobLogsAsync: ProcessId={processId}, URL={APIURL}");
            return await DownloadFileAsync(APIURL);
        }

        private async Task<string> DownloadSnapshotAsync(string processId)
        {
            var APIURL = $"{AppState.Instance.LoginUrl}{AppConstants.RestSecure}snapshot-output?processId={processId}";
            LogUtility.LogDebug($"GLSubmittedJobsViewModel.DownloadSnapshotAsync: ProcessId={processId}, URL={APIURL}");
            return await DownloadFileAsync(APIURL);
        }
        private async Task<string> DownloadFileAsync(string url)
        {
            LogUtility.LogDebug($"GLSubmittedJobsViewModel.DownloadFileAsync: entry, url={url}");
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
                LogUtility.LogDebug($"GLSubmittedJobsViewModel.DownloadFileAsync: HTTP status={response.StatusCode} for url={url}");

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

                    LogUtility.LogDebug($"GLSubmittedJobsViewModel.DownloadFileAsync: file '{fname}' saved to '{strDestDir}'");
                    return fname;
                }
                else
                {
                    // Log failure details
                    LogUtility.LogWarn($"{MethodBase.GetCurrentMethod().Name}|HTTP Error: {response.StatusCode}");
                    LogUtility.LogWarn($"{MethodBase.GetCurrentMethod().Name}|Error Description: {response.ReasonPhrase}");

                    return string.Empty;
                }

            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex, $"GLSubmittedJobsViewModel.DownloadFileAsync - failed downloading file from url={url}");
            }
            return string.Empty;
        }
        private async Task DownloadDrilldownAsync(GLJobModel job)
        {
            LogUtility.LogDebug($"GLSubmittedJobsViewModel.DownloadDrilldownAsync: entry, ProcessId={job?.ProcessId}, JobDescription={job?.JobDescription}, DrillType={job?.DrillType}");
            try
            {
                string apiUrl = $"{AppState.Instance.LoginUrl}{AppConstants.RestSecure}drilldown-data?processId={job.ProcessId}";
                using var cts = new CancellationTokenSource(TimeSpan.FromSeconds(3600));
                string response = await ApiHelper.ServerAPI(apiUrl, "Form", "", "GET", cts.Token);

                LogUtility.LogDebug(response);

                ValidateTransportResponse(response);

                var result = ApiResponseHelper.Parse<JsonElement>(response, JsonGlobals.Options);

                if (!result.IsSuccess)
                {
                    LogUtility.LogWarn(
                        $"{nameof(GetProcessRecordsAsync)}|API: {apiUrl}");
                    LogUtility.LogWarn(
                        $"{nameof(GetProcessRecordsAsync)}|Response: {response}");

                    return ;
                }

                LogUtility.LogDebug($"GLSubmittedJobsViewModel.DownloadDrilldownAsync: writing drilldown data to worksheet for ProcessId={job.ProcessId}");
                var dataToSheet = new DDDatatoWorksheet(_excelApp, response, job.DrillType.Replace("_TL_", "_UF_"), job.JobDescription, cts.Token, null);
                await dataToSheet.DD_DatetoWorksheet();
                LogUtility.LogDebug($"GLSubmittedJobsViewModel.DownloadDrilldownAsync: worksheet write completed for ProcessId={job.ProcessId}");
            }
            catch (OperationCanceledException)
            {
                await ShowErrorMessage("Drilldown data download was canceled.");
                LogUtility.LogWarn($"GLSubmittedJobsViewModel.DownloadDrilldownAsync: cancelled/timeout for ProcessId={job?.ProcessId}.");
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex, $"GLSubmittedJobsViewModel.DownloadDrilldownAsync - failed downloading drilldown data for ProcessId={job?.ProcessId}");
                await ShowErrorMessage($"Error downloading drilldown data: {ex.Message}");
            }
        }
        private void DeleteExcelNamedRange(string processId)
        {
            LogUtility.LogDebug($"GLSubmittedJobsViewModel.DeleteExcelNamedRange: ProcessId={processId}");
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
                LogUtility.LogException(ex, $"GLSubmittedJobsViewModel.DeleteExcelNamedRange - failed to delete named range for ProcessId={processId}");
            }
        }

        private void DeleteAllExcelNamedRanges()
        {
            LogUtility.LogDebug($"GLSubmittedJobsViewModel.DeleteAllExcelNamedRanges: entry, Names.Count={_excelApp?.ActiveWorkbook?.Names?.Count ?? 0}");
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
                        LogUtility.LogException(ex, $"GLSubmittedJobsViewModel.DeleteAllExcelNamedRanges - failed to delete named range at index {i}");
                        // Continue
                    }
                }
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex, "GLSubmittedJobsViewModel.DeleteAllExcelNamedRanges - failed to enumerate workbook named ranges");
            }
        }

        // Busy overlay helpers
        private async Task ShowBusyAsync(string message)
        {
            LogUtility.LogDebug($"GLSubmittedJobsViewModel.ShowBusyAsync: message='{message}'");
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
            LogUtility.LogDebug("GLSubmittedJobsViewModel.HideBusyAsync: entry");
            if (HideBusyAsyncAction != null)
            {
                await HideBusyAsyncAction.Invoke();
            }
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

        // ---------------- Excel interop for refedit controls ----------------
        private Microsoft.Office.Interop.Excel.Application _excelApp;
        public Microsoft.Office.Interop.Excel.Application ExcelApp
        {
            get => _excelApp;
            set
            {
                LogUtility.LogDebug($"GLSubmittedJobsViewModel.ExcelApp (set): value is {(value == null ? "null" : "non-null")}");
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
