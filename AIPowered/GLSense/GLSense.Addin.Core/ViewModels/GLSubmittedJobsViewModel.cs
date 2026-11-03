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
        // OISR-22830: step line shown above the busy timer (AppOverlay.SetBusyStep).
        public Action<string> BusyStepAction { get; set; }
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
                BusyToken.ThrowIfCancellationRequested();
                ServiceLocator.Logger?.LogDebug($"GLSubmittedJobsViewModel.LoadJobsAsync: GetProcessRecordsAsync returned {(string.IsNullOrWhiteSpace(processedJobs) ? "empty" : $"{processedJobs.Length} char(s)")} response.");

                if (!string.IsNullOrWhiteSpace(processedJobs))
                {
                    await ParseAndDisplayJobs(processedJobs);
                }

            }
            catch (OperationCanceledException) when (UserCancelled)
            {
                await ShowCancelledMessageAsync("Loading jobs");
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
        // OISR-22830: job names are read through Application.Names (the active workbook's names)
        // instead of ActiveWorkbook.Names. Right after a drilldown, Addin.Core already holds a
        // generic COM wrapper for the active workbook, and getting ActiveWorkbook across the
        // AppDomain boundary then fails (InvalidCastException: __ComObject to WorkbookClass) until
        // that wrapper is garbage-collected. Names is an interface type, so it can't hit that.
        private bool DrillDownJobsExists()
        {
            try
            {
                if (_excelApp?.Names == null)
                    return false;

                foreach (Excel.Name name in _excelApp.Names)
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
                foreach (Excel.Name name in _excelApp.Names)
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

        private async Task<string> GetProcessRecordsAsync()
        {
            string apiUrl =
                $"{AppState.Instance.LoginUrl}/rest/secure/finance/drilldown-processes?limit=100&page=1";


            try
            {
                using var cts = CancellationTokenSource.CreateLinkedTokenSource(BusyToken);
                cts.CancelAfter(TimeSpan.FromSeconds(300));

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
            catch (OperationCanceledException) when (UserCancelled)
            {
                throw;
            }
            catch (OperationCanceledException) when (!UserCancelled)
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
                    BusyToken.ThrowIfCancellationRequested();
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
                    catch (OperationCanceledException) when (UserCancelled)
                    {
                        throw;
                    }
                    catch (Exception ex)
                    {
                        ServiceLocator.Logger?.LogException(ex, $"GLSubmittedJobsViewModel.DownloadLogsAsync: job {job.ProcessId}");
                        failedJobs.Add(job.ProcessId);
                    }
                }

                BusyToken.ThrowIfCancellationRequested();

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
            catch (OperationCanceledException) when (UserCancelled)
            {
                await ShowCancelledMessageAsync("Download logs");
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

                BusyToken.ThrowIfCancellationRequested();

                // Hide busy overlay before showing toast to avoid race condition
                await HideBusyAsync();

                if (downloadedFiles.Length > 0)
                {
                    string strDestDir = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.UserProfile), "Downloads");
                    await ShowInfoMessageAsync($"Output(s) downloaded to folder {strDestDir}: {downloadedFiles}");
                }
            }
            catch (OperationCanceledException) when (UserCancelled)
            {
                await ShowCancelledMessageAsync("Download outputs");
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
                BusyToken.ThrowIfCancellationRequested();
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

                var response = await client.GetAsync(url, HttpCompletionOption.ResponseHeadersRead, BusyToken);

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
                    try
                    {
                        using (Stream responseStream = await response.Content.ReadAsStreamAsync())
                        using (var fs = new FileStream(strDestDir, FileMode.Create, FileAccess.Write, FileShare.None, 8192, true))
                        {
                            await responseStream.CopyToAsync(fs, 81920, BusyToken);
                        }
                    }
                    catch (OperationCanceledException)
                    {
                        // Don't leave a half-written file in Downloads.
                        try { File.Delete(strDestDir); } catch (IOException) { } catch (UnauthorizedAccessException) { }
                        throw;
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
            catch (OperationCanceledException) when (UserCancelled)
            {
                throw;
            }
            catch (Exception ex)
            {
                LogError(ex);
            }
            return string.Empty;
        }
        private async Task DownloadDrilldownAsync(GLJobModel job)
        {
            ServiceLocator.Logger?.LogDebug($"GLSubmittedJobsViewModel.DownloadDrilldownAsync: entry, ProcessId={job?.ProcessId}, JobDescription={job?.JobDescription}, DrillType={job?.DrillType}");
            string outputFile = null;
            try
            {
                string apiUrl = $"{AppState.Instance.LoginUrl}/rest/secure/finance/drilldown-data?processId={job.ProcessId}";
                using var cts = CancellationTokenSource.CreateLinkedTokenSource(BusyToken);
                cts.CancelAfter(TimeSpan.FromSeconds(3600));

                // OISR-22830: a Process-screen output can be hundreds of MB (process 2348724:
                // 647 MB). Reading it into one string, logging it and parsing it twice on the UI
                // thread froze Excel. Stream it to a file in the Temp folder instead, deserialize
                // it from the file on a worker thread, and hand the writer the parsed data. The
                // file is deleted afterwards unless Debug is on (then it is kept for diagnosis and
                // removed by the 7-day clean-up of the Temp folder).
                outputFile = GLSense.Shared.LargeLogText.NewFilePath(ServiceLocator.Paths.Temp, $"drilldown-data_{job.ProcessId}");
                BusyStepAction?.Invoke($"Process {job.ProcessId}: waiting for the server...");
                long bytes = await DownloadToFileAsync(apiUrl, outputFile, cts.Token, BusyStepAction);
                BusyToken.ThrowIfCancellationRequested();
                if (bytes < 0)
                {
                    await ShowErrorMessage($"Unable to download the output of process {job.ProcessId}.");
                    return;
                }
                ServiceLocator.Logger?.LogDebug($"GLSubmittedJobsViewModel.DownloadDrilldownAsync: ProcessId={job.ProcessId} downloaded {bytes:N0} bytes to '{outputFile}'");

                BusyStepAction?.Invoke($"Reading the downloaded data ({bytes / (1024L * 1024L):N0} MB)...");
                string fileToRead = outputFile;
                DrillDownQueryData data = await Task.Run(async () =>
                {
                    using var fs = new FileStream(fileToRead, FileMode.Open, FileAccess.Read, FileShare.Read, 1 << 16, FileOptions.Asynchronous | FileOptions.SequentialScan);
                    // Low-memory form: each value kept as its text only (see StreamedDrillDownData).
                    var streamed = await JsonSerializer.DeserializeAsync<StreamedDrillDownData>(fs, JsonGlobals.Options, cts.Token);
                    return streamed?.ToQueryData();
                }, cts.Token);
                BusyToken.ThrowIfCancellationRequested();

                if (data == null ||
                    (!string.IsNullOrWhiteSpace(data.status) &&
                     !string.Equals(data.status, AppConstants.Success, StringComparison.OrdinalIgnoreCase)))
                {
                    ServiceLocator.Logger?.LogWarn($"GLSubmittedJobsViewModel.DownloadDrilldownAsync: ProcessId={job.ProcessId} returned status='{data?.status}', msg='{data?.msg}'");
                    await ShowErrorMessage(string.IsNullOrWhiteSpace(data?.msg) ? "Server returned failure status." : data.msg);
                    return;
                }

                ServiceLocator.Logger?.LogDebug($"GLSubmittedJobsViewModel.DownloadDrilldownAsync: writing {data.records?.Length ?? 0:N0} record(s) to worksheet for ProcessId={job.ProcessId}");
                BusyStepAction?.Invoke($"Writing {data.records?.Length ?? 0:N0} record(s) to Excel...");
                var dataToSheet = new DDDatatoWorksheet(_excelApp, data, job.DrillType.Replace("_TL_", "_UF_"), job.JobDescription, cts.Token, null)
                {
                    StatusCallback = BusyStepAction
                };
                await dataToSheet.DD_DatetoWorksheet();
                if (dataToSheet.WasCancelled)
                    throw new OperationCanceledException(cts.Token);
                ServiceLocator.Logger?.LogDebug($"GLSubmittedJobsViewModel.DownloadDrilldownAsync: worksheet write completed for ProcessId={job.ProcessId}");
            }
            catch (OperationCanceledException) when (UserCancelled)
            {
                throw;
            }
            catch (OperationCanceledException) when (!UserCancelled)
            {
                await ShowErrorMessage("Drilldown data download was canceled.");
                ServiceLocator.Logger?.LogWarn($"GLSubmittedJobsViewModel.DownloadDrilldownAsync: cancelled/timeout for ProcessId={job?.ProcessId}.");
            }
            catch (Exception ex)
            {
                ServiceLocator.Logger?.LogException(ex, $"GLSubmittedJobsViewModel.DownloadDrilldownAsync - failed downloading drilldown data for ProcessId={job?.ProcessId}");
                await ShowErrorMessage($"Error downloading drilldown data: {ex.Message}");
            }
            finally
            {
                DeleteOutputFile(outputFile);
            }
        }

        // OISR-22830: streams an HTTP GET body straight to a file. Returns the byte count, or -1
        // when the server answers with an error status. Same client setup as DownloadFileAsync.
        private static async Task<long> DownloadToFileAsync(string url, string filePath, CancellationToken token, Action<string> step = null)
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

            using var response = await client.GetAsync(url, HttpCompletionOption.ResponseHeadersRead, token);
            if (!response.IsSuccessStatusCode)
            {
                ServiceLocator.Logger?.LogWarn($"GLSubmittedJobsViewModel.DownloadToFileAsync: HTTP {(int)response.StatusCode} {response.ReasonPhrase} for {url}");
                return -1;
            }

            using (Stream body = await response.Content.ReadAsStreamAsync())
            using (var fs = new FileStream(filePath, FileMode.Create, FileAccess.Write, FileShare.None, 1 << 16, true))
            {
                // Copy in chunks so the step line can show the progress (every ~5 MB).
                long total = response.Content.Headers.ContentLength ?? 0;
                long done = 0, nextReport = 0;
                var buffer = new byte[1 << 16];
                int read;
                while ((read = await body.ReadAsync(buffer, 0, buffer.Length, token)) > 0)
                {
                    await fs.WriteAsync(buffer, 0, read, token);
                    done += read;
                    if (step != null && done >= nextReport)
                    {
                        nextReport = done + 5L * 1024 * 1024;
                        long doneMb = done / (1024L * 1024L);
                        step(total > 0
                            ? $"Downloading: {doneMb:N0} of {total / (1024L * 1024L):N0} MB"
                            : $"Downloading: {doneMb:N0} MB");
                    }
                }
                return fs.Length;
            }
        }

        private static void DeleteOutputFile(string filePath)
        {
            if (string.IsNullOrEmpty(filePath))
                return;

            if (AppState.Instance.DebugLogs && File.Exists(filePath))
            {
                ServiceLocator.Logger?.LogDebug($"GLSubmittedJobsViewModel.DeleteOutputFile: Debug is on - full output kept at '{filePath}'");
                return;
            }

            try
            {
                if (File.Exists(filePath))
                    File.Delete(filePath);
            }
            catch (Exception ex)
            {
                ServiceLocator.Logger?.LogWarn($"GLSubmittedJobsViewModel.DeleteOutputFile: could not delete '{filePath}': {ex.Message}");
            }
        }
        private void DeleteExcelNamedRange(string processId)
        {
            try
            {
                var rangeName = $"GLSense_DD_{processId}";
                foreach (Excel.Name name in _excelApp.Names)
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
                for (int i = _excelApp.Names.Count; i >= 1; i--)
                {
                    try
                    {
                        var name = _excelApp.Names.Item(i);
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
                // OISR-22828: Cancel used to be a no-op here - it hid the overlay while the
                // request (and, for Download outputs, the sheet write) carried on. It now
                // cancels BusyToken, which every long call below is given.
                var cts = new CancellationTokenSource();
                _busyCts = cts;
                await ShowBusyAction.Invoke(message, () =>
                {
                    if (!cts.IsCancellationRequested)
                        cts.Cancel();
                    return Task.CompletedTask;
                });
            }
        }

        private CancellationTokenSource _busyCts;
        private CancellationToken BusyToken => _busyCts?.Token ?? CancellationToken.None;
        private bool UserCancelled => _busyCts?.IsCancellationRequested == true;

        private async Task ShowCancelledMessageAsync(string task)
        {
            await HideBusyAsync();
            await ShowInfoMessageAsync($"{task} task cancelled.");
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
