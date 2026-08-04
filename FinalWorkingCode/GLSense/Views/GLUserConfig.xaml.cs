#nullable enable
using GLSense.Helpers;
using GLSense.Models;
using GLSense.Utilities;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Linq;
using System.Text.Json;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Forms;

namespace GLSense.Views
{
    /// <summary>
    /// Interaction logic for GLUserConfig.xaml
    /// </summary>
    public partial class GLUserConfig : DpiAwareWindow
    {
        private string? _dataOption;
        private bool? _suppressZeroes;
        private int? _refreshCells;
        private int? _recordsPerPage;
        private bool? _validateCube;
        private bool _runSubLedgerDrilldownAsJob;
        private bool _runBalDrilldownAsJob;
        private bool _runTotalDrilldownAsJob;
        private bool _runJournalDrilldownAsJob;
        private bool _includeManualJournal;
        private bool? _overwriteDrilldownMetadata;
        private CancellationHelper? _activeCancellation;
        private string _baselineDataOption = string.Empty;
        private bool _baselineSuppressZeroes;
        private int _baselineRefreshCells = 100;
        private int _baselineRecordsPerPage = 100;
        private bool _baselineValidateCube;
        private bool _baselineRunSubLedgerDrilldownAsJob;
        private bool _baselineRunBalDrilldownAsJob;
        private bool _baselineRunTotalDrilldownAsJob;
        private bool _baselineRunJournalDrilldownAsJob;
        private bool _baselineIncludeManualJournal;
        private bool _baselineOverwriteDrilldownMetadata;
        private readonly string balanceDrilldownName = "Balance Drilldown";
        private readonly string journalDrilldownName = "Journal Drilldown";
        private readonly string subLedgerDrilldownName = "SubLedger Drilldown";
        private readonly string unifiedDrilldownName = "Unified Drilldown";

        public ObservableCollection<DrillDownOption> DrillDowns { get; set; } = new ObservableCollection<DrillDownOption>();
        private readonly List<OptionItem> _options = new()
        {
            new OptionItem { Value = "#Missing" },
            new OptionItem { Value = "#Blank" },
            new OptionItem { Value = "#Zero" },
            new OptionItem { Value = "#Hyphen" }
        };

        public GLUserConfig()
        {
            LogUtility.LogDebug("GLUserConfig.ctor invoked");
            InitializeComponent();
            EnhancedDragDropHelper.EnableWindowDrag(this);
            DataContext = this;

            CmbOptions.SelectionCommitted += CmbOptions_SelectionCommitted;
            HookComboTextChanges();

            InitializeDrillDownGrid();

            Loaded += Window_Loaded;
        }
        private void InitializeDrillDownGrid()
        {

            DrillDowns = new ObservableCollection<DrillDownOption>();

            DrillDowns.Add(new DrillDownOption
            {
                Name = balanceDrilldownName,
                RunAsJob = false,
                CanEditRunAsJob = true,
                IncludeManualJournal = false,
                CanEditManualJournals = false,
                ShowManualJournalsColumn = false
            });

            DrillDowns.Add(new DrillDownOption
            {
                Name = journalDrilldownName,
                RunAsJob = false,
                CanEditRunAsJob = true,
                IncludeManualJournal = false,
                CanEditManualJournals = false,
                ShowManualJournalsColumn = false
            });

            DrillDowns.Add(new DrillDownOption
            {
                Name = subLedgerDrilldownName,
                RunAsJob = false,
                CanEditRunAsJob = true,
                IncludeManualJournal = false,
                CanEditManualJournals = true,
                ShowManualJournalsColumn = true
            });

            if (!IsViewBasedCube()) //show unified drilldown only for non-view based cubes as view based cubes have limitations that prevent running unified drilldown as a job
            {
                DrillDowns.Add(new DrillDownOption
                {
                    Name = unifiedDrilldownName,
                    RunAsJob = false,
                    CanEditRunAsJob = true,  // Can be edited when visible
                    IncludeManualJournal = false,
                    CanEditManualJournals = false,
                    ShowManualJournalsColumn = false
                });
            }

            dgDrillDowns.ItemsSource = DrillDowns;

            UserConfig.DrillDownSettings = DrillDowns.ToList();
        }

        private static bool IsViewBasedCube()
        {
            return (AppState.Instance.SelectedCube?.ViewBased ?? false)
                || string.Equals(AppState.Instance.SelectedCube?.ErpType, "EBS", StringComparison.OrdinalIgnoreCase);
        }

        private void ApplyUnifiedDrilldownRestrictions()
        {
            var totalDrilldown = DrillDowns.FirstOrDefault(d => d.Name == unifiedDrilldownName);
            if (totalDrilldown == null)
            {
                return;
            }

            bool isViewBased = IsViewBasedCube();
            totalDrilldown.CanEditRunAsJob = !isViewBased;

            if (isViewBased)
            {
                _runTotalDrilldownAsJob = false;
                totalDrilldown.RunAsJob = false;
            }
            else
            {
                totalDrilldown.RunAsJob = _runTotalDrilldownAsJob;
            }

            UserConfig.Unified_RunAsJob = _runTotalDrilldownAsJob;
        }
        private void CmbOptions_Loaded(object sender, RoutedEventArgs e)
        {
            EnsureOptionsSource();
            SetOptionsComboFromState();
            RefreshComboDisplay();
        }

        private void EnsureOptionsSource()
        {
            if (CmbOptions.ItemsSource == null)
            {
                CmbOptions.ItemsSource = _options;
                LogUtility.LogDebug("CmbOptions initialized on tab selection");
            }
        }
        private async void Window_Loaded(object sender, RoutedEventArgs e)
        {
            LogUtility.LogDebug("GLUserConfig.Window_Loaded invoked");
            _activeCancellation?.Cancel();
            _activeCancellation = null;

            using var cts = new CancellationHelper();
            _activeCancellation = cts;

            try
            {
                await LoadPreferencesAsync(cts);
            }
            finally
            {
                if (_activeCancellation == cts)
                {
                    _activeCancellation = null;
                }
            }
        }

        // Extracted from Window_Loaded so other handlers (e.g. Reset) can reload the
        // preferences from server and update the UI in the same way.
        private async Task LoadPreferencesAsync(CancellationHelper cts)
        {
            LogUtility.LogDebug("GLUserConfig.LoadPreferencesAsync invoked");
            try
            {
                await ShowBusyOverlayAsync(cts, "Loading user preferences...");

                await Task.Delay(100, cts.GetToken());

                LogUtility.LogDebug("GLUserConfig.LoadPreferencesAsync: calling GetUserPreferences");
                var (isSuccess, message) = await GetUserPreferences(cts.GetToken());
                LogUtility.LogDebug($"GLUserConfig.LoadPreferencesAsync: GetUserPreferences returned isSuccess={isSuccess}");

                if (!isSuccess)
                {
                    await HandleLoadFailureAsync(message);
                    LoadUserPreferences();
                    return;
                }

                var userConfigs = ParseUserConfigs(message);

                if (userConfigs == null)
                {
                    LogUtility.LogDebug("GLUserConfig.LoadPreferencesAsync: ParseUserConfigs returned null, treating as parse failure");
                    await HandleParseFailureAsync(message);
                    return;
                }

                await HideBusyAndShowErrorAsync(string.Empty);
                await ApplyUserPreferences(userConfigs);

                LoadUserPreferences();
                LogUtility.LogDebug("GLUserConfig.LoadPreferencesAsync: preferences loaded and applied successfully");
            }
            catch (OperationCanceledException)
            {
                LogUtility.LogWarn("Window load cancelled by user");
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex, "GLUserConfig.LoadPreferencesAsync");
                await HideBusyAndShowErrorAsync(ex.Message);
            }
        }
        private static UserConfigResponse? ParseUserConfigs(string message)
        {
            if (string.IsNullOrWhiteSpace(message))
            {
                LogUtility.LogWarn("Empty response received.");
                return null;
            }

            try
            {
                var result =
                    ApiResponseHelper.Parse<UserConfigResponse>(
                        message,
                        JsonGlobals.Options);

                if (result.IsSuccess && result.Value != null)
                {
                    return result.Value;
                }

                // Fallback: handle raw payloads that already contain preferences/status
                var fallback = JsonSerializer.Deserialize<UserConfigResponse>(message, JsonGlobals.Options);
                if (fallback != null)
                {
                    if (string.IsNullOrWhiteSpace(fallback.Status))
                    {
                        fallback.Status = "success";
                    }
                    return fallback;
                }

                LogUtility.LogWarn(
                    $"API returned non-success: {result.ErrorMessage}");
                return null;
            }
            catch (JsonException jex)
            {
                LogUtility.LogWarn(
                    $"JSON parse failed: {jex.Message}\nRaw: {message}");
                LogUtility.LogRawJson("GLUserConfig.ParseUserConfigs", message);
                return null;
            }
            catch (Exception ex)
            {
                LogUtility.LogException(
                    ex,
                    "Unexpected error in ParseUserConfigs");
                LogUtility.LogRawJson("GLUserConfig.ParseUserConfigs", message);
                return null;
            }
        }

        private async Task ApplyUserPreferences(UserConfigResponse firstRecord)
        {
            LogUtility.LogDebug("GLUserConfig.ApplyUserPreferences invoked");
            if (firstRecord == null || firstRecord.Preferences == null)
            {
                LogUtility.LogDebug("GLUserConfig.ApplyUserPreferences: validation failed - firstRecord or Preferences is null, nothing to apply");
                return;
            }

            _dataOption = firstRecord.Preferences.DataOption ?? "#Blank";
            _suppressZeroes = firstRecord.Preferences.SupressZeroBalDrilldown ?? false;
            _refreshCells = firstRecord.Preferences.RefreshCells ?? 100;
            _validateCube = firstRecord.Preferences.ValidateCube ?? false;
            _recordsPerPage = firstRecord.Preferences.RecordsPerPage ?? 100;
            _runBalDrilldownAsJob = firstRecord.Preferences.RunBalDrilldownAsJob ?? false;
            _runJournalDrilldownAsJob = firstRecord.Preferences.RunJournalDrilldownAsJob ?? false;
            _runSubLedgerDrilldownAsJob = firstRecord.Preferences.RunSubLedgerDrilldownAsJob ?? false;
            _runTotalDrilldownAsJob = firstRecord.Preferences.RunTotalDrilldownAsJob ?? false;
            _includeManualJournal = firstRecord.Preferences.IncludeManualJournal ?? false;
            _overwriteDrilldownMetadata = firstRecord.Preferences.OverwriteDrilldownMetadata ?? false;

            UserConfig.DataOption = _dataOption ?? string.Empty;
            UserConfig.SupressZeroBalDrilldown = _suppressZeroes ?? false;
            UserConfig.RefreshCells = _refreshCells ?? UserConfig.RefreshCells;
            UserConfig.RecordsPerPage = _recordsPerPage ?? UserConfig.RecordsPerPage;
            UserConfig.ValidateCube = _validateCube ?? false;
            UserConfig.Balance_RunAsJob = _runBalDrilldownAsJob;
            UserConfig.Journal_RunAsJob = _runJournalDrilldownAsJob;
            UserConfig.SubLedger_RunAsJob = _runSubLedgerDrilldownAsJob;
            UserConfig.SubLedger_Manual_Journal = _includeManualJournal;
            UserConfig.Unified_RunAsJob = _runTotalDrilldownAsJob;
            UserConfig.OverwriteDrilldownMetadata = _overwriteDrilldownMetadata ?? false;
            UserConfig.DrillDownSettings = DrillDowns.ToList();

            await Dispatcher.InvokeAsync(() =>
            {
                SetOptionsComboFromState();
                ChkZeroes.IsChecked = _suppressZeroes;
                SpinnerRefreshCells.Value = _refreshCells ?? 100;
                SpinnerRecordsPerPage.Value = _recordsPerPage ?? 100;
                ChkValidateCube.IsChecked = _validateCube;
                ChkOverwriteDrilldownMetadata.IsChecked = _overwriteDrilldownMetadata;
                var balanceDrilldown = DrillDowns.FirstOrDefault(d => d.Name == balanceDrilldownName);
                if (balanceDrilldown != null)
                {
                    balanceDrilldown.RunAsJob = _runBalDrilldownAsJob;
                }
                var journalDrilldown = DrillDowns.FirstOrDefault(d => d.Name == journalDrilldownName);
                if (journalDrilldown != null)
                {
                    journalDrilldown.RunAsJob = _runJournalDrilldownAsJob;
                }
                var subLedgerDrilldown = DrillDowns.FirstOrDefault(d => d.Name == subLedgerDrilldownName);
                if (subLedgerDrilldown != null)
                {
                    subLedgerDrilldown.RunAsJob = _runSubLedgerDrilldownAsJob;
                    subLedgerDrilldown.IncludeManualJournal = _includeManualJournal;
                }

                if (!IsViewBasedCube())
                {
                    ApplyUnifiedDrilldownRestrictions();
                }

                dgDrillDowns.Items.Refresh();
                UserConfig.DrillDownSettings = DrillDowns.ToList();
            });
        }

        private async Task HandleLoadFailureAsync(string message)
        {
            LogUtility.LogWarn($"Failed to load/read user preferences. {message}");
            await Dispatcher.InvokeAsync(async () =>
            {
                await AppOverlayControl.HideBusyAsync();
                await AppOverlayControl.ShowErrorAsync("Failed to load user preferences.");
            });
        }

        private async Task HandleParseFailureAsync(string message)
        {
            LogUtility.LogDebug($"GLUserConfig.HandleParseFailureAsync invoked - message={message}");
            var displayMsg = "Invalid user preferences data.";

            if (!string.IsNullOrWhiteSpace(message))
            {
                displayMsg = message;
            }

            await Dispatcher.InvokeAsync(async () =>
            {
                await AppOverlayControl.HideBusyAsync();
                await AppOverlayControl.ShowErrorAsync($"Failed to load preferences: {displayMsg}");
            });
        }

        private async Task HideBusyAndShowErrorAsync(string errorMsg)
        {
            await Dispatcher.InvokeAsync(async () =>
            {
                await AppOverlayControl.HideBusyAsync();
                if (!string.IsNullOrWhiteSpace(errorMsg))
                {
                    await AppOverlayControl.ShowErrorAsync(errorMsg);
                }
            });
        }
        private async Task HideBusyAndShowSuccessAsync(string successMsg)
        {
            await Dispatcher.InvokeAsync(async () =>
            {
                await AppOverlayControl.HideBusyAsync();
                if (!string.IsNullOrWhiteSpace(successMsg))
                {
                    await AppOverlayControl.ShowSuccessAsync(successMsg);
                }
            });
        }
        private async Task HideBusyAndShowWarnAsync(string warnMsg)
        {
            await Dispatcher.InvokeAsync(async () =>
            {
                await AppOverlayControl.HideBusyAsync();
                if (!string.IsNullOrWhiteSpace(warnMsg))
                {
                    await AppOverlayControl.ShowWarningAsync(warnMsg);
                }
            });
        }
        private async Task HideBusyAndShowInfoAsync(string infoMsg)
        {
            await Dispatcher.InvokeAsync(async () =>
            {
                await AppOverlayControl.HideBusyAsync();
                if (!string.IsNullOrWhiteSpace(infoMsg))
                {
                    await AppOverlayControl.ShowInfoAsync(infoMsg);
                }
            });
        }
        private void CmbOptions_SelectionCommitted(object obj)
        {
            if (obj is OptionItem opt)
            {
                UpdateDataOption(opt.Value);
            }
            else
            {
                UpdateDataOption(obj?.ToString());
            }
        }

        private void ChkZeroes_Click(object sender, RoutedEventArgs e)
        {
            _suppressZeroes = ChkZeroes.IsChecked ?? false;
        }
        private async Task ShowBusyOverlayAsync(CancellationHelper helper, string message)
        {
            await Dispatcher.InvokeAsync(() =>
            {
                AppOverlayControl.ShowBusyasyn(
                    message: message,
                    cancelAction: async () =>
                    {
                        if (!helper.IsCancellationRequested)
                        {
                            helper.Cancel();
                            LogUtility.LogWarn($"Operation cancelled by user: {message}");
                        }

                        await Task.Delay(80);
                    }
                );
            });
        }
        private static async Task<(bool success, string result)> GetUserPreferences(
            CancellationToken ct)
        {
            string apiUrl =
                $"{AppState.Instance.LoginUrl}{AppConstants.RestSecure}user-config/get" +
                $"?cubeId={AppState.Instance.SelectedCube.CubeId}";

            LogUtility.LogDebug($"GLUserConfig.GetUserPreferences invoked - apiUrl={apiUrl}");
            try
            {
                string response =
                    await ApiHelper.ServerAPI(apiUrl, "Form", "", "GET", ct);
                LogUtility.LogDebug("GLUserConfig.GetUserPreferences: ApiHelper.ServerAPI call completed");

                ValidateTransportResponse(response);

                var parsed = ApiResponseHelper.Parse<UserConfigResponse>(response, JsonGlobals.Options);

                if (!parsed.IsSuccess)
                {
                    LogUtility.LogDebug($"GLUserConfig.GetUserPreferences: parse failed - {parsed.ErrorMessage}");
                    return (false, parsed.ErrorMessage ?? "Failed to fetch preferences.");
                }

                return (true, response);
            }
            catch (OperationCanceledException exOp)
            {
                LogUtility.LogWarn(exOp.Message);
                return (false, "User cancelled operation");
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex, "GLUserConfig.GetUserPreferences");
                return (false, ex.Message);
            }
        }
        private static void ValidateTransportResponse(string response)
        {
            if (string.IsNullOrWhiteSpace(response))
            {
                throw new InvalidOperationException("Empty API response.");
            }

            if (response.IndexOf("(401)", StringComparison.OrdinalIgnoreCase) >= 0)
            {
                throw new UnauthorizedAccessException("Session expired.");
            }

            if (response.StartsWith("Error", StringComparison.OrdinalIgnoreCase) ||
                response.IndexOf("<!DOCTYPE", StringComparison.OrdinalIgnoreCase) >= 0)
            {
                throw new InvalidOperationException(response);
            }
        }
        private async void CmdSave_Click(object sender, RoutedEventArgs e)
        {
            LogUtility.LogDebug("GLUserConfig.CmdSave_Click invoked");
            _activeCancellation?.Cancel();
            _activeCancellation = null;

            using var cts = new CancellationHelper();
            _activeCancellation = cts;

            try
            {
                if (cts.IsCancellationRequested)
                {
                    LogUtility.LogDebug("GLUserConfig.CmdSave_Click: cancellation already requested, aborting");
                    return;
                }

                if (!ValidateSpinnerInputs())
                {
                    LogUtility.LogDebug("GLUserConfig.CmdSave_Click: validation failed - invalid Range Cells or Records per page value");
                    await HideBusyAndShowWarnAsync("Please enter valid values for Range Cells and Records per page.");
                    return;
                }

                CaptureCurrentPreferencesFromUi();

                if (!HasPreferenceChanges())
                {
                    LogUtility.LogDebug("GLUserConfig.CmdSave_Click: no preference changes detected, skipping save");
                    await HideBusyAndShowInfoAsync("No changes to save.");
                    return;
                }

                UserConfig.RefreshCells = SpinnerRefreshCells.Value;
                UserConfig.RecordsPerPage = SpinnerRecordsPerPage.Value;
                UserConfig.DataOption = _dataOption ?? string.Empty;
                UserConfig.SupressZeroBalDrilldown = _suppressZeroes ?? false;
                UserConfig.ValidateCube = _validateCube ?? false;
                UserConfig.Balance_RunAsJob = _runBalDrilldownAsJob;
                UserConfig.Journal_RunAsJob = _runJournalDrilldownAsJob;
                UserConfig.SubLedger_RunAsJob = _runSubLedgerDrilldownAsJob;
                UserConfig.SubLedger_Manual_Journal = _includeManualJournal;
                UserConfig.Unified_RunAsJob = _runTotalDrilldownAsJob;
                UserConfig.OverwriteDrilldownMetadata = _overwriteDrilldownMetadata ?? false;
                UserConfig.DrillDownSettings = DrillDowns.ToList();

                await ShowBusyOverlayAsync(cts, "Saving user preferences...");

                var payloadObject = new
                {
                    cubeId = AppState.Instance.SelectedCube.CubeId,
                    preferences = new
                    {
                        validateCube = (_validateCube ?? false).ToString().ToLowerInvariant(),
                        supressZeroBalDrilldown = (_suppressZeroes ?? false).ToString().ToLowerInvariant(),
                        runSubLedgerDrilldownAsJob = _runSubLedgerDrilldownAsJob.ToString().ToLowerInvariant(),
                        runBalDrilldownAsJob = _runBalDrilldownAsJob.ToString().ToLowerInvariant(),
                        runTotalDrilldownAsJob = _runTotalDrilldownAsJob.ToString().ToLowerInvariant(),
                        recordsPerPage = (_recordsPerPage ?? 100).ToString(),
                        refreshCells = (_refreshCells ?? 100).ToString(),
                        runJournalDrilldownAsJob = _runJournalDrilldownAsJob.ToString().ToLowerInvariant(),
                        dataOption = _dataOption ?? string.Empty,
                        includeManualJournal = _includeManualJournal.ToString().ToLowerInvariant(),
                        overwriteDrilldownMetadata = (_overwriteDrilldownMetadata ?? false).ToString().ToLowerInvariant()
                    }
                };

                string payload =
                    JsonSerializer.Serialize(payloadObject, JsonGlobals.Options);

                string apiUrl =
                    $"{AppState.Instance.LoginUrl}{AppConstants.RestSecure}user-config/save?cubeId={AppState.Instance.SelectedCube.CubeId}";

                LogUtility.LogDebug(apiUrl);
                LogUtility.LogDebug($"Saving user config: {payload}");

                string response =
                    await ApiHelper.ServerAPI(
                        apiUrl,
                        "JSON",
                        payload,
                        "POST",
                        cts.GetToken());
                LogUtility.LogDebug("GLUserConfig.CmdSave_Click: ApiHelper.ServerAPI call completed");

                if (string.IsNullOrWhiteSpace(response))
                {
                    LogUtility.LogDebug("GLUserConfig.CmdSave_Click: no response received from server");
                    await HideBusyAndShowErrorAsync("No response from server.");
                    return;
                }

                ValidateTransportResponse(response);

                var parsed =
                    ApiResponseHelper.Parse<JsonElement>(response, JsonGlobals.Options);

                if (!parsed.IsSuccess)
                {
                    LogUtility.LogDebug($"GLUserConfig.CmdSave_Click: save failed - {parsed.ErrorMessage}");
                    await HideBusyAndShowWarnAsync(
                        parsed.ErrorMessage ?? "Save failed.");
                    return;
                }

                string message = ExtractMessage(parsed.Value)
                                 ?? "Saved successfully.";

                LogUtility.LogDebug("GLUserConfig.CmdSave_Click: save completed successfully");
                await HideBusyAndShowSuccessAsync(message);
                CaptureCurrentPreferencesAsBaseline();
            }
            catch (OperationCanceledException)
            {
                LogUtility.LogWarn("GLUserConfig.CmdSave_Click: save operation cancelled by user");
                await HideBusyAndShowWarnAsync("Save operation cancelled.");
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex, "GLUserConfig.CmdSave_Click");
                await HideBusyAndShowErrorAsync($"Save failed: {ex.Message}");
            }
            finally
            {
                if (_activeCancellation == cts)
                {
                    _activeCancellation = null;
                }
            }
        }
        private static readonly string[] _messageKeys = new[]
        {
            "msg",
            "message",
            "detail"
        };

        private static string? ExtractMessage(JsonElement root)
        {
            var matchingProp = root.EnumerateObject()
                .FirstOrDefault(prop => _messageKeys.Contains(prop.Name, StringComparer.OrdinalIgnoreCase));

            return !matchingProp.Equals(default) ? matchingProp.Value.GetString() : null;
        }
        private async void CmdReset_Click(object sender, RoutedEventArgs e)
        {
            LogUtility.LogDebug("GLUserConfig.CmdReset_Click invoked");
            _activeCancellation?.Cancel();
            _activeCancellation = null;

            using var cts = new CancellationHelper();
            _activeCancellation = cts;

            string resultString = string.Empty;
            try
            {
                if (cts.IsCancellationRequested)
                {
                    LogUtility.LogDebug("GLUserConfig.CmdReset_Click: cancellation already requested, aborting");
                    return;
                }

                await ShowBusyOverlayAsync(cts, "Resetting user preferences...");

                string apiUrl = AppState.Instance.LoginUrl
                    + $"{AppConstants.RestSecure}user-config/reset?cubeId={AppState.Instance.SelectedCube.CubeId}";

                LogUtility.LogDebug($"GLUserConfig.CmdReset_Click: calling reset API - {apiUrl}");
                resultString = await ApiHelper.ServerAPI(
                    apiUrl,
                    "Form",
                    "",
                    "GET",
                    cts.GetToken()
                );
                LogUtility.LogDebug("GLUserConfig.CmdReset_Click: ApiHelper.ServerAPI call completed");

                if (string.IsNullOrWhiteSpace(resultString) || !resultString.Contains("status"))
                {
                    LogUtility.LogDebug("GLUserConfig.CmdReset_Click: reset failed - empty or malformed response");
                    await HideBusyAndShowErrorAsync("Failed to reset user preferences." + Environment.NewLine + resultString);
                    return;
                }

                var resetResponse = JsonSerializer.Deserialize<UserConfigResetResponse>(
                    resultString,
                    JsonGlobals.Options
                );

                if (resetResponse?.status == AppConstants.Success)
                {
                    LogUtility.LogDebug("GLUserConfig.CmdReset_Click: reset successful, reloading preferences from server");
                    // After a successful reset on the server, reload preferences from server
                    await LoadPreferencesAsync(cts);
                    await HideBusyAndShowSuccessAsync(resetResponse.msg ?? "Reset successful.");
                }
                else
                {
                    string errorMsg = resetResponse?.msg
                        ?? (resultString.Contains("message") ? resetResponse?.message : null)
                        ?? "Failed to reset user preferences.";
                    LogUtility.LogDebug($"GLUserConfig.CmdReset_Click: reset failed - {errorMsg}");
                    await HideBusyAndShowErrorAsync(errorMsg);
                }
            }
            catch (OperationCanceledException)
            {
                LogUtility.LogWarn("GLUserConfig.CmdReset_Click: reset operation cancelled by user");
                await HideBusyAndShowWarnAsync("Reset cancelled.");
            }
            catch (JsonException ex)
            {
                LogUtility.LogException(ex, "Failed to parse reset user preferences response");
                LogUtility.LogRawJson("GLUserConfig.ResetUserPreferences", resultString);

                string apiMessage = string.Empty;

                if (resultString.IndexOf("<!doctype html>", StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    apiMessage = "Invalid response format. Received HTML instead of JSON.";
                }
                else if (resultString.IndexOf("InternalServerError", StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    apiMessage = "Server encountered an error. Please try again later.";
                }
                else
                {
                    apiMessage = "Received non-JSON response from server.";
                }

                await HideBusyAndShowErrorAsync($"Reset failed: {apiMessage}");
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex, "GLUserConfig.CmdReset_Click");
                LogUtility.LogRawJson("GLUserConfig.ResetUserPreferences", resultString);
                await HideBusyAndShowErrorAsync($"Reset failed: {ex.Message}");
            }
            finally
            {
                if (_activeCancellation == cts)
                {
                    _activeCancellation = null;
                }
            }
        }
        private void BtnClose_Click(object sender, RoutedEventArgs e)
        {
            LogUtility.LogDebug("GLUserConfig.BtnClose_Click invoked");
            Close();
        }

        private void ChkZeroes_Checked(object sender, RoutedEventArgs e)
        {
            _suppressZeroes = true;
        }

        private void ChkZeroes_Unchecked(object sender, RoutedEventArgs e)
        {
            _suppressZeroes = false;
        }
        private void ChkValidateCube_Checked(object sender, RoutedEventArgs e)
        {
            _validateCube = true;
        }

        private void ChkValidateCube_Unchecked(object sender, RoutedEventArgs e)
        {
            _validateCube = false;
        }
        private void ChkOverwriteDrilldownMetadata_Checked(object sender, RoutedEventArgs e)
        {
            _overwriteDrilldownMetadata = true;
        }

        private void ChkOverwriteDrilldownMetadata_Unchecked(object sender, RoutedEventArgs e)
        {
            _overwriteDrilldownMetadata = false;
        }
        private void SpinnerRefreshCells_ValueChanged(object sender, RoutedPropertyChangedEventArgs<int> e)
        {
            _refreshCells = e.NewValue;
        }

        private void SpinnerRecordsPerPage_ValueChanged(object sender, RoutedPropertyChangedEventArgs<int> e)
        {
            _recordsPerPage = e.NewValue;
        }
        private void LoadUserPreferences()
        {
            _dataOption = UserConfig.DataOption;
            _suppressZeroes = UserConfig.SupressZeroBalDrilldown;
            _refreshCells = UserConfig.RefreshCells;
            _recordsPerPage = UserConfig.RecordsPerPage;
            _validateCube = UserConfig.ValidateCube;
            _runBalDrilldownAsJob = UserConfig.Balance_RunAsJob;
            _runJournalDrilldownAsJob = UserConfig.Journal_RunAsJob;
            _runSubLedgerDrilldownAsJob = UserConfig.SubLedger_RunAsJob;
            _includeManualJournal = UserConfig.SubLedger_Manual_Journal;
            _runTotalDrilldownAsJob = UserConfig.Unified_RunAsJob;
            _overwriteDrilldownMetadata = UserConfig.OverwriteDrilldownMetadata;

            SpinnerRefreshCells.Value = _refreshCells ?? 100;
            SpinnerRecordsPerPage.Value = _recordsPerPage ?? 100;
            SetOptionsComboFromState();
            ChkZeroes.IsChecked = _suppressZeroes;
            ChkValidateCube.IsChecked = _validateCube;
            ChkOverwriteDrilldownMetadata.IsChecked = _overwriteDrilldownMetadata;

            var balanceDrilldown = DrillDowns.FirstOrDefault(d => d.Name == balanceDrilldownName);
            if (balanceDrilldown != null)
            {
                balanceDrilldown.RunAsJob = _runBalDrilldownAsJob;
            }
            var journalDrilldown = DrillDowns.FirstOrDefault(d => d.Name == journalDrilldownName);
            if (journalDrilldown != null)
            {
                journalDrilldown.RunAsJob = _runJournalDrilldownAsJob;
            }
            var subLedgerDrilldown = DrillDowns.FirstOrDefault(d => d.Name == subLedgerDrilldownName);
            if (subLedgerDrilldown != null)
            {
                subLedgerDrilldown.RunAsJob = _runSubLedgerDrilldownAsJob;
                subLedgerDrilldown.IncludeManualJournal = _includeManualJournal;
            }

            if (!IsViewBasedCube())
            {
                var totalDrilldown = DrillDowns.FirstOrDefault(d => d.Name == unifiedDrilldownName);
                if (totalDrilldown != null)
                {
                    totalDrilldown.RunAsJob = _runTotalDrilldownAsJob;
                }
            }

            dgDrillDowns.Items.Refresh();
            UserConfig.DrillDownSettings = DrillDowns.ToList();
            CaptureCurrentPreferencesAsBaseline();
        }

        private void ApplyComboOption(string? option)
        {
            if (string.IsNullOrWhiteSpace(option))
            {
                _dataOption = string.Empty;
                UserConfig.DataOption = string.Empty;
                return;
            }

            EnsureOptionsSource();
            var match = _options.FirstOrDefault(o => string.Equals(o.Value, option, StringComparison.OrdinalIgnoreCase));
            if (match != null)
            {
                CmbOptions.SelectedItem = match;
            }
            CmbOptions.Text = option;

            UpdateDataOption(option);
            RefreshComboDisplay();
        }

        private void SetOptionsComboFromState()
        {
            var option = !string.IsNullOrWhiteSpace(UserConfig.DataOption)
                ? UserConfig.DataOption
                : _dataOption;

            if (!string.IsNullOrWhiteSpace(option))
            {
                ApplyComboOption(option);
            }
        }

        private void CmbOptions_LostFocus(object sender, RoutedEventArgs e)
        {
            UpdateDataOption(CmbOptions.Text);
        }

        private void HookComboTextChanges()
        {
            DependencyPropertyDescriptor.FromProperty(Controls.SuggestAppendComboBox.TextProperty, typeof(Controls.SuggestAppendComboBox))
                ?.AddValueChanged(CmbOptions, (s, e) => UpdateDataOption(CmbOptions.Text));
            CmbOptions.LostFocus += CmbOptions_LostFocus;
        }

        private void UpdateDataOption(string? value)
        {
            var text = value ?? string.Empty;
            _dataOption = text;
        }

        private void RefreshComboDisplay()
        {
            Dispatcher.InvokeAsync(() =>
            {
                CmbOptions.ApplyTemplate();
                CmbOptions.UpdateLayout();
                if (!string.IsNullOrEmpty(_dataOption))
                {
                    CmbOptions.Text = _dataOption;
                    CmbOptions.SelectedItem = _options.FirstOrDefault(o => string.Equals(o.Value, _dataOption, StringComparison.OrdinalIgnoreCase));
                }
            }, System.Windows.Threading.DispatcherPriority.Loaded);
        }

        private void MainTabControl_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (sender is System.Windows.Controls.TabControl tab && tab.SelectedItem is TabItem selected && selected.Header is string header && string.Equals(header, "Options", StringComparison.OrdinalIgnoreCase))
            {
                SetOptionsComboFromState();
                RefreshComboDisplay();
            }
        }

        private sealed class OptionItem
        {
            public string Value { get; set; } = string.Empty;
        }

        private void ApplyDefaultPreferences()
        {
            _dataOption = string.Empty;
            _suppressZeroes = false;
            _refreshCells = 100;
            _recordsPerPage = 100;
            _validateCube = false;
            _runBalDrilldownAsJob = false;
            _runJournalDrilldownAsJob = false;
            _runSubLedgerDrilldownAsJob = false;
            _includeManualJournal = false;
            _runTotalDrilldownAsJob = false;
            _overwriteDrilldownMetadata = false;

            CmbOptions.Text = _dataOption;
            ChkZeroes.IsChecked = _suppressZeroes;
            ChkValidateCube.IsChecked = _validateCube;
            ChkOverwriteDrilldownMetadata.IsChecked = _overwriteDrilldownMetadata;
            SpinnerRefreshCells.Value = _refreshCells ?? 100;
            SpinnerRecordsPerPage.Value = _recordsPerPage ?? 100;

            var balanceDrilldown = DrillDowns.FirstOrDefault(d => d.Name == balanceDrilldownName);
            if (balanceDrilldown != null)
            {
                balanceDrilldown.RunAsJob = _runBalDrilldownAsJob;
            }

            var journalDrilldown = DrillDowns.FirstOrDefault(d => d.Name == journalDrilldownName);
            if (journalDrilldown != null)
            {
                journalDrilldown.RunAsJob = _runJournalDrilldownAsJob;
            }

            var subLedgerDrilldown = DrillDowns.FirstOrDefault(d => d.Name == subLedgerDrilldownName);
            if (subLedgerDrilldown != null)
            {
                subLedgerDrilldown.RunAsJob = _runSubLedgerDrilldownAsJob;
                subLedgerDrilldown.IncludeManualJournal = _includeManualJournal;
            }

            if (!IsViewBasedCube())
            {
                var totalDrilldown = DrillDowns.FirstOrDefault(d => d.Name == unifiedDrilldownName);
                if (totalDrilldown != null)
                {
                    totalDrilldown.RunAsJob = _runTotalDrilldownAsJob;
                }
            }

            dgDrillDowns.Items.Refresh();

            UserConfig.DataOption = _dataOption;
            UserConfig.SupressZeroBalDrilldown = _suppressZeroes ?? false;
            UserConfig.RefreshCells = _refreshCells ?? UserConfig.RefreshCells;
            UserConfig.RecordsPerPage = _recordsPerPage ?? UserConfig.RecordsPerPage;
            UserConfig.ValidateCube = _validateCube ?? false;
            UserConfig.Balance_RunAsJob = _runBalDrilldownAsJob;
            UserConfig.Journal_RunAsJob = _runJournalDrilldownAsJob;
            UserConfig.SubLedger_RunAsJob = _runSubLedgerDrilldownAsJob;
            UserConfig.SubLedger_Manual_Journal = _includeManualJournal;
            UserConfig.Unified_RunAsJob = _runTotalDrilldownAsJob;
            UserConfig.OverwriteDrilldownMetadata = _overwriteDrilldownMetadata ?? false;
            UserConfig.DrillDownSettings = DrillDowns.ToList();
            CaptureCurrentPreferencesAsBaseline();
        }

        private void CaptureCurrentPreferencesFromUi()
        {
            _refreshCells = SpinnerRefreshCells.Value;
            _recordsPerPage = SpinnerRecordsPerPage.Value;
            _suppressZeroes = ChkZeroes.IsChecked ?? false;
            _validateCube = ChkValidateCube.IsChecked ?? false;
            _overwriteDrilldownMetadata = ChkOverwriteDrilldownMetadata.IsChecked ?? false;

            if (CmbOptions.SelectedItem is OptionItem opt)
            {
                _dataOption = opt.Value;
            }
            else
            {
                _dataOption = CmbOptions.Text;
            }

            SyncDrilldownSelections();
        }

        private void CaptureCurrentPreferencesAsBaseline()
        {
            _baselineDataOption = _dataOption ?? string.Empty;
            _baselineSuppressZeroes = _suppressZeroes ?? false;
            _baselineRefreshCells = _refreshCells ?? 100;
            _baselineRecordsPerPage = _recordsPerPage ?? 100;
            _baselineValidateCube = _validateCube ?? false;
            _baselineRunBalDrilldownAsJob = _runBalDrilldownAsJob;
            _baselineRunJournalDrilldownAsJob = _runJournalDrilldownAsJob;
            _baselineRunSubLedgerDrilldownAsJob = _runSubLedgerDrilldownAsJob;
            _baselineIncludeManualJournal = _includeManualJournal;
            _baselineRunTotalDrilldownAsJob = _runTotalDrilldownAsJob;
            _baselineOverwriteDrilldownMetadata = _overwriteDrilldownMetadata ?? false;
        }

        private bool HasPreferenceChanges()
        {
            return !string.Equals(_baselineDataOption, _dataOption ?? string.Empty, StringComparison.Ordinal)
                || _baselineSuppressZeroes != (_suppressZeroes ?? false)
                || _baselineRefreshCells != (_refreshCells ?? 100)
                || _baselineRecordsPerPage != (_recordsPerPage ?? 100)
                || _baselineValidateCube != (_validateCube ?? false)
                || _baselineRunBalDrilldownAsJob != _runBalDrilldownAsJob
                || _baselineRunJournalDrilldownAsJob != _runJournalDrilldownAsJob
                || _baselineRunSubLedgerDrilldownAsJob != _runSubLedgerDrilldownAsJob
                || _baselineIncludeManualJournal != _includeManualJournal
                || _baselineRunTotalDrilldownAsJob != _runTotalDrilldownAsJob
                || _baselineOverwriteDrilldownMetadata != (_overwriteDrilldownMetadata ?? false);
        }

        private void SyncDrilldownSelections()
        {
            // Ensure any in-progress edits are committed before reading values
            dgDrillDowns.CommitEdit(System.Windows.Controls.DataGridEditingUnit.Cell, true);
            dgDrillDowns.CommitEdit(System.Windows.Controls.DataGridEditingUnit.Row, true);

            var balanceDrilldown = DrillDowns.FirstOrDefault(d => d.Name == balanceDrilldownName);
            if (balanceDrilldown != null)
            {
                _runBalDrilldownAsJob = balanceDrilldown.RunAsJob;
            }

            var journalDrilldown = DrillDowns.FirstOrDefault(d => d.Name == journalDrilldownName);
            if (journalDrilldown != null)
            {
                _runJournalDrilldownAsJob = journalDrilldown.RunAsJob;
            }

            var subLedgerDrilldown = DrillDowns.FirstOrDefault(d => d.Name == subLedgerDrilldownName);
            if (subLedgerDrilldown != null)
            {
                _runSubLedgerDrilldownAsJob = subLedgerDrilldown.RunAsJob;
                _includeManualJournal = subLedgerDrilldown.IncludeManualJournal;
            }

            var totalDrilldown = DrillDowns.FirstOrDefault(d => d.Name == unifiedDrilldownName);
            if (totalDrilldown != null)
            {
                _runTotalDrilldownAsJob = totalDrilldown.RunAsJob;
            }
        }

        private bool ValidateSpinnerInputs()
        {
            if (!SpinnerRefreshCells.ValidateText(true))
            {
                LogUtility.LogDebug("GLUserConfig.ValidateSpinnerInputs: validation failed - SpinnerRefreshCells text is invalid");
                return false;
            }

            if (!SpinnerRecordsPerPage.ValidateText(true))
            {
                LogUtility.LogDebug("GLUserConfig.ValidateSpinnerInputs: validation failed - SpinnerRecordsPerPage text is invalid");
                return false;
            }

            return true;
        }
    }
#nullable restore
}
