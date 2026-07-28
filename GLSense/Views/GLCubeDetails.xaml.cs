using AddinExpress.MSO;
using ControlzEx.Standard;
using GLSense.Helpers;
using GLSense.Models;
using GLSense.Repositories;
using GLSense.Utilities;
using MahApps.Metro.IconPacks;
using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Globalization;
using System.Linq;
using System.Text.Json;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using System.Windows.Threading;

namespace GLSense.Views
{
#nullable enable
    /// <summary>
    /// Interaction logic for GLCubeDetails.xaml
    /// </summary>
    public partial class GLCubeDetails : DpiAwareWindow
    {
        private enum ValidationStatus
        {
            Success,
            Failed,
            NotValidated
        }
        private CancellationHelper? _activeCancellation;

        private CubeRecord? _currentCube;
        private LedgerModel? _selectedLedger;

        public GLCubeDetails()
        {
            LogUtility.LogDebug("GLCubeDetails.ctor invoked");
            InitializeComponent();
            EnhancedDragDropHelper.EnableWindowDrag(this);


            DataContext = this;

            dgCubes.Columns[0].Visibility = Visibility.Collapsed;

            cmbCubes.SelectionCommitted += async (_) => await CmbCubes_SelectionCommitted();
            cmbCubes.InvalidSelection += (invalidText) =>
            {
                AppOverlayControl.ShowWarning($"Invalid cube name: '{invalidText}'. Please select a valid one.");
                _currentCube = null;
                UpdateValidationControls(null);
            };

            UpdateValidationControls(null);
        }

        private void UpdateValidationControls(CubeRecord? cube)
        {
            bool hasCube = cube != null;
            bool isViewBased = cube?.ViewBased == true;
            bool canValidate = hasCube && !isViewBased;

            chkValidateCube.IsEnabled = canValidate;
            btnValidateCube.IsEnabled = canValidate;
            chkValidateCube.IsChecked = canValidate && UserConfig.ValidateCube;
        }

        private static CubeValidationResult? GetValidation(CubeRecord cube)
        {
            if (cube == null) return null;

            return CubeCache.Validations.TryGetValue(cube.CubeId, out var validation)
                ? validation
                : null;
        }
        private void TryShowCachedValidationStatus(CubeRecord cube)
        {
            LogUtility.LogDebug($"GLCubeDetails.TryShowCachedValidationStatus invoked - cubeId={cube?.CubeId}");
            if (cube == null)
            {
                LogUtility.LogWarn("GLCubeDetails.TryShowCachedValidationStatus: cube is null, aborting");
                return;
            }

            if (!CubeCache.Validations.TryGetValue(cube.CubeId, out var validation))
            {
                SetValidationStatus(ValidationStatus.NotValidated, "Cube data is not validated");
                return;
            }

            if (!validation.IsValidated)
            {
                SetValidationStatus(ValidationStatus.NotValidated, "Cube data is not validated");
                return;
            }

            if (string.IsNullOrWhiteSpace(validation.Message))
            {
                return;
            }

            string displayMessage = GetValidationDisplayMessage(validation);

            if (validation.IsInSync)
            {
                SetValidationStatus(ValidationStatus.Success, displayMessage);
            }
            else
            {
                SetValidationStatus(ValidationStatus.Failed, displayMessage);
            }
        }

        private void InitializeComboBox()
        {
            LogUtility.LogDebug($"GLCubeDetails.InitializeComboBox invoked - cubeCount={CubeCache.AllCubes?.Count ?? 0}");
            cmbCubes.ItemsSource = CubeCache.AllCubes;
            cmbCubes.DisplayMemberPath = "CubeName";
        }

        private async void Window_Loaded(object sender, RoutedEventArgs e)
        {
            LogUtility.LogDebug("GLCubeDetails.Window_Loaded invoked");
            try
            {
                InitializeComboBox();

                if (AppState.Instance.SelectedCube?.CubeName is { } selectedName &&
                    CubeCache.AllCubes?.FirstOrDefault(c => c.CubeName == selectedName) is { } cube)
                {
                    _currentCube = cube;
                    cmbCubes.SelectedItem = cube;
                    cmbCubes.Text = cube.CubeName;

                    _activeCancellation?.Cancel();
                    _activeCancellation = null;

                    using var cts = new CancellationHelper();
                    _activeCancellation = cts;

                    try
                    {
                        await LoadUserPreferencesForCube(cube.CubeId, cts.GetToken());
                        UpdateValidationControls(cube);

                        if (!cube.ViewBased)
                        {
                            TryShowCachedValidationStatus(cube);
                        }

                        await LoadCubeData(cube, cts);
                    }
                    catch (Exception ex)
                    {
                        LogUtility.LogException(ex);
                    }
                    finally
                    {
                        if (_activeCancellation == cts)
                        {
                            _activeCancellation = null;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                await AppOverlayControl.ShowErrorAsync(ex.Message);
                LogUtility.LogException(ex);
            }
        }
        private async Task ShowBusyOverlayAsync(CancellationHelper helper, string message, string additionalInfo = "")
        {
            await Dispatcher.InvokeAsync(() =>
            {
                string fullMessage = message + " (Click cancel to stop)";
                if (!string.IsNullOrWhiteSpace(additionalInfo))
                {
                    fullMessage += Environment.NewLine + additionalInfo;
                }
                AppOverlayControl.ShowBusyasyn(
                    message: fullMessage,
                    cancelAction: async () =>
                    {
                        if (!helper.IsCancellationRequested)
                        {
                            helper.Cancel();
                            LogUtility.LogWarn($"Operation cancelled by user: {fullMessage}");
                        }
                        await Task.Delay(80); // small feedback delay
                                              // HideBusyAsync() is already called in the overlay's handler
                    }
                );
            }, DispatcherPriority.Background);
        }
        private async Task CmbCubes_SelectionCommitted()
        {
            LogUtility.LogDebug("GLCubeDetails.CmbCubes_SelectionCommitted invoked");

            if (cmbCubes.SelectedItem is not CubeRecord selected)
            {
                LogUtility.LogDebug("GLCubeDetails.CmbCubes_SelectionCommitted: no cube selected, aborting");
                return;
            }

            HideStatus();

            // 1. Stop & clean up any previous (very rare in Loaded, but good habit)
            _activeCancellation?.Cancel();
            _activeCancellation = null;

            using var cts = new CancellationHelper();
            _activeCancellation = cts;

            try
            {
                _currentCube = selected;
                await LoadUserPreferencesForCube(selected.CubeId, cts.GetToken());
                UpdateValidationControls(selected);

                if (!_currentCube.ViewBased)
                {
                    TryShowCachedValidationStatus(_currentCube);
                }

                await LoadCubeData(selected, cts);
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex);
            }
            finally
            {
                if (_activeCancellation == cts)
                {
                    _activeCancellation = null;
                }
            }
        }

        private async Task LoadCubeData(CubeRecord cube, CancellationHelper ctx)
        {
            LogUtility.LogDebug($"GLCubeDetails.LoadCubeData invoked - cubeId={cube?.CubeId}, cubeName={cube?.CubeName}");
            if (cube == null)
            {
                LogUtility.LogWarn("GLCubeDetails.LoadCubeData: cube is null, aborting");
                return;
            }

            try
            {
                await ShowBusyOverlayAsync(ctx, "Loading ledger lists...");

                var token = ctx.GetToken();

                var repository = new DataRepository();
                var dbLedgers = repository.GetLedgers(cube.CubeId);
                var refreshData = await FetchCubeLedgerResponse(cube.CubeId, token);

                if (refreshData is null) return;

                var gridData = MapLedgerData(dbLedgers, refreshData);

                bool hasWarnings = ApplyValidationWarnings(cube.CubeId, gridData);

                await UpdateGridAsync(gridData, hasWarnings);
            }
            catch (TaskCanceledException)
            {
                LogUtility.LogWarn("Cube data loading cancelled by user");
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex);
                await AppOverlayControl.ShowErrorAsync(ex.Message);
            }
            finally
            {
                await AppOverlayControl.HideBusyAsync();
            }
        }

        private static bool ApplyValidationWarnings(long cubeId, List<LedgerModel> gridData)
        {
            if (!CubeCache.Validations.TryGetValue(cubeId, out var validation)) return false;

            bool hasWarnings = false;
            foreach (var row in gridData)
            {
                var ledgerValidation = validation.Ledgers.FirstOrDefault(
                    l => string.Equals(l.LedgerName, row.LedgerName, StringComparison.OrdinalIgnoreCase));

                row.HasWarnings = ledgerValidation != null && !ledgerValidation.IsValid;
                hasWarnings |= row.HasWarnings;
            }
            return hasWarnings;
        }

        private async Task UpdateGridAsync(List<LedgerModel> data, bool hasWarnings = false)
        {
            await Dispatcher.InvokeAsync(async () =>
            {
                dgCubes.ItemsSource = data;

                if (dgCubes.Columns.Count > 0)
                {
                    dgCubes.Columns[0].Visibility = hasWarnings ? Visibility.Visible : Visibility.Collapsed;
                    if (hasWarnings) dgCubes.Columns[0].Width = new DataGridLength(42);
                }

                if (dgCubes.Columns.Count > 1) dgCubes.Columns[1].Width = new DataGridLength(2, DataGridLengthUnitType.Star);
                if (dgCubes.Columns.Count > 2) dgCubes.Columns[2].Width = new DataGridLength(1.5, DataGridLengthUnitType.Star);
                if (dgCubes.Columns.Count > 3) dgCubes.Columns[3].Width = new DataGridLength(3, DataGridLengthUnitType.Star);

                dgCubes.Items.Refresh();

                await DgGridUpdate(data);

            }, DispatcherPriority.Normal);
        }
        private bool changeSelection()
        {
            if (AppState.Instance.SelectedCube != null &&
                _currentCube != null && AppState.Instance.SelectedCube.CubeId == _currentCube.CubeId &&
                AppState.Instance.SelectedLedger != null)
            {
                return true;
            }

            return false;
        }
        private async Task DgGridUpdate(List<LedgerModel> data)
        {
            if (changeSelection())
            {
                var selectedLedger = data.FirstOrDefault(l =>
                    string.Equals(l.LedgerName, AppState.Instance.SelectedLedger.LedgerName, StringComparison.OrdinalIgnoreCase));
                if (selectedLedger != null)
                {
                    await Dispatcher.InvokeAsync(() =>
                    {
                        dgCubes.SelectedItem = selectedLedger;
                        dgCubes.ScrollIntoView(selectedLedger);
                    }, DispatcherPriority.Background);
                }
            }
        }
        private static async Task<CubeLedgerResponse?> FetchCubeLedgerResponse(long cubeId, CancellationToken token)
        {
            string apiUrl = $"{AppState.Instance.LoginUrl}{AppConstants.RestSecure}cube-refreshed-date?cubeId={cubeId}";
            string response = string.Empty;
            try
            {
                LogUtility.LogDebug($"GLCubeDetails.FetchCubeLedgerResponse: calling API {apiUrl}");
                response = await ApiHelper.ServerAPI(apiUrl, "Form", "", "GET", token);

                if (string.IsNullOrEmpty(response) || response.StartsWith("Error"))
                {
                    LogUtility.LogWarn($"GLCubeDetails.FetchCubeLedgerResponse: API {apiUrl} returned empty/error response");
                    return null;
                }

                LogUtility.LogDebug("GLCubeDetails.FetchCubeLedgerResponse: API call succeeded, parsing response");
                return JsonSerializer.Deserialize<CubeLedgerResponse>(response, JsonGlobals.Options);
            }
            catch (JsonException ex)
            {
                LogUtility.LogException(ex, $"Invalid JSON received from cube refresh date API: {apiUrl}");
                LogUtility.LogRawJson("GLCubeDetails.FetchCubeLedgerResponse", response);
                return null;
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex);
                if (!string.IsNullOrWhiteSpace(response))
                {
                    LogUtility.LogRawJson("GLCubeDetails.FetchCubeLedgerResponse", response);
                }
                return null;
            }
        }

        private static List<LedgerModel> MapLedgerData(ObservableCollection<LedgerModel> dbLedgers, CubeLedgerResponse apiResponse)
        {
            var result = new List<LedgerModel>();
            var refreshLookup = apiResponse?.Records?.ToDictionary(r => r.LedgerId) ?? new Dictionary<long, CubeLedgerRecord>();

            foreach (var ledger in dbLedgers)
            {
                var mergedLedger = new LedgerModel
                {
                    LedgerId = ledger.LedgerId,
                    CubeId = ledger.CubeId,
                    LedgerName = ledger.LedgerName,
                    CoaId = ledger.CoaId,
                    PeriodSetName = ledger.PeriodSetName,
                    CurrencyCode = ledger.CurrencyCode,
                    TimeZone = TimeZoneInfo.Local.DisplayName
                };

                if (refreshLookup.TryGetValue(ledger.LedgerId, out var refreshData))
                {
                    mergedLedger.LastRefreshedDate = GLSenseFormatUTCDate(ledger,"last refreshed date",refreshData.LastRefreshedDateInUTC);
                    mergedLedger.ADMRefreshedDate = GLSenseFormatUTCDate(ledger,"last refreshed adaptive memory date",refreshData.LastRefreshedAdaptiveMemInUTC);
                }
                else
                {
                    mergedLedger.LastRefreshedDate = string.Empty;
                    mergedLedger.ADMRefreshedDate = string.Empty;
                }

                result.Add(mergedLedger);
            }

            return result;
        }

        private static string GLSenseFormatUTCDate(LedgerModel ledger, string label, string? dateStr)
        {
            if (string.IsNullOrWhiteSpace(dateStr)) return string.Empty;

            LogUtility.LogDebug($"{label} for ledger \"{ledger.LedgerName}\": {dateStr}");
            string[] formats = [
                "yyyy-MM-ddTHH:mm:ss.fffffff",
                    "yyyy-MM-ddTHH:mm:ss",
                    "dd-MM-yyyy HH:mm:ss",
                    "yyyy-MM-dd HH:mm:ss",
                    "yyyy/MM/dd HH:mm:ss",
                    "dd/MM/yyyy HH:mm:ss"
            ];

            var provider = CultureInfo.InvariantCulture;
            var styles = DateTimeStyles.AssumeUniversal | DateTimeStyles.AdjustToUniversal;

            foreach (var fmt in formats)
            {
                if (DateTime.TryParseExact(dateStr, fmt, provider, styles, out var parsedDate))
                {
                    var localTime = TimeZoneInfo.ConvertTimeFromUtc(parsedDate, TimeZoneInfo.Local);
                    return localTime.ToString("dd-MMM-yyyy HH:mm:ss", provider);
                }
            }

            if (DateTime.TryParse(dateStr, provider, styles, out var parsed))
            {
                var localTime = TimeZoneInfo.ConvertTimeFromUtc(parsed, TimeZoneInfo.Local);
                return localTime.ToString("dd-MMM-yyyy HH:mm:ss", provider);
            }

            LogUtility.LogError($"Date parsing failed for: {dateStr}");
            return string.Empty;
        }

        private void BtnClose_Click(object sender, RoutedEventArgs e)
        {
            LogUtility.LogDebug("GLCubeDetails.BtnClose_Click invoked");
            Close();
        }

        private static List<string> GetFailedLedgersFromValidation(CubeValidationResult validationResult)
        {
            var failed = new List<string>();
            if (validationResult == null || validationResult.Ledgers == null) return failed;

            failed.AddRange(validationResult.Ledgers
                .Where(l => !l.IsValid && !string.IsNullOrEmpty(l.LedgerName))
                .Select(l => l.LedgerName!));

            return failed;
        }

        private bool UpdateGridWithWarnings(List<string> failedLedgers)
        {
            if (dgCubes.ItemsSource is not List<LedgerModel> rows || rows.Count == 0) return false;

            bool hasWarnings = false;
            foreach (var row in rows)
            {
                row.HasWarnings = failedLedgers.Contains(row.LedgerName);
                hasWarnings |= row.HasWarnings;
            }

            dgCubes.Items.Refresh();
            return hasWarnings;
        }

        private async Task UpdateGridColumnsForWarningsAsync()
        {
            await Dispatcher.InvokeAsync(() =>
            {
                if (dgCubes.Columns.Count > 0)
                {
                    dgCubes.Columns[0].Visibility = Visibility.Visible;
                    dgCubes.Columns[0].Width = new DataGridLength(42);
                }

                if (dgCubes.Columns.Count > 1) dgCubes.Columns[1].Width = new DataGridLength(3, DataGridLengthUnitType.Star);
                if (dgCubes.Columns.Count > 2) dgCubes.Columns[2].Width = new DataGridLength(1, DataGridLengthUnitType.Star);
                if (dgCubes.Columns.Count > 3) dgCubes.Columns[3].Width = new DataGridLength(2, DataGridLengthUnitType.Star);

            }, DispatcherPriority.Background);
        }

        private async Task<CubeValidationResult> CubeDataValidation(long cubeId, CancellationToken token)
        {
            string apiUrl = $"{AppState.Instance.LoginUrl}{AppConstants.RestSecure}cube-dimension-status?cubeId={cubeId}";
            string outputString = string.Empty;
            try
            {
                LogUtility.LogDebug($"API for getting cube validation data : {apiUrl}");
                outputString = await ApiHelper.ServerAPI(apiUrl, "Form", "", "GET", token);
                LogUtility.LogDebug($"Response received : {outputString}");

                if (string.IsNullOrWhiteSpace(outputString) || outputString.StartsWith("Error", StringComparison.OrdinalIgnoreCase))
                {
                    return new CubeValidationResult
                    {
                        CubeId = cubeId,
                        IsValidated = false,
                        ErrorMessage = outputString,
                        ValidatedAt = DateTime.Now,
                        Ledgers = new List<LedgerValidationResult>(),
                        Message = string.Empty
                    };
                }

                using JsonDocument doc = JsonDocument.Parse(outputString);
                JsonElement root = doc.RootElement;

                string message = root.TryGetProperty("msg", out var msgElem) ? msgElem.GetString() ?? string.Empty : string.Empty;

                if (!CubeCache.Validations.ContainsKey(cubeId))
                {
                    CubeCache.Validations[cubeId] = new CubeValidationResult { CubeId = cubeId };
                }

                var cubeResult = CubeCache.Validations[cubeId];
                    cubeResult.Message = message;
                    cubeResult.ErrorMessage = string.Empty;
                    cubeResult.IsValidated = true;
                    cubeResult.ValidatedAt = DateTime.Now;
                    cubeResult.Ledgers.Clear();

                if (root.TryGetProperty("records", out var recordsElem) && recordsElem.ValueKind == JsonValueKind.Array)
                {
                    foreach (var record in recordsElem.EnumerateArray())
                    {
                        string? ledgerName = record.TryGetProperty("ledgerName", out var nameElem) ? nameElem.GetString() : null;
                        bool isValid = record.TryGetProperty("status", out var statusElem) && statusElem.ValueKind == JsonValueKind.True;

                        cubeResult.Ledgers.Add(new LedgerValidationResult
                        {
                            LedgerName = ledgerName ?? "(unknown)",
                            IsValid = isValid
                        });
                    }
                }

                return cubeResult;
            }
            catch (OperationCanceledException ex)
            {
                LogUtility.LogException(ex, "User cancelled the operation");
            }
            catch (JsonException jsonEx)
            {
                LogUtility.LogException(jsonEx, $"Invalid JSON received from cube validation API: {apiUrl}");
                LogUtility.LogRawJson("GLCubeDetails.CubeDataValidation", outputString);
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex);
                LogUtility.LogRawJson("GLCubeDetails.CubeDataValidation", outputString);
            }
            return new CubeValidationResult
            {
                CubeId = cubeId,
                IsValidated = false,
                ErrorMessage = "Failed to validate cube data.",
                ValidatedAt = DateTime.Now,
                Ledgers = new List<LedgerValidationResult>(),
                Message = string.Empty
            };
        }
        private void SetValidationStatus(ValidationStatus status, string? message = null)
        {
            var (iconKind, fgColor) = status switch
            {
                ValidationStatus.Success => (PackIconFontAwesomeKind.CircleCheckSolid, "#059669"),
                ValidationStatus.Failed => (PackIconFontAwesomeKind.CircleXmarkSolid, "#DC2626"),
                ValidationStatus.NotValidated => (PackIconFontAwesomeKind.CircleInfoSolid, "#2563EB"),
                _ => throw new ArgumentException("Invalid status")
            };

            statusIcon.Kind = iconKind;
            statusIcon.Foreground = (Brush)new BrushConverter().ConvertFrom(fgColor);
            lblCubeStatus.Text = message ?? GetDefaultMessage(status);
            lblCubeStatus.Foreground = (Brush)new BrushConverter().ConvertFrom(fgColor);
            statusPanel.Visibility = Visibility.Visible;
        }
        private static string GetDefaultMessage(ValidationStatus status)
        {
            return status switch
            {
                ValidationStatus.Success => "Cube validated successfully!",
                ValidationStatus.Failed => "Validation failed!",
                ValidationStatus.NotValidated => "Cube data is not validated",
                _ => "Unknown status"
            };
        }

        private static string GetValidationDisplayMessage(CubeValidationResult validation)
        {
            if (string.IsNullOrWhiteSpace(validation.Message))
            {
                return string.Empty;
            }

            if (!validation.ValidatedAt.HasValue)
            {
                return validation.Message;
            }

            return $"{validation.Message} Validated on : {validation.ValidatedAt.Value.ToString("dd-MMM-yyyy HH:mm:ss", CultureInfo.InvariantCulture)}";
        }
        private void HideStatus()
        {
            statusPanel.Visibility = Visibility.Collapsed;
        }

        private async void BtnValidateCube_Click(object sender, RoutedEventArgs e)
        {
            LogUtility.LogDebug($"GLCubeDetails.BtnValidateCube_Click invoked - cube={_currentCube?.CubeName}, viewBased={_currentCube?.ViewBased}");
            if (_currentCube == null || _currentCube.ViewBased)
            {
                LogUtility.LogDebug("GLCubeDetails.BtnValidateCube_Click: no cube selected or cube is view-based, aborting validation");
                return;
            }

            _activeCancellation?.Cancel();
            _activeCancellation = null;

            using var cts = new CancellationHelper();
            _activeCancellation = cts;

            try
            {

                LedgerModel? selectedLedger = GetSelectedLedger();

                string cubeTitle = _currentCube.CubeName;
                string ledgerTitle = selectedLedger?.LedgerName ?? string.Empty;

                string msg = "Cube: " + cubeTitle + Environment.NewLine + "Ledger: " + ledgerTitle;

                await ShowBusyOverlayAsync(cts, "Validating cube data...", msg);
                await ValidateCubeAsync(_currentCube, cts.GetToken());
                TryShowCachedValidationStatus(_currentCube);
            }
            catch (OperationCanceledException)
            {
                LogUtility.LogWarn("Cube validation cancelled by user");
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex);
            }
            finally
            {
                await AppOverlayControl.HideBusyAsync();
                if (_activeCancellation == cts)
                {
                    _activeCancellation = null;
                }
            }
        }

        private void ChkValidateCube_Changed(object sender, RoutedEventArgs e)
        {
            LogUtility.LogDebug($"GLCubeDetails.ChkValidateCube_Changed invoked - IsChecked={chkValidateCube.IsChecked}");
            if (_currentCube == null || _currentCube.ViewBased)
            {
                chkValidateCube.IsChecked = false;
                return;
            }

            UserConfig.ValidateCube = chkValidateCube.IsChecked ?? false;
        }
        private async void BtnOK_Click(object sender, RoutedEventArgs e)
        {
            LogUtility.LogDebug($"GLCubeDetails.BtnOK_Click invoked - cube={_currentCube?.CubeName}");
            if (_currentCube == null)
            {
                LogUtility.LogDebug("GLCubeDetails.BtnOK_Click: no cube selected, aborting");
                AppOverlayControl.ShowWarning("Please select a cube before continuing.");
                return;
            }

            _activeCancellation?.Cancel();
            _activeCancellation = null;

            using var cts = new CancellationHelper();
            _activeCancellation = cts;
            var token = cts.GetToken();
            OperationResult result = new();

            try
            {
                _selectedLedger = GetSelectedLedger();
                if (_selectedLedger == null)
                {
                    LogUtility.LogDebug("GLCubeDetails.BtnOK_Click: no ledger selected, aborting");
                    AppOverlayControl.ShowWarning("Please select a ledger before continuing.");
                    return;
                }

                var cubeValidation = GetValidation(_currentCube);

                if (chkValidateCube.IsChecked == true && !_currentCube.ViewBased && (cubeValidation == null || !cubeValidation.IsValidated))
                {
                    try
                    {
                        string cubeTitle = _currentCube.CubeName;
                        string ledgerTitle = _selectedLedger.LedgerName;
                        string msg = "Cube: " + cubeTitle + Environment.NewLine + "Ledger: " + ledgerTitle;
                        await ShowBusyOverlayAsync(cts, "Validating cube data...", msg);

                        CubeValidationResult validation = await ValidateCubeAsync(_currentCube, token);

                        TryShowCachedValidationStatus(_currentCube);

                        if (validation.NeedsConfirmation)
                        {
                            await AppOverlayControl.HideBusyAsync();

                            var proceed = await AppOverlayControl.ShowConfirmAsync(
                                "The cube data is not sync with the source. Do you want to proceed with cube selection?");

                            LogUtility.LogDebug($"GLCubeDetails.BtnOK_Click: out-of-sync confirmation dialog result={proceed}");
                            if (proceed != true)
                                return;
                        }
                    }
                    catch (OperationCanceledException)
                    {
                        LogUtility.LogWarn("Cube validation cancelled by user");
                    }
                    finally
                    {
                        await AppOverlayControl.HideBusyAsync();
                    }
                }
                else if (!_currentCube.ViewBased && cubeValidation != null && cubeValidation.IsValidated && !cubeValidation.IsInSync)
                {
                    var proceed = await AppOverlayControl.ShowConfirmAsync("The cube data is not sync with the source. Do you want to proceed with cube selection?");

                    LogUtility.LogDebug($"GLCubeDetails.BtnOK_Click: cached out-of-sync confirmation dialog result={proceed}");
                    if (proceed != true)
                        return;
                }


                await ShowBusyOverlayAsync(cts, "Please wait while we fetch the ledger data...");

                if (AppState.Instance.SelectedCube?.CubeId != _currentCube.CubeId)
                {
                    result = await ProcessCubeSelectionNew(_currentCube, _selectedLedger, token);
                    ConfiguratorRelaunch();
                }
                else if (AppState.Instance.SelectedLedger?.LedgerId != _selectedLedger.LedgerId)
                {
                    result = await ProcessCubeSelectionReload(_currentCube, _selectedLedger, cts);
                    ConfiguratorRelaunch();
                }
                else
                {
                    result.IsSuccess = true;
                }

                if (result.IsSuccess)
                {
                    AppState.Instance.SelectedCube = _currentCube;
                    AppState.Instance.SelectedLedger = _selectedLedger;
                }
            }
            catch (OperationCanceledException)
            {
                LogUtility.LogWarn("GLCubeDetails.BtnOK_Click: cube selection operation was canceled by user");
                await AppOverlayControl.HideBusyAsync();
                await AppOverlayControl.ShowErrorAsync("Operation was canceled by user.");
            }
            catch (Exception ex)
            {
                await AppOverlayControl.HideBusyAsync();
                await AppOverlayControl.ShowErrorAsync($"Error while selecting cube: {ex.Message}");
                LogUtility.LogException(ex, "GLCubeDetails.BtnOK_Click");
            }
            finally
            {
                await AppOverlayControl.HideBusyAsync();

                if (!string.IsNullOrWhiteSpace(AppState.Instance.SelectedCube?.UserName))
                {
                    AppState.Instance.LoginUserName = AppState.Instance.SelectedCube?.UserName;
                }

                if (result.IsSuccess)
                {
                    Close();
                }
                else if (!string.IsNullOrEmpty(result.Message))
                {
                    await AppOverlayControl.ShowErrorAsync(result.Message);
                }
            }
        }

        private async Task<CubeValidationResult> ValidateCubeAsync(CubeRecord cube, CancellationToken token)
        {
            LogUtility.LogDebug($"GLCubeDetails.ValidateCubeAsync invoked - cubeId={cube?.CubeId}, cubeName={cube?.CubeName}");
            if (cube == null)
            {
                LogUtility.LogWarn("GLCubeDetails.ValidateCubeAsync: cube is null, cannot validate");
                return new CubeValidationResult
                {
                    CubeId = 0,
                    ErrorMessage = "Failed to validate cube data.",
                    ValidatedAt = DateTime.Now,
                    Message = string.Empty,
                    IsValidated = true
                };
            }

            try
            {
                CubeValidationResult validationResult = await CubeDataValidation(cube.CubeId, token);

                if (validationResult == null)
                {
                    return new CubeValidationResult
                    {
                        CubeId = cube.CubeId,
                        ErrorMessage = "Failed to validate cube data.",
                        ValidatedAt = DateTime.Now,
                        Message = string.Empty,
                        IsValidated = true
                    };
                }

                token.ThrowIfCancellationRequested();

                var failedLedgers = GetFailedLedgersFromValidation(validationResult);
                token.ThrowIfCancellationRequested();

                bool hasWarnings = UpdateGridWithWarnings(failedLedgers);
                token.ThrowIfCancellationRequested();

                if (hasWarnings)
                {
                    await UpdateGridColumnsForWarningsAsync();
                    token.ThrowIfCancellationRequested();
                }

                return validationResult;
            }
            catch (OperationCanceledException)
            {
                LogUtility.LogWarn("Cube validation cancelled by user");
                throw;
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex);
                return new CubeValidationResult
                {
                    CubeId = cube.CubeId,
                    ValidatedAt = DateTime.Now,
                    ErrorMessage = "Failed to validate cube data.",
                    Message = string.Empty,
                    IsValidated = true
                };
            }
        }
        private static async Task LoadUserPreferencesForCube(long cubeId, CancellationToken ct)
        {
            LogUtility.LogDebug($"GLCubeDetails.LoadUserPreferencesForCube invoked - cubeId={cubeId}");
            try
            {
                string apiUrl =
                    $"{AppState.Instance.LoginUrl}{AppConstants.RestSecure}user-config/get?cubeId={cubeId}";

                LogUtility.LogDebug($"GLCubeDetails.LoadUserPreferencesForCube: calling API {apiUrl}");
                string response = await ApiHelper.ServerAPI(apiUrl, "Form", string.Empty, "GET", ct);

                var parsed = ApiResponseHelper.Parse<UserConfigResponse>(response, JsonGlobals.Options);

                if (!parsed.IsSuccess || parsed.Value?.Preferences == null)
                {
                    LogUtility.LogWarn($"GLCubeDetails.LoadUserPreferencesForCube: failed to parse user preferences response - {response}");
                    return;
                }

                LogUtility.LogDebug("GLCubeDetails.LoadUserPreferencesForCube: preferences loaded successfully");

                var prefs = parsed.Value.Preferences;
                UserConfig.DataOption = prefs.DataOption ?? string.Empty;
                UserConfig.SupressZeroBalDrilldown = prefs.SupressZeroBalDrilldown ?? false;
                UserConfig.RefreshCells = prefs.RefreshCells ?? UserConfig.RefreshCells;
                UserConfig.RecordsPerPage = prefs.RecordsPerPage ?? UserConfig.RecordsPerPage;
                UserConfig.ValidateCube = prefs.ValidateCube ?? false;
                UserConfig.Balance_RunAsJob = prefs.RunBalDrilldownAsJob ?? false;
                UserConfig.Journal_RunAsJob = prefs.RunJournalDrilldownAsJob ?? false;
                UserConfig.SubLedger_RunAsJob = prefs.RunSubLedgerDrilldownAsJob ?? false;
                UserConfig.SubLedger_Manual_Journal = prefs.IncludeManualJournal ?? false;
                UserConfig.Unified_RunAsJob = prefs.RunTotalDrilldownAsJob ?? false;
            }
            catch (OperationCanceledException)
            {
                LogUtility.LogWarn("GLCubeDetails.LoadUserPreferencesForCube: operation was canceled");
                throw;
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex, "Error loading user preferences for cube selection");
            }
        }

        private LedgerModel? GetSelectedLedger()
        {
            if (dgCubes.SelectedItem is LedgerModel selected) return selected;

            if (dgCubes.ItemsSource is IEnumerable<LedgerModel> source && source.Any())
            {
                return source.First();
            }

            return null;
        }

        private async Task<OperationResult> ProcessCubeSelectionNew(CubeRecord cube, LedgerModel ledger, CancellationToken token)
        {
            LogUtility.LogDebug($"GLCubeDetails.ProcessCubeSelectionNew invoked - cubeId={cube?.CubeId}, ledger={ledger?.LedgerName}");
            if (cube == null || ledger == null)
            {
                LogUtility.LogWarn("GLCubeDetails.ProcessCubeSelectionNew: cube or ledger is null, aborting cube selection");
                return new OperationResult { IsSuccess = false, Message = "Cube or ledger selection is invalid." };
            }

            try
            {
                token.ThrowIfCancellationRequested();

                await CommonFunctions.FillResponsibilitiesAsync(ledger.LedgerId, cube.CubeId, token);
                await LoadCubeLedgers(cube, ledger.LedgerName, token);

                FinalizeLogin();
                UpdateRibbonForCube(cube);

                return new OperationResult { IsSuccess = true };
            }
            catch (OperationCanceledException)
            {
                LogUtility.LogWarn("GLCubeDetails.ProcessCubeSelectionNew: operation was canceled");
                throw;
            }
            catch (InvalidOperationException ex)
            {
                LogUtility.LogException(ex, "GLCubeDetails.ProcessCubeSelectionNew");
                return new OperationResult { IsSuccess = false, Message = ex.Message, Exception = ex };
            }
            catch (UnauthorizedAccessException ex)
            {
                LogUtility.LogException(ex, "GLCubeDetails.ProcessCubeSelectionNew");
                return new OperationResult { IsSuccess = false, Message = $"An unexpected error occurred while loading data. {ex.Message}", Exception = ex };
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex, "GLCubeDetails.ProcessCubeSelectionNew");
                return new OperationResult { IsSuccess = false, Message = "An unexpected error occurred while loading data.", Exception = ex };
            }
        }

        private async Task<OperationResult> ProcessCubeSelectionReload(CubeRecord cube, LedgerModel ledger, CancellationHelper cts)
        {
            LogUtility.LogDebug($"GLCubeDetails.ProcessCubeSelectionReload invoked - cubeId={cube?.CubeId}, ledger={ledger?.LedgerName}");
            if (cube == null || ledger == null)
            {
                LogUtility.LogWarn("GLCubeDetails.ProcessCubeSelectionReload: cube or ledger is null, aborting cube selection");
                return new OperationResult { IsSuccess = false, Message = "Cube or ledger selection is invalid." };
            }

            try
            {
                cts.GetToken().ThrowIfCancellationRequested();
                if (!SameChartofAccounts())
                {
                    LogUtility.LogDebug("GLCubeDetails.ProcessCubeSelectionReload: chart of accounts differs from previous, prompting user to confirm data wipe");
                    await AppOverlayControl.HideBusyAsync();
                    string quest = "Selected chart of account is different from the previous one.\n" +
                                   "This will wipe out the entire data of the sheet!\n" +
                                   "Do you wish to continue with the data cleanup? Press 'Cancel' to escape cube change.";

                    var userAction = await AppOverlayControl.ShowConfirmAsync(quest);
                    LogUtility.LogDebug($"GLCubeDetails.ProcessCubeSelectionReload: data-wipe confirmation result={userAction}");
                    if (!userAction.HasValue) return new OperationResult { IsSuccess = false };

                    if (userAction.Value)
                    {
                        CommonMethods.Clear_Sheet();
                    }

                    await ShowBusyOverlayAsync(cts, "Please wait while we fetch the ledger data...");
                }

                await CommonFunctions.FillResponsibilitiesAsync(ledger.LedgerId, cube.CubeId, cts.GetToken());
                await LoadCubeLedgers(cube, ledger.LedgerName, cts.GetToken());

                AddinModule.CurrentInstance.Ribledger.Text = ledger.LedgerName;
                FinalizeLogin();
                UpdateRibbonForCube(cube);

                return new OperationResult { IsSuccess = true };
            }
            catch (OperationCanceledException)
            {
                LogUtility.LogWarn("GLCubeDetails.ProcessCubeSelectionReload: operation was canceled");
                throw;
            }
            catch (InvalidOperationException ex)
            {
                LogUtility.LogException(ex, "GLCubeDetails.ProcessCubeSelectionReload");
                return new OperationResult { IsSuccess = false, Message = ex.Message, Exception = ex };
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex, "GLCubeDetails.ProcessCubeSelectionReload");
                return new OperationResult { IsSuccess = false, Message = ex.Message, Exception = ex };
            }
        }

        private static void ConfiguratorRelaunch()
        {
            AppState.Instance.BalancePane = AddinModule.CurrentInstance.GetPaneInstance();

            if (AppState.Instance.BalancePane != null && AppState.Instance.BalancePane.Visible)
            {
                _ = AppState.Instance.BalancePane.RelaunchPane();
            }
        }

        private bool SameChartofAccounts()
        {
            if (AppState.Instance.SelectedLedger == null) return true;

            return AppState.Instance.SelectedLedger.CoaId == _selectedLedger?.CoaId;
        }

        private static void FinalizeLogin()
        {
            AppState.Instance.IsLoginCompleted = true;
            AddinModule.RibbonHelper.ApplyState("LoggedIn");
        }

        private async Task LoadCubeLedgers(CubeRecord cube, string selectedLedgerName, CancellationToken token)
        {
            LogUtility.LogDebug($"GLCubeDetails.LoadCubeLedgers invoked - cubeId={cube?.CubeId}, selectedLedgerName={selectedLedgerName}");
            if (cube == null)
            {
                LogUtility.LogWarn("GLCubeDetails.LoadCubeLedgers: cube is null, cannot load ledgers");
                throw new InvalidOperationException("Cube is required to load ledgers.");
            }

            var selectedRecord = CubeCache.AllCubes?.FirstOrDefault(c => c.CubeId == cube.CubeId)
                ?? throw new InvalidOperationException("Cube or ledgers not found for the given cube.");

            if (selectedRecord.Ledgers == null || selectedRecord.Ledgers.Count == 0)
            {
                LogUtility.LogWarn($"GLCubeDetails.LoadCubeLedgers: no ledgers available for cubeId={cube.CubeId}");
                throw new InvalidOperationException("No ledgers available for the selected cube.");
            }

            var ledgerNames = await Task.Run(() =>
                selectedRecord.Ledgers
                    .OrderBy(l => l.LedgerName)
                    .Select(l => l.LedgerName)
                    .ToList(),
                token);

            await Dispatcher.InvokeAsync(() =>
            {
                var ribbon = AddinModule.CurrentInstance.Ribledger;
                ribbon.Items.Clear();

                foreach (var name in ledgerNames)
                {
                    token.ThrowIfCancellationRequested();
                    ribbon.Items.Add(new ADXRibbonItem { Caption = name });
                }

                if (!string.IsNullOrWhiteSpace(selectedLedgerName) &&
                    ledgerNames.Contains(selectedLedgerName, StringComparer.OrdinalIgnoreCase))
                {
                    ribbon.Text = selectedLedgerName;
                }
            }, DispatcherPriority.Background);
        }

        private void UpdateRibbonForCube(CubeRecord cube)
        {
            LogUtility.LogDebug($"GLCubeDetails.UpdateRibbonForCube invoked - cube={cube?.CubeName}");
            AddinModule.CurrentInstance.RibGetCube.Caption = cube != null ? "Cube: " + cube.CubeName : "Cube: Select cube";

            AddinModule.CurrentInstance.RibSegS.Items.Clear();
            AddinModule.CurrentInstance.RibSegS.Text = string.Empty;

            var repository = new DataRepository();
            if (_selectedLedger != null && cube != null)
            {
                var segs = repository.GetSegments(cube.CubeId, _selectedLedger.LedgerId);

                foreach (var s in segs)
                {
                    var item = new ADXRibbonItem { Caption = s.SegmentName };
                    AddinModule.CurrentInstance.RibSegS.Items.Add(item);
                }
            }
        }
    }
#nullable disable
}

