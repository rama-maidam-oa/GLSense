// GLCubeDetails.xaml.cs in GLSense.Addin.Core
// Port of GLSense\Views\GLCubeDetails.xaml.cs (FinalWorkingCode) - the cube/ledger
// selection window opened by the RibGetCube ribbon button. This is where a bare
// successful login (Group A / PartialLoggedIn) actually becomes "LoggedIn": FinalizeLogin
// sets AppState.Instance.IsLoginCompleted = true once a cube+ledger has been committed
// via BtnOK_Click.
//
// Adjustments made when porting into this project's architecture (see PORTING_GUIDE.md
// for the general rules referenced below):
//   - Base class DpiAwareWindow -> DpiAwareWindow (same as GLLogin/GLWaitWindow). DpiAwareWindow
//     already centers/modals against the Excel owner, so AddinEntry's ShowCubeDetails()
//     just does `new GLCubeDetails { CenterInExcel = true, ModalToExcel = true }.ShowDialog()`.
//   - EnhancedDragDropHelper.EnableWindowDrag(this) -> the dedicated
//     TitleBar_MouseLeftButtonDown handler already present on every other window in this
//     project (see GLLogin.xaml.cs for the identical pattern).
//   - LogUtility.* (static) -> ServiceLocator.Logger.*.
//   - AddinModule.CurrentInstance.Ribledger / RibSegS / RibGetCube (direct ADX control
//     access) -> ServiceLocator.RibbonController.SetComboItems/ClearComboItems/
//     SetComboText/SetControlLabel. This is exactly the gap that motivated extending
//     IRibbonController for this group - see GLSense.Contracts\IRibbonController.cs and
//     GLSense\RibbonController.cs (host) for the reflection-based ADXRibbonItem
//     population this replaces.
//   - AddinModule.RibbonHelper.ApplyState("LoggedIn") -> ServiceLocator.RibbonController?.SetState("LoggedIn").
//   - new GLSense.Repositories.DataRepository() -> GLSense.Addin.Core.Repositories.DataRepository
//     (right-sized port: only GetLedgers/GetSegments are needed here - see that file's
//     header comment for what was deliberately left out).
//   - ConfiguratorRelaunch() (Group H, resolved): the original re-launched the Balance
//     Configurator task pane directly (AddinModule.CurrentInstance.GetPaneInstance()/
//     GLConfiguratorPane.RelaunchPane()). GLConfiguratorPane is a host-only ADX/WinForms
//     construct that Addin.Core can't reference, so this now goes through
//     IRibbonController.RelaunchConfiguratorPaneIfVisible() (the host does the Visible
//     check) - see GLSense.Addin.Core\Views\ConfiguratorPaneHost.cs for the full
//     HWND-reparenting bridge this and RibFSG_OnClick both rely on.
using GLSense.Addin.Core.Helpers;
using GLSense.Addin.Core.Infrastructure;
using GLSense.Addin.Core.Models;
using GLSense.Addin.Core.Repositories;
using MahApps.Metro.IconPacks;
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
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Threading;
using GLSense.Addin.Core.Utilities;

namespace GLSense.Addin.Core.Views
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
            InitializeComponent();

            DataContext = this;

            dgCubes.Columns[0].Visibility = Visibility.Collapsed;

            // "Ledger Name" (index 1) fills any left-over width instead of leaving a blank
            // gap now that every column is Width="Auto" (see DataGridColumnFillHelper for why
            // the star-width columns were removed).
            DataGridColumnFillHelper.EnableFillColumn(dgCubes, dgCubes.Columns[1]);

            cmbCubes.SelectionCommitted += async (_) => await CmbCubes_SelectionCommitted();
            cmbCubes.InvalidSelection += (invalidText) =>
            {
                AppOverlayControl.ShowWarning($"Invalid cube name: '{invalidText}'. Please select a valid one.");
                _currentCube = null;
                UpdateValidationControls(null);
            };

            UpdateValidationControls(null);
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

        private void UpdateValidationControls(CubeRecord? cube)
        {
            bool hasCube = cube != null;
            bool isViewBased = cube?.ViewBased == true;
            bool canValidate = hasCube && !isViewBased;

            chkValidateCube.IsEnabled = canValidate;
            btnValidateCube.IsEnabled = canValidate;
            chkValidateCube.IsChecked = canValidate && Utilities.UserConfig.ValidateCube;
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
            cmbCubes.ItemsSource = CubeCache.AllCubes;
            cmbCubes.DisplayMemberPath = "CubeName";
        }

        private async void Window_Loaded(object sender, RoutedEventArgs e)
        {
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
                        ServiceLocator.Logger?.LogException(ex);
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
                ServiceLocator.Logger?.LogException(ex);
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
                            ServiceLocator.Logger?.LogWarn($"Operation cancelled by user: {fullMessage}");
                        }
                        await Task.Delay(80); // small feedback delay
                                              // HideBusyAsync() is already called in the overlay's handler
                    }
                );
            }, DispatcherPriority.Background);
        }

        private async Task CmbCubes_SelectionCommitted()
        {
            if (cmbCubes.SelectedItem is not CubeRecord selected) return;

            HideStatus();

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
                ServiceLocator.Logger?.LogException(ex);
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
                ServiceLocator.Logger?.LogWarn("Cube data loading cancelled by user");
            }
            catch (Exception ex)
            {
                ServiceLocator.Logger?.LogException(ex);
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

                // DpiAwareWindow.OnLoaded's SizeToContent resettle (see CLAUDE.md section 1)
                // runs synchronously as soon as this window's Loaded event fires - which
                // is BEFORE this method's caller chain (Window_Loaded -> ...
                // -> LoadCubeData -> UpdateGridAsync, several awaits deep) has actually
                // populated dgCubes with any rows. That resettle therefore always measures
                // an empty grid, producing the reported "gap until a cube is selected"
                // symptom regardless of whether a cube was pre-selected on open - manually
                // picking a cube afterwards just happens to run this same UpdateGridAsync
                // while the window is already visible, which WPF's normal live-content
                // layout handles correctly without needing the toggle trick. Now that the
                // grid actually has real row data and its own layout has settled, resettle
                // again so the window grows to the correct height immediately instead of
                // waiting for the user to touch the cube combo.
                ForceSizeToContentResettle();
                PumpDispatcherFrame();

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
            string apiUrl = $"{AppState.Instance.LoginUrl}/rest/secure/finance/cube-refreshed-date?cubeId={cubeId}";
            string response = string.Empty;
            try
            {
                response = await ApiHelper.ServerAPI(apiUrl, "Form", "", "GET", token);

                if (string.IsNullOrEmpty(response) || response.StartsWith("Error")) return null;

                return JsonSerializer.Deserialize<CubeLedgerResponse>(response, JsonGlobals.Options);
            }
            catch (JsonException ex)
            {
                ServiceLocator.Logger?.LogException(ex, $"Invalid JSON received from cube refresh date API: {apiUrl}");
                ServiceLocator.Logger?.LogRawJson("GLCubeDetails.FetchCubeLedgerResponse", response);
                return null;
            }
            catch (Exception ex)
            {
                ServiceLocator.Logger?.LogException(ex);
                if (!string.IsNullOrWhiteSpace(response))
                {
                    ServiceLocator.Logger?.LogRawJson("GLCubeDetails.FetchCubeLedgerResponse", response);
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
                    mergedLedger.LastRefreshedDate = GLSenseFormatUTCDate(ledger, "last refreshed date", refreshData.LastRefreshedDateInUTC);
                    mergedLedger.ADMRefreshedDate = GLSenseFormatUTCDate(ledger, "last refreshed adaptive memory date", refreshData.LastRefreshedAdaptiveMemInUTC);
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

            ServiceLocator.Logger?.LogDebug($"{label} for ledger \"{ledger.LedgerName}\": {dateStr}");
            string[] formats =
            {
                "yyyy-MM-ddTHH:mm:ss.fffffff",
                "yyyy-MM-ddTHH:mm:ss",
                "dd-MM-yyyy HH:mm:ss",
                "yyyy-MM-dd HH:mm:ss",
                "yyyy/MM/dd HH:mm:ss",
                "dd/MM/yyyy HH:mm:ss"
            };

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

            ServiceLocator.Logger?.LogError($"Date parsing failed for: {dateStr}");
            return string.Empty;
        }

        private void BtnClose_Click(object sender, RoutedEventArgs e)
        {
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
            string apiUrl = $"{AppState.Instance.LoginUrl}/rest/secure/finance/cube-dimension-status?cubeId={cubeId}";
            string outputString = string.Empty;
            try
            {
                ServiceLocator.Logger?.LogDebug($"API for getting cube validation data : {apiUrl}");
                outputString = await ApiHelper.ServerAPI(apiUrl, "Form", "", "GET", token);
                ServiceLocator.Logger?.LogDebug($"Response received : {outputString}");

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
                ServiceLocator.Logger?.LogException(ex, "User cancelled the operation");
            }
            catch (JsonException jsonEx)
            {
                ServiceLocator.Logger?.LogException(jsonEx, $"Invalid JSON received from cube validation API: {apiUrl}");
                ServiceLocator.Logger?.LogRawJson("GLCubeDetails.CubeDataValidation", outputString);
            }
            catch (Exception ex)
            {
                ServiceLocator.Logger?.LogException(ex);
                ServiceLocator.Logger?.LogRawJson("GLCubeDetails.CubeDataValidation", outputString);
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
            if (_currentCube == null || _currentCube.ViewBased)
            {
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
                ServiceLocator.Logger?.LogWarn("Cube validation cancelled by user");
            }
            catch (Exception ex)
            {
                ServiceLocator.Logger?.LogException(ex);
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
            if (_currentCube == null || _currentCube.ViewBased)
            {
                chkValidateCube.IsChecked = false;
                return;
            }

            Utilities.UserConfig.ValidateCube = chkValidateCube.IsChecked ?? false;
        }

        private async void BtnOK_Click(object sender, RoutedEventArgs e)
        {
            if (_currentCube == null)
            {
                await AppOverlayControl.ShowWarningAsync("Please select a cube before continuing.");
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
                    await AppOverlayControl.ShowWarningAsync("Please select a ledger before continuing.");
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

                            if (proceed != true)
                                return;
                        }
                    }
                    catch (OperationCanceledException)
                    {
                        ServiceLocator.Logger?.LogWarn("Cube validation cancelled by user");
                    }
                    finally
                    {
                        await AppOverlayControl.HideBusyAsync();
                    }
                }
                else if (!_currentCube.ViewBased && cubeValidation != null && cubeValidation.IsValidated && !cubeValidation.IsInSync)
                {
                    var proceed = await AppOverlayControl.ShowConfirmAsync("The cube data is not sync with the source. Do you want to proceed with cube selection?");

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
                await AppOverlayControl.HideBusyAsync();
                await AppOverlayControl.ShowErrorAsync("Operation was canceled by user.");
            }
            catch (Exception ex)
            {
                await AppOverlayControl.HideBusyAsync();
                await AppOverlayControl.ShowErrorAsync($"Error while selecting cube: {ex.Message}");
                ServiceLocator.Logger?.LogException(ex);
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
                ServiceLocator.Logger?.LogWarn("Cube validation cancelled by user");
                throw;
            }
            catch (Exception ex)
            {
                ServiceLocator.Logger?.LogException(ex);
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
            try
            {
                string apiUrl =
                    $"{AppState.Instance.LoginUrl}/rest/secure/finance/user-config/get?cubeId={cubeId}";

                string response = await ApiHelper.ServerAPI(apiUrl, "Form", string.Empty, "GET", ct);

                var parsed = ApiResponseHelper.Parse<UserConfigResponse>(response, JsonGlobals.Options);

                if (!parsed.IsSuccess || parsed.Value?.Preferences == null)
                {
                    return;
                }

                var prefs = parsed.Value.Preferences;
                Utilities.UserConfig.DataOption = prefs.DataOption ?? string.Empty;
                Utilities.UserConfig.SupressZeroBalDrilldown = prefs.SupressZeroBalDrilldown ?? false;
                Utilities.UserConfig.RefreshCells = prefs.RefreshCells ?? Utilities.UserConfig.RefreshCells;
                Utilities.UserConfig.RecordsPerPage = prefs.RecordsPerPage ?? Utilities.UserConfig.RecordsPerPage;
                Utilities.UserConfig.ValidateCube = prefs.ValidateCube ?? false;
                Utilities.UserConfig.Balance_RunAsJob = prefs.RunBalDrilldownAsJob ?? false;
                Utilities.UserConfig.Journal_RunAsJob = prefs.RunJournalDrilldownAsJob ?? false;
                Utilities.UserConfig.SubLedger_RunAsJob = prefs.RunSubLedgerDrilldownAsJob ?? false;
                Utilities.UserConfig.SubLedger_Manual_Journal = prefs.IncludeManualJournal ?? false;
                Utilities.UserConfig.Unified_RunAsJob = prefs.RunTotalDrilldownAsJob ?? false;
            }
            catch (OperationCanceledException)
            {
                throw;
            }
            catch (Exception ex)
            {
                ServiceLocator.Logger?.LogException(ex, "Error loading user preferences for cube selection");
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
            try
            {
                token.ThrowIfCancellationRequested();

                await Utilities.CommonFunctions.FillResponsibilitiesAsync(ledger.LedgerId, cube.CubeId, token);

                // Commit the new selection to AppState BEFORE touching ANY ribbon control -
                // including LoadCubeLedgers below, not just UpdateRibbonForCube afterwards
                // (BtnOK_Click used to commit only after this method returned). Excel's
                // WorkbookActivate handler (AddinEntry.SyncRibbonSelectionWithAppState)
                // re-populates the same ribbon controls (RibGetCube's caption AND
                // Ribledger's item list) from whatever AppState.Instance.SelectedCube/
                // SelectedLedger currently hold, and can fire while this async method is
                // still mid-flight (e.g. from focus changes caused by the busy overlay).
                // LoadCubeLedgers itself only marshals its SetComboItems("Ribledger", ...)
                // call onto the dispatcher at Background priority - a low, preemptible
                // priority - so a mid-flight SyncRibbonSelectionWithAppState firing before
                // AppState is updated would read the OLD cube's (possibly single-ledger)
                // list and silently overwrite the new cube's full ledger list right after
                // LoadCubeLedgers sets it - exactly the "ribbon ledger dropdown only shows
                // one value" symptom. Safe to commit here: every awaited step that could
                // still fail has already completed without throwing.
                AppState.Instance.SelectedCube = cube;
                AppState.Instance.SelectedLedger = ledger;

                await LoadCubeLedgers(cube, ledger.LedgerName, token);

                FinalizeLogin();
                UpdateRibbonForCube(cube);

                return new OperationResult { IsSuccess = true };
            }
            catch (OperationCanceledException)
            {
                throw;
            }
            catch (InvalidOperationException ex)
            {
                return new OperationResult { IsSuccess = false, Message = ex.Message, Exception = ex };
            }
            catch (UnauthorizedAccessException ex)
            {
                return new OperationResult { IsSuccess = false, Message = $"An unexpected error occurred while loading data. {ex.Message}", Exception = ex };
            }
            catch (Exception ex)
            {
                ServiceLocator.Logger?.LogException(ex);
                return new OperationResult { IsSuccess = false, Message = "An unexpected error occurred while loading data.", Exception = ex };
            }
        }

        private async Task<OperationResult> ProcessCubeSelectionReload(CubeRecord cube, LedgerModel ledger, CancellationHelper cts)
        {
            try
            {
                cts.GetToken().ThrowIfCancellationRequested();
                if (!SameChartofAccounts())
                {
                    await AppOverlayControl.HideBusyAsync();
                    string quest = "Selected chart of account is different from the previous one.\n" +
                                   "This will wipe out the entire data of the sheet!\n" +
                                   "Do you wish to continue with the data cleanup? Press 'Cancel' to escape cube change.";

                    var userAction = await AppOverlayControl.ShowConfirmAsync(quest);
                    if (!userAction.HasValue) return new OperationResult { IsSuccess = false };

                    if (userAction.Value)
                    {
                        Utilities.CommonMethods.Clear_Sheet();
                    }

                    await ShowBusyOverlayAsync(cts, "Please wait while we fetch the ledger data...");
                }

                await Utilities.CommonFunctions.FillResponsibilitiesAsync(ledger.LedgerId, cube.CubeId, cts.GetToken());

                // See the identical comment in ProcessCubeSelectionNew - commit to AppState
                // before touching ANY ribbon control, including LoadCubeLedgers below, so a
                // concurrently-firing WorkbookActivate sync can't read a stale cube/ledger
                // and overwrite the new cube's full Ribledger item list with the old cube's.
                AppState.Instance.SelectedCube = cube;
                AppState.Instance.SelectedLedger = ledger;

                await LoadCubeLedgers(cube, ledger.LedgerName, cts.GetToken());

                ServiceLocator.RibbonController?.SetComboText("Ribledger", ledger.LedgerName);
                FinalizeLogin();
                UpdateRibbonForCube(cube);

                return new OperationResult { IsSuccess = true };
            }
            catch (OperationCanceledException)
            {
                throw;
            }
            catch (InvalidOperationException ex)
            {
                return new OperationResult { IsSuccess = false, Message = ex.Message, Exception = ex };
            }
            catch (Exception ex)
            {
                ServiceLocator.Logger?.LogException(ex);
                return new OperationResult { IsSuccess = false, Message = ex.Message, Exception = ex };
            }
        }

        /// <summary>
        /// Group H (resolved): the original relaunched the Balance Configurator task pane
        /// here (AddinModule.CurrentInstance.GetPaneInstance() -> GLConfiguratorPane.
        /// RelaunchPane() if it was visible). GLConfiguratorPane is a host-only ADX/
        /// WinForms construct, so this goes through IRibbonController.
        /// RelaunchConfiguratorPaneIfVisible() (the host does the Visible check) instead
        /// of reaching into AddinModule directly.
        /// </summary>
        private static void ConfiguratorRelaunch()
        {
            try
            {
                ServiceLocator.RibbonController?.RelaunchConfiguratorPaneIfVisible();
            }
            catch (Exception ex)
            {
                ServiceLocator.Logger?.LogException(ex, "ConfiguratorRelaunch failed");
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
            ServiceLocator.RibbonController?.SetState("LoggedIn");
        }

        private async Task LoadCubeLedgers(CubeRecord cube, string selectedLedgerName, CancellationToken token)
        {
            var selectedRecord = CubeCache.AllCubes?.FirstOrDefault(c => c.CubeId == cube.CubeId)
                ?? throw new InvalidOperationException("Cube or ledgers not found for the given cube.");

            if (selectedRecord.Ledgers == null || selectedRecord.Ledgers.Count == 0)
            {
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
                token.ThrowIfCancellationRequested();

                ServiceLocator.Logger?.LogDebug($"LoadCubeLedgers: cubeId={cube.CubeId}, ledgerNames.Count={ledgerNames.Count}, selectedLedgerName={selectedLedgerName}");
                ServiceLocator.RibbonController?.SetComboItems("Ribledger", ledgerNames);

                if (!string.IsNullOrWhiteSpace(selectedLedgerName) &&
                    ledgerNames.Contains(selectedLedgerName, StringComparer.OrdinalIgnoreCase))
                {
                    ServiceLocator.RibbonController?.SetComboText("Ribledger", selectedLedgerName);
                }

                // Blanket _ribbon.Invalidate() (fired later by SetState("LoggedIn")) reliably
                // refreshes getEnabled/getText for most controls, but a dropdown's cached
                // item list (getItemCount/getItemLabel) is a separate, more static piece of
                // Ribbon state - some RibbonX hosts only re-pull it when THIS control is
                // invalidated specifically, not on a whole-ribbon invalidate. Belt-and-
                // suspenders: force it here too, right where the items were actually set,
                // instead of relying solely on the later blanket invalidate.
                ServiceLocator.RibbonController?.Invalidate("Ribledger");
            }, DispatcherPriority.Background);
        }

        private void UpdateRibbonForCube(CubeRecord cube)
        {
            ServiceLocator.RibbonController?.SetControlLabel("RibGetCube", cube != null ? "Cube : " + cube.CubeName : "Cube : Select Cube");

            ServiceLocator.RibbonController?.ClearComboItems("RibSegS");

            var repository = new DataRepository();
            if (_selectedLedger != null && cube != null)
            {
                var segs = repository.GetSegments(cube.CubeId, _selectedLedger.LedgerId);
                // .ToList() matters here, not just style: SetComboItems is a cross-AppDomain
                // IRibbonController call (this project's host/hot-reload split), and the
                // lazy WhereSelectEnumerableIterator .Select() produces on its own isn't
                // [Serializable] - passing it directly throws a SerializationException
                // during remoting argument marshaling, before SetComboItems even runs.
                ServiceLocator.RibbonController?.SetComboItems("RibSegS", segs.Select(s => s.SegmentName).ToList());
                ServiceLocator.RibbonController?.Invalidate("RibSegS");
            }
        }
    }
#nullable disable
}
