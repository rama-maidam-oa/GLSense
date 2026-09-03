// GLLogin.xaml.cs in GLSense.Addin.Core
// Real port of GLSense\Views\GLLogin.xaml.cs (FinalWorkingCode) - replaces the earlier
// placeholder version of this file (which only demoed AppOverlay/toast behavior and had
// no actual login logic). The XAML for this window already had the WebView2 control,
// SuggestAppendComboBox and AppOverlay wired up ahead of this pass; only the code-behind
// needed the real flow.
//
// Adjustments made when porting into this project's architecture (see
// PORTING_GUIDE.md for the general rules referenced below):
//   - Base class DpiAwareWindow -> BaseWindow. BaseWindow already sets the Excel owner
//     and handles DPI/work-area sizing, so there's no SetExcelOwner()/ShowWithOwner()
//     call here - AddinEntry.Login() just calls ShowDialog().
//   - EnhancedDragDropHelper.EnableWindowDrag(this) -> the dedicated
//     TitleBar_MouseLeftButtonDown handler already present on this window (matches the
//     pattern used by every other window in this project).
//   - LogUtility.* (static) -> ServiceLocator.Logger.*.
//   - AppPaths.TempUrlsPath -> ServiceLocator.Paths.UrlsDirectory (this project's
//     IPathProvider - already used this way by the previous version of this file).
//   - AddinModule.CurrentInstance.RibLogin.Enabled = ... / AddinModule.RibbonHelper.
//     ApplyState(...) -> ServiceLocator.RibbonController?.SetControlEnabled("RibLogin", ...)
//     / ServiceLocator.RibbonController?.SetState(...). Addin.Core cannot reference the
//     host's AddinModule/ribbon-designer types at all - IRibbonController is the only
//     door across that boundary, and it already has everything this file needs
//     (SetControlEnabled/SetState), see GLSense\RibbonController.cs.
//   - AddinModule.GetEdgeAddinInstance() -> EdgeAddinHelper.GetEdgeAddinInstance() (new
//     in this project - see Helpers\EdgeAddinHelper.cs - since AddinModule isn't
//     referenceable here either).
//   - CommonFunctions.GLSenseMessage(string, MessageBoxIcon, MessageBoxButtons) (WinForms
//     enums) is gone in this project; nothing in this file actually called it directly
//     (HandleApiError only logged + disabled the ribbon), so no replacement was needed.
//   - IsLoginCompleted is intentionally never set to true anywhere in this file - in the
//     original project that only happens once a cube+ledger is chosen in GLCubeDetails
//     (Group B), confirmed by grepping the original source. Setting it here on a bare
//     successful login (as an earlier draft of the host's AddinEntry.Login() briefly
//     did) would be wrong - the ribbon only reaches "LoggedIn" after Group B runs.
using GLSense.Addin.Core.Helpers;
using GLSense.Addin.Core.Infrastructure;
using GLSense.Addin.Core.Models;
using GLSense.Addin.Core.Repositories;
using GLSense.Addin.Core.Utilities;
using Microsoft.Web.WebView2.Core;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Text;
using System.Text.Json;
using System.Threading;
using System.Threading.Tasks;
using System.Web;
using System.Windows;
using System.Windows.Input;
using System.Xml.Linq;

namespace GLSense.Addin.Core.Views
{
#nullable enable
    /// <summary>
    /// Interaction logic for GLLogin.xaml
    /// </summary>
    public partial class GLLogin : BaseWindow, INotifyPropertyChanged
    {
        private bool _xlEdgePermission;
        private Task? _webViewInitTask;
        private CancellationHelper? _activeCancellation;
        private WebView2NavigationResilience? _resilience;

        // Backing field
        private ServerInfo? _selectedServer;
        public List<ServerInfo>? ServerList { get; set; }

        public ServerInfo? SelectedServer
        {
            get => _selectedServer;
            set
            {
                _selectedServer = value;
                txtServer.Text = value?.Address ?? string.Empty;
                OnPropertyChanged();
            }
        }

        public event PropertyChangedEventHandler? PropertyChanged;
        protected virtual void OnPropertyChanged([CallerMemberName] string? propertyName = null)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }

        public GLLogin()
        {
            InitializeComponent();

            DataContext = this;

            // When user confirms a valid selection
            cmbServer.SelectionCommitted += async (obj) =>
            {
                await CmbServer_SelectionCommitted(obj);
            };

            // When user leaves an invalid entry
            cmbServer.InvalidSelection += async (invalidText) =>
            {
                await Dispatcher.InvokeAsync(() =>
                {
                    AppOverlayControl.ShowWarning($"Invalid server name: '{invalidText}'. Please select a valid one.");
                });

                SelectedServer = null;
                txtServer.Text = string.Empty;
                await NavigateToBlankPageAsync();
            };

            Loaded += GLLogin_Loaded;
            webView.Loaded += WebView_Loaded;
            Closed += GLLogin_Closed;
        }

        // WebView2 spins up a real Chromium browser-process tree (browser + GPU + network
        // service + renderer(s) + crashpad handler) per environment - without disposing the
        // control, that whole tree is orphaned every time this window closes, since nothing
        // else in this app's lifecycle ever tears it down. Confirmed via Task Manager (see
        // FinalWorkingCode's identical fix): dozens of stray msedgewebview2.exe processes
        // accumulate across sessions with this missing. Unsubscribing the CoreWebView2 event
        // handlers first avoids them firing against a control that's mid-disposal.
        private void GLLogin_Closed(object? sender, EventArgs e)
        {
            try
            {
                ServiceLocator.Logger?.LogDebug("GLLogin.GLLogin_Closed: disposing WebView2 control");

                if (webView.CoreWebView2 != null)
                {
                    _resilience?.Detach(webView.CoreWebView2);
                    webView.CoreWebView2.PermissionRequested -= CoreWebView2_PermissionRequested;
                    webView.CoreWebView2.NavigationCompleted -= WebView_NavigationCompleted;
                }

                webView.Dispose();
            }
            catch (Exception ex)
            {
                ServiceLocator.Logger?.LogException(ex, "GLLogin.GLLogin_Closed: WebView2 dispose failed");
            }
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

        private void CloseButton_Click(object sender, RoutedEventArgs e)
        {
            ServiceLocator.Logger?.LogDebug("GLLogin.CloseButton_Click invoked - closing window");
            Close();
        }

        private void BtnClose_Click(object sender, RoutedEventArgs e)
        {
            ServiceLocator.Logger?.LogDebug("GLLogin.BtnClose_Click invoked - closing window");
            Close();
        }

        // ---------- Server list / server combo ----------

        private void GLLogin_Loaded(object sender, RoutedEventArgs e)
        {
            if (!File.Exists(ServiceLocator.Paths.UrlsDirectory))
            {
                Dispatcher.InvokeAsync(() =>
                {
                    AppOverlayControl.ShowWarning("Server list path is not configured.");
                });

                return;
            }
            LoadServerList();
        }

        private void LoadServerList()
        {
            try
            {
                var doc = XDocument.Load(ServiceLocator.Paths.UrlsDirectory);
                ServerList = doc.Descendants("URL")
                    .Select(x => new ServerInfo
                    {
                        Name = (string)x.Element("Name"),
                        Address = (string)x.Element("Address"),
                        DefaultURL = (bool?)x.Element("DefaultURL") ?? false
                    })
                    .ToList();

                cmbServer.ItemsSource = ServerList;
            }
            catch (Exception ex)
            {
                ServiceLocator.Logger?.LogException(ex, "Failed to load server list");
                Dispatcher.InvokeAsync(() =>
                {
                    AppOverlayControl.ShowWarning("Failed to load server list.");
                });
            }
        }

        // ---------- WebView2 lifecycle ----------

        private async void WebView_Loaded(object sender, RoutedEventArgs e)
        {
            webView.Loaded -= WebView_Loaded; // prevent double init

            try
            {
                // 1) Ensure a writable user data folder (logs/profile)
                string logDir = ServiceLocator.Paths.LoginBrowserPath;
                DirectoryInfo di = new(logDir);
                if (!di.Exists)
                    di.Create();

                string webViewLogsPath = di.FullName;

                // 2) Create environment options FIRST
                var envOptions = new CoreWebView2EnvironmentOptions
                {
                    // Enable SSO if your scenario needs it
                    AllowSingleSignOnUsingOSPrimaryAccount = true
                };

                // 3) Create the environment with options + user data folder
                var env = await CoreWebView2Environment.CreateAsync(
                    browserExecutableFolder: null,
                    userDataFolder: webViewLogsPath,
                    options: envOptions);

                // 4) Initialize WebView2 with that environment
                _webViewInitTask = webView.EnsureCoreWebView2Async(env);
                await _webViewInitTask;

                // 5) Hook device permission handler and diagnostics after CoreWebView2 is ready
                webView.CoreWebView2.PermissionRequested += CoreWebView2_PermissionRequested;
                webView.CoreWebView2.NavigationCompleted += WebView_NavigationCompleted;

                // Cert-error bypass (scoped to the customer's own configured server) and
                // retry-once navigation for SSO/SAML/OIDC redirects. Ported from
                // FinalWorkingCode's WebView2NavigationResilience (popup-hosting piece
                // excluded - see that class's header comment).
                _resilience = new WebView2NavigationResilience(nameof(GLLogin), this);
                _resilience.Attach(webView.CoreWebView2, GetTrustedHosts);

                // Optional: turn on DevTools during development
                webView.CoreWebView2.Settings.AreDevToolsEnabled = true;

                // 6) Diagnostics: log WebView2 runtime and SSO setting
                var version = webView.CoreWebView2.Environment.BrowserVersionString;
                ServiceLocator.Logger?.LogDebug($"WebView2 BrowserVersion={version}");
                ServiceLocator.Logger?.LogDebug($"AllowSingleSignOnUsingOSPrimaryAccount={envOptions.AllowSingleSignOnUsingOSPrimaryAccount}");
            }
            catch (Exception ex)
            {
                ServiceLocator.Logger?.LogException(ex, "WebView2 initialization failed in GLLogin");
                Close();
            }
        }

        private void CoreWebView2_PermissionRequested(object? sender, CoreWebView2PermissionRequestedEventArgs e)
        {
            try
            {
                ServiceLocator.Logger?.LogDebug($"Permission requested: Kind={e.PermissionKind}, Uri={e.Uri}");

                switch (e.PermissionKind)
                {
                    case CoreWebView2PermissionKind.Microphone:
                    case CoreWebView2PermissionKind.Camera:
                    case CoreWebView2PermissionKind.Geolocation:
                    case CoreWebView2PermissionKind.MidiSystemExclusiveMessages:
                    case CoreWebView2PermissionKind.ClipboardRead:
                        {
                            e.State = CoreWebView2PermissionState.Allow;
                            e.Handled = true;
                            ServiceLocator.Logger?.LogDebug($"Permission allowed: {e.PermissionKind}");
                            break;
                        }

                    default:
                        e.State = CoreWebView2PermissionState.Deny;
                        e.Handled = true;
                        ServiceLocator.Logger?.LogWarn($"Permission denied: {e.PermissionKind}");
                        break;
                }
            }
            catch (Exception ex)
            {
                ServiceLocator.Logger?.LogException(ex, "PermissionRequested handler error");
                // Fail closed on error
                e.State = CoreWebView2PermissionState.Deny;
                e.Handled = true;
            }
        }

        private async Task NavigateToBlankPageAsync()
        {
            try
            {
                await EnsureWebViewInitializedAsync();

                var tcs = new TaskCompletionSource<bool>();

                void Handler(object? sender, CoreWebView2NavigationCompletedEventArgs e)
                {
                    webView.CoreWebView2.NavigationCompleted -= Handler;
                    tcs.TrySetResult(true);
                }

                webView.CoreWebView2.NavigationCompleted += Handler;
                webView.CoreWebView2.Navigate("about:blank");

                await tcs.Task;
            }
            catch (Exception ex)
            {
                ServiceLocator.Logger?.LogException(ex, "Login: NavigateToBlankPage");
            }
        }

        private async Task CmbServer_SelectionCommitted(object obj)
        {
            if (obj is not ServerInfo selected)
            {
                ServiceLocator.Logger?.LogDebug("GLLogin.CmbServer_SelectionCommitted: committed object was not a ServerInfo, ignoring");
                return;
            }

            ServiceLocator.Logger?.LogDebug($"GLLogin.CmbServer_SelectionCommitted invoked: server={selected.Name}, address={selected.Address}");
            SelectedServer = selected;
            txtServer.Text = selected.Address;

            _activeCancellation?.Cancel();
            _activeCancellation = null;

            using var cts = new CancellationHelper();
            _activeCancellation = cts;

            try
            {
                await ShowBusyOverlayAsync(cts, "Connecting to server...");

                await NavigateToBlankPageAsync();

                if (selected.Address != null)
                    await NavigateToServerAsync(selected.Address, cts);
            }
            catch (TaskCanceledException)
            {
                await NavigateToBlankPageAsync();
            }
            catch (Exception ex)
            {
                ServiceLocator.Logger?.LogException(ex, "Login: SelectionCommitted navigation failed");
            }
            finally
            {
                if (_activeCancellation == cts)
                {
                    _activeCancellation = null;
                }
            }
        }

        private async Task ShowBusyOverlayAsync(CancellationHelper helper, string message)
        {
            await Dispatcher.InvokeAsync(() =>
            {
                webView.Visibility = Visibility.Collapsed;

                AppOverlayControl.ShowBusyasyn(
                    message: message + " (Click cancel to stop)",
                    cancelAction: async () =>
                    {
                        if (!helper.IsCancellationRequested)
                        {
                            helper.Cancel();
                            ServiceLocator.Logger?.LogWarn($"Operation cancelled by user: {message}");
                        }
                        await Task.Delay(80); // small feedback delay
                                              // HideBusyAsync() is already called in the overlay's handler
                    }
                );
            });
        }

        private async Task NavigateToServerAsync(string address, CancellationHelper cts)
        {
            try
            {
                if (!string.IsNullOrWhiteSpace(address))
                {
                    await EnsureWebViewInitializedAsync();

                    try
                    {
                        var result = await _resilience!.NavigateWithRetryAsync(
                            webView.CoreWebView2, address.Trim() + "?finance_excel=Y", cts.GetToken());

                        if (!result.IsSuccess)
                        {
                            ServiceLocator.Logger?.LogWarn($"Login: navigation to '{address}' failed after retry ({result.WebErrorStatus}).");
                            await NavigateToBlankPageAsync();
                            await AppOverlayControl.ShowErrorAsync("Unable to load the page. Please try again.");
                        }
                    }
                    catch (TaskCanceledException)
                    {
                        ServiceLocator.Logger?.LogWarn("Navigation was cancelled by the user.");
                        await NavigateToBlankPageAsync();
                    }
                }
                else
                {
                    await NavigateToBlankPageAsync();
                }
            }
            catch (Exception ex)
            {
                ServiceLocator.Logger?.LogException(ex, "Login: NavigateToServer");
            }
            finally
            {
                await Dispatcher.InvokeAsync(async () =>
                {
                    await AppOverlayControl.HideBusyAsync();
                    webView.Visibility = Visibility.Visible;
                });
            }
        }

        // Trusted-host set for WebView2NavigationResilience's certificate-error bypass -
        // re-read at bypass time, so switching the server dropdown selection updates it
        // without needing to re-attach anything.
        private IReadOnlyCollection<string> GetTrustedHosts()
        {
            try
            {
                var address = SelectedServer?.Address;
                if (string.IsNullOrWhiteSpace(address))
                    return Array.Empty<string>();

                return new[] { new Uri(address, UriKind.Absolute).Host };
            }
            catch (Exception ex)
            {
                ServiceLocator.Logger?.LogDebug($"GLLogin.GetTrustedHosts: could not resolve host from '{SelectedServer?.Address}': {ex.Message}");
                return Array.Empty<string>();
            }
        }

        private async Task EnsureWebViewInitializedAsync()
        {
            if (_webViewInitTask != null)
                await _webViewInitTask;

            while (webView.CoreWebView2 == null)
                await Task.Delay(50); // Wait until CoreWebView2 is ready
        }

        // ---------- Login-success detection + cube fetch ----------

        private async void WebView_NavigationCompleted(object? sender, CoreWebView2NavigationCompletedEventArgs e)
        {
            if (!string.IsNullOrWhiteSpace(AppState.Instance.LoginToken))
                return;

            if (!e.IsSuccess)
            {
                ServiceLocator.Logger?.LogWarn("Navigation failed: " + e.WebErrorStatus + " | Source : " + (webView.Source?.ToString() ?? string.Empty));
                return;
            }

            var src = webView.Source?.ToString() ?? string.Empty;
            if (!IsLoginSuccessView(src))
                return;

            ServiceLocator.Logger?.LogDebug("Document Title: " + (webView.CoreWebView2?.DocumentTitle ?? ""));
            ServiceLocator.Logger?.LogDebug("URL: " + (webView.Source?.ToString() ?? ""));

            _activeCancellation?.Cancel();
            _activeCancellation = null;

            using var cts = new CancellationHelper();
            _activeCancellation = cts;

            try
            {
                try
                {
                    await ShowBusyOverlayAsync(cts, "Login Success! Fetching Cube Details...");
                    AppState.Instance.LoginUrl = CleanLoginUrl(txtServer.Text.Trim());

                    await ExtractLoginCookies(src);
                    if (string.IsNullOrEmpty(AppState.Instance.LoginToken))
                    {
                        await ShowErrorAndCloseAsync("Cookies configuration missing.");
                        return;
                    }

                    var apiResult = await ProcessApiAfterLogin(cts.GetToken());

                    // XLEdge login if the sibling add-in is present
                    InvokeXlEdgeLogin();

                    await HandleApiResult(apiResult, cts.GetToken());
                }
                finally
                {
                    await AppOverlayControl.HideBusyAsync();
                }
            }
            catch (Exception ex)
            {
                ServiceLocator.Logger?.LogException(ex, "Login: Navigation Completed");
            }
            finally
            {
                if (_activeCancellation == cts)
                {
                    _activeCancellation = null;
                }
            }
        }

        private static bool IsLoginSuccessView(string src)
        {
            return !string.IsNullOrEmpty(src) &&
                   src.IndexOf("finance-excel-login-success-view", StringComparison.OrdinalIgnoreCase) >= 0;
        }

        private static string CleanLoginUrl(string url)
        {
            var patterns = new[] { "/bypass-saml-login-flow", "/bypass-sso-login-flow" };
            var cleanUrl = url.ToLowerInvariant();
            foreach (var pattern in patterns)
                cleanUrl = cleanUrl.Replace(pattern, "");
            return cleanUrl.TrimEnd('/', '\\');  // Note: May need case-preserving logic if original casing matters
        }

        private async Task ExtractLoginCookies(string src)
        {
            if (webView.CoreWebView2 == null) return;
            bool xlEdgeKeyFound = false;
            try
            {
                var browserCookies = await webView.CoreWebView2.CookieManager.GetCookiesAsync(src);
                foreach (var c in browserCookies)
                {
                    if (c == null || string.IsNullOrEmpty(c.Name) || string.IsNullOrEmpty(c.Value)) continue;
                    var name = c.Name.ToUpper();
                    switch (name)
                    {
                        case "XL-AUTH-TOKEN":
                        case "ORB-AUTH-TOKEN":
                            AppState.Instance.LoginToken = c.Value;
                            break;
                        case "X-ORB-USERNAME":
                            AppState.Instance.LoginUserName = HttpUtility.UrlDecode(c.Value) ?? string.Empty;
                            break;
                        case "XLEDGE-USER-ACCESS":
                            {
                                xlEdgeKeyFound = true;
                                if (bool.TryParse(c.Value, out bool tempValue))
                                {
                                    _xlEdgePermission = tempValue;
                                }
                            }
                            break;
                    }
                }

                if (!xlEdgeKeyFound)
                {
                    ServiceLocator.Logger?.LogWarn("XLEDGE-USER-ACCESS cookie not found. Possible older version of GL-Sense. Using default, user has access to XLEdge.");
                }
            }
            catch (Exception ex)
            {
                ServiceLocator.Logger?.LogException(ex, "Failed to extract login token from page");
                throw;
            }
        }

        // ---------- XLEdge hand-off ----------

        private void InvokeXlEdgeLogin(bool loginFailed = false)
        {
            try
            {
                object? edgeAddinInstance = EdgeAddinHelper.GetEdgeAddinInstance();
                if (edgeAddinInstance == null)
                    return;

                if (loginFailed && !_xlEdgePermission)
                    return;

                if (loginFailed && _xlEdgePermission)
                {
                    ServiceLocator.RibbonController?.SetControlEnabled("RibLogin", false);
                }

                string instanceName = SelectedServer?.Name ?? "UnKnown";

                ServiceLocator.Logger?.LogDebug("Invoking XLEdge login with parameters: " +
                    $"InstanceName={instanceName}, " +
                    $"LoginUrl={AppState.Instance.LoginUrl}, " +
                    $"LoginToken={(string.IsNullOrEmpty(AppState.Instance.LoginToken) ? "null/empty" : "exists")}, " +
                    $"LoginUserName={(string.IsNullOrEmpty(AppState.Instance.LoginUserName) ? "null/empty" : AppState.Instance.LoginUserName)}, " +
                    $"XLEdgePermission={_xlEdgePermission}");

                edgeAddinInstance.GetType().InvokeMember(
                    "InvokedFromGLSense",
                    System.Reflection.BindingFlags.InvokeMethod,
                    null,
                    edgeAddinInstance,
                    new object[]
                    {
                        instanceName,
                        AppState.Instance.LoginUrl!,
                        AppState.Instance.LoginToken!,
                        AppState.Instance.LoginUserName!,
                        _xlEdgePermission
                    });
            }
            catch (Exception ex)
            {
                ServiceLocator.Logger?.LogException(ex);
            }
        }

        // ---------- Cube fetch + result handling ----------

        private static async Task<ApiResult<string>> ProcessApiAfterLogin(CancellationToken token)
        {
            return await GetDataFromApi(token);
        }

        private async Task HandleApiResult(ApiResult<string> apiResult, CancellationToken token)
        {
            ServiceLocator.Logger?.LogDebug($"GLLogin.HandleApiResult invoked: IsSuccess={apiResult.IsSuccess}");
            if (apiResult.IsSuccess)
            {
                AppState.Instance.IsLoggedIn = true;
                ServiceLocator.RibbonController?.SetState("PartialLoggedIn");
                var result = ApiResponseHelper.Parse<List<CubeRecord>>(apiResult.Value!, JsonGlobals.Options);

                if (!result.IsSuccess)
                {
                    ServiceLocator.Logger?.LogError($"GLLogin.HandleApiResult: cube response parse failed: {result.ErrorMessage}");
                    await ShowErrorAndCloseAsync(result.ErrorMessage);
                    return;
                }

                await SuccessCube(result, token);
            }
            else
            {
                var error = apiResult.Exception;
                if (error != null)
                {
                    ServiceLocator.Logger?.LogError($"GLLogin.HandleApiResult: API result failed: {error.Message}", error);
                    await ShowErrorAndCloseAsync(error.Message);
                }
            }
        }

        private async Task SuccessCube(ApiResult<List<CubeRecord>> result, CancellationToken token)
        {
            CubeCache.AllCubes = result.Value!.OrderBy(c => c.CubeName).ToList();
            ServiceLocator.Logger?.LogDebug($"GLLogin.SuccessCube: loaded {CubeCache.AllCubes.Count} cubes, inserting cube data");

            await CubeDataRepository.InsertCubeDataAsync();

            var broadcastMsg = await BroadcastMessageFromApi(token);

            await AppOverlayControl.HideBusyAsync();
            if (broadcastMsg != null && !string.IsNullOrWhiteSpace(broadcastMsg))
                await AppOverlayControl.ShowInfoAsync(broadcastMsg);

            Close();
        }

        private async Task ShowErrorAndCloseAsync(string message)
        {
            ServiceLocator.Logger?.LogDebug($"GLLogin.ShowErrorAndCloseAsync invoked: {message}");
            await AppOverlayControl.HideBusyAsync();
            await AppOverlayControl.ShowErrorAsync(message);
            ServiceLocator.RibbonController?.SetControlEnabled("RibLogin", true);
            await NavigateToLoginAsync();
        }

        private async Task NavigateToLoginAsync()
        {
            await NavigateToBlankPageAsync();
            try
            {
                if (!string.IsNullOrWhiteSpace(AppState.Instance.LoginUrl))
                {
                    AppState.Instance.LoginUrl = CleanLoginUrl(AppState.Instance.LoginUrl);
                    AppState.Instance.LoginToken = null;
                    await EnsureWebViewInitializedAsync();

                    try
                    {
                        var result = await _resilience!.NavigateWithRetryAsync(
                            webView.CoreWebView2, AppState.Instance.LoginUrl + "?finance_excel=Y", CancellationToken.None);

                        if (!result.IsSuccess)
                        {
                            ServiceLocator.Logger?.LogWarn($"Login: navigation to '{AppState.Instance.LoginUrl}' failed after retry ({result.WebErrorStatus}).");
                            await NavigateToBlankPageAsync();
                            await AppOverlayControl.ShowErrorAsync("Unable to load the page. Please try again.");
                        }
                    }
                    catch (TaskCanceledException)
                    {
                        ServiceLocator.Logger?.LogWarn("Navigation was cancelled by the user.");
                        await NavigateToBlankPageAsync();
                    }
                }
                else
                {
                    await NavigateToBlankPageAsync();
                }
            }
            catch (Exception ex)
            {
                ServiceLocator.Logger?.LogException(ex, "Login: NavigateToServer");
            }
            finally
            {
                await Dispatcher.InvokeAsync(async () =>
                {
                    await AppOverlayControl.HideBusyAsync();
                    webView.Visibility = Visibility.Visible;
                });
            }
        }

        private static async Task<ApiResult<string>> GetDataFromApi(CancellationToken ct)
        {
            try
            {
                string apiUrl = AppState.Instance.LoginUrl + "/rest/secure/finance/finance-cubes";
                string cubeResponse = await ApiHelper.ServerAPI(apiUrl, "Form", "", "POST", ct);

                string? errorMsg = ExtractErrorMessage(cubeResponse);
                if (!string.IsNullOrEmpty(errorMsg))
                {
                    HandleApiError(new InvalidOperationException(errorMsg), apiUrl, cubeResponse);
                    errorMsg ??= string.Empty;
                    return ApiResult<string>.Failure(errorMsg);
                }

                if (!IsValidCubeResponse(cubeResponse, apiUrl))
                    return ApiResult<string>.Failure($"No cubes assigned to user \"{AppState.Instance.LoginUserName}\"");

                return ApiResult<string>.Success(cubeResponse);
            }
            catch (OperationCanceledException ex)
            {
                HandleApiError(ex);
                return ApiResult<string>.Failure(ex);
            }
            catch (Exception ex)
            {
                HandleApiError(ex);
                return ApiResult<string>.Failure(ex);
            }
        }

        private static string? ExtractErrorMessage(string response)
        {
            try
            {
                using JsonDocument doc = JsonDocument.Parse(response);
                JsonElement root = doc.RootElement;

                if (root.TryGetProperty("status", out JsonElement statusElem) &&
                    statusElem.GetString()?.Equals("failed", StringComparison.OrdinalIgnoreCase) == true &&
                    root.TryGetProperty("msg", out JsonElement msgElem))
                {
                    return msgElem.GetString() ?? "Unknown error from server";
                }
            }
            catch (JsonException ex)
            {
                ServiceLocator.Logger?.LogException(ex, "Failed to parse error message JSON");
                ServiceLocator.Logger?.LogRawJson("GLLogin.ExtractErrorMessage", response);
            }

            return null;
        }

        private static bool IsValidCubeResponse(string response, string apiUrl)
        {
            if (response.StartsWith(AppConstants.ErrorPrefix, StringComparison.OrdinalIgnoreCase) ||
                response.IndexOf(AppConstants.UnauthorizedMessage, StringComparison.OrdinalIgnoreCase) >= 0 ||
                response.IndexOf("<!DOCTYPE", StringComparison.OrdinalIgnoreCase) >= 0)
            {
                HandleApiError(new InvalidOperationException("Cube data error."), apiUrl, response);
                return false;
            }

            if (string.IsNullOrWhiteSpace(response))
            {
                HandleApiError(new InvalidOperationException("Cube data format is invalid."), apiUrl, response);
                return false;
            }

            try
            {
                using JsonDocument doc = JsonDocument.Parse(response);
                JsonElement root = doc.RootElement;

                if (!HasValidStatus(root, apiUrl, response))
                    return false;

                if (!HasValidRecords(root))
                    return false;

                return true;
            }
            catch (JsonException jsonEx)
            {
                ServiceLocator.Logger?.LogRawJson("CubeResponse", response);
                HandleApiError(jsonEx, apiUrl, response);
                return false;
            }
        }

        private static bool HasValidStatus(JsonElement root, string apiUrl, string response)
        {
            if (!root.TryGetProperty(AppConstants.Status, out JsonElement statusElem) ||
                statusElem.GetString() != AppConstants.Success)
            {
                HandleApiError(new InvalidOperationException("Cube data status is failed."), apiUrl, response);
                return false;
            }
            return true;
        }

        private static bool HasValidRecords(JsonElement root)
        {
            if (root.TryGetProperty(AppConstants.Records, out JsonElement recordsElem) &&
                recordsElem.ValueKind == JsonValueKind.Array &&
                recordsElem.GetArrayLength() > 0)
            {
                return true;
            }

            var ex = new InvalidOperationException("No cubes assigned to the current user.");
            HandleApiError(ex);
            return false;
        }

        private static void HandleApiError(Exception ex, string? apiUrl = null, string? response = null)
        {
            ServiceLocator.Logger?.LogWarn($"GetDataFromApi|API: {apiUrl}");
            if (ex != null)
            {
                ServiceLocator.Logger?.LogException(ex, "GetDataFromApi|Exception");
            }
            if (!string.IsNullOrEmpty(response))
                ServiceLocator.Logger?.LogWarn($"GetDataFromApi|Response: {response}");

            AppState.Instance.IsLoggedIn = false;
            SafeDisableLoginRibbon();
        }

        private static async Task<string> BroadcastMessageFromApi(CancellationToken ct)
        {
            try
            {
                string apiUrl = AppState.Instance.LoginUrl + AppConstants.WebSecure + "get-broadcast-msg";
                string rawResponse = await FetchApiResponseAsync(apiUrl, ct);

                if (string.IsNullOrWhiteSpace(rawResponse))
                    return string.Empty;

                return FormatBroadcastMessages(rawResponse);
            }
            catch (OperationCanceledException ex)
            {
                ServiceLocator.Logger?.LogWarn(ex.Message);
                return string.Empty;
            }
            catch (Exception ex)
            {
                ServiceLocator.Logger?.LogException(ex);
                return string.Empty;
            }
        }

        private static string FormatBroadcastMessages(string rawResponse)
        {
            try
            {
                var result = ApiResponseHelper.Parse<List<BroadcastMessage>>(rawResponse, JsonGlobals.Options);

                if (!result.IsSuccess)
                {
                    ServiceLocator.Logger?.LogWarn($"Broadcast parsing failed: {result.ErrorMessage}");
                    return string.Empty;
                }

                if (result.Value == null || result.Value.Count == 0)
                    return string.Empty;

                return BuildMessageString(result.Value);
            }
            catch (Exception ex)
            {
                ServiceLocator.Logger?.LogException(ex, "Error processing broadcast messages.");
                return string.Empty;
            }
        }

        private static async Task<string> FetchApiResponseAsync(string apiUrl, CancellationToken ct)
        {
            try
            {
                return await ApiHelper.ServerAPI(apiUrl, "Form", "", "POST", ct);
            }
            catch (OperationCanceledException ex)
            {
                ServiceLocator.Logger?.LogWarn(ex.Message);
                return string.Empty;
            }
            catch (Exception ex)
            {
                ServiceLocator.Logger?.LogException(ex, $"Error fetching broadcast messages from {apiUrl}: {ex.Message}");
                return string.Empty;
            }
        }

        private static string BuildMessageString(List<BroadcastMessage> messages)
        {
            if (messages == null || messages.Count == 0)
                return string.Empty;

            var sb = new StringBuilder();

            for (int i = 0; i < messages.Count; i++)
            {
                var msg = messages[i];

                sb.Append(i + 1)
                  .Append(".) ")
                  .Append(msg.MsgType ?? "Info")
                  .Append(" : ")
                  .Append(msg.Message ?? string.Empty);

                if (i < messages.Count - 1)
                    sb.AppendLine();
            }

            return sb.ToString();
        }

        private static void SafeDisableLoginRibbon()
        {
            try
            {
                if (EdgeAddinHelper.GetEdgeAddinInstance() != null)
                    ServiceLocator.RibbonController?.SetControlEnabled("RibLogin", false);
            }
            catch (Exception ex)
            {
                ServiceLocator.Logger?.LogException(ex);
            }
        }
    }

    public class ServerInfo
    {
        public string? Name { get; set; }
        public string? Address { get; set; }
        public bool DefaultURL { get; set; }
    }
#nullable disable
}
