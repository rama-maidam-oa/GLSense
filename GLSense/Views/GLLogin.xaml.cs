using ControlzEx.Standard;
using GLSense.Helpers;
using GLSense.Models;
using GLSense.Repositories;
using GLSense.Utilities;
using Microsoft.Web.WebView2.Core;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.CompilerServices;
using System.Text;
using System.Text.Json;
using System.Threading;
using System.Threading.Tasks;
using System.Web;
using System.Windows;
using System.Xml.Linq;

namespace GLSense.Views
{
#nullable enable
    /// <summary>
    /// Interaction logic for GLLogin.xaml
    /// </summary>
    public partial class GLLogin : DpiAwareWindow, INotifyPropertyChanged
    {
        private bool xlEdgePermission;
        private Task? _webViewInitTask;
        private CancellationHelper? _activeCancellation;

        // Backing field
        private ServerInfo? _selectedServer;
        public List<ServerInfo>? ServerList { get; set; }
        // Property with custom setter logic
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
        public GLLogin()
        {
            InitializeComponent();
            EnhancedDragDropHelper.EnableWindowDrag(this);

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
        }

        public event PropertyChangedEventHandler? PropertyChanged;
        protected virtual void OnPropertyChanged([CallerMemberName] string? propertyName = null)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
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
                            LogUtility.LogWarn($"Operation cancelled by user: {message}");
                        }
                        await Task.Delay(80); // small feedback delay
                                              // HideBusyAsync() is already called in the overlay's handler
                    }
                );
            });
        }
        private void GLLogin_Loaded(object sender, RoutedEventArgs e)
        {
            if (!File.Exists(AppPaths.TempUrlsPath))
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
                var doc = XDocument.Load(AppPaths.TempUrlsPath);
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
                LogUtility.LogException(ex, "Failed to load server list");
                Dispatcher.InvokeAsync(() =>
                {
                    AppOverlayControl.ShowWarning("Failed to load server list.");
                });

            }
        }
        private async void WebView_Loaded(object sender, RoutedEventArgs e)
        {
            webView.Loaded -= WebView_Loaded; // prevent double init

            using (new LogUtility.LogScope("WebView2 Initialization"))
            {
                try
                {
                    // 1) Ensure a writable user data folder (logs/profile)
                    string logDir = AppPaths.LoginBrowserLogsPath;
                    DirectoryInfo di = new(logDir);
                    if (!di.Exists)
                        di.Create();

                    string webViewLogsPath = di.FullName;

                    // 2) Create environment options FIRST
                    var envOptions = new CoreWebView2EnvironmentOptions
                    {
                        // Enable SSO if your scenario needs it
                        AllowSingleSignOnUsingOSPrimaryAccount = true

                        // Optional: enable features / diagnostics
                        // AdditionalBrowserArguments = "--enable-logging=stderr --v=1"
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
                    webView.CoreWebView2.ProcessFailed += CoreWebView2_ProcessFailed;
                    webView.CoreWebView2.NavigationCompleted += WebView_NavigationCompleted;

                    // Optional: turn on DevTools during development
                    webView.CoreWebView2.Settings.AreDevToolsEnabled = true;

                    // 6) Diagnostics: log WebView2 runtime and SSO setting
                    var version = webView.CoreWebView2.Environment.BrowserVersionString;
                    LogUtility.LogDebug($"WebView2 BrowserVersion={version}");
                    LogUtility.LogDebug($"AllowSingleSignOnUsingOSPrimaryAccount={envOptions.AllowSingleSignOnUsingOSPrimaryAccount}");

                }
                catch (Exception ex)
                {
                    LogUtility.LogException(ex, "WebView2 initialization failed in GLLogin");
                    Close();
                }
            }
        }

        private void CoreWebView2_PermissionRequested(object? sender, CoreWebView2PermissionRequestedEventArgs e)
        {
            using (new LogUtility.LogScope("CoreWebView2_PermissionRequested"))
            {
                try
                {
                    LogUtility.LogDebug($"Permission requested: Kind={e.PermissionKind}, Uri={e.Uri}");

                    // Decide what to allow
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
                                LogUtility.LogDebug($"Permission allowed: {e.PermissionKind}");
                                break;
                            }

                        default:
                            e.State = CoreWebView2PermissionState.Deny;
                            e.Handled = true;
                            LogUtility.LogWarn($"Permission denied: {e.PermissionKind}");
                            break;
                    }
                }
                catch (Exception ex)
                {
                    LogUtility.LogException(ex, "PermissionRequested handler error");
                    // Fail closed on error
                    e.State = CoreWebView2PermissionState.Deny;
                    e.Handled = true;
                }
            }
        }
        private void CoreWebView2_ProcessFailed(object? sender, CoreWebView2ProcessFailedEventArgs e)
        {
            LogUtility.LogWarn($"WebView2 process failed. Kind={e.ProcessFailedKind}");
        }
        private async Task NavigateToBlankPageAsync()
        {
            try
            {
                await EnsureWebViewInitializedAsync();

                var tcs = new TaskCompletionSource<bool>();

                void Handler(object sender, CoreWebView2NavigationCompletedEventArgs e)
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
                LogUtility.LogException(ex, "Login: NavigateToBlankPage");
            }
        }

        private async Task CmbServer_SelectionCommitted(object obj)
        {
            if (obj is not ServerInfo selected)
                return;

            SelectedServer = selected;
            txtServer.Text = selected.Address;

            // 1. Stop & clean up any previous (very rare in Loaded, but good habit)
            _activeCancellation?.Cancel();
            _activeCancellation = null;

            using var cts = new CancellationHelper();
            _activeCancellation = cts;

            try
            {
                await ShowBusyOverlayAsync(cts, "Connecting to server...");

                await NavigateToBlankPageAsync();

                // 🟢 Actually navigate — will trigger busy period
                if (selected.Address != null)
                    await NavigateToServerAsync(selected.Address, cts);
            }
            catch (TaskCanceledException)
            {
                // Handle cancel (optional)
                LogUtility.LogDebug("GLLogin.CmbServer_SelectionCommitted: navigation cancelled (TaskCanceledException) - returning to blank page");
                await NavigateToBlankPageAsync();
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex, "Login: SelectionCommitted navigation failed");
            }
            finally
            {
                if (_activeCancellation == cts)
                {
                    _activeCancellation = null;
                }
            }
        }

        private async Task NavigateToServerAsync(string address, CancellationHelper cts)
        {
            try
            {
                if (!string.IsNullOrWhiteSpace(address))
                {
                    await EnsureWebViewInitializedAsync();

                    var tcs = new TaskCompletionSource<bool>();

                    void Handler(object sender, CoreWebView2NavigationCompletedEventArgs e)
                    {
                        webView.CoreWebView2.NavigationCompleted -= Handler;
                        tcs.TrySetResult(true);
                    }

                    webView.CoreWebView2.NavigationCompleted += Handler;

                    try
                    {
                        webView.CoreWebView2.Navigate(address.Trim() + "?finance_excel=Y");

                        using (cts.GetToken().Register(() => tcs.TrySetCanceled()))
                        {
                            await tcs.Task;
                        }
                    }
                    catch (TaskCanceledException)
                    {
                        LogUtility.LogWarn("Navigation was cancelled by the user.");
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
                LogUtility.LogException(ex, "Login: NavigateToServer");
            }
            finally
            {
                // 🟢 Hide busy overlay safely on UI thread
                await Dispatcher.InvokeAsync(async () =>
                {
                    await AppOverlayControl.HideBusyAsync();
                    webView.Visibility = Visibility.Visible;
                });
            }
        }
        private async Task EnsureWebViewInitializedAsync()
        {
            if (_webViewInitTask != null)
                await _webViewInitTask;

            while (webView.CoreWebView2 == null)
                await Task.Delay(50); // Wait until CoreWebView2 is ready
        }

        private async void WebView_NavigationCompleted(object sender, CoreWebView2NavigationCompletedEventArgs e)
        {
            if (!string.IsNullOrWhiteSpace(AppState.Instance.LoginToken))
                return;

            if (!e.IsSuccess)
            {
                LogUtility.LogWarn("Navigation failed: " + e.WebErrorStatus + " | Source : " + webView.Source?.ToString() ?? string.Empty);
                return;
            }

            var src = webView.Source?.ToString() ?? string.Empty;
            if (!IsLoginSuccessView(src))
                return;

            using (new LogUtility.LogScope("WebView_NavigationCompleted"))
            {
                LogUtility.LogDebug("Document Title: " + webView.CoreWebView2?.DocumentTitle ?? "");
                LogUtility.LogDebug("URL: " + webView.Source?.ToString() ?? "");

                // 1. Stop & clean up any previous (very rare in Loaded, but good habit)
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

                        //XlEdge Login If Exists
                        InvokeXlEdgeLogin();
                        //Ends XLEdge Login

                        await HandleApiResult(apiResult, cts.GetToken());
                    }
                    finally
                    {
                        await AppOverlayControl.HideBusyAsync();
                    }
                }
                catch (Exception ex)
                {
                    LogUtility.LogException(ex, "Login: Navigation Completed");
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
            bool XLEdgeKey = false;
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
                                XLEdgeKey = true;
                                if (bool.TryParse(c.Value, out bool tempValue))
                                {
                                    xlEdgePermission = tempValue;
                                }
                            }
                            break;
                    }
                }

                if (!XLEdgeKey)
                {
                    LogUtility.LogWarn("XLEDGE-USER-ACCESS cookie not found. Possible older version of GL-Sense. Using default, user has access to XLEdge.");
                }

            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex, "Failed to extract login token from page");
                throw;
            }
        }

        private void InvokeXlEdgeLogin(bool loginFailed = false)
        {
            try
            {
                object? edgeAddinInstance = GetEdgeAddinInstance();
                if (edgeAddinInstance == null)
                    return;

                if (loginFailed && !xlEdgePermission)
                    return;

                if (loginFailed && xlEdgePermission && AddinModule.CurrentInstance != null)
                {
                    AddinModule.CurrentInstance.RibLogin.Enabled = false;
                }

                string instanceName = SelectedServer?.Name ?? "UnKnown";

                LogUtility.LogDebug("Invoking XLEdge login with parameters: " +
                    $"InstanceName={instanceName}, " +
                    $"LoginUrl={AppState.Instance.LoginUrl}, " +
                    $"LoginToken={(string.IsNullOrEmpty(AppState.Instance.LoginToken) ? "null/empty" : "exists")}, " +
                    $"LoginUserName={(string.IsNullOrEmpty(AppState.Instance.LoginUserName) ? "null/empty" : AppState.Instance.LoginUserName)}, " +
                    $"XLEdgePermission={xlEdgePermission}");

                edgeAddinInstance.GetType().InvokeMember(
                    "InvokedFromGLSense",
                    BindingFlags.InvokeMethod,
                    null,
                    edgeAddinInstance,
                    new object[]
                    {
                        instanceName,
                        AppState.Instance.LoginUrl,
                        AppState.Instance.LoginToken,
                        AppState.Instance.LoginUserName,
                        xlEdgePermission
                    });
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex);
            }
        }

        public static object? GetEdgeAddinInstance()
        {
            return AddinModule.GetEdgeAddinInstance();
        }
        private static async Task<ApiResult<string>> ProcessApiAfterLogin(CancellationToken token)
        {
            return await GetDataFromApi(token);
        }

        private async Task HandleApiResult(ApiResult<string> apiResult, CancellationToken token)
        {
            if (apiResult.IsSuccess)
            {
                AddinModule.RibbonHelper.ApplyState("PartialLoggedIn");
                var result = ApiResponseHelper.Parse<List<CubeRecord>>(apiResult.Value, JsonGlobals.Options);

                if (!result.IsSuccess)
                {
                    if (AppOverlayControl != null)
                    {
                        await ShowErrorAndCloseAsync(result.ErrorMessage);
                    }
                    return;
                }

                await SuccessCube(result, token);
            }
            else
            {
                await Dispatcher.InvokeAsync(() => { });
                var error = apiResult.Exception;
                if (AppOverlayControl != null && error != null)
                {
                    await ShowErrorAndCloseAsync(error.Message);
                }
            }
        }
        private async Task SuccessCube(ApiResult<List<CubeRecord>> result, CancellationToken token)
        {
            CubeCache.AllCubes = result.Value!.OrderBy(c => c.CubeName).ToList();

            await CubeDataRepository.InsertCubeDataAsync();

            var broadcastMsg = await BroadcastMessageFromApi(token);
            await Dispatcher.InvokeAsync(() => { });
            if (AppOverlayControl != null)
            {
                await AppOverlayControl.HideBusyAsync();
                if (broadcastMsg != null && !string.IsNullOrWhiteSpace(broadcastMsg))
                    await AppOverlayControl.ShowInfoAsync(broadcastMsg);
            }
            Close();
        }
        private async Task ShowErrorAndCloseAsync(string message)
        {
            await AppOverlayControl.HideBusyAsync();
            await AppOverlayControl.ShowErrorAsync(message);
            AddinModule.CurrentInstance.RibLogin.Enabled = true;
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

                    var tcs = new TaskCompletionSource<bool>();

                    void Handler(object sender, CoreWebView2NavigationCompletedEventArgs e)
                    {
                        webView.CoreWebView2.NavigationCompleted -= Handler;
                        tcs.TrySetResult(true);
                    }

                    webView.CoreWebView2.NavigationCompleted += Handler;

                    try
                    {
                        webView.CoreWebView2.Navigate(AppState.Instance.LoginUrl + "?finance_excel=Y");
                    }
                    catch (TaskCanceledException)
                    {
                        LogUtility.LogWarn("Navigation was cancelled by the user.");
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
                LogUtility.LogException(ex, "Login: NavigateToServer");
            }
            finally
            {
                // 🟢 Hide busy overlay safely on UI thread
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
                string apiUrl = $"{AppState.Instance.LoginUrl}{AppConstants.RestSecure}finance-cubes";
                string cubeResponse = await ApiHelper.ServerAPI(apiUrl, "Form", "", "POST", ct);

                // Extract error message FIRST if it's a failed JSON response
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
        // Add this helper method
        private static string? ExtractErrorMessage(string response)
        {
            try
            {
                using JsonDocument doc = JsonDocument.Parse(response);
                JsonElement root = doc.RootElement;

                // Check if status is failed AND has msg
                if (root.TryGetProperty("status", out JsonElement statusElem) &&
                    statusElem.GetString()?.Equals("failed", StringComparison.OrdinalIgnoreCase) == true &&
                    root.TryGetProperty("msg", out JsonElement msgElem))
                {
                    return msgElem.GetString() ?? "Unknown error from server";
                }
            }
            catch (JsonException ex)
            {
                LogUtility.LogException(ex, "Failed to parse error message JSON");
                LogUtility.LogRawJson("GLLogin.ExtractErrorMessage", response);
                // Not valid JSON, return null
            }

            return null;
        }
        private static bool IsValidCubeResponse(string response, string apiUrl)
        {
            // Quick content checks first (fast fails)
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

            // JSON validation
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
                LogUtility.LogRawJson("CubeResponse", response);
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
            LogUtility.LogWarn($"GetDataFromApi|API: {apiUrl}");
            if (ex != null)
            {
                LogUtility.LogException(ex, "GetDataFromApi|Exception");
            }
            if (!string.IsNullOrEmpty(response))
                LogUtility.LogWarn($"GetDataFromApi|Response: {response}");

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
                LogUtility.LogWarn(ex.Message);
                return string.Empty;
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex);
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
                    LogUtility.LogWarn($"Broadcast parsing failed: {result.ErrorMessage}");
                    return string.Empty;
                }

                if (result.Value == null || result.Value.Count == 0)
                    return string.Empty;

                return BuildMessageString(result.Value);
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex, "Error processing broadcast messages.");
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
                LogUtility.LogWarn(ex.Message);
                return string.Empty;
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex, $"Error fetching broadcast messages from {apiUrl}: {ex.Message}");
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
                if (GetEdgeAddinInstance() != null)
                    AddinModule.CurrentInstance.RibLogin.Enabled = false;
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex);
            }
        }

        private void BtnClose_Click(object sender, RoutedEventArgs e)
        {
            LogUtility.LogDebug("GLLogin.BtnClose_Click invoked");
            Close();
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

