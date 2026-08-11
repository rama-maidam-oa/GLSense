// GLAbout.xaml.cs in GLSense.Addin.Core
// Port of GLSense\Views\GLAbout.xaml.cs (FinalWorkingCode) - the about/version-
// compatibility-checker dialog opened by the RibAbout ribbon button (ribbon wiring
// itself is out of scope here - see PORTING_GUIDE.md / this group's task notes).
//
// Adjustments made when porting into this project's architecture:
//   - Base class DpiAwareWindow -> BaseWindow (same as every other window in this
//     project).
//   - EnhancedDragDropHelper.EnableWindowDrag(this) -> the dedicated
//     TitleBar_MouseLeftButtonDown handler already present on every other window here.
//   - LogUtility.* (static) -> ServiceLocator.Logger?.*.
//   - AppPaths.TempUrlsPath -> ServiceLocator.Paths.UrlsDirectory (this project's
//     IPathProvider - already used this way by GLLogin.xaml.cs / GLServerConfiguration).
//   - data:AppConstants.DefaultVersion / DefaultCommitDate (XAML {x:Static} bindings to
//     hardcoded constants that don't exist in this project's AppConstants.cs) -> this
//     project already exposes the *actual* running version/release date through
//     ServiceLocator.Version / ServiceLocator.ReleaseDate (populated from the host's
//     IGLSenseContext - see Infrastructure\ServiceLocator.cs, already used the same way by
//     Helpers\SQLiteHelper.cs to resolve the versioned DLL folder). Using those instead of
//     reintroducing a hardcoded AppConstants.DefaultVersion is both more correct (reflects
//     the real running version) and avoids duplicating version data that already has a
//     single source of truth in this project. txtVersion/txtBuildDate are set in the
//     constructor (with a try/catch, since ServiceLocator throws if not yet initialized)
//     instead of via XAML static binding.
//   - The version-compatibility check (CheckUrlCompatibility) compares the server's
//     reported glSenseVersion against ServiceLocator.Version for the same reason.
//   - The logo <Image Source="/GLSense;component/Images/orbit_logo.png"/> was originally
//     replaced with a plain PackIconFontAwesome badge because no Images folder existed in
//     this project. Per user request to sync this window with FinalWorkingCode, the actual
//     orbit_logo.png asset has since been copied into GLSense.Addin.Core\Images\ (see the
//     .csproj's Resource item) and GLAbout.xaml now references it via
//     "/GLSense.Addin.Core;component/Images/orbit_logo.png", matching the original.
using GLSense.Addin.Core.Helpers;
using GLSense.Addin.Core.Infrastructure;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Net.Http;
using System.Net.Sockets;
using System.Security.Authentication;
using System.Text.Json;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Input;
using System.Xml.Linq;
using GLSense.Addin.Core.Utilities;

namespace GLSense.Addin.Core.Views
{
#nullable enable
    /// <summary>
    /// Interaction logic for GLAbout.xaml
    /// </summary>
    public partial class GLAbout : BaseWindow
    {
        private readonly ObservableCollection<InstanceCompatibility> instances;
        private readonly string xmlFilePath = ServiceLocator.Paths.UrlsDirectory;

        public GLAbout()
        {
            InitializeComponent();
            ServiceLocator.Logger?.LogDebug("GLAbout constructor invoked");

            DataContext = this;

            // "Instance" (the only column) fills the whole grid width instead of leaving a
            // blank gap now that it's Width="Auto" (see DataGridColumnFillHelper for why the
            // star-width column was removed).
            DataGridColumnFillHelper.EnableFillColumn(dgInstances, dgInstances.Columns[0]);

            // Initialize the collection correctly
            instances = new ObservableCollection<InstanceCompatibility>();
            dgInstances.ItemsSource = instances;

            SetVersionAndBuildDateText();
        }

        private void SetVersionAndBuildDateText()
        {
            try
            {
                txtVersion.Text = $"Version : {ServiceLocator.Version}";
                txtBuildDate.Text = $"Build Date : {FormatBuildDate(ServiceLocator.ReleaseDate)}";
            }
            catch (Exception ex)
            {
                txtVersion.Text = "Version : Unknown";
                txtBuildDate.Text = "Build Date : Unknown";
                ServiceLocator.Logger?.LogException(ex, "Failed to resolve version/release date for GLAbout");
            }
        }

        /// <summary>
        /// Manifest.json's releaseDate is written by GLSense.Addin.Core\post_build.cmd as a
        /// local-time "yyyy-MM-ddTHH:mm:ss" string (see CLAUDE.md - deliberately local, no
        /// "Z"/UTC suffix, so it always reads as this machine's own clock). Parsed here with
        /// no timezone-adjustment styles (DateTimeStyles.None) so it's displayed exactly as
        /// the local time it already is, not shifted again. Falls back to the raw string
        /// as-is if it doesn't parse (e.g. an older manifest.json, or the "Unknown" seed
        /// value before ServiceLocator/Context.ReleaseDate is populated).
        /// </summary>
        private static string FormatBuildDate(string rawReleaseDate)
        {
            if (string.IsNullOrWhiteSpace(rawReleaseDate))
                return "Unknown";

            if (DateTime.TryParse(
                    rawReleaseDate,
                    CultureInfo.InvariantCulture,
                    DateTimeStyles.None,
                    out DateTime parsed))
            {
                return parsed.ToString("dd-MMM-yyyy hh:mm tt", CultureInfo.InvariantCulture);
            }

            // Not a parseable date/time (e.g. a legacy "dd-MMM-yyyy"-only value) - show as-is
            // rather than hiding a real value behind "Unknown".
            return rawReleaseDate;
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

        private async void AboutWindow_Loaded(object sender, RoutedEventArgs e)
        {
            ServiceLocator.Logger?.LogDebug("GLAbout.AboutWindow_Loaded invoked");
            // Start compatibility checking
            await CheckInstanceCompatibility();

            // BaseWindow.OnLoaded's SizeToContent resettle already ran (synchronously)
            // before this async chain populated dgInstances - so it measured an empty
            // grid. Resettle again now that real rows are in place. See CLAUDE.md
            // section 1.4b (GLCubeDetails) for the full history of this pattern.
            ForceSizeToContentResettle();
            PumpDispatcherFrame();
        }

        private async Task CheckInstanceCompatibility()
        {
            try
            {
                // Ensure UI shows the indeterminate marquee before doing network work
                await Application.Current.Dispatcher.InvokeAsync(() =>
                {
                    progressPanel.Visibility = Visibility.Visible;
                    progressBar.IsIndeterminate = true;
                    txtProgress.Text = "Starting compatibility checks...";
                }, System.Windows.Threading.DispatcherPriority.Background);

                // Give the UI a tiny moment to render the marquee
                await Task.Yield();

                var urlInstances = LoadInstancesFromXml();
                int totalInstances = urlInstances.Count;

                // Clear existing data on UI thread
                await Application.Current.Dispatcher.InvokeAsync(() =>
                {
                    instances.Clear();
                }, System.Windows.Threading.DispatcherPriority.Background);

                if (totalInstances == 0)
                {
                    await Application.Current.Dispatcher.InvokeAsync(() =>
                    {
                        instances.Add(new InstanceCompatibility { Instance = "No instances configured", IsCompatible = false });
                        progressPanel.Visibility = Visibility.Collapsed;
                        progressBar.IsIndeterminate = false;
                        txtProgress.Text = "No instances configured";
                    }, System.Windows.Threading.DispatcherPriority.Background);
                    return;
                }

                int processed = 0;

                // Use indeterminate marquee while doing the checks (network I/O off the UI thread)
                foreach (var instance in urlInstances)
                {
                    // Update status text on UI thread before the network call so user sees it immediately
                    await Application.Current.Dispatcher.InvokeAsync(() =>
                    {
                        txtProgress.Text = $"Checking {instance.Name}... ({processed + 1}/{totalInstances})";
                    }, System.Windows.Threading.DispatcherPriority.Background);

                    // Run the potentially slow operation off the UI thread
                    bool isCompatible = false;
                    try
                    {
                        isCompatible = await Task.Run(async () => await CheckUrlCompatibility(instance.Address).ConfigureAwait(false)).ConfigureAwait(false);
                    }
                    catch (Exception ex)
                    {
                        // Unexpected per-instance error: log and continue (do not add a global "Error" row)
                        ServiceLocator.Logger?.LogError($"Per-instance compatibility check failed for '{instance.Address}': {ex.Message}");
                        isCompatible = false;
                    }

                    processed++;

                    // Add result to collection on UI thread
                    await Application.Current.Dispatcher.InvokeAsync(() =>
                    {
                        instances.Add(new InstanceCompatibility { Instance = instance.Address, IsCompatible = isCompatible });
                    }, System.Windows.Threading.DispatcherPriority.Background);

                    // Small pause to allow UI to breathe and show updates
                    await Task.Delay(150).ConfigureAwait(false);
                }

                // Finish: stop marquee and hide panel
                await Application.Current.Dispatcher.InvokeAsync(() =>
                {
                    progressBar.IsIndeterminate = false;
                    progressPanel.Visibility = Visibility.Collapsed;
                    txtProgress.Text = "Compatibility check completed";
                }, System.Windows.Threading.DispatcherPriority.Background);
            }
            catch (Exception ex)
            {
                // Log and update UI but DO NOT append an "Error" row automatically.
                // Adding an "Error" item previously caused an extra row to appear in successful runs when an outer exception occurred.
                ServiceLocator.Logger?.LogException(ex);
                await Application.Current.Dispatcher.InvokeAsync(() =>
                {
                    progressPanel.Visibility = Visibility.Collapsed;
                    progressBar.IsIndeterminate = false;
                    txtProgress.Text = "Error checking instances (see logs)";
                }, System.Windows.Threading.DispatcherPriority.Background);
            }
        }

        private List<UrlInstance> LoadInstancesFromXml()
        {
            var instances1 = new List<UrlInstance>();

            try
            {
                if (!File.Exists(xmlFilePath))
                    return instances1;

                XDocument doc = XDocument.Load(xmlFilePath);

                foreach (var urlElement in doc.Descendants("URL"))
                {
                    instances1.Add(new UrlInstance
                    {
                        Name = urlElement.Element("Name")?.Value ?? "",
                        Address = urlElement.Element("Address")?.Value ?? "",
                        IsDefault = bool.Parse(urlElement.Element("DefaultURL")?.Value ?? "false")
                    });
                }
            }
            catch (Exception ex)
            {
                // Log error but don't throw - we'll show empty grid
                ServiceLocator.Logger?.LogError($"Error loading instances: {ex.Message}");
            }

            return instances1;
        }

        private async Task<bool> CheckUrlCompatibility(string url)
        {
            try
            {
                var handler = new HttpClientHandler()
                {
                    SslProtocols = System.Security.Authentication.SslProtocols.Tls12 | System.Security.Authentication.SslProtocols.Tls13,
                    ServerCertificateCustomValidationCallback = StrictCertificateValidator.Validate
                };

                using var httpClient = new HttpClient(handler);
                // Was Timeout.InfiniteTimeSpan - bounded so a transient blip triggers the
                // retry below instead of hanging indefinitely. Ported from
                // FinalWorkingCode's identical fix.
                httpClient.Timeout = TimeSpan.FromSeconds(30);
                httpClient.DefaultRequestHeaders.ExpectContinue = false;

                var ReqURL = url.Trim();

                var patterns = new string[] { "/bypass-saml-login-flow", "/bypass-sso-login-flow" };

                foreach (var pattern in patterns)
                {
                    ReqURL = Regex.Replace(ReqURL, Regex.Escape(pattern), "", RegexOptions.IgnoreCase);
                }

                // Bounded timeout + one retry on a transient failure, matching the
                // resilience ApiHelper.ServerAPI already has for the main API path -
                // this check previously had neither, so any transient blip showed up
                // identically to a genuinely incompatible/unreachable instance. Ported
                // from FinalWorkingCode's identical fix.
                const int maxAttempts = 2;
                for (int attempt = 1; attempt <= maxAttempts; attempt++)
                {
                try
                {
                    // Log request
                    ServiceLocator.Logger?.LogDebug($"Sending request: {ReqURL}");

                    // Create request object
                    var request = new HttpRequestMessage(HttpMethod.Get, $"{ReqURL}/rest/public/orbit-version");

                    // Send request and get response
                    var responseMessage = await httpClient.SendAsync(request).ConfigureAwait(false);

                    // Capture status code and headers
                    var statusCode = (int)responseMessage.StatusCode;
                    var responseHeaders = responseMessage.Headers.ToString();

                    // Read full response body
                    var responseBody = await responseMessage.Content.ReadAsStringAsync().ConfigureAwait(false);

                    // Log detailed response
                    ServiceLocator.Logger?.LogDebug($"Response from: {ReqURL}/rest/public/orbit-version");
                    ServiceLocator.Logger?.LogDebug($"Status Code: {statusCode}");
                    ServiceLocator.Logger?.LogDebug($"Headers: {responseHeaders}");
                    ServiceLocator.Logger?.LogDebug($"Response Body: {responseBody}");

                    // Parse JSON (with additional logging for unexpected structures)
                    try
                    {
                        if (string.IsNullOrWhiteSpace(responseBody))
                        {
                            ServiceLocator.Logger?.LogError($"Empty response body from {ReqURL}");
                            return false;
                        }

                        using var doc = JsonDocument.Parse(responseBody);
                        var root = doc.RootElement;

                        ServiceLocator.Logger?.LogDebug($"Parsed JSON: {root.GetRawText()}");

                        var glSenseVersion = root
                            .EnumerateObject()
                            .FirstOrDefault(p =>
                                string.Equals(p.Name, "verionInfo",
                                              StringComparison.OrdinalIgnoreCase))
                            .Value
                            .EnumerateObject()
                            .FirstOrDefault(p =>
                                string.Equals(p.Name, "glSenseVersion",
                                              StringComparison.OrdinalIgnoreCase))
                            .Value
                            .GetString();

                        if (string.Equals(glSenseVersion,
                                          ServiceLocator.Version,
                                          StringComparison.Ordinal))
                        {
                            return true;
                        }

                        ServiceLocator.Logger?.LogDebug(
                            $"Version mismatch or missing: Expected='{ServiceLocator.Version}', " +
                            $"Received='{glSenseVersion ?? "(null)"}'");

                        return false;
                    }
                    catch (JsonException jsonEx)
                    {
                        ServiceLocator.Logger?.LogError(
                            $"JSON Parsing Error at {ReqURL}: {jsonEx.Message} | Raw Response: {responseBody}");
                        return false;
                    }
                    catch (Exception ex)
                    {
                        ServiceLocator.Logger?.LogError(
                            $"Unexpected error parsing response from {ReqURL}: {ex.Message} | Raw Response: {responseBody}");
                        return false;
                    }
                }
                catch (HttpRequestException ex) when (attempt < maxAttempts)
                {
                    ServiceLocator.Logger?.LogWarn($"Network error for {ReqURL} (attempt {attempt}/{maxAttempts}) - retrying in 1s: {ex.Message}");
                    await Task.Delay(1000).ConfigureAwait(false);
                }
                // HttpClient throws this same exception type both for a genuine request
                // timeout AND for an explicit CancellationToken cancellation - only retry
                // the former (a user-cancelled check should stop immediately, not retry).
                catch (TaskCanceledException ex) when (attempt < maxAttempts && !ex.CancellationToken.IsCancellationRequested)
                {
                    ServiceLocator.Logger?.LogWarn($"Timeout for {ReqURL} (attempt {attempt}/{maxAttempts}) - retrying in 1s: {ex.Message}");
                    await Task.Delay(1000).ConfigureAwait(false);
                }
                catch (HttpRequestException ex)
                {
                    LogInstanceCheckFailure(ReqURL, ex);
                    return false;
                }
                catch (TaskCanceledException ex)
                {
                    if (ex.CancellationToken.IsCancellationRequested)
                    {
                        ServiceLocator.Logger?.LogWarn($"Instance check for {ReqURL} was cancelled: {ex.Message}");
                    }
                    else
                    {
                        ServiceLocator.Logger?.LogError($"Instance not reachable (request timed out) for {ReqURL}: {ex.Message} | StackTrace: {ex.StackTrace}");
                    }
                    return false;
                }
                catch (Exception ex)
                {
                    ServiceLocator.Logger?.LogError($"Unexpected error checking instance {ReqURL}: {ex.GetType().Name}: {ex.Message} | StackTrace: {ex.StackTrace}");
                    return false;
                }
                }

                return false;
            }
            catch (Exception ex)
            {
                ServiceLocator.Logger?.LogException(ex, "GLAbout.CheckUrlCompatibility");
                return false;
            }
        }

        /// <summary>
        /// HttpClient always wraps the real transport/TLS failure inside an
        /// HttpRequestException whose own .Message ("An error occurred while sending the
        /// request.") is rarely useful on its own - the actual cause lives in the
        /// InnerException chain. Walks that chain and logs one clearly-labeled line so the
        /// logs distinguish "instance not reachable" (DNS/connection-refused/host down),
        /// "certificate/security" failures (TLS handshake / StrictCertificateValidator
        /// rejection - see that class's own detailed chain-status log lines just above this
        /// one in the log), and anything else, instead of one generic "Network error" for
        /// every case.
        /// </summary>
        private static void LogInstanceCheckFailure(string url, HttpRequestException ex)
        {
            Exception root = ex;
            while (root.InnerException != null)
                root = root.InnerException;

            bool looksLikeCertOrSecurityIssue =
                root is AuthenticationException ||
                root is System.Security.Cryptography.CryptographicException ||
                ContainsAny(root.Message, "certificate", "SSL", "TLS", "trust", "authentication");

            if (looksLikeCertOrSecurityIssue)
            {
                ServiceLocator.Logger?.LogError(
                    $"Certificate/security error connecting to {url}: {root.GetType().Name}: {root.Message} " +
                    "(see the TLS validation log lines above for chain/policy details) | " +
                    $"StackTrace: {ex.StackTrace}");
                return;
            }

            if (root is SocketException socketEx)
            {
                ServiceLocator.Logger?.LogError(
                    $"Instance not reachable at {url}: {socketEx.SocketErrorCode} - {socketEx.Message} | " +
                    $"StackTrace: {ex.StackTrace}");
                return;
            }

            ServiceLocator.Logger?.LogError(
                $"Network error connecting to {url}: {ex.Message} | " +
                $"InnerException: {ex.InnerException?.Message ?? "(none)"} | StackTrace: {ex.StackTrace}");
        }

        private static bool ContainsAny(string haystack, params string[] needles)
        {
            if (string.IsNullOrEmpty(haystack))
                return false;

            foreach (var needle in needles)
            {
                if (haystack.IndexOf(needle, StringComparison.OrdinalIgnoreCase) >= 0)
                    return true;
            }

            return false;
        }

        private void SupportLink_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            try
            {
                // Open support URL in default browser
                System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo
                {
                    FileName = "https://www.orbitanalytics.com",
                    UseShellExecute = true
                });
            }
            catch (Exception ex)
            {
                ServiceLocator.Logger?.LogError($"Cannot open browser: {ex.Message}");
            }
        }

        private void BtnClose_Click(object sender, RoutedEventArgs e)
        {
            ServiceLocator.Logger?.LogDebug("GLAbout.BtnClose_Click invoked - closing window");
            this.Close();
        }
    }

    public class InstanceCompatibility : INotifyPropertyChanged
    {
        private string? _instance;
        private bool _isCompatible;

        public string? Instance
        {
            get => _instance;
            set
            {
                _instance = value;
                OnPropertyChanged(nameof(Instance));
            }
        }
        public bool IsCompatible
        {
            get => _isCompatible;
            set
            {
                _isCompatible = value;
                OnPropertyChanged(nameof(IsCompatible));
            }
        }

        public event PropertyChangedEventHandler? PropertyChanged;

        protected virtual void OnPropertyChanged(string propertyName)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
    }
#nullable disable
}
