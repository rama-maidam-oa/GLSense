// GLReloadSourcePicker.xaml.cs in GLSense\Views
using GLSense.Contracts;
using GLSense.Shared;
using System;
using System.IO;
using System.Linq;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Runtime.InteropServices;
using System.Security.Cryptography;
using System.Threading.Tasks;
using System.Windows;

namespace GLSense
{
    public partial class GLReloadSourcePicker : Window
    {
        public string SelectedSource { get; private set; }

        private string _candidateManifestPath;
        private string _candidateZipPath;
        private bool _isValidated;
        private bool _onlineAvailable;

        // Small, deliberate pacing between Offline validation steps (which are
        // otherwise near-instant local file checks) so the step log is actually
        // readable instead of flashing by in one frame - same technique already
        // used elsewhere in this codebase for busy-overlay visibility (see
        // CLAUDE.md 27.2's Dispatcher.Yield pattern). Online's own steps never
        // need this - real network/download latency already paces them.
        private const int OfflineStepDelayMs = 150;

        public GLReloadSourcePicker()
        {
            InitializeComponent();
            SourceInitialized += (s, e) => WindowChromeHelper.RemoveMinimizeMaximizeButtons(this);
            InitializeModeAvailability();
        }

        private void InitializeModeAvailability()
        {
            try
            {
                var loginInfo = GlobalsEx.Addin?.GetLoginInfo();
                _onlineAvailable = loginInfo != null && loginInfo.IsLoggedIn && !string.IsNullOrWhiteSpace(loginInfo.LoginUrl);
            }
            catch
            {
                // Defensive: an older historical Addin.Core build reloaded via the
                // Release History browser may not implement GetLoginInfo at all.
                _onlineAvailable = false;
            }

            RbOnline.IsEnabled = _onlineAvailable;
            if (_onlineAvailable) RbOnline.IsChecked = true;
            else RbOffline.IsChecked = true;
        }

        // Not async: switching mode no longer triggers a scan itself (see the Offline
        // branch below), so there is nothing left in this method to await.
        private void Mode_Checked(object sender, RoutedEventArgs e)
        {
            if (OnlinePanel == null || OfflinePanel == null) return; // fires during InitializeComponent

            bool isOnline = RbOnline.IsChecked == true;
            OnlinePanel.Visibility = isOnline ? Visibility.Visible : Visibility.Collapsed;
            OfflinePanel.Visibility = isOnline ? Visibility.Collapsed : Visibility.Visible;

            ResetValidation();

            if (isOnline)
            {
                TxtStatus.Text = "Click \"Check for Update\" to begin.";
            }
            else
            {
                // Deliberately no default folder and no auto-scan here anymore - this
                // used to default straight to the Downloads folder and immediately
                // scan it the instant the window opened (Downloads is very often the
                // largest, most heavily-populated folder on a machine, so this could
                // visibly delay the window even showing up). Offline mode now starts
                // empty; scanning only ever happens when the user explicitly clicks
                // Browse... (which still conveniently defaults ITS OWN starting folder
                // to Downloads - see BtnBrowse_Click - without scanning anything until
                // a folder is actually chosen).
                TxtStatus.Text = "Click \"Browse...\" to select a folder to scan for a manifest.json + zip pair.";
            }
        }

        // Clears to empty (not a static placeholder) since TxtStatus is now an
        // append-based step log, not a single overwritten status line - a
        // placeholder set here would sit above the real log lines as a stale
        // first entry. Callers that need a placeholder (e.g. switching to
        // Online mode, which doesn't immediately run anything) set one
        // explicitly right after calling this.
        private void ResetValidation()
        {
            _isValidated = false;
            _candidateManifestPath = null;
            _candidateZipPath = null;
            BtnReload.IsEnabled = false;
            TxtStatus.Text = string.Empty;
        }

        // ------------------------------------------------------------------
        // Busy state + step log helpers
        // ------------------------------------------------------------------

        // This host's WPF window doesn't reliably run under a DispatcherSynchronizationContext
        // (see the reference memory on WPF dispatcher-thread-loss in this VSTO add-in), so an
        // `await` continuation - including the plain `await Task.Delay(...)` pacing used in the
        // Offline flow - can resume on a background ThreadPool thread instead of hopping back to
        // the UI thread automatically. Every method here that touches a UI element must be
        // callable from any thread - guard with CheckAccess()/synchronous Invoke (not
        // InvokeAsync) rather than assuming the caller is already on the UI thread.
        private void SetBusy(bool busy)
        {
            if (!Dispatcher.CheckAccess())
            {
                Dispatcher.Invoke(() => SetBusy(busy));
                return;
            }

            BusyProgress.Visibility = busy ? Visibility.Visible : Visibility.Collapsed;
            BtnCheckOnline.IsEnabled = !busy;
            BtnBrowse.IsEnabled = !busy;
            BtnReload.IsEnabled = !busy && _isValidated;
            BtnCancel.IsEnabled = !busy;
            RbOnline.IsEnabled = !busy && _onlineAvailable;
            RbOffline.IsEnabled = !busy;
        }

        private void AppendLine(string prefix, string message)
        {
            if (!Dispatcher.CheckAccess())
            {
                Dispatcher.Invoke(() => AppendLine(prefix, message));
                return;
            }

            if (!string.IsNullOrEmpty(TxtStatus.Text)) TxtStatus.AppendText(Environment.NewLine);
            TxtStatus.AppendText($"{prefix} {message}");
            TxtStatus.ScrollToEnd();
        }

        /// <summary>Routine progress narration. UI-only plus LogDebug (gated on
        /// DebugMode, matching every other routine narration line in this host
        /// project) - not the kind of thing worth always-on logging.</summary>
        private void LogStep(string message)
        {
            AppendLine("•", message);
            GlobalsEx.Context?.Logger?.LogDebug($"GLReloadSourcePicker: {message}");
        }

        /// <summary>A step that completed successfully - same log level as
        /// LogStep, just a distinct glyph so the final "ready to reload" line
        /// stands out from the routine narration above it.</summary>
        private void LogSuccess(string message)
        {
            AppendLine("✓", message);
            GlobalsEx.Context?.Logger?.LogDebug($"GLReloadSourcePicker: {message}");
        }

        /// <summary>A blocking-but-expected condition (not logged in, no
        /// manifest+zip pair in the folder, missing checksum) - always written
        /// to the log file (LogWarn), not just DebugMode-gated.</summary>
        private void LogWarning(string message)
        {
            AppendLine("⚠", message);
            GlobalsEx.Context?.Logger?.LogWarn($"GLReloadSourcePicker: {message}");
        }

        /// <summary>A genuine failure (parse error, checksum mismatch, network
        /// exception, staging copy failure) - always written to the log file.</summary>
        private void LogFailure(string message, Exception ex = null)
        {
            AppendLine("✗", message);
            if (ex != null)
                GlobalsEx.Context?.Logger?.LogException(ex, $"GLReloadSourcePicker: {message}");
            else
                GlobalsEx.Context?.Logger?.LogError($"GLReloadSourcePicker: {message}");
        }

        private void PromptNoUpdate(string candidateVersion, string candidateReleaseDate)
        {
            if (!Dispatcher.CheckAccess())
            {
                Dispatcher.Invoke(() => PromptNoUpdate(candidateVersion, candidateReleaseDate));
                return;
            }

            MessageBox.Show(
                this,
                $"No updates available.\n\nYou already have the latest version: {GlobalsEx.Context?.Version} (released {GlobalsEx.Context?.ReleaseDate}).\n\nChecked release: {candidateVersion} ({candidateReleaseDate}).",
                "No Updates Available",
                MessageBoxButton.OK,
                MessageBoxImage.Information);
        }

        // ------------------------------------------------------------------
        // Offline flow
        // ------------------------------------------------------------------

        private async void BtnBrowse_Click(object sender, RoutedEventArgs e)
        {
            using (var dialog = new System.Windows.Forms.FolderBrowserDialog())
            {
                dialog.SelectedPath = Directory.Exists(TxtFolder.Text) ? TxtFolder.Text : GetDownloadsFolder();
                if (dialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    TxtFolder.Text = dialog.SelectedPath;
                    await ScanFolderAsync(dialog.SelectedPath);
                }
            }
        }

        private async Task ScanFolderAsync(string folder)
        {
            ResetValidation();
            SetBusy(true);

            try
            {
                LogStep($"Scanning folder: {folder}");
                await Task.Delay(OfflineStepDelayMs);

                if (!Directory.Exists(folder))
                {
                    LogWarning($"Folder not found: {folder}");
                    return;
                }

                string manifestPath = Directory.GetFiles(folder, "manifest*.json")
                    .OrderByDescending(File.GetLastWriteTimeUtc)
                    .FirstOrDefault();
                string zipPath = Directory.GetFiles(folder, "v*.zip")
                    .OrderByDescending(File.GetLastWriteTimeUtc)
                    .FirstOrDefault();

                if (manifestPath == null || zipPath == null)
                {
                    LogWarning("No manifest.json + zip pair found in this folder.");
                    return;
                }

                LogStep($"Found {Path.GetFileName(manifestPath)} and {Path.GetFileName(zipPath)}.");
                await Task.Delay(OfflineStepDelayMs);

                await ValidateCandidateAsync(manifestPath, zipPath);
            }
            catch (Exception ex)
            {
                LogFailure($"Folder scan failed: {ex.Message}", ex);
            }
            finally
            {
                SetBusy(false);
            }
        }

        private async Task ValidateCandidateAsync(string manifestPath, string zipPath)
        {
            LogStep("Parsing manifest.json...");
            await Task.Delay(OfflineStepDelayMs);

            var parser = new VersionParser();
            var result = parser.ParseVersionFile(manifestPath);

            if (!result.Success)
            {
                LogFailure($"Could not parse manifest.json: {result.ErrorMessage}");
                return;
            }

            if (string.IsNullOrWhiteSpace(result.Checksum))
            {
                LogWarning("manifest.json has no checksum recorded - cannot verify, refusing to reload.");
                return;
            }

            LogStep("Verifying checksum...");
            await Task.Delay(OfflineStepDelayMs);
            string actualChecksum = ComputeSha256(zipPath);
            if (!string.Equals(actualChecksum, result.Checksum, StringComparison.OrdinalIgnoreCase))
            {
                LogFailure($"Checksum mismatch - the zip may be corrupt or incomplete. Expected {result.Checksum}, got {actualChecksum}.");
                return;
            }
            LogSuccess("Checksum verified.");

            LogStep($"Checking version - candidate is {result.Version} (released {result.ReleaseDate})...");
            await Task.Delay(OfflineStepDelayMs);
            if (!IsStrictlyNewer(result.ReleaseDate))
            {
                LogStep("No newer version available.");
                PromptNoUpdate(result.Version, result.ReleaseDate.ToString());
                return;
            }

            _candidateManifestPath = manifestPath;
            _candidateZipPath = zipPath;
            _isValidated = true;
            LogSuccess($"Ready to reload: version {result.Version}, released {result.ReleaseDate} ({new FileInfo(zipPath).Length / 1024} KB).");
        }

        // VersionParseResult.ReleaseDate is already a parsed DateTime (see
        // VersionParser.ParseVersionJson/ParseVersionFile) - takes DateTime directly,
        // does not re-parse a string.
        private bool IsStrictlyNewer(DateTime candidateReleaseDate)
        {
            string baseline = GlobalsEx.Context?.ReleaseDate;
            if (string.IsNullOrWhiteSpace(baseline)) return true; // nothing loaded yet

            if (!DateTime.TryParse(baseline, out var baselineDate)) return true;

            return candidateReleaseDate > baselineDate;
        }

        // ------------------------------------------------------------------
        // Online flow
        // ------------------------------------------------------------------

        private async void BtnCheckOnline_Click(object sender, RoutedEventArgs e)
        {
            ResetValidation();
            SetBusy(true);

            try
            {
                LogStep("Checking login status...");

                LoginInfo loginInfo;
                try { loginInfo = GlobalsEx.Addin?.GetLoginInfo(); }
                catch (Exception ex)
                {
                    loginInfo = null;
                    LogWarning($"Could not read login info from Addin.Core: {ex.Message}");
                }

                if (loginInfo == null || !loginInfo.IsLoggedIn || string.IsNullOrWhiteSpace(loginInfo.LoginUrl))
                {
                    LogWarning("Not logged in - switch to Offline mode.");
                    return;
                }

                string url = loginInfo.LoginUrl.TrimEnd('/') + "/glsense/projectdlls";
                LogStep($"Fetching latest release info from {url}...");

                using (var client = new HttpClient())
                {
                    client.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", loginInfo.LoginToken);

                    string responseJson = await client.GetStringAsync(url);
                    LogStep("Parsing server response...");

                    var parser = new VersionParser();
                    var result = parser.ParseVersionJson(responseJson);

                    if (!result.Success)
                    {
                        LogFailure($"Could not parse server response: {result.ErrorMessage}");
                        return;
                    }

                    LogStep($"Server has version {result.Version} (released {result.ReleaseDate}). Comparing against the currently loaded release ({GlobalsEx.Context?.Version} / {GlobalsEx.Context?.ReleaseDate})...");

                    if (!IsStrictlyNewer(result.ReleaseDate))
                    {
                        LogStep("No newer version available.");
                        PromptNoUpdate(result.Version, result.ReleaseDate.ToString());
                        return;
                    }

                    LogStep($"New version found: {result.Version}. Downloading update...");
                    string tempZip = Path.Combine(Path.GetTempPath(), $"GLSenseOnline_{Guid.NewGuid():N}.zip");
                    var zipBytes = await client.GetByteArrayAsync(result.DownloadUrl);
                    File.WriteAllBytes(tempZip, zipBytes);
                    LogStep($"Downloaded {zipBytes.Length / 1024} KB.");

                    LogStep("Verifying checksum...");
                    string actualChecksum = ComputeSha256(tempZip);
                    if (!string.Equals(actualChecksum, result.Checksum, StringComparison.OrdinalIgnoreCase))
                    {
                        LogFailure($"Downloaded zip failed checksum verification - expected {result.Checksum}, got {actualChecksum}. Not reloading.");
                        File.Delete(tempZip);
                        return;
                    }
                    LogSuccess("Checksum verified.");

                    string tempManifest = Path.Combine(Path.GetTempPath(), $"GLSenseOnline_{Guid.NewGuid():N}.json");
                    File.WriteAllText(tempManifest, responseJson);

                    _candidateManifestPath = tempManifest;
                    _candidateZipPath = tempZip;
                    _isValidated = true;
                    LogSuccess($"Ready to reload: version {result.Version}, released {result.ReleaseDate}.");
                }
            }
            catch (Exception ex)
            {
                LogFailure($"Online check failed: {ex.Message}", ex);
            }
            finally
            {
                SetBusy(false);
            }
        }

        // ------------------------------------------------------------------
        // Reload / Cancel
        // ------------------------------------------------------------------

        private void BtnReload_Click(object sender, RoutedEventArgs e)
        {
            if (!_isValidated || _candidateManifestPath == null || _candidateZipPath == null) return;

            try
            {
                string manifestDir = GlobalsEx.Context.Paths.ManifestDirectory;
                if (!Directory.Exists(manifestDir)) Directory.CreateDirectory(manifestDir);

                string zipDestination = Path.Combine(manifestDir, Path.GetFileName(_candidateZipPath));
                bool zipAlreadyStaged = string.Equals(
                    Path.GetFullPath(_candidateZipPath),
                    Path.GetFullPath(zipDestination),
                    StringComparison.OrdinalIgnoreCase);

                // Delete every OTHER zip in the Manifest folder, but never the candidate
                // itself - if the user browsed Offline directly to the Manifest folder,
                // the candidate zip may already be sitting there, and deleting it before
                // the copy below would destroy the very file being staged.
                foreach (var oldZip in Directory.GetFiles(manifestDir, "*.zip"))
                {
                    if (!string.Equals(Path.GetFullPath(oldZip), Path.GetFullPath(_candidateZipPath), StringComparison.OrdinalIgnoreCase))
                        File.Delete(oldZip);
                }

                if (!zipAlreadyStaged)
                    File.Copy(_candidateZipPath, zipDestination, true);

                string manifestDestination = GlobalsEx.Context.Paths.ManifestFile;
                bool manifestAlreadyStaged = string.Equals(
                    Path.GetFullPath(_candidateManifestPath),
                    Path.GetFullPath(manifestDestination),
                    StringComparison.OrdinalIgnoreCase);

                if (!manifestAlreadyStaged)
                    File.Copy(_candidateManifestPath, manifestDestination, true);

                SelectedSource = RbOnline.IsChecked == true ? "Online" : "Offline";
                GlobalsEx.Context?.Logger?.LogDebug($"GLReloadSourcePicker: staged release into Manifest folder from {SelectedSource} mode - proceeding with reload.");
                DialogResult = true;
                Close();
            }
            catch (Exception ex)
            {
                LogFailure($"Failed to stage the new release: {ex.Message}", ex);
            }
        }

        private void BtnCancel_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = false;
            Close();
        }

        private static string ComputeSha256(string filePath)
        {
            using (var sha256 = SHA256.Create())
            using (var stream = File.OpenRead(filePath))
            {
                var hash = sha256.ComputeHash(stream);
                return BitConverter.ToString(hash).Replace("-", "");
            }
        }

        [DllImport("shell32.dll")]
        private static extern int SHGetKnownFolderPath(ref Guid rfid, uint dwFlags, IntPtr hToken, out IntPtr pszPath);

        private static readonly Guid FolderIdDownloads = new Guid("374DE290-123F-4565-9164-39C4925E467B");

        private static string GetDownloadsFolder()
        {
            try
            {
                var guid = FolderIdDownloads;
                if (SHGetKnownFolderPath(ref guid, 0, IntPtr.Zero, out IntPtr pathPtr) == 0)
                {
                    string path = Marshal.PtrToStringUni(pathPtr);
                    Marshal.FreeCoTaskMem(pathPtr);
                    return path;
                }
            }
            catch
            {
                // fall through to the profile-based fallback below
            }
            return Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.UserProfile), "Downloads");
        }
    }
}
