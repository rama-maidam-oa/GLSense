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
using System.Windows;

namespace GLSense
{
    public partial class GLReloadSourcePicker : Window
    {
        public string SelectedSource { get; private set; }

        private string _candidateManifestPath;
        private string _candidateZipPath;
        private bool _isValidated;

        public GLReloadSourcePicker()
        {
            InitializeComponent();
            InitializeModeAvailability();
        }

        private void InitializeModeAvailability()
        {
            bool onlineAvailable;
            try
            {
                var loginInfo = GlobalsEx.Addin?.GetLoginInfo();
                onlineAvailable = loginInfo != null && loginInfo.IsLoggedIn && !string.IsNullOrWhiteSpace(loginInfo.LoginUrl);
            }
            catch
            {
                // Defensive: an older historical Addin.Core build reloaded via the
                // Release History browser may not implement GetLoginInfo at all.
                onlineAvailable = false;
            }

            RbOnline.IsEnabled = onlineAvailable;
            if (onlineAvailable) RbOnline.IsChecked = true;
            else RbOffline.IsChecked = true;
        }

        private void Mode_Checked(object sender, RoutedEventArgs e)
        {
            if (OnlinePanel == null || OfflinePanel == null) return; // fires during InitializeComponent

            bool isOnline = RbOnline.IsChecked == true;
            OnlinePanel.Visibility = isOnline ? Visibility.Visible : Visibility.Collapsed;
            OfflinePanel.Visibility = isOnline ? Visibility.Collapsed : Visibility.Visible;

            ResetValidation();

            if (!isOnline)
            {
                TxtFolder.Text = GetDownloadsFolder();
                ScanFolder(TxtFolder.Text);
            }
        }

        private void ResetValidation()
        {
            _isValidated = false;
            _candidateManifestPath = null;
            _candidateZipPath = null;
            BtnReload.IsEnabled = false;
            TxtStatus.Text = "Select Online or Offline to begin.";
        }

        private void BtnBrowse_Click(object sender, RoutedEventArgs e)
        {
            using (var dialog = new System.Windows.Forms.FolderBrowserDialog())
            {
                dialog.SelectedPath = Directory.Exists(TxtFolder.Text) ? TxtFolder.Text : GetDownloadsFolder();
                if (dialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    TxtFolder.Text = dialog.SelectedPath;
                    ScanFolder(dialog.SelectedPath);
                }
            }
        }

        private void ScanFolder(string folder)
        {
            ResetValidation();

            if (!Directory.Exists(folder))
            {
                TxtStatus.Text = $"Folder not found: {folder}";
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
                TxtStatus.Text = "No manifest.json + zip pair found in this folder.";
                return;
            }

            ValidateCandidate(manifestPath, zipPath);
        }

        private void ValidateCandidate(string manifestPath, string zipPath)
        {
            var parser = new VersionParser();
            var result = parser.ParseVersionFile(manifestPath);

            if (!result.Success)
            {
                TxtStatus.Text = $"Could not parse manifest.json: {result.ErrorMessage}";
                return;
            }

            if (string.IsNullOrWhiteSpace(result.Checksum))
            {
                TxtStatus.Text = "manifest.json has no checksum recorded - cannot verify, refusing to reload.";
                return;
            }

            string actualChecksum = ComputeSha256(zipPath);
            if (!string.Equals(actualChecksum, result.Checksum, StringComparison.OrdinalIgnoreCase))
            {
                TxtStatus.Text = $"Checksum mismatch - the zip may be corrupt or incomplete. Expected {result.Checksum}, got {actualChecksum}.";
                return;
            }

            if (!IsStrictlyNewer(result.ReleaseDate))
            {
                TxtStatus.Text = $"No update available - {result.Version} ({result.ReleaseDate}) is not newer than the currently loaded release ({GlobalsEx.Context?.Version} / {GlobalsEx.Context?.ReleaseDate}).";
                return;
            }

            _candidateManifestPath = manifestPath;
            _candidateZipPath = zipPath;
            _isValidated = true;
            BtnReload.IsEnabled = true;
            TxtStatus.Text = $"Ready to reload: version {result.Version}, released {result.ReleaseDate}.\n{Path.GetFileName(zipPath)} ({new FileInfo(zipPath).Length / 1024} KB)";
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

        private async void BtnCheckOnline_Click(object sender, RoutedEventArgs e)
        {
            ResetValidation();

            LoginInfo loginInfo;
            try { loginInfo = GlobalsEx.Addin?.GetLoginInfo(); }
            catch { loginInfo = null; }

            if (loginInfo == null || !loginInfo.IsLoggedIn || string.IsNullOrWhiteSpace(loginInfo.LoginUrl))
            {
                TxtStatus.Text = "Not logged in - switch to Offline mode.";
                return;
            }

            OnlineProgress.Visibility = Visibility.Visible;
            BtnCheckOnline.IsEnabled = false;

            try
            {
                string url = loginInfo.LoginUrl.TrimEnd('/') + "/glsense/projectdlls";
                using (var client = new HttpClient())
                {
                    client.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", loginInfo.LoginToken);

                    string responseJson = await client.GetStringAsync(url);

                    var parser = new VersionParser();
                    var result = parser.ParseVersionJson(responseJson);

                    if (!result.Success)
                    {
                        TxtStatus.Text = $"Could not parse server response: {result.ErrorMessage}";
                        return;
                    }

                    if (!IsStrictlyNewer(result.ReleaseDate))
                    {
                        TxtStatus.Text = $"No updates available - server has {result.Version} ({result.ReleaseDate}), which is not newer than the currently loaded release.";
                        return;
                    }

                    string tempZip = Path.Combine(Path.GetTempPath(), $"GLSenseOnline_{Guid.NewGuid():N}.zip");
                    var zipBytes = await client.GetByteArrayAsync(result.DownloadUrl);
                    File.WriteAllBytes(tempZip, zipBytes);

                    string actualChecksum = ComputeSha256(tempZip);
                    if (!string.Equals(actualChecksum, result.Checksum, StringComparison.OrdinalIgnoreCase))
                    {
                        TxtStatus.Text = $"Downloaded zip failed checksum verification - expected {result.Checksum}, got {actualChecksum}. Not reloading.";
                        File.Delete(tempZip);
                        return;
                    }

                    string tempManifest = Path.Combine(Path.GetTempPath(), $"GLSenseOnline_{Guid.NewGuid():N}.json");
                    File.WriteAllText(tempManifest, responseJson);

                    _candidateManifestPath = tempManifest;
                    _candidateZipPath = tempZip;
                    _isValidated = true;
                    BtnReload.IsEnabled = true;
                    TxtStatus.Text = $"Ready to reload: version {result.Version}, released {result.ReleaseDate}.";
                }
            }
            catch (Exception ex)
            {
                TxtStatus.Text = $"Online check failed: {ex.Message}";
            }
            finally
            {
                OnlineProgress.Visibility = Visibility.Collapsed;
                BtnCheckOnline.IsEnabled = true;
            }
        }

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
                DialogResult = true;
                Close();
            }
            catch (Exception ex)
            {
                TxtStatus.Text = $"Failed to stage the new release: {ex.Message}";
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
