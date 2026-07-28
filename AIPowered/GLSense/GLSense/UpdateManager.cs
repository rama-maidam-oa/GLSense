//UpdateManager.cs in GLSense
using GLSense.Contracts;
using GLSense.Shared;
using System;
using System.IO;
using System.IO.Compression;
using System.Net.Http;
using System.Threading.Tasks;

namespace GLSense
{
    public class UpdateManager : IDisposable
    {
        private readonly IGLSenseContext _context;
        private readonly IVersionParser _versionParser;
        private HttpClient _httpClient;
        private bool _disposed;

        public UpdateManager(IGLSenseContext context)
        {
            _context = context ?? throw new ArgumentNullException(nameof(context));
            _versionParser = new VersionParser(_context.Logger);
            _httpClient = new HttpClient
            {
                Timeout = TimeSpan.FromMinutes(5)
            };
            _context.Logger?.LogDebug("UpdateManager: instance created.");
        }

        /// <summary>
        /// Downloads manifest.json from domain and returns version info
        /// </summary>
        public async Task<(string Version, DateTime ReleaseDate)> DownloadVersionJsonAsync(string domainUrl)
        {
            ThrowIfDisposed();

            string jsonUrl = $"{domainUrl}/manifest.json";

            try
            {
                _context.Logger?.LogDebug($"UpdateManager.DownloadVersionJsonAsync: requesting '{jsonUrl}'.");
                string jsonContent = await _httpClient.GetStringAsync(jsonUrl);

                // ✅ Use VersionParser
                var result = _versionParser.ParseVersionJson(jsonContent);

                if (!result.Success)
                {
                    _context.Logger.LogError($"Failed to parse manifest.json: {result.ErrorMessage}");
                    return (null, DateTime.MinValue);
                }

                _context.Logger?.LogDebug($"UpdateManager.DownloadVersionJsonAsync: parsed version={result.Version}, releaseDate={result.ReleaseDate}.");
                return (result.Version, result.ReleaseDate);
            }
            catch (Exception ex)
            {
                _context.Logger.LogError($"Failed to download manifest.json: {ex.Message}", ex);
                _context.Logger?.LogException(ex, $"UpdateManager.DownloadVersionJsonAsync(domainUrl='{domainUrl}')");
                throw;
            }
        }

        /// <summary>
        /// Downloads zip and extracts to version folder
        /// </summary>
        public async Task<bool> DownloadAndExtractVersionAsync(string domainUrl, string version)
        {
            ThrowIfDisposed();

            string versionFolderName = $"V{version}";
            string versionPath = Path.Combine(_context.Paths.VersionsPath, versionFolderName);
            string tempPath = Path.Combine(Path.GetTempPath(), $"GLSense_{version}_{Guid.NewGuid()}");

            Directory.CreateDirectory(tempPath);

            try
            {
                _context.Logger?.LogDebug($"UpdateManager.DownloadAndExtractVersionAsync: installing version '{version}' from '{domainUrl}' into '{versionPath}' (temp='{tempPath}').");

                string zipUrl = $"{domainUrl}/versions/V{version}/release.zip";
                string zipPath = Path.Combine(tempPath, "release.zip");

                using (var response = await _httpClient.GetAsync(zipUrl))
                {
                    response.EnsureSuccessStatusCode();

                    using (var fs = new FileStream(zipPath, FileMode.Create, FileAccess.Write, FileShare.None))
                    using (var stream = await response.Content.ReadAsStreamAsync())
                    {
                        await stream.CopyToAsync(fs);
                    }
                }

                _context.Logger?.LogDebug($"UpdateManager.DownloadAndExtractVersionAsync: downloaded '{zipUrl}' to '{zipPath}', extracting to '{versionPath}'.");

                if (Directory.Exists(versionPath))
                    Directory.Delete(versionPath, true);

                Directory.CreateDirectory(versionPath);
                ZipFile.ExtractToDirectory(zipPath, versionPath);

                // Download and save manifest.json
                string jsonUrl = $"{domainUrl}/manifest.json";
                string jsonContent = await _httpClient.GetStringAsync(jsonUrl);
                string localJsonPath = Path.Combine(versionPath, "manifest.json");

                using (var writer = new StreamWriter(localJsonPath, false))
                {
                    await writer.WriteAsync(jsonContent);
                }

                _context.Logger?.LogDebug($"UpdateManager.DownloadAndExtractVersionAsync: version '{version}' installed successfully.");
                return true;
            }
            catch (Exception ex)
            {
                _context.Logger.LogError($"Failed to install version {version}: {ex.Message}", ex);
                _context.Logger?.LogException(ex, $"UpdateManager.DownloadAndExtractVersionAsync(domainUrl='{domainUrl}', version='{version}')");

                try
                {
                    if (Directory.Exists(versionPath))
                        Directory.Delete(versionPath, true);
                }
                catch (Exception cleanupEx)
                {
                    _context.Logger?.LogException(cleanupEx, $"UpdateManager.DownloadAndExtractVersionAsync: cleanup of versionPath '{versionPath}' failed after install error");
                }

                return false;
            }
            finally
            {
                try
                {
                    if (Directory.Exists(tempPath))
                        Directory.Delete(tempPath, true);
                }
                catch (Exception cleanupEx)
                {
                    _context.Logger?.LogException(cleanupEx, $"UpdateManager.DownloadAndExtractVersionAsync: cleanup of tempPath '{tempPath}' failed");
                }
            }
        }

        /// <summary>
        /// Reads local manifest.json using VersionParser
        /// </summary>
        public async Task<(string Version, DateTime ReleaseDate)> GetLocalVersionInfoAsync()
        {
            ThrowIfDisposed();

            await Task.Yield(); //This to prevent async method lacks await operators warning

            try
            {
                // ✅ Use VersionParser
                string versionJsonPath = _context.Paths.ManifestFile;
                _context.Logger?.LogDebug($"UpdateManager.GetLocalVersionInfoAsync: reading local manifest file '{versionJsonPath}'.");
                var result = _versionParser.ParseVersionFile(versionJsonPath);

                if (result.Success)
                {
                    _context.Logger?.LogDebug($"UpdateManager.GetLocalVersionInfoAsync: local version={result.Version}, releaseDate={result.ReleaseDate}.");
                    return (result.Version, result.ReleaseDate);
                }

                _context.Logger.LogWarn($"No valid version found locally: {result.ErrorMessage}");
                return (null, DateTime.MinValue);
            }
            catch (Exception ex)
            {
                _context.Logger.LogError($"Failed to get local version info: {ex.Message}", ex);
                _context.Logger?.LogException(ex, "UpdateManager.GetLocalVersionInfoAsync");
                return (null, DateTime.MinValue);
            }
        }

        /// <summary>
        /// Gets the current installed version from PathProvider
        /// </summary>
        public string GetCurrentVersion()
        {
            string version = _context.Paths.LatestVersion;
            _context.Logger?.LogDebug($"UpdateManager.GetCurrentVersion: '{version}'.");
            return version;
        }

        /// <summary>
        /// Gets the current installed version release date
        /// </summary>
        public string GetCurrentReleaseDate()
        {
            string releaseDate = _context.Paths.LatestReleaseDate;
            _context.Logger?.LogDebug($"UpdateManager.GetCurrentReleaseDate: '{releaseDate}'.");
            return releaseDate;
        }

        /// <summary>
        /// Checks if an update is available
        /// </summary>
        public async Task<bool> IsUpdateAvailableAsync(string domainUrl)
        {
            try
            {
                _context.Logger?.LogDebug($"UpdateManager.IsUpdateAvailableAsync: checking '{domainUrl}' for updates.");
                var (serverVersion, _) = await DownloadVersionJsonAsync(domainUrl);
                string currentVersion = GetCurrentVersion();

                if (string.IsNullOrEmpty(serverVersion) || string.IsNullOrEmpty(currentVersion))
                {
                    _context.Logger?.LogWarn($"UpdateManager.IsUpdateAvailableAsync: cannot compare versions (serverVersion='{serverVersion}', currentVersion='{currentVersion}').");
                    return false;
                }

                // Compare versions
                var serverVer = new Version(serverVersion);
                var currentVer = new Version(currentVersion);

                bool updateAvailable = serverVer > currentVer;
                _context.Logger?.LogDebug($"UpdateManager.IsUpdateAvailableAsync: serverVersion={serverVer}, currentVersion={currentVer}, updateAvailable={updateAvailable}.");
                return updateAvailable;
            }
            catch (Exception ex)
            {
                _context.Logger.LogError($"Failed to check for updates: {ex.Message}", ex);
                _context.Logger?.LogException(ex, $"UpdateManager.IsUpdateAvailableAsync(domainUrl='{domainUrl}')");
                return false;
            }
        }

        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }

        protected virtual void Dispose(bool disposing)
        {
            if (_disposed)
                return;

            if (disposing)
            {
                _httpClient?.Dispose();
                _httpClient = null;
            }

            _disposed = true;
        }

        protected void ThrowIfDisposed()
        {
            if (_disposed)
                throw new ObjectDisposedException(nameof(UpdateManager));
        }
    }
}
