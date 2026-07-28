// PathProvider.cs in GLSense.Shared 
using GLSense.Contracts;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.Json;

namespace GLSense.Shared
{
    public class PathProvider : MarshalByRefObject, IPathProvider
    {
        private readonly string _root;
        private readonly string _basePath;
        private static PathProvider _instance;
        public static PathProvider Instance => _instance ??= new PathProvider();

        private static string _latestVersion = "Unknown";
        private static string _latestReleaseDate = "Unknown";
        private static string _latestDownloadUrl = string.Empty;
        private static string _latestChecksum = string.Empty;
        private static string _latestNotes = string.Empty;
        private static bool _latestMandatory = false;
        private static List<VersionInfo> _allVersions = new();

        public PathProvider()
        {
            _basePath = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                "ORBIT", "Excel_Logs");

            _root = Path.Combine(_basePath, "GLSense_Logs_New");

            // Auto-initialize version when PathProvider is created
            InitializeVersion();
        }

        public string Root => _root;
        public string UrlsDirectory => Path.Combine(_basePath, "ORBIT_URLS.xml");
        public string Logs => Path.Combine(_root, "Logs");
        public string Database => Path.Combine(_root, "Database");
        public string Temp => Path.Combine(_root, "Temp");

        public string LoginBrowserPath => Path.Combine(_root, "BrowserLogs", "Login");
        public string DrilldownBrowserPath => Path.Combine(_root, "BrowserLogs", "Drilldown");
        public string VersionsPath => Path.Combine(_root, "Versions");
        public string Resources => Path.Combine(_root, "Resources");

        // "Manifest" (not "Version") since this folder/file is the update-tracking
        // record (releaseDate/version/downloadUrl/etc.), distinct from "Versions" (plural)
        // which holds the actual hot-reloadable DLL payloads.
        public string ManifestDirectory => Path.Combine(_root, "Manifest");
        public string ManifestFile => Path.Combine(ManifestDirectory, "manifest.json");

        // Version properties
        public string LatestVersion => _latestVersion;
        public string LatestReleaseDate => _latestReleaseDate;
        public IReadOnlyList<VersionInfo> AllVersions => _allVersions;

        // Manifest schema fields for the latest version (see manifest.json)
        public string LatestDownloadUrl => _latestDownloadUrl;
        public string LatestChecksum => _latestChecksum;
        public string LatestNotes => _latestNotes;
        public bool LatestMandatory => _latestMandatory;

        public void Ensure()
        {
            Directory.CreateDirectory(Logs);
            Directory.CreateDirectory(Database);
            Directory.CreateDirectory(Temp);
            Directory.CreateDirectory(LoginBrowserPath);
            Directory.CreateDirectory(DrilldownBrowserPath);
            Directory.CreateDirectory(VersionsPath);
            Directory.CreateDirectory(Resources);

            // Ensure Manifest directory exists
            var manifestDir = Path.GetDirectoryName(ManifestFile);
            if (!string.IsNullOrEmpty(manifestDir))
                Directory.CreateDirectory(manifestDir);
        }

        /// <summary>
        /// Re-parses manifest.json and refreshes LatestVersion/LatestReleaseDate/etc.
        /// Call this after something outside PathProvider's own constructor has written
        /// a new manifest.json (e.g. the update-bootstrap flow adopting a newly
        /// downloaded/extracted version) - without this, the static fields would keep
        /// reflecting whatever was on disk when this PathProvider instance was created.
        /// </summary>
        public void Refresh()
        {
            InitializeVersion();
        }

        /// <summary>
        /// Overwrites manifest.json with a single VersionInfo entry (the newly-adopted
        /// "current" version) and immediately refreshes the cached Latest* fields so
        /// callers see the update without needing a new PathProvider instance. Mirrors
        /// CreateDefaultManifestFile()'s single-entry-array shape - manifest.json is a
        /// record of "what's currently installed," not a growing history.
        /// </summary>
        public void WriteManifest(VersionInfo info)
        {
            if (info == null) throw new ArgumentNullException(nameof(info));

            if (!Directory.Exists(ManifestDirectory))
                Directory.CreateDirectory(ManifestDirectory);

            var options = new JsonSerializerOptions { WriteIndented = true };
            var json = JsonSerializer.Serialize(new[] { info }, options);
            File.WriteAllText(ManifestFile, json);

            Refresh();
        }

        private void InitializeVersion()
        {
            try
            {
                // Ensure directories exist first
                Ensure();

                // Check if manifest file exists
                if (!File.Exists(ManifestFile))
                {
                    // Create default manifest file if missing
                    CreateDefaultManifestFile();
                }

                // Read and parse manifest file
                var json = File.ReadAllText(ManifestFile);
                var options = new JsonSerializerOptions { PropertyNameCaseInsensitive = true };

                var versions = JsonSerializer.Deserialize<List<VersionInfo>>(json, options);

                if (versions != null && versions.Any())
                {
                    _allVersions = versions;

                    // Find latest version using semantic version comparison
                    var latest = versions
                        .Select(v => new { v, ver = new Version(v.Version) })
                        .OrderByDescending(x => x.ver)
                        .First().v;

                    _latestVersion = latest.Version;
                    _latestReleaseDate = latest.ReleaseDate;
                    _latestDownloadUrl = latest.DownloadUrl ?? string.Empty;
                    _latestChecksum = latest.Checksum ?? string.Empty;
                    _latestNotes = latest.Notes ?? string.Empty;
                    _latestMandatory = latest.Mandatory;

                }
            }
            catch (Exception ex)
            {
                try
                {
                    string tempLog = Path.Combine(Path.GetTempPath(), $"GLSense_Error_{DateTime.Now:yyyyMMdd_HHmmss}.log");
                    File.WriteAllText(tempLog, $"{DateTime.Now:yyyy-MM-dd HH:mm:ss}|ERROR|PathProvider.InitializeVersion|{ex.Message}|{ex.StackTrace}");
                }
                catch { /* Give up */ }
            }
        }

        private void CreateDefaultManifestFile()
        {
            var defaultVersions = new[]
            {
                new VersionInfo
                {
                    Version = "11.1.0",
                    // Local time, same "yyyy-MM-ddTHH:mm:ss" shape post_build.cmd writes
                    // (see GLSense.Addin.Core\post_build.cmd) - keeps both manifest.json
                    // writers on one consistent, parseable, always-local format.
                    ReleaseDate = DateTime.Now.ToString("yyyy-MM-ddTHH:mm:ss"),
                    DownloadUrl = string.Empty,
                    Checksum = string.Empty,
                    Notes = "Default manifest (auto-created)",
                    Mandatory = false
                }
            };

            var options = new JsonSerializerOptions { WriteIndented = true };
            var json = JsonSerializer.Serialize(defaultVersions, options);

            var manifestDir = Path.GetDirectoryName(ManifestFile);
            if (!string.IsNullOrEmpty(manifestDir))
                Directory.CreateDirectory(manifestDir);

            File.WriteAllText(ManifestFile, json);
        }
        public override object InitializeLifetimeService() => null;
    }
}
