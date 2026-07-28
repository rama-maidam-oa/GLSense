// VersionParser.cs in GLSense.Shared 
using GLSense.Contracts;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.Json;

namespace GLSense.Shared
{
    public interface IVersionParser
    {
        /// <summary>
        /// Parses version JSON string and returns version info
        /// </summary>
        VersionParseResult ParseVersionJson(string jsonContent);

        /// <summary>
        /// Reads and parses manifest.json from a file path
        /// </summary>
        VersionParseResult ParseVersionFile(string filePath);

        /// <summary>
        /// Reads and parses manifest.json from a directory
        /// </summary>
        VersionParseResult ParseVersionFromDirectory(string directoryPath);

        /// <summary>
        /// Gets the latest version from a list of versions
        /// </summary>
        VersionInfo GetLatestVersion(IEnumerable<VersionInfo> versions);
    }

    public class VersionParseResult
    {
        public bool Success { get; set; }
        public string Version { get; set; }
        public DateTime ReleaseDate { get; set; }
        public string DownloadUrl { get; set; }
        public string Checksum { get; set; }
        public string Notes { get; set; }
        public bool Mandatory { get; set; }
        public List<VersionInfo> AllVersions { get; set; } = new();
        public string ErrorMessage { get; set; }
    }

    public class VersionParser : IVersionParser
    {
        private readonly ILogger _logger;
        private readonly JsonSerializerOptions _options;

        public VersionParser(ILogger logger = null)
        {
            _logger = logger;
            _options = new JsonSerializerOptions
            {
                PropertyNameCaseInsensitive = true,
                PropertyNamingPolicy = JsonNamingPolicy.CamelCase
            };
        }

        /// <summary>
        /// Parses version JSON array and returns the latest version info
        /// </summary>
        public VersionParseResult ParseVersionJson(string jsonContent)
        {
            var result = new VersionParseResult();

            try
            {
                if (string.IsNullOrWhiteSpace(jsonContent))
                {
                    result.ErrorMessage = "JSON content is empty or null";
                    _logger?.LogError(result.ErrorMessage);
                    return result;
                }

                // Parse the JSON array
                var versions = JsonSerializer.Deserialize<List<VersionInfo>>(jsonContent, _options);

                if (versions == null || !versions.Any())
                {
                    result.ErrorMessage = "No version data found in JSON";
                    _logger?.LogError(result.ErrorMessage);
                    return result;
                }

                result.AllVersions = versions;

                // Find the latest version using semantic versioning
                var latest = GetLatestVersion(versions);

                if (latest != null)
                {
                    result.Version = latest.Version;
                    result.ReleaseDate = ParseReleaseDate(latest.ReleaseDate);
                    result.DownloadUrl = latest.DownloadUrl;
                    result.Checksum = latest.Checksum;
                    result.Notes = latest.Notes;
                    result.Mandatory = latest.Mandatory;
                    result.Success = true;

                    _logger?.LogDebug($"Parsed version: {result.Version}, Release Date: {result.ReleaseDate:dd-MMM-yyyy}");
                }
                else
                {
                    result.ErrorMessage = "Failed to determine latest version";
                    _logger?.LogError(result.ErrorMessage);
                }
            }
            catch (JsonException ex)
            {
                result.ErrorMessage = $"Invalid JSON format: {ex.Message}";
                _logger?.LogError(result.ErrorMessage, ex);
            }
            catch (Exception ex)
            {
                result.ErrorMessage = $"Failed to parse version JSON: {ex.Message}";
                _logger?.LogError(result.ErrorMessage, ex);
            }

            return result;
        }

        /// <summary>
        /// Reads and parses manifest.json from a file
        /// </summary>
        public VersionParseResult ParseVersionFile(string filePath)
        {
            try
            {
                _logger?.LogDebug($"ParseVersionFile: reading manifest file '{filePath}'.");

                if (!File.Exists(filePath))
                {
                    var error = $"Manifest file not found: {filePath}";
                    _logger?.LogError(error);
                    return new VersionParseResult { ErrorMessage = error };
                }

                var jsonContent = File.ReadAllText(filePath);
                return ParseVersionJson(jsonContent);
            }
            catch (Exception ex)
            {
                var error = $"Failed to read manifest file: {ex.Message}";
                _logger?.LogError(error, ex);
                return new VersionParseResult { ErrorMessage = error };
            }
        }

        /// <summary>
        /// Reads and parses manifest.json from a directory
        /// </summary>
        public VersionParseResult ParseVersionFromDirectory(string directoryPath)
        {
            try
            {
                _logger?.LogDebug($"ParseVersionFromDirectory: resolving manifest.json under '{directoryPath}'.");

                if (!Directory.Exists(directoryPath))
                {
                    var error = $"Directory not found: {directoryPath}";
                    _logger?.LogError(error);
                    return new VersionParseResult { ErrorMessage = error };
                }

                var jsonPath = Path.Combine(directoryPath, "manifest.json");
                return ParseVersionFile(jsonPath);
            }
            catch (Exception ex)
            {
                var error = $"Failed to parse version from directory: {ex.Message}";
                _logger?.LogError(error, ex);
                return new VersionParseResult { ErrorMessage = error };
            }
        }

        /// <summary>
        /// Gets the latest version using semantic version comparison
        /// </summary>
        public VersionInfo GetLatestVersion(IEnumerable<VersionInfo> versions)
        {
            if (versions == null || !versions.Any())
                return null;

            return versions
                .Where(v => !string.IsNullOrEmpty(v.Version))
                .Select(v => new { VersionInfo = v, VersionObj = TryParseVersion(v.Version) })
                .Where(x => x.VersionObj != null)
                .OrderByDescending(x => x.VersionObj)
                .FirstOrDefault()?.VersionInfo;
        }

        /// <summary>
        /// Parses release date string to DateTime
        /// </summary>
        private DateTime ParseReleaseDate(string releaseDate)
        {
            if (string.IsNullOrEmpty(releaseDate))
                return DateTime.MinValue;

            // Try common formats
            string[] formats = {
                "dd-MMM-yyyy",
                "dd MMM yyyy",
                "yyyy-MM-dd",
                "MM/dd/yyyy",
                "dd/MM/yyyy"
            };

            if (DateTime.TryParseExact(releaseDate, formats,
                System.Globalization.CultureInfo.InvariantCulture,
                System.Globalization.DateTimeStyles.None,
                out DateTime result))
            {
                return result;
            }

            // Fallback to generic parse
            if (DateTime.TryParse(releaseDate, out result))
                return result;

            _logger?.LogWarn($"Could not parse release date: {releaseDate}");
            return DateTime.MinValue;
        }

        /// <summary>
        /// Tries to parse a version string to Version object
        /// </summary>
        private Version TryParseVersion(string versionString)
        {
            try
            {
                return new Version(versionString);
            }
            catch (Exception ex)
            {
                _logger?.LogWarn($"Could not parse version string '{versionString}' as a System.Version: {ex.Message}");
                return null;
            }
        }
    }
}
