using GLSense.Contracts;
using GLSense.Shared;
using System;
using System.Collections.Generic;
using System.Linq;

namespace GLSense
{
    /// <summary>
    /// Pure classification logic for GLReloadSourcePicker's Online-mode release
    /// list - no WPF, no HttpClient, so it's directly testable. See
    /// docs/superpowers/specs/2026-09-18-online-reload-multiversion-design.md
    /// section 2.
    /// </summary>
    public static class OnlineReleaseClassifier
    {
        public static List<OnlineReleaseRow> Classify(
            IEnumerable<VersionInfo> serverEntries,
            IReadOnlyList<ReleaseEntry> localEntries,
            string currentVersion,
            string currentReleaseDate)
        {
            var rows = new List<OnlineReleaseRow>();
            if (serverEntries == null) return rows;

            localEntries = localEntries ?? new List<ReleaseEntry>();

            foreach (var entry in serverEntries)
            {
                if (entry == null || string.IsNullOrWhiteSpace(entry.Version)) continue;

                bool isCurrentlyLoaded =
                    string.Equals(entry.Version, currentVersion, StringComparison.OrdinalIgnoreCase) &&
                    string.Equals(entry.ReleaseDate, currentReleaseDate, StringComparison.OrdinalIgnoreCase);

                bool isAlreadyDownloaded = !isCurrentlyLoaded && localEntries.Any(l =>
                    string.Equals(l.Version, entry.Version, StringComparison.OrdinalIgnoreCase) &&
                    string.Equals(l.ReleaseDate, entry.ReleaseDate, StringComparison.OrdinalIgnoreCase));

                rows.Add(new OnlineReleaseRow
                {
                    Version = entry.Version,
                    ReleaseDate = entry.ReleaseDate,
                    Mandatory = entry.Mandatory,
                    FolderName = entry.FolderName,
                    FileName = entry.FileName,
                    Checksum = entry.Checksum,
                    Notes = entry.Notes,
                    IsCurrentlyLoaded = isCurrentlyLoaded,
                    IsAlreadyDownloaded = isAlreadyDownloaded
                });
            }

            var newestSelectable = rows
                .Where(r => r.IsSelectable)
                .OrderByDescending(r => ParseReleaseDateOrMin(r.ReleaseDate))
                .FirstOrDefault();

            if (newestSelectable != null)
                newestSelectable.IsChecked = true;

            return rows;
        }

        /// <summary>Shared by Classify's default-selection sort and
        /// GLReloadSourcePicker's ascending download-order sort, so "what counts as
        /// newest/oldest" is defined in exactly one place.</summary>
        public static DateTime ParseReleaseDateOrMin(string releaseDate)
        {
            return DateTime.TryParse(releaseDate, out var dt) ? dt : DateTime.MinValue;
        }
    }
}
