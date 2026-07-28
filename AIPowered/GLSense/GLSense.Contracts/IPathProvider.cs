// IPathProvider.cs in GLSense.Contracts
using System.Collections.Generic;

namespace GLSense.Contracts
{
    public interface IPathProvider
    {
        string Root { get; }
        string Logs { get; }
        string Database { get; }
        string LoginBrowserPath { get; }
        string DrilldownBrowserPath { get; }
        string Temp { get; }
        string UrlsDirectory { get; }
        string VersionsPath { get; }
        string Resources { get; }
        string ManifestFile { get; }
        string ManifestDirectory { get; }

        // Version properties
        string LatestVersion { get; }        // ✅ Added
        string LatestReleaseDate { get; }    // ✅ Added
        IReadOnlyList<VersionInfo> AllVersions { get; }  // ✅ Added

        // Manifest schema fields for the latest version (see manifest.json)
        string LatestDownloadUrl { get; }
        string LatestChecksum { get; }
        string LatestNotes { get; }
        bool LatestMandatory { get; }

        void Ensure();

        /// <summary>Re-parses manifest.json and refreshes the Latest* properties.</summary>
        void Refresh();

        /// <summary>Overwrites manifest.json with a single newly-adopted VersionInfo entry, then Refresh()es.</summary>
        void WriteManifest(VersionInfo info);
    }
}
