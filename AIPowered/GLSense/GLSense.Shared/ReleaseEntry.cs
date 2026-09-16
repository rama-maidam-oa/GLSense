// ReleaseEntry.cs in GLSense.Shared
using System;

namespace GLSense.Shared
{
    // [Serializable] even though this doesn't currently cross the AppDomain boundary -
    // it's read/written purely host-side and inside UpdateBootstrapper (also host-side,
    // GLSense.Loader.Core) - kept Serializable anyway for consistency with VersionInfo
    // and in case a future call site needs to hand one across.
    [Serializable]
    public class ReleaseEntry
    {
        public string Version { get; set; }
        public string ReleaseDate { get; set; }
        public string FolderName { get; set; }
        public string Checksum { get; set; }
        public string Notes { get; set; }

        /// <summary>"Install" (MSI-seeded first run, or an ordinary local dev-loop
        /// rebuild picked up automatically at Excel startup), "Online", or "Offline".
        /// See docs/superpowers/specs/2026-08-30-hotreload-release-history-design.md
        /// section 4.</summary>
        public string Source { get; set; }
    }
}
