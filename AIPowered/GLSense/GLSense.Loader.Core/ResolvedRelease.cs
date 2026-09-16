// ResolvedRelease.cs in GLSense.Loader.Core
namespace GLSense.Loader.Core
{
    /// <summary>
    /// Identifies exactly which Addin.Core release to load: FolderName is the only
    /// thing AddinDomainLoader actually needs (Versions\{FolderName}\ is where the DLLs
    /// live); Version/ReleaseDate are kept for display (GLAbout, log lines, ribbon
    /// messages). See
    /// docs/superpowers/specs/2026-08-30-hotreload-release-history-design.md section 3.1.
    /// </summary>
    public class ResolvedRelease
    {
        public string Version { get; set; }
        public string ReleaseDate { get; set; }
        public string FolderName { get; set; }
    }
}
