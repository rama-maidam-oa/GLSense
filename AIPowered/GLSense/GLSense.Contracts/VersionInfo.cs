// VersionInfo.cs in GLSense.Contracts
using System;

namespace GLSense.Contracts
{
    // [Serializable] because this crosses the host<->Addin.Core AppDomain boundary -
    // both as IPathProvider.AllVersions (a List<VersionInfo>) and as the
    // IPathProvider.WriteManifest(VersionInfo) parameter. VersionInfo isn't a
    // MarshalByRefObject (it's a plain data record, not remoted-by-reference), so
    // .NET Remoting requires it to be marked Serializable to pass by value across
    // domains - without this, any cross-domain call touching it would throw a
    // SerializationException at runtime.
    [Serializable]
    public class VersionInfo
    {
        public string Version { get; set; }
        public string ReleaseDate { get; set; }

        // Manifest schema fields (see manifest.json under the "Manifest" folder).
        // System.Text.Json binds these case-insensitively against the JSON keys
        // "downloadUrl"/"checksum"/"notes"/"mandatory" - no attributes needed.
        public string DownloadUrl { get; set; }
        public string Checksum { get; set; }
        public string Notes { get; set; }
        public bool Mandatory { get; set; }

        // Added for the Online multi-version reload feature - binds against the
        // "folderName"/"fileName" keys every manifest.json already writes (see
        // GLSense.Addin.Core\post_build.cmd STEP 2b/4). FolderName is always
        // lowercase "v{version}_{releaseDateSafe}" - see
        // ReleaseHistoryStore.BuildFolderName, the authoritative generator this
        // mirrors.
        public string FolderName { get; set; }
        public string FileName { get; set; }
    }
}
