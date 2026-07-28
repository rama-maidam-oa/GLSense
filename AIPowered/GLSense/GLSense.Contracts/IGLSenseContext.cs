// IGLSenseContext.cs in GLSense.Contracts
using System;
using System.Collections.Generic;

namespace GLSense.Contracts
{
    public interface IGLSenseContext
    {
        object ExcelApp { get; }
        ILogger Logger { get; }
        IPathProvider Paths { get; }
        IRibbonController RibbonController { get; }
        IntPtr ExcelHandle { get; set; }

        bool DebugMode { get; set; }

        // Version information
        string Version { get; set; }
        string ReleaseDate { get; set; }  // ✅ Added this
        IReadOnlyList<VersionInfo> AllVersions { get; }
        void SetRibbonController(IRibbonController controller);

        object EdgeAddinInstance { get; set; }
        object GetEdgeAddinInstance();
    }
}
