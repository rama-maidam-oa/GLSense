// IGLSenseContext.cs in GLSense.Contracts
using System;
using System.Collections.Generic;

namespace GLSense.Contracts
{
    public interface IGLSenseContext
    {
        object ExcelApp { get; }

        // Excel's raw COM IUnknown pointer, AddRef'd for the caller (who must Marshal.Release it).
        // Addin.Core turns it into its OWN Excel object (Marshal.GetObjectForIUnknown) so its Excel
        // calls stay inside its AppDomain instead of going through a cross-AppDomain proxy to the
        // host's object. Through that proxy every returned Excel object (workbook, sheet, range...)
        // was also held by the host side, which kept Excel running with no window after it was
        // closed once a drilldown sheet had been written, and caused the "__ComObject to
        // WorkbookClass" cast errors (CLAUDE.md 24.2/25).
        IntPtr AcquireExcelUnknown();
        ILogger Logger { get; }
        IPathProvider Paths { get; }
        IRibbonController RibbonController { get; }
        IntPtr ExcelHandle { get; set; }

        bool DebugMode { get; set; }

        // Version information
        string Version { get; set; }
        string ReleaseDate { get; set; }  // ✅ Added this
        string ActiveFolderName { get; set; } // Versions\{ActiveFolderName}\ is where the currently loaded release's DLLs live - see ResolvedRelease.
        IReadOnlyList<VersionInfo> AllVersions { get; }
        void SetRibbonController(IRibbonController controller);

        object EdgeAddinInstance { get; set; }
        object GetEdgeAddinInstance();
    }
}
