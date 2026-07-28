//GLSenseContext.cs in GLSense
using GLSense.Contracts;
using GLSense.Shared;
using System;
using System.Collections.Generic;
using System.Reflection;

namespace GLSense
{
    public class GLSenseContext : MarshalByRefObject, IGLSenseContext
    {
        public object ExcelApp { get; }
        public ILogger Logger { get; }
        public IPathProvider Paths { get; }
        public IntPtr ExcelHandle { get; set; }
        public bool DebugMode { get; set; } = false;

        // Version properties - get from PathProvider
        public string Version { get; set; }
        public string ReleaseDate { get; set; }
        public IReadOnlyList<VersionInfo> AllVersions => Paths?.AllVersions ?? new List<VersionInfo>();

        // Ribbon controller - stored in a field
        private IRibbonController _ribbonController;
        public IRibbonController RibbonController => _ribbonController;

        // Edge Add-in instance
        public object EdgeAddinInstance { get; set; }

        public GLSenseContext(object app)
        {
            ExcelApp = app ?? throw new ArgumentNullException(nameof(app));

            // Initialize PathProvider
            Paths = new PathProvider();
            Paths.Ensure();

            // Create Logger with THIS context reference
            Logger = new Logger(this);

            // Ribbon controller is set later via SetRibbonController
            _ribbonController = null;

            //XLEdge addin instance
            EdgeAddinInstance = null;
        }
        // NEW: Method to set the ribbon controller after creation
        public void SetRibbonController(IRibbonController controller)
        {
            _ribbonController = controller;
        }

        /// <summary>
        /// Gets the XLEdge COM add-in instance
        /// </summary>
        public object GetEdgeAddinInstance()
        {
            try
            {
                Logger?.LogDebug("GLSenseContext.GetEdgeAddinInstance: looking up XLEdge COM add-in.");

                // If already cached, return it
                if (EdgeAddinInstance != null)
                    return EdgeAddinInstance;

                // Get the host application
                var hostApplication = ExcelApp;
                if (hostApplication == null)
                    return null;

                // Get COMAddIns collection
                var comAddIns = hostApplication.GetType().InvokeMember(
                    "COMAddIns",
                    BindingFlags.GetProperty,
                    null,
                    hostApplication,
                    Array.Empty<object>());

                if (comAddIns is System.Collections.IEnumerable addIns)
                {
                    foreach (var addIn in addIns)
                    {
                        if (addIn == null)
                            continue;

                        var addInType = addIn.GetType();
                        var progId = addInType.InvokeMember(
                            "ProgId",
                            BindingFlags.GetProperty,
                            null,
                            addIn,
                            Array.Empty<object>()) as string;

                        var description = addInType.InvokeMember(
                            "Description",
                            BindingFlags.GetProperty,
                            null,
                            addIn,
                            Array.Empty<object>()) as string;

                        if (!IsXlEdgeAddin(progId) && !IsXlEdgeAddin(description))
                            continue;

                        var addinObject = addInType.InvokeMember(
                            "Object",
                            BindingFlags.GetProperty,
                            null,
                            addIn,
                            Array.Empty<object>());

                        if (addinObject != null)
                        {
                            EdgeAddinInstance = addinObject;
                            Logger?.LogDebug("XLEdge COM add-in found and cached");
                            return addinObject;
                        }
                    }
                }

                Logger?.LogWarn("XLEdge COM add-in not found");
                return null;
            }
            catch (Exception ex)
            {
                Logger?.LogError("Failed to locate XLEdge COM add-in", ex);
                return null;
            }
        }

        private static bool IsXlEdgeAddin(string value)
        {
            return !string.IsNullOrWhiteSpace(value) &&
                   value.IndexOf("XLEdge", StringComparison.OrdinalIgnoreCase) >= 0;
        }

        public override object InitializeLifetimeService()
        {
            return null; // Unlimited lifetime
        }
    }
}