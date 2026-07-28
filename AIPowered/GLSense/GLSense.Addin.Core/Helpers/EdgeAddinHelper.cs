// EdgeAddinHelper.cs in GLSense.Addin.Core
// New helper extracted from GLSense\AddinModule.cs's static GetEdgeAddinInstance()/
// IsXlEdgeAddin() (FinalWorkingCode). The original cached the discovered XLEdge COM
// add-in object on AddinModule.CurrentInstance.EdgeAddinInstance - that cache can't be
// replicated here since GLSense.Addin.Core has no reference to the host AddinModule
// type (only the stable IGLSenseContext/ServiceLocator surface crosses that boundary).
// This version always re-scans Excel's COMAddIns collection via ServiceLocator.ExcelApp;
// that is acceptable because it is only ever called from Login/Logout (not a hot path).
using GLSense.Addin.Core.Infrastructure;
using System;
using System.Collections;
using System.Reflection;

namespace GLSense.Addin.Core.Helpers
{
    /// <summary>
    /// Locates the XLEdge COM add-in (a sibling Excel add-in this project cross-invokes
    /// for single-sign-on hand-off) via late-bound reflection over Excel's COMAddIns
    /// collection, since XLEdge's assembly isn't referenced directly.
    /// </summary>
    public static class EdgeAddinHelper
    {
        public static object GetEdgeAddinInstance()
        {
            ServiceLocator.Logger?.LogDebug("EdgeAddinHelper.GetEdgeAddinInstance: scanning Excel COMAddIns for XLEdge add-in");

            var hostApplication = ServiceLocator.ExcelApp;
            if (hostApplication == null)
            {
                return null;
            }

            try
            {
                var comAddIns = hostApplication.GetType().InvokeMember("COMAddIns", BindingFlags.GetProperty, null, hostApplication, Array.Empty<object>());
                if (comAddIns is IEnumerable addIns)
                {
                    foreach (var addIn in addIns)
                    {
                        if (addIn == null)
                            continue;

                        var addInType = addIn.GetType();
                        var progId = addInType.InvokeMember("ProgId", BindingFlags.GetProperty, null, addIn, Array.Empty<object>()) as string;
                        var description = addInType.InvokeMember("Description", BindingFlags.GetProperty, null, addIn, Array.Empty<object>()) as string;

                        if (!IsXlEdgeAddin(progId) && !IsXlEdgeAddin(description))
                            continue;

                        ServiceLocator.Logger?.LogDebug($"EdgeAddinHelper.GetEdgeAddinInstance: found XLEdge add-in (ProgId: {progId})");
                        return addInType.InvokeMember("Object", BindingFlags.GetProperty, null, addIn, Array.Empty<object>());
                    }
                }
            }
            catch (Exception ex)
            {
                ServiceLocator.Logger?.LogException(ex, "EdgeAddinHelper.GetEdgeAddinInstance - Failed to locate XLEdge COM add-in.");
            }

            return null;
        }

        private static bool IsXlEdgeAddin(string value)
        {
            return !string.IsNullOrWhiteSpace(value) && value.IndexOf("XLEdge", StringComparison.OrdinalIgnoreCase) >= 0;
        }
    }
}
