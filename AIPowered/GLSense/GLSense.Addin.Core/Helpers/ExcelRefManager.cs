// ExcelRefManager.cs in GLSense.Addin.Core
// Port of GLSense\Helpers\ExcelRefManager.cs (FinalWorkingCode) for Group C
// (Segment/Period pickers) - static cross-control duplicate-reference tracker used by
// ExcelRefEditControl (Views\ExcelRefEditControl.xaml.cs) to warn when the same cell is
// selected for two different RefEdit fields within the same host window.
// Re-pointed vs. the original:
//   - GLSense.Interfaces.IWarningHost -> GLSense.Addin.Core.Interfaces.IWarningHost
//     (already ported).
//   - GLSense.Utilities.AppState.Instance.ExcelApp -> GLSense.Addin.Core.Infrastructure.
//     ServiceLocator.ExcelApp (this project's AppState has no ExcelApp property - Excel
//     access always goes through ServiceLocator, supplied by the host via IGLSenseContext).
//   - CommonMethods.EnsureExcelApp() call dropped: per CommonMethods.cs's own header
//     comment, that method was the old project's way of recovering a stale
//     AddinModule.CurrentInstance.HostApplication reference - not applicable here since
//     ServiceLocator.ExcelApp is always current.
using GLSense.Addin.Core.Infrastructure;
using GLSense.Addin.Core.Interfaces;
using GLSense.Addin.Core.Views;
using System;
using System.Collections.Generic;
using System.Linq;
using Excel = Microsoft.Office.Interop.Excel;

namespace GLSense.Addin.Core.Helpers
{
    public static class ExcelRefManager
    {
        // Key: Host identifier (combination of host and tag name)
        // Value: Normalized cell reference
        private static readonly Dictionary<string, string> SelectedReferences = new(StringComparer.OrdinalIgnoreCase);

        // Track which hosts are using which references to prevent duplicates within same host
        private static readonly Dictionary<string, HashSet<string>> HostReferences = new(StringComparer.OrdinalIgnoreCase);

        public static void SetupControl(ExcelRefEditControl ctrl, string tagName, IWarningHost hostWindow)
        {
            ServiceLocator.Logger?.LogDebug($"ExcelRefManager.SetupControl: wiring up control for tag '{tagName}'");

            ctrl.ExcelApp = ServiceLocator.ExcelApp;
            ctrl.TagName = tagName;

            // Store host window reference in control's Tag
            ctrl.Tag = hostWindow;

            ctrl.CellReferenceChanged -= OnCellReferenceChanged;
            ctrl.CellReferenceChanged += OnCellReferenceChanged;
        }

        private static void OnCellReferenceChanged(object sender, CellReferenceChangedEventArgs e)
        {
            var ctrl = (ExcelRefEditControl)sender;
            var hostWindow = ctrl.Tag as IWarningHost;

            if (hostWindow == null)
                return;

            var normalizedNewRef = NormalizeAddress(e.NewReference);
            var hostKey = GetHostKey(hostWindow);
            var controlKey = $"{hostKey}:{e.TagName}";

            // If reference is empty, just remove it and return
            if (string.IsNullOrWhiteSpace(normalizedNewRef))
            {
                RemoveReference(controlKey, hostKey);
                return;
            }

            // Check for duplicates within the same host
            if (HostReferences.TryGetValue(hostKey, out var hostRefs) &&
                hostRefs.Contains(normalizedNewRef))
            {
                // Find which control in this host is using this reference
                var conflictingControl = SelectedReferences
                    .FirstOrDefault(kvp =>
                        kvp.Key.StartsWith(hostKey + ":") &&
                        kvp.Key != controlKey &&
                        kvp.Value == normalizedNewRef)
                    .Key;

                if (!string.IsNullOrEmpty(conflictingControl))
                {
                    var conflictingTagName = conflictingControl.Split(':')[1];
                    hostWindow.CellSelectionWarning(
                        $"Reference '{e.NewReference}' is already used by '{conflictingTagName}' in this host. " +
                        "Please select a different cell.");
                }

                ctrl.Text = string.Empty; // Reset the control value
                return;
            }

            // Remove old reference if it exists
            if (SelectedReferences.TryGetValue(controlKey, out var oldRef) && !string.IsNullOrWhiteSpace(oldRef))
            {
                if (HostReferences.TryGetValue(hostKey, out var refs))
                {
                    refs.Remove(oldRef);
                }
            }

            // Store the new reference
            if (!string.IsNullOrWhiteSpace(e.TagName))
            {
                SelectedReferences[controlKey] = normalizedNewRef;

                // Track by host
                if (!HostReferences.ContainsKey(hostKey))
                {
                    HostReferences[hostKey] = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                }
                HostReferences[hostKey].Add(normalizedNewRef);
            }
        }

        private static string GetHostKey(IWarningHost host)
        {
            // Use host's unique identifier (you might need to add a property to IWarningHost)
            // This could be a Name, ID, or instance hash code
            return host.GetHashCode().ToString();
        }

        private static void RemoveReference(string controlKey, string hostKey)
        {
            if (SelectedReferences.TryGetValue(controlKey, out var oldRef) && !string.IsNullOrWhiteSpace(oldRef))
            {
                SelectedReferences.Remove(controlKey);

                if (HostReferences.TryGetValue(hostKey, out var refs))
                {
                    refs.Remove(oldRef);
                    if (refs.Count == 0)
                    {
                        HostReferences.Remove(hostKey);
                    }
                }
            }
        }

        private static string NormalizeAddress(string address)
        {
            if (string.IsNullOrWhiteSpace(address))
                return null;

            try
            {
                Excel.Range range = ServiceLocator.ExcelApp.Range[address];

                if (range == null)
                    return null;

                Excel.Workbook workbook = (Excel.Workbook)range.Worksheet.Parent;
                string bookName = workbook.Name;
                string sheetName = range.Worksheet.Name;

                string r1c1 = range.get_Address(
                        true, true,
                        Excel.XlReferenceStyle.xlR1C1, false);

                return $"{bookName}!{sheetName}!{r1c1}";
            }
            catch (Exception ex)
            {
                // Return null if range is invalid
                ServiceLocator.Logger?.LogDebug($"ExcelRefManager.NormalizeAddress: could not normalize '{address}' - {ex.Message}");
                return null;
            }
        }

        public static void Reset()
        {
            ServiceLocator.Logger?.LogDebug("ExcelRefManager.Reset: clearing all tracked cell references for all hosts.");
            SelectedReferences.Clear();
            HostReferences.Clear();
        }

        // Optional: Clear references for a specific host
        public static void ClearHostReferences(IWarningHost host)
        {
            if (host == null)
                return;

            ServiceLocator.Logger?.LogDebug("ExcelRefManager.ClearHostReferences: clearing tracked references for a host window.");

            var hostKey = GetHostKey(host);

            // Remove all references for this host
            var keysToRemove = SelectedReferences.Keys
                .Where(k => k.StartsWith(hostKey + ":"))
                .ToList();

            foreach (var key in keysToRemove)
            {
                SelectedReferences.Remove(key);
            }

            HostReferences.Remove(hostKey);
        }
    }
}
