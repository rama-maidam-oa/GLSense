// BalanceConfigXmlStore.cs in GLSense.Addin.Core
// Ported from GLSense\Common\BalanceConfigXmlStore.cs (FinalWorkingCode).
// Namespace changed from GLSense.Common -> GLSense.Addin.Core.Common. Re-pointing applied
// here (logic unchanged): LogUtility.* (static) -> ServiceLocator.Logger?.* - same convention
// as this project's DrilldownMetadataXmlStore.cs, which this file is directly modeled on.
using GLSense.Addin.Core.Helpers;
using GLSense.Addin.Core.Infrastructure;
using GLSense.Addin.Core.Models;
using System;
using System.Collections.Generic;
using System.Text.Json;
using System.Xml.Linq;
using Excel = Microsoft.Office.Interop.Excel;

namespace GLSense.Addin.Core.Common
{
    /// <summary>
    /// CubeId-keyed CustomXMLPart storage for the full list of a cube's saved Balance
    /// Configurator configurations. Modeled directly on DrilldownMetadataXmlStore.cs
    /// (delete-then-recreate on every write, 1-based backward iteration, cheap substring
    /// check before XDocument.Parse) - same idiom, a distinct root marker so this store's
    /// lookups can never collide with DrilldownMetadataXmlStore's DRILLDOWNMETADATA parts.
    ///
    /// The whole per-cube list lives in ONE part (not one part per saved configuration) -
    /// every Save/Update/Delete of a single entry is: TryRead the current list, mutate it
    /// in memory, Save the mutated list back. Safe at the realistic scale here (a handful
    /// to a few dozen manually-saved entries per cube).
    /// </summary>
    public static class BalanceConfigXmlStore
    {
        private const string RootElementName = "BALANCECONFIGSAVE";
        private const string CubeIdElementName = "CUBEID";
        private const string CubeNameElementName = "CUBENAME";
        private const string PayloadElementName = "PAYLOAD";

        /// <summary>
        /// Deletes any existing BALANCECONFIGSAVE part for this cube, then stores the
        /// given list (serialized as JSON) as-is in a fresh part.
        /// </summary>
        public static void Save(Excel.Workbook wb, long cubeId, string cubeName, List<SavedBalanceConfig> configs)
        {
            if (wb == null)
            {
                ServiceLocator.Logger?.LogWarn("BalanceConfigXmlStore.Save: no active workbook, cannot save configurations.");
                return;
            }

            try
            {
                // Build the new payload FIRST, before touching the existing part - this store
                // holds hand-authored, non-recoverable user data (unlike
                // DrilldownMetadataXmlStore, which holds a re-fetchable API response), so a
                // serialization/XDocument failure here must not leave the user with nothing.
                // Only the RemoveExisting + Add pair below (the actual replace) happens after
                // the new content is known-good.
                string rawJson = JsonSerializer.Serialize(configs ?? new List<SavedBalanceConfig>(), JsonGlobals.Options);

                var root = new XElement(
                    RootElementName,
                    new XElement(CubeIdElementName, cubeId),
                    new XElement(CubeNameElementName, cubeName ?? string.Empty),
                    new XElement(PayloadElementName, new XCData(rawJson)));

                string newPartXml = new XDocument(root).ToString();

                RemoveExisting(wb, cubeId, out _);
                wb.CustomXMLParts.Add(newPartXml);
                ServiceLocator.Logger?.LogDebug($"BalanceConfigXmlStore.Save: stored {configs?.Count ?? 0} saved configuration(s) for cubeId={cubeId}");
            }
            catch (Exception ex)
            {
                ServiceLocator.Logger?.LogException(ex, "BalanceConfigXmlStore.Save");
            }
        }

        /// <summary>
        /// Reads back the saved-configuration list for this cube, if any. Returns an
        /// empty list (not null) via the out parameter when nothing has been saved yet,
        /// or when the stored payload is corrupt.
        /// </summary>
        public static bool TryRead(Excel.Workbook wb, long cubeId, out List<SavedBalanceConfig> configs)
        {
            configs = new List<SavedBalanceConfig>();

            if (wb?.CustomXMLParts == null || wb.CustomXMLParts.Count == 0)
                return false;

            try
            {
                var cxps = wb.CustomXMLParts;

                // CustomXMLParts collections are 1-based
                for (int i = cxps.Count; i >= 1; i--)
                {
                    var xml = cxps[i]?.XML;
                    if (!ContainsConfigForCube(xml, cubeId))
                        continue;

                    var doc = XDocument.Parse(xml);
                    string rawJson = doc.Root?.Element(PayloadElementName)?.Value;
                    if (string.IsNullOrEmpty(rawJson))
                        return false;

                    var deserialized = JsonSerializer.Deserialize<List<SavedBalanceConfig>>(rawJson, JsonGlobals.Options);
                    configs = deserialized ?? new List<SavedBalanceConfig>();
                    return true;
                }
            }
            catch (Exception ex)
            {
                ServiceLocator.Logger?.LogException(ex, "BalanceConfigXmlStore.TryRead");
                configs = new List<SavedBalanceConfig>();
            }

            return false;
        }

        private static bool RemoveExisting(Excel.Workbook wb, long cubeId, out bool deletedAny)
        {
            deletedAny = false;

            if (wb?.CustomXMLParts == null || wb.CustomXMLParts.Count == 0)
                return true;

            try
            {
                var cxps = wb.CustomXMLParts;

                for (int i = cxps.Count; i >= 1; i--)
                {
                    var part = cxps[i];
                    if (part != null && ContainsConfigForCube(part.XML, cubeId))
                    {
                        part.Delete();
                        deletedAny = true;
                    }
                }

                return true;
            }
            catch (Exception ex)
            {
                ServiceLocator.Logger?.LogException(ex, "BalanceConfigXmlStore.RemoveExisting");
                return false;
            }
        }

        private static bool ContainsConfigForCube(string xml, long cubeId)
        {
            if (string.IsNullOrEmpty(xml) || xml.IndexOf(RootElementName, StringComparison.OrdinalIgnoreCase) < 0)
                return false;

            try
            {
                var doc = XDocument.Parse(xml);
                var cubeIdValue = doc.Root?.Element(CubeIdElementName)?.Value;
                return !string.IsNullOrEmpty(cubeIdValue)
                    && long.TryParse(cubeIdValue, out long parsedCubeId)
                    && parsedCubeId == cubeId;
            }
            catch (Exception ex)
            {
                ServiceLocator.Logger?.LogException(ex, "BalanceConfigXmlStore.ContainsConfigForCube");
                return false;
            }
        }
    }
}
