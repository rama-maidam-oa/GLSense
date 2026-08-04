using GLSense.Helpers;
using GLSense.Utilities;
using System;
using System.Collections.Generic;
using System.Text.Json;
using System.Xml.Linq;
using Excel = Microsoft.Office.Interop.Excel;

namespace GLSense.Common
{
    /// <summary>
    /// CubeId-keyed CustomXMLPart storage for the raw drilldown-metadata API response, saved
    /// locally via GLDrilldownCustomization's "Save Locally" button
    /// (Views\GLDrilldownCustomization.xaml.cs::BtnSaveLocally_Click) and read back by
    /// Drilldowns\DDDatatoWorksheet.cs::ExtractMetadata when the user's "Overwrite drilldown
    /// metadata with locally saved" preference is enabled (Utilities\UserConfig.cs::
    /// OverwriteDrilldownMetadata, checkbox in Views\GLUserConfig.xaml).
    ///
    /// THIS FILE is the canonical place to look/update if the drilldown-metadata API's exact
    /// JSON shape ever needs adjusting:
    ///   - ExtractDrilldownTypeMetadata below assumes the response looks like
    ///     {"msg":"...","records":{"BALANCE":[...],"JOURNAL":[...],"SUBLEDGER":[...],
    ///     "UNIFIED":[...]}} - confirmed against a live non-fusion cube's response on
    ///     01-Aug-2026 (BALANCE/JOURNAL/SUBLEDGER present, UNIFIED absent - per the API, UNIFIED
    ///     is only present for fusion-based cubes, so this should be re-confirmed against a
    ///     fusion cube's response once available).
    ///   - The DrilldownType -> records-key mapping table lives in DDDatatoWorksheet.cs
    ///     (GetLocalMetadataRecordsKey), since it needs to sit next to DD_Type/DrilldownHelpers.
    ///   - Each entry under a type key has the same per-column shape as
    ///     DrillDownQueryData.metadata (id/drilldownType/viewName/columnName/displayName/
    ///     dataType/format/enabledFlag/displaySequence/customFormula/customDrilldownConfig/
    ///     calculated/subtotalFunction/etc.), so ExtractDrilldownTypeMetadata deserializes
    ///     straight into the same Dictionary&lt;string, object&gt;[] shape and DDDatatoWorksheet's
    ///     existing BuildMetadataDictionary/FillColumnAndTypeInfo need no changes at all.
    ///
    /// Mirrors the older sheet-name-keyed pattern already in DDDatatoWorksheet.cs
    /// (CreateCustomDrilldownXMLPart/RemoveExistingDrilldownParts/AddDrilldownPart), but keyed by
    /// CubeId instead of sheet name since this metadata is per-cube, not per-sheet - these are
    /// two entirely separate CustomXMLPart mechanisms (root element DRILLDOWNSHEET vs
    /// DRILLDOWNMETADATA), each with its own marker so RemoveExistingDrilldownParts's sheet-name
    /// substring check can never collide with this one (see the guard added to
    /// ContainsDrilldownSheet in DDDatatoWorksheet.cs).
    /// </summary>
    public static class DrilldownMetadataXmlStore
    {
        private const string RootElementName = "DRILLDOWNMETADATA";
        private const string CubeIdElementName = "CUBEID";
        private const string PayloadElementName = "PAYLOAD";

        /// <summary>
        /// Deletes any existing DRILLDOWNMETADATA part for this cube, then stores rawJson
        /// (the untouched raw API response body) as-is in a fresh part.
        /// </summary>
        public static void Save(Excel.Workbook wb, long cubeId, string rawJson)
        {
            if (wb == null)
            {
                LogUtility.LogWarn("DrilldownMetadataXmlStore.Save: no active workbook, cannot save metadata.");
                return;
            }

            RemoveExisting(wb, cubeId, out _);

            try
            {
                var root = new XElement(
                    RootElementName,
                    new XElement(CubeIdElementName, cubeId),
                    new XElement(PayloadElementName, new XCData(rawJson ?? string.Empty)));

                wb.CustomXMLParts.Add(new XDocument(root).ToString());
                LogUtility.LogDebug($"DrilldownMetadataXmlStore.Save: stored drilldown metadata CustomXMLPart for cubeId={cubeId}");
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex, "DrilldownMetadataXmlStore.Save");
            }
        }

        /// <summary>
        /// Reads back the raw JSON previously stored via Save for this cube, if any.
        /// </summary>
        public static bool TryRead(Excel.Workbook wb, long cubeId, out string rawJson)
        {
            rawJson = null;

            if (wb?.CustomXMLParts == null || wb.CustomXMLParts.Count == 0)
                return false;

            try
            {
                var cxps = wb.CustomXMLParts;

                // CustomXMLParts collections are 1-based
                for (int i = cxps.Count; i >= 1; i--)
                {
                    var xml = cxps[i]?.XML;
                    if (!ContainsMetadataForCube(xml, cubeId))
                        continue;

                    var doc = XDocument.Parse(xml);
                    rawJson = doc.Root?.Element(PayloadElementName)?.Value;
                    return !string.IsNullOrEmpty(rawJson);
                }
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex, "DrilldownMetadataXmlStore.TryRead");
            }

            return false;
        }

        /// <summary>
        /// Deletes the saved DRILLDOWNMETADATA CustomXMLPart for this cube, if one exists.
        /// Used by GLDrilldownCustomization's "Save Locally" flow (Save above, which always
        /// clears any prior part before writing a fresh one) and by the ribbon's "Delete
        /// Customization" button (AddinModule.cs::RibDDDeleteConfiguration_OnClick), which lets
        /// the user remove a saved customization without saving a new one in its place.
        /// Returns true if a part existed and was deleted, false if there was nothing to delete.
        /// </summary>
        public static bool Delete(Excel.Workbook wb, long cubeId)
        {
            return RemoveExisting(wb, cubeId, out bool deletedAny) && deletedAny;
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
                    if (ContainsMetadataForCube(part?.XML, cubeId))
                    {
                        part.Delete();
                        deletedAny = true;
                    }
                }

                return true;
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex, "DrilldownMetadataXmlStore.RemoveExisting");
                return false;
            }
        }

        private static bool ContainsMetadataForCube(string xml, long cubeId)
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
                LogUtility.LogException(ex, "DrilldownMetadataXmlStore.ContainsMetadataForCube");
                return false;
            }
        }

        /// <summary>
        /// Pulls the array of column-metadata dictionaries for a single drilldown type (e.g.
        /// "BALANCE") out of the raw drilldown-metadata API response, matching
        /// DrillDownQueryData.metadata's Dictionary&lt;string, object&gt;[] shape exactly so
        /// callers can feed it straight into DDDatatoWorksheet's existing
        /// BuildMetadataDictionary. Returns null if rawJson is malformed, or if "records" or the
        /// requested recordsKey isn't present (e.g. UNIFIED on a non-fusion cube).
        /// </summary>
        public static Dictionary<string, object>[] ExtractDrilldownTypeMetadata(string rawJson, string recordsKey)
        {
            if (string.IsNullOrWhiteSpace(rawJson) || string.IsNullOrWhiteSpace(recordsKey))
                return null;

            try
            {
                using var doc = JsonDocument.Parse(rawJson);
                var root = doc.RootElement;

                if (root.ValueKind != JsonValueKind.Object)
                    return null;

                if (!TryGetPropertyIgnoreCase(root, "records", out var recordsElement) ||
                    recordsElement.ValueKind != JsonValueKind.Object)
                {
                    return null;
                }

                if (!TryGetPropertyIgnoreCase(recordsElement, recordsKey, out var typeElement) ||
                    typeElement.ValueKind != JsonValueKind.Array)
                {
                    return null;
                }

                return JsonSerializer.Deserialize<Dictionary<string, object>[]>(typeElement.GetRawText(), JsonGlobals.Options);
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex, "DrilldownMetadataXmlStore.ExtractDrilldownTypeMetadata");
                return null;
            }
        }

        private static bool TryGetPropertyIgnoreCase(JsonElement element, string name, out JsonElement value)
        {
            foreach (var prop in element.EnumerateObject())
            {
                if (string.Equals(prop.Name, name, StringComparison.OrdinalIgnoreCase))
                {
                    value = prop.Value;
                    return true;
                }
            }

            value = default;
            return false;
        }
    }
}
