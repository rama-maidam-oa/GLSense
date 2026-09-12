// DrilldownMetadataXmlStore.cs in GLSense.Addin.Core
// Ported from GLSense\Common\DrilldownMetadataXmlStore.cs (FinalWorkingCode, OISR-22390 -
// per-DD-type CustomXMLPart splitting + granular delete, replacing the old single-part
// whole-response storage this file previously had).
// Namespace changed from GLSense.Common -> GLSense.Addin.Core.Common.
// Re-pointing applied here (logic unchanged): LogUtility.* (static) -> ServiceLocator.Logger.*
// (instance via context) - see the same convention already used throughout
// Drilldowns\DDDatatoWorksheet.cs.
using GLSense.Addin.Core.Helpers;
using GLSense.Addin.Core.Infrastructure;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.Json;
using System.Text.Json.Nodes;
using System.Xml.Linq;
using Excel = Microsoft.Office.Interop.Excel;

namespace GLSense.Addin.Core.Common
{
    /// <summary>
    /// CubeId+DDType-keyed CustomXMLPart storage for the raw drilldown-metadata API response,
    /// saved locally via GLDrilldownCustomization's "Save Locally" button
    /// (Views\GLDrilldownCustomization.xaml.cs::BtnSaveLocally_Click) and read back by
    /// Drilldowns\DDDatatoWorksheet.cs::ExtractMetadata when the user's "Overwrite drilldown
    /// metadata with locally saved" preference is enabled (Utilities\UserConfig.cs::
    /// OverwriteDrilldownMetadata, checkbox in Views\GLUserConfig.xaml).
    ///
    /// THIS FILE is the canonical place to look/update if the drilldown-metadata API's exact
    /// JSON shape ever needs adjusting:
    ///   - Save below assumes the response looks like {"msg":"...","records":{"BALANCE":[...],
    ///     "JOURNAL":[...],"SUBLEDGER":[...],"UNIFIED":[...]}} - confirmed against a live
    ///     non-fusion cube's response on 01-Aug-2026 (BALANCE/JOURNAL/SUBLEDGER present,
    ///     UNIFIED absent) and a fusion cube's response (all four present).
    ///   - The DrilldownType -> records-key mapping table lives in DDDatatoWorksheet.cs
    ///     (GetLocalMetadataRecordsKey), since it needs to sit next to DD_Type/DrilldownHelpers.
    ///   - Each entry under a type key has the same per-column shape as
    ///     DrillDownQueryData.metadata (id/drilldownType/viewName/columnName/displayName/
    ///     dataType/format/enabledFlag/displaySequence/customFormula/customDrilldownConfig/
    ///     calculated/subtotalFunction/etc.), so ExtractDrilldownTypeMetadata deserializes
    ///     straight into the same Dictionary&lt;string, object&gt;[] shape and DDDatatoWorksheet's
    ///     existing BuildMetadataDictionary/FillColumnAndTypeInfo need no changes at all.
    ///
    /// Storage granularity: one CustomXMLPart *per DD type* (records-key) instead of one part
    /// holding the entire API response - so a cube saved from an EBS source gets up to 3 parts
    /// (BALANCE/JOURNAL/SUBLEDGER) and a Fusion source up to 4 (adds UNIFIED), and
    /// Views\GLDrilldownDeleteCustomization.xaml.cs can delete a subset of them without
    /// disturbing the rest. Each part's PAYLOAD is still a full "{"records":{"<TYPE>":[...]}}"
    /// document (just narrowed to one type) so ExtractDrilldownTypeMetadata's existing
    /// records-object-then-type-key parsing needs no changes. Save only clears/rewrites the
    /// types actually present in a given response - a type saved earlier but absent from a
    /// later response (e.g. a stale UNIFIED part from a prior Fusion save, now re-saving from
    /// an EBS response) is left untouched.
    ///
    /// Mirrors the older sheet-name-keyed pattern already in DDDatatoWorksheet.cs
    /// (BuildDrilldownDocument/RemoveExistingDrilldownParts/AddDrilldownPart), but keyed by
    /// CubeId instead of sheet name since this metadata is per-cube, not per-sheet - these are
    /// two entirely separate CustomXMLPart mechanisms (root element DRILLDOWNSHEET vs
    /// DRILLDOWNMETADATA), each with its own marker so ContainsDrilldownSheet's sheet-name
    /// substring check can never collide with this one.
    /// </summary>
    public static class DrilldownMetadataXmlStore
    {
        private const string RootElementName = "DRILLDOWNMETADATA";
        private const string CubeIdElementName = "CUBEID";
        private const string DdTypeElementName = "DDTYPE";
        private const string PayloadElementName = "PAYLOAD";

        // Canonical display order for the records-keys this store ever splits a response
        // into - matches the order Views\GLDrilldownDeleteCustomization's picker grid lists
        // saved types in, regardless of the order CustomXMLParts happen to enumerate in.
        private static readonly string[] KnownRecordsKeysOrder = { "BALANCE", "JOURNAL", "SUBLEDGER", "UNIFIED" };

        /// <summary>
        /// Splits rawJson (the untouched raw API response body) by records-key and stores each
        /// present type in its own fresh part. Only the DD types actually present in rawJson
        /// are touched - any previously-saved type for this cube that ISN'T present in this
        /// response is left alone rather than deleted.
        /// </summary>
        public static void Save(Excel.Workbook wb, long cubeId, string rawJson)
        {
            if (wb == null)
            {
                ServiceLocator.Logger?.LogWarn("DrilldownMetadataXmlStore.Save: no active workbook, cannot save metadata.");
                return;
            }

            try
            {
                using var doc = JsonDocument.Parse(rawJson);
                var root = doc.RootElement;

                if (root.ValueKind != JsonValueKind.Object ||
                    !TryGetPropertyIgnoreCase(root, "records", out var recordsElement) ||
                    recordsElement.ValueKind != JsonValueKind.Object)
                {
                    ServiceLocator.Logger?.LogWarn("DrilldownMetadataXmlStore.Save: rawJson has no 'records' object, nothing saved.");
                    return;
                }

                var presentTypes = recordsElement.EnumerateObject()
                    .Where(p => p.Value.ValueKind == JsonValueKind.Array)
                    .Select(p => p.Name.ToUpperInvariant())
                    .ToList();

                if (presentTypes.Count == 0)
                {
                    ServiceLocator.Logger?.LogWarn("DrilldownMetadataXmlStore.Save: rawJson's 'records' object has no array entries, nothing saved.");
                    return;
                }

                // Clear only the types this response is about to rewrite - anything else
                // saved earlier for this cube (a type absent from this response) is untouched.
                Delete(wb, cubeId, presentTypes);

                int savedCount = 0;
                foreach (var prop in recordsElement.EnumerateObject())
                {
                    if (prop.Value.ValueKind != JsonValueKind.Array)
                        continue;

                    string ddType = prop.Name.ToUpperInvariant();

                    var wrapped = new JsonObject
                    {
                        ["records"] = new JsonObject { [ddType] = JsonNode.Parse(prop.Value.GetRawText()) }
                    };

                    var partRoot = new XElement(
                        RootElementName,
                        new XElement(CubeIdElementName, cubeId),
                        new XElement(DdTypeElementName, ddType),
                        new XElement(PayloadElementName, new XCData(wrapped.ToJsonString())));

                    wb.CustomXMLParts.Add(new XDocument(partRoot).ToString());
                    savedCount++;
                }

                ServiceLocator.Logger?.LogDebug($"DrilldownMetadataXmlStore.Save: stored {savedCount} drilldown-type CustomXMLPart(s) for cubeId={cubeId} (types: {string.Join(",", presentTypes)}).");
            }
            catch (Exception ex)
            {
                ServiceLocator.Logger?.LogException(ex, "DrilldownMetadataXmlStore.Save");
            }
        }

        /// <summary>
        /// Reads back the raw JSON previously stored via Save for this cube and DD type
        /// (records-key, e.g. "BALANCE"), if any.
        /// </summary>
        public static bool TryRead(Excel.Workbook wb, long cubeId, string ddType, out string rawJson)
        {
            rawJson = null;

            if (wb?.CustomXMLParts == null || wb.CustomXMLParts.Count == 0 || string.IsNullOrWhiteSpace(ddType))
                return false;

            try
            {
                var cxps = wb.CustomXMLParts;

                // CustomXMLParts collections are 1-based
                for (int i = cxps.Count; i >= 1; i--)
                {
                    var xml = cxps[i]?.XML;
                    if (!TryGetCubeAndType(xml, out long partCubeId, out string partDdType) ||
                        partCubeId != cubeId ||
                        !string.Equals(partDdType, ddType, StringComparison.OrdinalIgnoreCase))
                    {
                        continue;
                    }

                    var doc = XDocument.Parse(xml);
                    rawJson = doc.Root?.Element(PayloadElementName)?.Value;
                    return !string.IsNullOrEmpty(rawJson);
                }
            }
            catch (Exception ex)
            {
                ServiceLocator.Logger?.LogException(ex, "DrilldownMetadataXmlStore.TryRead");
            }

            return false;
        }

        /// <summary>
        /// Returns the DD types (records-keys) currently saved for this cube, each with how
        /// many column-metadata entries it holds, in KnownRecordsKeysOrder. Feeds both
        /// AddinEntry's "no customizations exist" safeguard and
        /// GLDrilldownDeleteCustomization's picker grid.
        /// </summary>
        public static IReadOnlyList<(string DdType, int RecordCount)> GetSavedTypeSummaries(Excel.Workbook wb, long cubeId)
        {
            var results = new List<(string DdType, int RecordCount)>();

            if (wb?.CustomXMLParts == null || wb.CustomXMLParts.Count == 0)
                return results;

            try
            {
                var cxps = wb.CustomXMLParts;

                for (int i = 1; i <= cxps.Count; i++)
                {
                    var xml = cxps[i]?.XML;
                    if (!TryGetCubeAndType(xml, out long partCubeId, out string partDdType) || partCubeId != cubeId)
                        continue;

                    var doc = XDocument.Parse(xml);
                    var payload = doc.Root?.Element(PayloadElementName)?.Value;
                    int recordCount = ExtractDrilldownTypeMetadata(payload, partDdType)?.Length ?? 0;

                    results.Add((partDdType, recordCount));
                }
            }
            catch (Exception ex)
            {
                ServiceLocator.Logger?.LogException(ex, "DrilldownMetadataXmlStore.GetSavedTypeSummaries");
            }

            return results
                .OrderBy(r =>
                {
                    int idx = Array.IndexOf(KnownRecordsKeysOrder, r.DdType);
                    return idx >= 0 ? idx : int.MaxValue;
                })
                .ToList();
        }

        /// <summary>
        /// Deletes only the saved parts for this cube whose DD type is in ddTypes - used by
        /// GLDrilldownDeleteCustomization to remove just the types the user checked. Returns
        /// true if at least one part was deleted.
        /// </summary>
        public static bool Delete(Excel.Workbook wb, long cubeId, IEnumerable<string> ddTypes)
        {
            if (wb?.CustomXMLParts == null || wb.CustomXMLParts.Count == 0 || ddTypes == null)
                return false;

            var typeSet = new HashSet<string>(ddTypes, StringComparer.OrdinalIgnoreCase);
            if (typeSet.Count == 0)
                return false;

            bool deletedAny = false;

            try
            {
                var cxps = wb.CustomXMLParts;

                for (int i = cxps.Count; i >= 1; i--)
                {
                    var part = cxps[i];
                    if (TryGetCubeAndType(part?.XML, out long partCubeId, out string partDdType) &&
                        partCubeId == cubeId &&
                        typeSet.Contains(partDdType))
                    {
                        part.Delete();
                        deletedAny = true;
                    }
                }
            }
            catch (Exception ex)
            {
                ServiceLocator.Logger?.LogException(ex, "DrilldownMetadataXmlStore.Delete(ddTypes)");
            }

            return deletedAny;
        }

        /// <summary>
        /// Parses a candidate CustomXMLPart's XML and, if it's one of ours, returns its
        /// CUBEID/DDTYPE. False for anything that isn't a DRILLDOWNMETADATA part of ours, or
        /// whose CUBEID/DDTYPE can't be read.
        /// </summary>
        private static bool TryGetCubeAndType(string xml, out long cubeId, out string ddType)
        {
            cubeId = 0;
            ddType = null;

            if (string.IsNullOrEmpty(xml) || xml.IndexOf(RootElementName, StringComparison.OrdinalIgnoreCase) < 0)
                return false;

            try
            {
                var doc = XDocument.Parse(xml);
                var cubeIdValue = doc.Root?.Element(CubeIdElementName)?.Value;
                ddType = doc.Root?.Element(DdTypeElementName)?.Value;

                return !string.IsNullOrEmpty(cubeIdValue)
                    && long.TryParse(cubeIdValue, out cubeId)
                    && !string.IsNullOrEmpty(ddType);
            }
            catch (Exception ex)
            {
                ServiceLocator.Logger?.LogException(ex, "DrilldownMetadataXmlStore.TryGetCubeAndType");
                return false;
            }
        }

        /// <summary>
        /// Pulls the array of column-metadata dictionaries for a single drilldown type (e.g.
        /// "BALANCE") out of a "{"records":{...}}" drilldown-metadata document, matching
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
                ServiceLocator.Logger?.LogException(ex, "DrilldownMetadataXmlStore.ExtractDrilldownTypeMetadata");
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
