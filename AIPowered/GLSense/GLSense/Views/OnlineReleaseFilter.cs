using GLSense.Contracts;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.Json;

namespace GLSense
{
    /// <summary>
    /// Keeps only the server releases marked for users. Each entry of the server's release list
    /// (glsenseManifests.json) can carry "release": a release is shown in the Online list when the
    /// flag is true (true, "true", "TRUE"), blank or missing. Anything else ("release": false) is
    /// hidden, so several fix builds can be published and only the last one offered.
    /// Read from the raw JSON, not from VersionInfo: VersionInfo crosses the AppDomain boundary,
    /// and a new field there could break an older core loaded from Release History.
    /// </summary>
    public static class OnlineReleaseFilter
    {
        /// <summary>
        /// The entries of <paramref name="entries"/> whose JSON entry in <paramref name="indexJson"/> is released.
        /// <paramref name="entries"/> must be <paramref name="indexJson"/> deserialized (same array, same order);
        /// if they don't line up, nothing is filtered.
        /// </summary>
        public static List<VersionInfo> ReleasedOnly(string indexJson, IEnumerable<VersionInfo> entries)
        {
            List<VersionInfo> all = entries?.ToList() ?? new List<VersionInfo>();
            try
            {
                using (JsonDocument doc = JsonDocument.Parse(indexJson ?? string.Empty))
                {
                    JsonElement root = doc.RootElement;
                    if (root.ValueKind != JsonValueKind.Array || root.GetArrayLength() != all.Count)
                        return all;

                    var kept = new List<VersionInfo>();
                    int i = 0;
                    foreach (JsonElement entry in root.EnumerateArray())
                    {
                        if (IsReleased(entry)) kept.Add(all[i]);
                        i++;
                    }
                    return kept;
                }
            }
            catch (JsonException)
            {
                return all;
            }
        }

        /// <summary>True when the entry's "release" (any case) is true, "true" (any case), blank or missing.</summary>
        public static bool IsReleased(JsonElement entry)
        {
            if (entry.ValueKind != JsonValueKind.Object) return true;
            foreach (JsonProperty property in entry.EnumerateObject())
            {
                if (!string.Equals(property.Name, "release", StringComparison.OrdinalIgnoreCase)) continue;
                switch (property.Value.ValueKind)
                {
                    case JsonValueKind.True:
                    case JsonValueKind.Null:
                        return true;
                    case JsonValueKind.String:
                        string text = property.Value.GetString()?.Trim();
                        return string.IsNullOrEmpty(text) || string.Equals(text, "true", StringComparison.OrdinalIgnoreCase);
                    default:
                        return false;
                }
            }
            return true;
        }
    }
}
