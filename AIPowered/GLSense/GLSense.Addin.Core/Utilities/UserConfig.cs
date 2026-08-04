// UserConfig.cs in GLSense.Addin.Core
// Port of GLSense\Utilities\UserConfig.cs (FinalWorkingCode) - static cached-preferences
// class backed by DataRepository.GetUserConfigs()/SaveUserConfigs(), consumed by
// GLCubeDetails (LoadUserPreferencesForCube reads server-pushed prefs into these
// properties; ChkValidateCube_Changed writes ValidateCube back out).
// Re-pointed vs. the original: LogUtility.* -> ServiceLocator.Logger.*;
// GLSense.Repositories.DataRepository -> GLSense.Addin.Core.Repositories.DataRepository
// (the right-sized port in this project).
// Group I (Config/Debug/About/Help) resolution: DrillDownSettings (List<DrillDownOption>)
// was previously deferred here with a note that it belonged to "Group E (Drilldown ribbon
// group), which hasn't been ported yet." That deferral is now stale - Group E (and every
// other group through H) is fully ported. GLUserConfig (Group I) is the first thing that
// actually needs DrillDownSettings, so it's added below as a simple in-memory property
// (not persisted through Get/Set/DataRepository like the other properties here - the
// original never persisted it either; GLUserConfig rebuilds/repopulates it directly from
// server preferences and its own DataGrid on every load). DrillDownOption itself lives in
// Models\DrillDownOption.cs (this project's Models folder convention - see that file's
// header for why no NotifyBase substitution was needed).
using GLSense.Addin.Core.Infrastructure;
using GLSense.Addin.Core.Models;
using GLSense.Addin.Core.Repositories;
using System;
using System.Collections.Generic;
using System.Linq;

namespace GLSense.Addin.Core.Utilities
{
    public static class UserConfig
    {
        private static Dictionary<string, string> _cache = new Dictionary<string, string>();
        private static bool _loaded = false;
        private static readonly string _stringFalse = "false";
        private static readonly string _stringTrue = "true";

        // Default values (used if not found in DB)
        private static readonly Dictionary<string, string> Defaults = new()
        {
            { "RefreshCells", "100" },
            { "RecordsPerPage", "100" },
            { "SupressZeroBalDrilldown", _stringFalse },
            { "DataOption", "#Blank" },
            // Drilldown defaults
            { "Balance_RunAsJob", _stringFalse },
            { "Journal_RunAsJob", _stringFalse },
            { "SubLedger_RunAsJob", _stringTrue },
            { "Unified_RunAsJob", _stringTrue },
            { "Manual_Journal", _stringFalse },
            // Cube validation defaults
            { "ValidateCube", _stringFalse },
            // Ported from GLSense\Utilities\UserConfig.cs (FinalWorkingCode): backs GLUserConfig's
            // "Overwrite drilldown metadata with locally saved" checkbox.
            { "OverwriteDrilldownMetadata", _stringFalse }
        };

        private static void EnsureLoaded()
        {
            if (_loaded) return;

            _cache = Defaults; // start with defaults

            try
            {
                var repo = new DataRepository();
                var dictConfigs = repo.GetUserConfigs();

                if (dictConfigs.Any())
                {
                    foreach (var kvp in dictConfigs)
                    {
                        _cache[kvp.Key] = kvp.Value;
                    }
                }
            }
            catch (Exception ex)
            {
                ServiceLocator.Logger?.LogException(ex, $"Failed to load user config: {ex.Message}");
            }

            _loaded = true;
        }

        // Generic getter
        private static string Get(string key)
        {
            EnsureLoaded();
            if (_cache.TryGetValue(key, out string value))
            {
                return value;
            }
            else if (Defaults.ContainsKey(key))
            {
                return Defaults[key];
            }
            else
            {
                return "";
            }
        }

        // Generic setter
        private static void Set(string key, string value)
        {
            EnsureLoaded();
            _cache[key] = value;

            // Save to SQLite immediately
            try
            {
                int rows = DataRepository.SaveUserConfigs(key, value);

                if (rows <= 0)
                {
                    ServiceLocator.Logger?.LogError($"Failed to save preference {key}: No rows affected.");
                }
            }
            catch (Exception ex)
            {
                ServiceLocator.Logger?.LogException(ex, $"Failed to save preference {key}: {ex.Message}");
            }
        }

        // === Strongly-typed properties ===

        public static int RefreshCells
        {
            get => int.TryParse(Get("RefreshCells"), out int v) ? v : 100;
            set => Set("RefreshCells", value.ToString());
        }
        public static int RecordsPerPage
        {
            get => int.TryParse(Get("RecordsPerPage"), out int v) ? v : 100;
            set => Set("RecordsPerPage", value.ToString());
        }
        public static bool SupressZeroBalDrilldown
        {
            get => bool.TryParse(Get("SupressZeroBalDrilldown"), out bool v) && v;
            set => Set("SupressZeroBalDrilldown", value.ToString());
        }
        public static string DataOption
        {
            get
            {
                var raw = Get("DataOption");
                return string.IsNullOrEmpty(raw) ? "#Blank" : raw;
            }
            set
            {
                Set("DataOption", string.IsNullOrEmpty(value) ? "#Blank" : value);
            }
        }

        // Cube validation
        public static bool ValidateCube
        {
            get => bool.TryParse(Get("ValidateCube"), out bool v) && v;
            set => Set("ValidateCube", value.ToString());
        }

        // Ported from GLSense\Utilities\UserConfig.cs (FinalWorkingCode): when enabled,
        // Drilldowns\DDDatatoWorksheet.cs's ExtractMetadata prefers the CustomXMLPart saved via
        // GLDrilldownCustomization's "Save Locally" button (Common\DrilldownMetadataXmlStore.cs)
        // over the server's drilldown metadata.
        public static bool OverwriteDrilldownMetadata
        {
            get => bool.TryParse(Get("OverwriteDrilldownMetadata"), out bool v) && v;
            set => Set("OverwriteDrilldownMetadata", value.ToString());
        }
        // Balance
        public static bool Balance_RunAsJob
        {
            get => bool.TryParse(Get("Balance_RunAsJob"), out bool v) && v;
            set => Set("Balance_RunAsJob", value.ToString());
        }

        // Journal
        public static bool Journal_RunAsJob
        {
            get => bool.TryParse(Get("Journal_RunAsJob"), out bool v) && v;
            set => Set("Journal_RunAsJob", value.ToString());
        }

        // SubLedger
        public static bool SubLedger_RunAsJob
        {
            get => bool.TryParse(Get("SubLedger_RunAsJob"), out bool v) && v;
            set => Set("SubLedger_RunAsJob", value.ToString());
        }

        // SubLedger include manual journals
        public static bool SubLedger_Manual_Journal
        {
            get => bool.TryParse(Get("Manual_Journal"), out bool v) && v;
            set => Set("Manual_Journal", value.ToString());
        }

        // Unified
        public static bool Unified_RunAsJob
        {
            get => bool.TryParse(Get("Unified_RunAsJob"), out bool v) && v;
            set => Set("Unified_RunAsJob", value.ToString());
        }

        // Drilldown grid settings (Group I / GLUserConfig). Not persisted via Get/Set -
        // matches the original, which only ever kept this in memory as a convenience cache
        // of GLUserConfig's DataGrid rows for other code to inspect.
        public static List<DrillDownOption> DrillDownSettings { get; set; } = new List<DrillDownOption>();
    }
}
