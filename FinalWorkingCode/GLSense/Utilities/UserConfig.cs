using GLSense.Models;
using GLSense.Repositories;
using System;
using System.Collections.Generic;
using System.Linq;

namespace GLSense.Utilities
{
    public static class UserConfig
    {
        private static Dictionary<string, string> _cache = new Dictionary<string, string>();
        private static bool _loaded = false;
        private static readonly string _stringFalse = "false";
        private static readonly string _stringTrue = "true";

        public static List<DrillDownOption> DrillDownSettings { get; set; } = new();

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
            { "ValidateCube", _stringFalse }
        };

        private static void EnsureLoaded()
        {
            if (_loaded) return;

            LogUtility.LogDebug("UserConfig.EnsureLoaded: loading user configs from DB (first access).");
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
                    LogUtility.LogDebug($"UserConfig.EnsureLoaded: loaded {dictConfigs.Count} config override(s) from DB.");
                }
                else
                {
                    LogUtility.LogDebug("UserConfig.EnsureLoaded: no config overrides found in DB, using defaults.");
                }
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex, "UserConfig.EnsureLoaded: failed to load user config");
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
            LogUtility.LogDebug($"UserConfig.Set: {key} = {value}");

            // Save to SQLite immediately
            try
            {
                int rows = DataRepository.SaveUserConfigs(key, value);

                if (rows <= 0)
                {
                    LogUtility.LogError($"Failed to save preference {key}: No rows affected.");
                }
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex, $"UserConfig.Set: failed to save preference {key}");
            }
        }

        // === Your Strongly-Typed Properties ===

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

        //SubLedger include manual journals
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
    }
}
