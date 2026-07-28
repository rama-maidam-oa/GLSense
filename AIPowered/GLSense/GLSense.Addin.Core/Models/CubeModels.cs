// CubeModels.cs in GLSense.Addin.Core
// Right-sized carve-out from GLSense\Models\AllModels.cs (FinalWorkingCode) - that file
// is ~1000 lines covering segment/period/config models that belong to other ribbon
// groups (Group C/D/H/I). Only the pieces the Login flow (GLLogin) and the cube cache it
// populates actually need are ported here: CubeRecord, LedgerRecord, BroadcastMessage,
// CubeCache. No logic changes vs. the original - these are plain data holders.
// Group B (Cube/Ledger selection) additions: CubeValidationResult, LedgerValidationResult
// (cube-dimension-status validation results, cached per-cube in CubeCache.Validations),
// CubeLedgerRecord/CubeLedgerResponse (cube-refreshed-date API response, merged with the
// SQLite-cached ledger list in GLCubeDetails.MapLedgerData).
using GLSense.Addin.Core.Infrastructure;
using System;
using System.Collections.Generic;
using System.Linq;

namespace GLSense.Addin.Core.Models
{
    public class CubeRecord
    {
        public long CubeId { get; set; }
        public string CubeName { get; set; }
        public string UserName { get; set; }
        public List<LedgerRecord> Ledgers { get; set; }
        public string LastRefreshedDate { get; set; }
        public bool BlazeEnabled { get; set; }
        public string ErpType { get; set; }
        public bool AdaptiveMemoryEnabled { get; set; }
        public string AdaptiveMemoryTableName { get; set; }
        public bool ViewBased { get; set; }

        public override string ToString()
        {
            return CubeName ?? string.Empty;
        }

        // Method to get LedgerID by ledger name
        public long? GetLedgerIdByName(string ledgerName)
        {
            if (string.IsNullOrEmpty(ledgerName) || Ledgers == null || !Ledgers.Any())
            {
                ServiceLocator.Logger?.LogDebug($"CubeRecord.GetLedgerIdByName: no ledgers to search (LedgerName={ledgerName}, CubeId={CubeId}).");
                return null;
            }

            // Remove surrounding quotes if present
            string cleanLedgerName = ledgerName.Trim().Trim('"');

            // Also handle escaped quotes
            cleanLedgerName = cleanLedgerName.Replace("\\\"", "\"");

            var ledger = Ledgers.FirstOrDefault(l =>
                string.Equals(l.LedgerName?.Trim(), cleanLedgerName, System.StringComparison.OrdinalIgnoreCase));

            if (ledger == null)
            {
                ServiceLocator.Logger?.LogDebug($"CubeRecord.GetLedgerIdByName: no ledger match for '{ledgerName}' in CubeId={CubeId}.");
            }

            return ledger?.LedgerId;
        }

        public LedgerRecord GetLedgerByName(string ledgerName)
        {
            if (string.IsNullOrEmpty(ledgerName) || Ledgers == null || !Ledgers.Any())
            {
                ServiceLocator.Logger?.LogDebug($"CubeRecord.GetLedgerByName: no ledgers to search (LedgerName={ledgerName}, CubeId={CubeId}).");
                return null;
            }

            var match = Ledgers.FirstOrDefault(l =>
                string.Equals(l.LedgerName?.Trim(), ledgerName.Trim(), System.StringComparison.OrdinalIgnoreCase));

            if (match == null)
            {
                ServiceLocator.Logger?.LogDebug($"CubeRecord.GetLedgerByName: no ledger match for '{ledgerName}' in CubeId={CubeId}.");
            }

            return match;
        }

        // Method to get all ledger names
        public List<string> GetLedgerNames()
        {
            return Ledgers?.Select(l => l.LedgerName).ToList() ?? new List<string>();
        }
    }

    public class LedgerRecord
    {
        public long LedgerId { get; set; }
        public string LedgerName { get; set; }
        public long Coaid { get; set; }
        public string PeriodSetName { get; set; }
        public string CurrencyCode { get; set; }
        public string PeriodType { get; set; }
        public string LedgerData { get; set; }
    }

    public class BroadcastMessage
    {
        public string MsgType { get; set; }
        public string Message { get; set; }
    }

    /// <summary>
    /// In-memory cache of the cubes returned by the last successful login. Populated by
    /// GLLogin after a successful /finance-cubes call; persisted to SQLite via
    /// CubeDataRepository.InsertCubeDataAsync.
    /// </summary>
    public static class CubeCache
    {
        public static List<CubeRecord> AllCubes { get; set; }

        /// <summary>
        /// Per-cube validation results, populated by GLCubeDetails.CubeDataValidation
        /// (cube-dimension-status API) and consulted on subsequent cube/ledger selections
        /// so a cube doesn't have to be re-validated every time the window opens.
        /// </summary>
        public static Dictionary<long, CubeValidationResult> Validations { get; set; } = new Dictionary<long, CubeValidationResult>();
    }

    public class LedgerValidationResult
    {
        public string LedgerName { get; set; }
        public bool IsValid { get; set; }
    }

    /// <summary>Cube-level validation info, keyed by CubeId in CubeCache.Validations.</summary>
    public class CubeValidationResult
    {
        public long CubeId { get; set; }
        public string Message { get; set; }
        public string ErrorMessage { get; set; }
        public bool IsValidated { get; set; }
        public DateTime? ValidatedAt { get; set; }
        public bool IsInSync => !string.IsNullOrWhiteSpace(Message) && (Message.IndexOf("in sync", StringComparison.OrdinalIgnoreCase) >= 0);
        public bool NeedsConfirmation => IsValidated && !IsInSync;
        public List<LedgerValidationResult> Ledgers { get; set; } = new List<LedgerValidationResult>();
    }

    public class CubeLedgerRecord
    {
        public long LedgerId { get; set; }
        public string LedgerName { get; set; }
        public string LastRefreshedDateInUTC { get; set; }
        public long LastRefreshedDateInMilliSecs { get; set; }
        public string LastRefreshedAdaptiveMemInUTC { get; set; }
        public long LastRefreshedAdaptiveMemDateInMilliSecs { get; set; }
        public string LastRefreshedSourceADMInUTC { get; set; }
        public long LastRefreshedSourceADMDateInMilliSecs { get; set; }
    }

    public class CubeLedgerResponse
    {
        public string Msg { get; set; }
        public List<CubeLedgerRecord> Records { get; set; }
        public string Status { get; set; }
    }
}
