using GLSense.Models;
using GLSense.Utilities;
using System.Data.SQLite;
using System;
using System.Linq;
using System.Threading.Tasks;

namespace GLSense.Repositories
{
    public class CubeDataRepository : BaseRepository
    {
        // No constructor needed — BaseRepository has no state

        public static async Task InsertCubeDataAsync()
        {
            if (CubeCache.AllCubes == null || CubeCache.AllCubes.Count == 0)
            {
                LogUtility.LogDebug("CubeDataRepository.InsertCubeDataAsync: CubeCache.AllCubes is empty/null, nothing to insert.");
                return;
            }

            LogUtility.LogDebug($"CubeDataRepository.InsertCubeDataAsync: starting insert for {CubeCache.AllCubes.Count} cube(s).");

            try
            {
                await Task.Run(() => // Required for sync DB access on UI thread (e.g., Excel COM)
                {
                    Execute(conn =>
                    {
                        using var transaction = conn.BeginTransaction();
                        try
                        {
                            foreach (var cube in CubeCache.AllCubes)
                            {
                                InsertOrReplaceCube(conn, transaction, cube);

                                if (cube.Ledgers != null && cube.Ledgers.Any())
                                {
                                    InsertOrReplaceLedgers(conn, transaction, cube);
                                }
                            }

                            transaction.Commit();
                            LogUtility.LogDebug($"CubeDataRepository.InsertCubeDataAsync: transaction committed for {CubeCache.AllCubes.Count} cube(s).");
                        }
                        catch (Exception ex)
                        {
                            LogUtility.LogException(ex, "CubeDataRepository.InsertCubeDataAsync: error during transaction, rolling back");
                            transaction.Rollback();
                            throw; // Will be caught by outer catch
                        }
                    });
                }).ConfigureAwait(false);
            }
            catch (Exception ex)
            {
                // Continue with in-memory data — do not crash add-in
                LogUtility.LogException(ex, "CubeDataRepository.InsertCubeDataAsync: failed to persist cube data, continuing with in-memory data only");
            }
        }
        private static void InsertOrReplaceCube(SQLiteConnection conn, SQLiteTransaction transaction, CubeRecord cube)
        {
            const string sql = @"
                INSERT OR REPLACE INTO CUBES
                (cubeId, cubeName, userName, lastRefreshedDate, blazeEnabled, erpType,
                 adaptiveMemoryEnabled, adaptiveMemoryTableName, viewBased)
                VALUES (@cubeId, @cubeName, @userName, @lastRefreshedDate, @blazeEnabled,
                        @erpType, @adaptiveMemoryEnabled, @adaptiveMemoryTableName, @viewBased);";

            LogUtility.LogDebug($"CubeDataRepository.InsertOrReplaceCube: CubeId={cube?.CubeId}, CubeName={cube?.CubeName}");
            try
            {
                using var cmd = conn.CreateCommand();
                cmd.Transaction = transaction;
                cmd.CommandText = sql;

                cmd.Parameters.AddWithValue("@cubeId", cube.CubeId);
                cmd.Parameters.AddWithValue("@cubeName", cube.CubeName ?? (object)DBNull.Value);
                cmd.Parameters.AddWithValue("@userName", cube.UserName ?? (object)DBNull.Value);
                cmd.Parameters.AddWithValue("@lastRefreshedDate", cube.LastRefreshedDate ?? (object)DBNull.Value);
                cmd.Parameters.AddWithValue("@blazeEnabled", cube.BlazeEnabled ? 1 : 0);
                cmd.Parameters.AddWithValue("@erpType", cube.ErpType ?? (object)DBNull.Value);
                cmd.Parameters.AddWithValue("@adaptiveMemoryEnabled", cube.AdaptiveMemoryEnabled ? 1 : 0);
                cmd.Parameters.AddWithValue("@adaptiveMemoryTableName", cube.AdaptiveMemoryTableName ?? (object)DBNull.Value);
                cmd.Parameters.AddWithValue("@viewBased", cube.ViewBased ? 1 : 0);

                cmd.ExecuteNonQuery();
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex, $"Error preparing to insert cube data, query : {sql}");
                throw; // Rethrow to be caught by outer catch
            }
        }
        private static void InsertOrReplaceLedgers(SQLiteConnection conn, SQLiteTransaction transaction, CubeRecord cube)
        {
            const string sql = @"
                INSERT OR REPLACE INTO LEDGERS
                (ledgerId, cubeId, ledgerName, coaid, periodSetName, currencyCode, cubeName)
                VALUES (@ledgerId, @cubeId, @ledgerName, @coaid, @periodSetName, @currencyCode, @cubeName);";

            LogUtility.LogDebug($"CubeDataRepository.InsertOrReplaceLedgers: CubeId={cube?.CubeId}, LedgerCount={cube?.Ledgers?.Count ?? 0}");
            try
            {
                using var cmd = conn.CreateCommand();
                cmd.Transaction = transaction;
                cmd.CommandText = sql;

                var pLedgerId = cmd.Parameters.Add("@ledgerId", System.Data.DbType.Int64);
                var pCubeId = cmd.Parameters.Add("@cubeId", System.Data.DbType.Int64);
                var pLedgerName = cmd.Parameters.Add("@ledgerName", System.Data.DbType.String);
                var pCoaid = cmd.Parameters.Add("@coaid", System.Data.DbType.Int64);
                var pPeriodSetName = cmd.Parameters.Add("@periodSetName", System.Data.DbType.String);
                var pCurrencyCode = cmd.Parameters.Add("@currencyCode", System.Data.DbType.String);
                var pCubeName = cmd.Parameters.Add("@cubeName", System.Data.DbType.String);

                pCubeId.Value = cube.CubeId;
                pCubeName.Value = cube.CubeName ?? (object)DBNull.Value;

                foreach (var ledger in cube.Ledgers)
                {
                    pLedgerId.Value = ledger.LedgerId;
                    pLedgerName.Value = ledger.LedgerName ?? (object)DBNull.Value;
                    pCoaid.Value = ledger.Coaid;
                    pPeriodSetName.Value = ledger.PeriodSetName ?? (object)DBNull.Value;
                    pCurrencyCode.Value = ledger.CurrencyCode ?? (object)DBNull.Value;

                    cmd.ExecuteNonQuery();
                }
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex, $"Error preparing to insert ledger data, query : {sql}");
                throw;
            }
        }
    }
}
