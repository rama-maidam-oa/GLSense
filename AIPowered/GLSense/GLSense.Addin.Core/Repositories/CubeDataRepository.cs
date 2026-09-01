// CubeDataRepository.cs in GLSense.Addin.Core
// Port of GLSense\Repositories\CubeDataRepository.cs (FinalWorkingCode).
// The original went through a BaseRepository abstraction (BaseRepository.Execute ->
// static SQLiteHelper.GetConnection()). This project's SQLiteHelper is an instance
// singleton (SQLiteHelper.Instance.GetConnection()), and there's no BaseRepository
// here yet, so this talks to it directly - same net effect (open connection, run
// inside a transaction, dispose), just without the extra abstraction layer since this
// is the only repository ported so far. If Group B/C/D add more repositories that want
// the same ExecuteNonQuery/ExecuteQuery helpers, factor a BaseRepository out then rather
// than duplicating this pattern by hand each time.
using GLSense.Addin.Core.Helpers;
using GLSense.Addin.Core.Infrastructure;
using GLSense.Addin.Core.Models;
using System;
using System.Data.SQLite;
using System.Linq;
using System.Threading.Tasks;

namespace GLSense.Addin.Core.Repositories
{
    public static class CubeDataRepository
    {
        public static async Task InsertCubeDataAsync()
        {
            if (CubeCache.AllCubes == null || CubeCache.AllCubes.Count == 0)
            {
                ServiceLocator.Logger?.LogDebug("CubeDataRepository.InsertCubeDataAsync: no cubes in CubeCache.AllCubes - nothing to persist.");
                return;
            }

            ServiceLocator.Logger?.LogDebug($"CubeDataRepository.InsertCubeDataAsync started. CubeCount={CubeCache.AllCubes.Count}");

            try
            {
                await Task.Run(() => // Required for sync DB access on UI thread (e.g., Excel COM)
                {
                    using var conn = SQLiteHelper.Instance.GetConnection();
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
                    }
                    catch
                    {
                        transaction.Rollback();
                        throw; // Will be caught by outer catch
                    }
                }).ConfigureAwait(false);

                ServiceLocator.Logger?.LogDebug($"CubeDataRepository.InsertCubeDataAsync completed successfully. CubeCount={CubeCache.AllCubes.Count}");
            }
            catch (Exception ex)
            {
                // Continue with in-memory data - do not crash the add-in
                ServiceLocator.Logger?.LogException(ex, "InsertCubeDataAsync failed - continuing with in-memory cube data only");
            }
        }

        private static void InsertOrReplaceCube(SQLiteConnection conn, SQLiteTransaction transaction, CubeRecord cube)
        {
            ServiceLocator.Logger?.LogDebug($"CubeDataRepository.InsertOrReplaceCube: CubeId={cube?.CubeId}, CubeName={cube?.CubeName}");
            const string sql = @"
                INSERT OR REPLACE INTO CUBES
                (cubeId, cubeName, userName, lastRefreshedDate, blazeEnabled, erpType,
                 adaptiveMemoryEnabled, adaptiveMemoryTableName, viewBased)
                VALUES (@cubeId, @cubeName, @userName, @lastRefreshedDate, @blazeEnabled,
                        @erpType, @adaptiveMemoryEnabled, @adaptiveMemoryTableName, @viewBased);";

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
                ServiceLocator.Logger?.LogException(ex, $"Error preparing to insert cube data, query : {sql}");
                throw; // Rethrow to be caught by outer catch
            }
        }

        private static void InsertOrReplaceLedgers(SQLiteConnection conn, SQLiteTransaction transaction, CubeRecord cube)
        {
            ServiceLocator.Logger?.LogDebug($"CubeDataRepository.InsertOrReplaceLedgers: CubeId={cube?.CubeId}, LedgerCount={cube?.Ledgers?.Count() ?? 0}");
            const string sql = @"
                INSERT OR REPLACE INTO LEDGERS
                (ledgerId, cubeId, ledgerName, coaid, periodSetName, currencyCode, cubeName)
                VALUES (@ledgerId, @cubeId, @ledgerName, @coaid, @periodSetName, @currencyCode, @cubeName);";

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
                ServiceLocator.Logger?.LogException(ex, $"Error preparing to insert ledger data, query : {sql}");
                throw;
            }
        }
    }
}
