// LedgerDataRepository.cs in GLSense.Addin.Core
// Port of GLSense\Repositories\LedgerDataRepository.cs (FinalWorkingCode). The original
// derived from BaseRepository (static SQLiteHelper.GetConnection()); this project has no
// BaseRepository yet, so - matching the CubeDataRepository.cs pattern already established
// in this project - it talks to SQLiteHelper.Instance.GetConnection() directly.
// Re-pointed vs. the original: LogUtility.* -> ServiceLocator.Logger.*; JsonGlobals/
// JsonException resolve via this project's Helpers namespace (already ported).
using GLSense.Addin.Core.Helpers;
using GLSense.Addin.Core.Infrastructure;
using GLSense.Addin.Core.Models;
using System;
using System.Data;
using System.Data.SQLite;
using System.Text.Json;
using System.Threading.Tasks;

namespace GLSense.Addin.Core.Repositories
{
    public static class LedgerDataRepository
    {
        public static async Task InsertLedgerDataAsync(long cubeId, long ledgerId, string ledgerJson)
        {
            ServiceLocator.Logger?.LogDebug($"LedgerDataRepository.InsertLedgerDataAsync started. CubeId={cubeId}, LedgerId={ledgerId}, JsonLength={ledgerJson?.Length ?? 0}");

            if (string.IsNullOrWhiteSpace(ledgerJson))
            {
                ServiceLocator.Logger?.LogWarn($"LedgerDataRepository.InsertLedgerDataAsync: empty ledgerJson for CubeId={cubeId}, LedgerId={ledgerId} - nothing to persist.");
                return;
            }

            try
            {
                await Task.Run(() => ProcessLedgerData(cubeId, ledgerId, ledgerJson));
                ServiceLocator.Logger?.LogDebug($"LedgerDataRepository.InsertLedgerDataAsync completed successfully. CubeId={cubeId}, LedgerId={ledgerId}");
            }
            catch (Exception ex)
            {
                ServiceLocator.Logger?.LogException(ex, $"Failed to insert ledger data for Cube {cubeId}, Ledger {ledgerId}");
                // Do not throw - allow add-in to continue with partial data
            }
        }

        private static void ProcessLedgerData(long cubeId, long ledgerId, string ledgerJson)
        {
            ServiceLocator.Logger?.LogDebug($"LedgerDataRepository.ProcessLedgerData started. CubeId={cubeId}, LedgerId={ledgerId}");

            LedgerQueryData recs;
            try
            {
                recs = JsonSerializer.Deserialize<LedgerQueryData>(ledgerJson, JsonGlobals.Options);
            }
            catch (JsonException ex)
            {
                ServiceLocator.Logger?.LogException(ex, "Invalid ledger JSON");
                ServiceLocator.Logger?.LogRawJson("LedgerDataRepository.ProcessLedgerData", ledgerJson);
                return;
            }

            if (recs == null || recs.records == null)
            {
                ServiceLocator.Logger?.LogWarn($"LedgerDataRepository.ProcessLedgerData: deserialized records were null for CubeId={cubeId}, LedgerId={ledgerId}.");
                return;
            }

            using var conn = SQLiteHelper.Instance.GetConnection();
            using var transaction = conn.BeginTransaction();

            try
            {
                ClearExistingData(conn, transaction, cubeId, ledgerId);
                InsertAllData(conn, transaction, cubeId, ledgerId, recs.records);
                transaction.Commit();
                ServiceLocator.Logger?.LogDebug($"LedgerDataRepository.ProcessLedgerData: transaction committed for CubeId={cubeId}, LedgerId={ledgerId}");
            }
            catch (Exception ex)
            {
                ServiceLocator.Logger?.LogException(ex, $"LedgerDataRepository.ProcessLedgerData: rolling back transaction for CubeId={cubeId}, LedgerId={ledgerId}");
                transaction.Rollback();
                throw;
            }
        }

        private static void ClearExistingData(SQLiteConnection conn, SQLiteTransaction transaction, long cubeId, long ledgerId)
        {
            ServiceLocator.Logger?.LogDebug($"LedgerDataRepository.ClearExistingData: CubeId={cubeId}, LedgerId={ledgerId}");
            try
            {
                var tablesToClear = new[]
                {
                    "ACTIVITY", "JOURNALSOURCES", "ENCUMBRANCES",
                    "JOURNALCATEGORIES", "CURRENCIES", "BUDGETS",
                    "PERIODS", "SEGMENTS", "SEGMENT_VALUES", "SEGMENT_HIERARCHY_CACHE"
                };

                using var deleteCmd = conn.CreateCommand();
                deleteCmd.Transaction = transaction;

                foreach (var table in tablesToClear)
                {
                    deleteCmd.CommandText = $"DELETE FROM [{table}] WHERE cubeId = @cubeId AND ledgerId = @ledgerId";
                    deleteCmd.Parameters.Clear();
                    deleteCmd.Parameters.AddWithValue("@cubeId", cubeId);
                    deleteCmd.Parameters.AddWithValue("@ledgerId", ledgerId);
                    deleteCmd.ExecuteNonQuery();
                }
            }
            catch (Exception ex)
            {
                ServiceLocator.Logger?.LogException(ex);
            }
        }

        private static void InsertAllData(SQLiteConnection conn, SQLiteTransaction transaction,
                                 long cubeId, long ledgerId, Records records)
        {
            ServiceLocator.Logger?.LogDebug($"LedgerDataRepository.InsertAllData: CubeId={cubeId}, LedgerId={ledgerId}");
            try
            {
                InsertActivity(conn, transaction, cubeId, ledgerId, records.activity);
                InsertJournalSources(conn, transaction, cubeId, ledgerId, records.journalsources);
                InsertEncumbrances(conn, transaction, cubeId, ledgerId, records.encumbrances);
                InsertJournalCategories(conn, transaction, cubeId, ledgerId, records.journalcategories);
                InsertCurrencies(conn, transaction, cubeId, ledgerId, records.currencies);

                if (records.ledgers?.ledgerData != null)
                {
                    InsertBudgets(conn, transaction, cubeId, ledgerId, records.ledgers.ledgerData.budgets);
                    InsertPeriods(conn, transaction, cubeId, ledgerId, records.ledgers.ledgerData.periods);
                    InsertSegmentsAndValues(conn, transaction, cubeId, ledgerId, records.ledgers.ledgerData.segments);
                }
            }
            catch (Exception ex)
            {
                ServiceLocator.Logger?.LogException(ex);
            }
        }

        private static void InsertActivity(SQLiteConnection conn, SQLiteTransaction transaction,
                                  long cubeId, long ledgerId, string[] activities)
        {
            try
            {
                if (activities == null || activities.Length == 0)
                {
                    ServiceLocator.Logger?.LogDebug($"LedgerDataRepository.InsertActivity: no activities for CubeId={cubeId}, LedgerId={ledgerId}.");
                    return;
                }

                ServiceLocator.Logger?.LogDebug($"LedgerDataRepository.InsertActivity: inserting {activities.Length} rows for CubeId={cubeId}, LedgerId={ledgerId}.");

                using var cmd = conn.CreateCommand();
                cmd.Transaction = transaction;
                cmd.CommandText = @"INSERT OR REPLACE INTO ACTIVITY (cubeId, ledgerId, activityType)
                          VALUES (@cubeId, @ledgerId, @activityType);";

                foreach (var act in activities)
                {
                    cmd.Parameters.Clear();
                    AddCommonParameters(cmd, cubeId, ledgerId);
                    cmd.Parameters.AddWithValue("@activityType", act ?? (object)DBNull.Value);
                    cmd.ExecuteNonQuery();
                }
            }
            catch (Exception ex)
            {
                ServiceLocator.Logger?.LogException(ex);
            }
        }

        private static void InsertJournalSources(SQLiteConnection conn, SQLiteTransaction transaction,
                                           long cubeId, long ledgerId, JESources[] sources)
        {
            try
            {
                if (sources == null || sources.Length <= 0)
                {
                    ServiceLocator.Logger?.LogDebug($"LedgerDataRepository.InsertJournalSources: no journal sources for CubeId={cubeId}, LedgerId={ledgerId}.");
                    return;
                }

                ServiceLocator.Logger?.LogDebug($"LedgerDataRepository.InsertJournalSources: inserting {sources.Length} rows for CubeId={cubeId}, LedgerId={ledgerId}.");

                using var cmd = conn.CreateCommand();
                cmd.Transaction = transaction;
                cmd.CommandText = @"INSERT OR REPLACE INTO JOURNALSOURCES
                          (cubeId, ledgerId, jeSourceName, sourceName)
                          VALUES (@cubeId, @ledgerId, @jeSourceName, @sourceName);";
                foreach (var source in sources)
                {
                    cmd.Parameters.Clear();
                    AddCommonParameters(cmd, cubeId, ledgerId);
                    cmd.Parameters.AddWithValue("@jeSourceName", source.jeSourceName ?? (object)DBNull.Value);
                    cmd.Parameters.AddWithValue("@sourceName", source.sourceName ?? (object)DBNull.Value);
                    cmd.ExecuteNonQuery();
                }
            }
            catch (Exception ex)
            {
                ServiceLocator.Logger?.LogException(ex);
            }
        }

        private static void InsertEncumbrances(SQLiteConnection conn, SQLiteTransaction transaction,
                                             long cubeId, long ledgerId, Encumbrance[] encumbrances)
        {
            try
            {
                if (encumbrances == null || encumbrances.Length == 0)
                {
                    ServiceLocator.Logger?.LogDebug($"LedgerDataRepository.InsertEncumbrances: no encumbrances for CubeId={cubeId}, LedgerId={ledgerId}.");
                    return;
                }

                ServiceLocator.Logger?.LogDebug($"LedgerDataRepository.InsertEncumbrances: inserting {encumbrances.Length} rows for CubeId={cubeId}, LedgerId={ledgerId}.");

                using var cmd = conn.CreateCommand();
                cmd.Transaction = transaction;
                cmd.CommandText = @"INSERT OR REPLACE INTO ENCUMBRANCES
                          (cubeId, ledgerId, encumbranceTypeId, encumbranceType)
                          VALUES (@cubeId, @ledgerId, @encumbranceTypeId, @encumbranceType);";
                foreach (var enc in encumbrances)
                {
                    cmd.Parameters.Clear();
                    AddCommonParameters(cmd, cubeId, ledgerId);
                    cmd.Parameters.AddWithValue("@encumbranceTypeId", enc.encumbranceTypeId);
                    cmd.Parameters.AddWithValue("@encumbranceType", enc.encumbranceType ?? (object)DBNull.Value);
                    cmd.ExecuteNonQuery();
                }
            }
            catch (Exception ex)
            {
                ServiceLocator.Logger?.LogException(ex);
            }
        }

        private static void InsertJournalCategories(SQLiteConnection conn, SQLiteTransaction transaction,
                                                  long cubeId, long ledgerId, JECategories[] categories)
        {
            try
            {
                if (categories == null || categories.Length == 0)
                {
                    ServiceLocator.Logger?.LogDebug($"LedgerDataRepository.InsertJournalCategories: no journal categories for CubeId={cubeId}, LedgerId={ledgerId}.");
                    return;
                }

                ServiceLocator.Logger?.LogDebug($"LedgerDataRepository.InsertJournalCategories: inserting {categories.Length} rows for CubeId={cubeId}, LedgerId={ledgerId}.");

                using var cmd = conn.CreateCommand();
                cmd.Transaction = transaction;
                cmd.CommandText = @"INSERT OR REPLACE INTO JOURNALCATEGORIES
                          (cubeId, ledgerId, jeCategoryName, categoryName)
                          VALUES (@cubeId, @ledgerId, @jeCategoryName, @categoryName);";

                foreach (var cat in categories)
                {
                    cmd.Parameters.Clear();
                    AddCommonParameters(cmd, cubeId, ledgerId);
                    cmd.Parameters.AddWithValue("@jeCategoryName", cat.jeCategoryName ?? (object)DBNull.Value);
                    cmd.Parameters.AddWithValue("@categoryName", cat.categoryName ?? (object)DBNull.Value);
                    cmd.ExecuteNonQuery();
                }
            }
            catch (Exception ex)
            {
                ServiceLocator.Logger?.LogException(ex);
            }
        }

        private static void InsertCurrencies(SQLiteConnection conn, SQLiteTransaction transaction,
                                           long cubeId, long ledgerId, string[] currencies)
        {
            try
            {
                if (currencies == null || currencies.Length == 0)
                {
                    ServiceLocator.Logger?.LogDebug($"LedgerDataRepository.InsertCurrencies: no currencies for CubeId={cubeId}, LedgerId={ledgerId}.");
                    return;
                }

                ServiceLocator.Logger?.LogDebug($"LedgerDataRepository.InsertCurrencies: inserting {currencies.Length} rows for CubeId={cubeId}, LedgerId={ledgerId}.");

                using var cmd = conn.CreateCommand();
                cmd.Transaction = transaction;
                cmd.CommandText = @"INSERT OR REPLACE INTO CURRENCIES (cubeId, ledgerId, currencyCode)
                          VALUES (@cubeId, @ledgerId, @currencyCode);";

                foreach (var currency in currencies)
                {
                    cmd.Parameters.Clear();
                    AddCommonParameters(cmd, cubeId, ledgerId);
                    cmd.Parameters.AddWithValue("@currencyCode", currency ?? (object)DBNull.Value);
                    cmd.ExecuteNonQuery();
                }
            }
            catch (Exception ex)
            {
                ServiceLocator.Logger?.LogException(ex);
            }
        }

        private static void InsertBudgets(SQLiteConnection conn, SQLiteTransaction transaction,
                                        long cubeId, long ledgerId, string[] budgets)
        {
            try
            {
                if (budgets == null || budgets.Length == 0)
                {
                    ServiceLocator.Logger?.LogDebug($"LedgerDataRepository.InsertBudgets: no budgets for CubeId={cubeId}, LedgerId={ledgerId}.");
                    return;
                }

                ServiceLocator.Logger?.LogDebug($"LedgerDataRepository.InsertBudgets: inserting {budgets.Length} rows for CubeId={cubeId}, LedgerId={ledgerId}.");
                using var cmd = conn.CreateCommand();
                cmd.Transaction = transaction;
                cmd.CommandText = @"INSERT OR REPLACE INTO BUDGETS (cubeId, ledgerId, budgetName)
                          VALUES (@cubeId, @ledgerId, @budgetName);";
                foreach (var budget in budgets)
                {
                    cmd.Parameters.Clear();
                    AddCommonParameters(cmd, cubeId, ledgerId);
                    cmd.Parameters.AddWithValue("@budgetName", budget ?? (object)DBNull.Value);
                    cmd.ExecuteNonQuery();
                }
            }
            catch (Exception ex)
            {
                ServiceLocator.Logger?.LogException(ex);
            }
        }

        private static void InsertPeriods(SQLiteConnection conn, SQLiteTransaction transaction,
                                        long cubeId, long ledgerId, Period[] periods)
        {
            try
            {
                if (periods == null || periods.Length == 0)
                {
                    ServiceLocator.Logger?.LogDebug($"LedgerDataRepository.InsertPeriods: no periods for CubeId={cubeId}, LedgerId={ledgerId}.");
                    return;
                }

                ServiceLocator.Logger?.LogDebug($"LedgerDataRepository.InsertPeriods: inserting {periods.Length} rows for CubeId={cubeId}, LedgerId={ledgerId}.");

                using var cmd = conn.CreateCommand();
                cmd.Transaction = transaction;
                cmd.CommandText = @"INSERT OR REPLACE INTO PERIODS
                          (cubeId, ledgerId, periodName, periodYear, periodNum, quarterNum,
                           periodSetName, periodType, startDate, endDate, adjustmentPeriodFlag)
                          VALUES (@cubeId, @ledgerId, @periodName, @periodYear, @periodNum,
                                  @quarterNum, @periodSetName, @periodType, @startDate,
                                  @endDate, @adjustmentPeriodFlag);";

                foreach (var period in periods)
                {
                    cmd.Parameters.Clear();
                    AddCommonParameters(cmd, cubeId, ledgerId);
                    cmd.Parameters.AddWithValue("@periodName", period.periodName ?? (object)DBNull.Value);
                    cmd.Parameters.AddWithValue("@periodYear", period.periodYear);
                    cmd.Parameters.AddWithValue("@periodNum", period.periodNum);
                    cmd.Parameters.AddWithValue("@quarterNum", period.quarterNum);
                    cmd.Parameters.AddWithValue("@periodSetName", period.periodSetName ?? (object)DBNull.Value);
                    cmd.Parameters.AddWithValue("@periodType", period.periodType ?? (object)DBNull.Value);
                    cmd.Parameters.AddWithValue("@startDate", period.startDate);
                    cmd.Parameters.AddWithValue("@endDate", period.endDate);
                    cmd.Parameters.AddWithValue("@adjustmentPeriodFlag", period.adjustmentPeriodFlag ?? (object)DBNull.Value);
                    cmd.ExecuteNonQuery();
                }
            }
            catch (Exception ex)
            {
                ServiceLocator.Logger?.LogException(ex);
            }
        }

        private static void InsertSegmentsAndValues(SQLiteConnection conn, SQLiteTransaction transaction,
                                                  long cubeId, long ledgerId, LedgerSegment[] segments)
        {
            try
            {
                if (segments == null || segments.Length == 0)
                {
                    ServiceLocator.Logger?.LogDebug($"LedgerDataRepository.InsertSegmentsAndValues: no segments for CubeId={cubeId}, LedgerId={ledgerId}.");
                    return;
                }

                ServiceLocator.Logger?.LogDebug($"LedgerDataRepository.InsertSegmentsAndValues: inserting {segments.Length} segments for CubeId={cubeId}, LedgerId={ledgerId}.");

                foreach (var segment in segments)
                {
                    InsertSegment(conn, transaction, cubeId, ledgerId, segment);
                    InsertSegmentValues(conn, transaction, cubeId, ledgerId, segment);
                }
            }
            catch (Exception ex)
            {
                ServiceLocator.Logger?.LogException(ex);
            }
        }

        private static void InsertSegment(
                    SQLiteConnection conn,
                    SQLiteTransaction transaction,
                    long cubeId,
                    long ledgerId,
                    LedgerSegment segment)
        {
            try
            {
                using var cmd = conn.CreateCommand();
                cmd.Transaction = transaction;

                cmd.CommandText = @"
                                    INSERT INTO SEGMENTS (
                                        cubeId, ledgerId, coaid, segmentName, segmentValueSetId,
                                        securityEnabledFlag, defaultType, defaultValue, displaySize,
                                        segmentDelimiter, applicationColumnName
                                    )
                                    VALUES (
                                        @cubeId, @ledgerId, @coaid, @segmentName, @segmentValueSetId,
                                        @securityEnabledFlag, @defaultType, @defaultValue, @displaySize,
                                        @segmentDelimiter, @applicationColumnName
                                    )
                                    ON CONFLICT(cubeId, ledgerId, applicationColumnName)
                                    DO UPDATE SET
                                        coaid               = excluded.coaid,
                                        segmentName         = excluded.segmentName,
                                        securityEnabledFlag = excluded.securityEnabledFlag,
                                        defaultType         = excluded.defaultType,
                                        defaultValue        = excluded.defaultValue,
                                        displaySize         = excluded.displaySize,
                                        segmentDelimiter    = excluded.segmentDelimiter,
                                        applicationColumnName = excluded.applicationColumnName;
                                ";

                AddCommonParameters(cmd, cubeId, ledgerId);
                cmd.Parameters.AddWithValue("@coaid", segment.coaid);
                cmd.Parameters.AddWithValue("@segmentName", segment.segmentName ?? (object)DBNull.Value);
                cmd.Parameters.AddWithValue("@segmentValueSetId", segment.segmentValueSetId);
                cmd.Parameters.AddWithValue("@securityEnabledFlag", segment.securityEnabledFlag ?? (object)DBNull.Value);
                cmd.Parameters.AddWithValue("@defaultType", segment.defaultType ?? (object)DBNull.Value);
                cmd.Parameters.AddWithValue("@defaultValue", segment.defaultValue ?? (object)DBNull.Value);
                cmd.Parameters.AddWithValue("@displaySize", segment.displaySize);
                cmd.Parameters.AddWithValue("@segmentDelimiter", segment.segmentDelimiter ?? (object)DBNull.Value);
                cmd.Parameters.AddWithValue("@applicationColumnName", segment.applicationColumnName ?? (object)DBNull.Value);

                cmd.ExecuteNonQuery();
            }
            catch (Exception ex)
            {
                ServiceLocator.Logger?.LogException(ex,
                    $"InsertSegment failed for segment: {segment.segmentName ?? "null"} (ValueSetId: {segment.segmentValueSetId})");
            }
        }

        private static void InsertSegmentValues(
                        SQLiteConnection conn,
                        SQLiteTransaction transaction,
                        long cubeId,
                        long ledgerId,
                        LedgerSegment segment)
        {
            if (segment.segmentValues == null || segment.segmentValues.Length == 0)
                return;

            try
            {
                // Reuse command object for better performance
                using var cmd = conn.CreateCommand();
                cmd.Transaction = transaction;

                cmd.CommandText = @"
                            INSERT INTO SEGMENT_VALUES (
                                cubeId, ledgerId, segmentName, segmentValue, description,
                                summaryFlag, enabledFlag, segmentValueSetId, applicationColumnName
                            )
                            VALUES (
                                @cubeId, @ledgerId, @segmentName, @segmentValue,
                                @description, @summaryFlag, @enabledFlag,
                                @segmentValueSetId, @applicationColumnName
                            )
                            ON CONFLICT(cubeId, ledgerId, segmentValueSetId, segmentValue, summaryFlag, applicationColumnName)
                            DO UPDATE SET
                                description     = excluded.description,
                                enabledFlag     = excluded.enabledFlag,
                                segmentName     = excluded.segmentName;
                                -- Do NOT update: cubeId, ledgerId, segmentValueSetId, segmentValue, summaryFlag, applicationColumnName
                        ";

                // Define parameters once
                cmd.Parameters.AddWithValue("@cubeId", cubeId);
                cmd.Parameters.AddWithValue("@ledgerId", ledgerId);
                cmd.Parameters.Add("@segmentName", DbType.String);
                cmd.Parameters.Add("@segmentValue", DbType.String);
                cmd.Parameters.Add("@description", DbType.String);
                cmd.Parameters.Add("@summaryFlag", DbType.String);
                cmd.Parameters.Add("@enabledFlag", DbType.String);
                cmd.Parameters.AddWithValue("@segmentValueSetId", 0); // will be set per row
                cmd.Parameters.Add("@applicationColumnName", DbType.String);

                foreach (var segValue in segment.segmentValues)
                {
                    try
                    {
                        cmd.Parameters["@segmentName"].Value = segment.segmentName ?? (object)DBNull.Value;
                        cmd.Parameters["@segmentValue"].Value = segValue.segmentValue ?? (object)DBNull.Value;
                        cmd.Parameters["@description"].Value = segValue.description ?? (object)DBNull.Value;
                        cmd.Parameters["@summaryFlag"].Value = segValue.summaryFlag ?? (object)DBNull.Value;
                        cmd.Parameters["@enabledFlag"].Value = segValue.enabledFlag ?? (object)DBNull.Value;
                        cmd.Parameters["@segmentValueSetId"].Value = segValue.segmentValueSetId;
                        cmd.Parameters["@applicationColumnName"].Value = segment.applicationColumnName ?? (object)DBNull.Value;

                        cmd.ExecuteNonQuery();
                    }
                    catch (Exception ex)
                    {
                        ServiceLocator.Logger?.LogException(ex,
                            $"Segment value failed: {segValue.segmentValue ?? "null"} / summaryFlag: {segValue.summaryFlag ?? "null"}");
                    }
                }
            }
            catch (Exception ex)
            {
                ServiceLocator.Logger?.LogException(ex,
                    $"InsertSegmentValues failed for segment: {segment.segmentName ?? "null"} (ValueSetId: {segment.segmentValueSetId})");
            }
        }

        private static void AddCommonParameters(SQLiteCommand cmd, long cubeId, long ledgerId)
        {
            cmd.Parameters.AddWithValue("@cubeId", cubeId);
            cmd.Parameters.AddWithValue("@ledgerId", ledgerId);
        }
    }
}
