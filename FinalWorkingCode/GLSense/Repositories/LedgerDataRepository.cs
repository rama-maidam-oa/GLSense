using GLSense.Helpers;
using GLSense.Models;
using GLSense.Utilities;
using System;
using System.Data;
using System.Data.SQLite;
using System.Text.Json;
using System.Threading.Tasks;

namespace GLSense.Repositories
{
    public class LedgerDataRepository : BaseRepository
    {
        // No constructor needed

        public static async Task InsertLedgerDataAsync(long cubeId, long ledgerId, string ledgerJson)
        {
            if (string.IsNullOrWhiteSpace(ledgerJson))
            {
                LogUtility.LogDebug($"LedgerDataRepository.InsertLedgerDataAsync: CubeId={cubeId}, LedgerId={ledgerId} - ledgerJson is null/empty, skipping insert.");
                return;
            }

            LogUtility.LogDebug($"LedgerDataRepository.InsertLedgerDataAsync: starting for CubeId={cubeId}, LedgerId={ledgerId}, jsonLength={ledgerJson.Length}");

            try
            {
                await Task.Run(() => ProcessLedgerData(cubeId, ledgerId, ledgerJson));
                LogUtility.LogDebug($"LedgerDataRepository.InsertLedgerDataAsync: completed for CubeId={cubeId}, LedgerId={ledgerId}");
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex, $"LedgerDataRepository.InsertLedgerDataAsync: failed for CubeId={cubeId}, Ledger={ledgerId}");
                // Do not throw — allow add-in to continue with partial data
            }
        }

        private static void ProcessLedgerData(long cubeId, long ledgerId, string ledgerJson)
        {
            LogUtility.LogDebug($"LedgerDataRepository.ProcessLedgerData: CubeId={cubeId}, LedgerId={ledgerId}");

            LedgerQueryData recs;
            try
            {
                recs = JsonSerializer.Deserialize<LedgerQueryData>(
                    ledgerJson,
                    JsonGlobals.Options);
            }
            catch (JsonException ex)
            {
                LogUtility.LogException(ex, "Invalid ledger JSON");
                LogUtility.LogRawJson("LedgerDataRepository.ProcessLedgerData", ledgerJson);
                return;
            }

            if (recs == null || recs.records == null)
            {
                LogUtility.LogWarn($"LedgerDataRepository.ProcessLedgerData: deserialized data or records is null for CubeId={cubeId}, LedgerId={ledgerId}. Skipping insert.");
                return;
            }

            Execute(conn =>
            {
                using var transaction = conn.BeginTransaction();

                try
                {
                    ClearExistingData(conn, transaction, cubeId, ledgerId);
                    InsertAllData(conn, transaction, cubeId, ledgerId, recs.records);
                    transaction.Commit();
                    LogUtility.LogDebug($"LedgerDataRepository.ProcessLedgerData: transaction committed for CubeId={cubeId}, LedgerId={ledgerId}");
                }
                catch (Exception ex)
                {
                    LogUtility.LogException(ex, $"LedgerDataRepository.ProcessLedgerData: error during transaction for CubeId={cubeId}, LedgerId={ledgerId}, rolling back");
                    transaction.Rollback();
                    throw;
                }
            });
        }
        private static void ClearExistingData(SQLiteConnection conn, SQLiteTransaction transaction, long cubeId, long ledgerId)
        {
            LogUtility.LogDebug($"LedgerDataRepository.ClearExistingData: CubeId={cubeId}, LedgerId={ledgerId}");
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
                LogUtility.LogException(ex);
            }


        }
        private static void InsertAllData(SQLiteConnection conn, SQLiteTransaction transaction,
                                 long cubeId, long ledgerId, Records records)
        {
            LogUtility.LogDebug($"LedgerDataRepository.InsertAllData: CubeId={cubeId}, LedgerId={ledgerId}");
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
                LogUtility.LogException(ex);
            }
        }
        private static void InsertActivity(SQLiteConnection conn, SQLiteTransaction transaction,
                                  long cubeId, long ledgerId, string[] activities)
        {
            try
            {
                if (activities == null || activities.Length == 0) return;

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
                LogUtility.LogException(ex);
            }


        }
        private static void InsertJournalSources(SQLiteConnection conn, SQLiteTransaction transaction,
                                           long cubeId, long ledgerId, JESources[] sources)
        {
            try
            {
                if (sources == null || sources.Length <= 0) return;

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
                LogUtility.LogException(ex);
            }
        }

        private static void InsertEncumbrances(SQLiteConnection conn, SQLiteTransaction transaction,
                                             long cubeId, long ledgerId, Encumbrance[] encumbrances)
        {
            try
            {

                if (encumbrances == null || encumbrances.Length == 0) return;

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
                LogUtility.LogException(ex);
            }
        }

        private static void InsertJournalCategories(SQLiteConnection conn, SQLiteTransaction transaction,
                                                  long cubeId, long ledgerId, JECategories[] categories)
        {
            try
            {
                if (categories == null || categories.Length == 0) return;

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
                LogUtility.LogException(ex);
            }
        }

        private static void InsertCurrencies(SQLiteConnection conn, SQLiteTransaction transaction,
                                           long cubeId, long ledgerId, string[] currencies)
        {
            try
            {
                if (currencies == null || currencies.Length == 0) return;

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
                LogUtility.LogException(ex);
            }
        }

        private static void InsertBudgets(SQLiteConnection conn, SQLiteTransaction transaction,
                                        long cubeId, long ledgerId, string[] budgets)
        {
            try
            {
                if (budgets == null || budgets.Length == 0) return;
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
                LogUtility.LogException(ex);
            }
        }

        private static void InsertPeriods(SQLiteConnection conn, SQLiteTransaction transaction,
                                        long cubeId, long ledgerId, Period[] periods)
        {
            try
            {
                if (periods == null || periods.Length == 0) return;

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
                LogUtility.LogException(ex);
            }
        }

        private static void InsertSegmentsAndValues(SQLiteConnection conn, SQLiteTransaction transaction,
                                                  long cubeId, long ledgerId, LedgerSegment[] segments)
        {
            try
            {
                if (segments == null || segments.Length == 0) return;

                foreach (var segment in segments)
                {
                    InsertSegment(conn, transaction, cubeId, ledgerId, segment);
                    InsertSegmentValues(conn, transaction, cubeId, ledgerId, segment);
                }
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex);
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
                LogUtility.LogException(ex,
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
                        LogUtility.LogException(ex,
                            $"Segment value failed: {segValue.segmentValue ?? "null"} / summaryFlag: {segValue.summaryFlag ?? "null"}");
                    }
                }
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex,
                    $"InsertSegmentValues failed for segment: {segment.segmentName ?? "null"} (ValueSetId: {segment.segmentValueSetId})");
            }
        }

        private static void InsertSingleSegmentValue(SQLiteConnection conn, SQLiteTransaction transaction,
                                                   long cubeId, long ledgerId, LedgerSegment segment, LedgerSegmentValue segValue)
        {
            try
            {
                using var cmd = conn.CreateCommand();
                cmd.Transaction = transaction;
                cmd.CommandText = @"INSERT OR REPLACE INTO SEGMENT_VALUES
                          (cubeId, ledgerId, segmentName, segmentValue, description,
                           summaryFlag, enabledFlag, segmentValueSetId, applicationColumnName)
                          VALUES (@cubeId, @ledgerId, @segmentName, @segmentValue,
                                  @description, @summaryFlag, @enabledFlag, 
                                  @segmentValueSetId, @applicationColumnName);";

                cmd.Parameters.Clear();
                AddCommonParameters(cmd, cubeId, ledgerId);
                cmd.Parameters.AddWithValue("@segmentName", segment.segmentName ?? (object)DBNull.Value);
                cmd.Parameters.AddWithValue("@segmentValue", segValue.segmentValue ?? (object)DBNull.Value);
                cmd.Parameters.AddWithValue("@description", segValue.description ?? (object)DBNull.Value);
                cmd.Parameters.AddWithValue("@summaryFlag", segValue.summaryFlag ?? (object)DBNull.Value);
                cmd.Parameters.AddWithValue("@enabledFlag", segValue.enabledFlag ?? (object)DBNull.Value);
                cmd.Parameters.AddWithValue("@segmentValueSetId", segValue.segmentValueSetId);
                cmd.Parameters.AddWithValue("@applicationColumnName", segment.applicationColumnName ?? (object)DBNull.Value);

                cmd.ExecuteNonQuery();
            }
            catch (Exception ex) { LogUtility.LogException(ex); }

        }

        private static void AddCommonParameters(SQLiteCommand cmd, long cubeId, long ledgerId)
        {
            cmd.Parameters.AddWithValue("@cubeId", cubeId);
            cmd.Parameters.AddWithValue("@ledgerId", ledgerId);
        }
    }
}
