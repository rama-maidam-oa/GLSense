using GLSense.Helpers;
using GLSense.Models;
using GLSense.Utilities;
using System.Data.SQLite;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Text.Json;
using System.Text.Json.Nodes;
using System.Globalization;

namespace GLSense.Repositories
{
    // Data repository for accessing ledger-related data from the SQLite database
    // Inherits from BaseRepository which provides common database access methods
    // such as ExecuteQueryObservable and ExecuteScalar
    // Each method corresponds to a specific query or operation on the database
    // and returns strongly typed models
    // Models used: LedgerModel, SegmentModel, SegmentValueModel, ActivityModel, PeriodModel, BudgetModel, CurrencyModel, EncumbranceModel, JournalCategoryModel, JournalSourceModel, GenericLedgerModel
    /// SQL parameters are used to prevent SQL injection and ensure safe queries
    /// usage example:
    /// var repo = new DataRepository();
    /// var ledgers = repo.GetLedgers(cubeId);
    /// var segments = repo.GetSegments(cubeId, ledgerId);
    public class DataRepository : BaseRepository
    {
        // Column constants
        private const string CubeIdCol = "cubeId";
        private const string LedgerIdCol = "ledgerId";
        private const string LedgerNameCol = "ledgerName";
        private const string CoaIdCol = "coaid";
        private const string PeriodSetNameCol = "periodSetName";
        private const string CurrencyCodeCol = "currencyCode";
        private const string SegmentNameCol = "segmentName";
        private const string SegmentValueSetIdCol = "segmentValueSetId";
        private const string SegmentValueCol = "segmentValue";
        private const string DescriptionCol = "description";
        private const string SummaryFlagCol = "summaryFlag";
        private const string EnabledFlagCol = "enabledFlag";
        private const string ApplicationColumnNameCol = "applicationColumnName";
        private const string ParentCol = "parent";
        private const string LevelCol = "lvl";

        // Parameter constants
        private const string CubeIdParam = "@cubeId";
        private const string LedgerIdParam = "@ledgerId";
        private const string SegmentNameParam = "@segmentName";
        private const string SvsIdParam = "@svsid";
        private const string ParentParam = "@parent";

        public ObservableCollection<LedgerModel> GetLedgers(long cubeId)
        {
            const string sql = @"
                SELECT ledgerId, cubeId, ledgerName, coaid, periodSetName, currencyCode
                FROM ledgers
                WHERE cubeId = @cubeId
                ORDER BY ledgerName ASC;";

            LogUtility.LogDebug($"DataRepository.GetLedgers: CubeId={cubeId}");
            try
            {
                return ExecuteQueryObservable(sql,
                reader => new LedgerModel
                {
                    LedgerId = reader.GetInt64(reader.GetOrdinal(LedgerIdCol)),
                    CubeId = reader.GetInt64(reader.GetOrdinal(CubeIdCol)),
                    LedgerName = reader.GetString(reader.GetOrdinal(LedgerNameCol)),
                    CoaId = reader.GetInt64(reader.GetOrdinal(CoaIdCol)),
                    PeriodSetName = reader.GetString(reader.GetOrdinal(PeriodSetNameCol)),
                    CurrencyCode = reader.GetString(reader.GetOrdinal(CurrencyCodeCol)),
                    LastRefreshedDate = "-1",
                    ADMRefreshedDate = "-1",
                    TimeZone = TimeZoneInfo.Local.DisplayName,
                    HasWarnings = false,
                    IsSelected = false
                },
                new SQLiteParameter(CubeIdParam, cubeId));
            }
            catch (Exception ex) 
            { 
                LogUtility.LogException(ex ,$"Query: {sql} , Parameter: CubeId={cubeId}");
                return new ObservableCollection<LedgerModel>();
            }
        }

        public ObservableCollection<SegmentModel> GetSegments(long cubeId, long ledgerId)
        {
            const string sql = @"
                SELECT cubeId, ledgerId, coaid, segmentName, segmentValueSetId, securityEnabledFlag,
                       defaultType, defaultValue, displaySize, segmentDelimiter, applicationColumnName
                FROM SEGMENTS
                WHERE cubeId = @cubeId AND ledgerId = @ledgerId;";

            LogUtility.LogDebug($"DataRepository.GetSegments: CubeId={cubeId}, LedgerId={ledgerId}");
            try
            {
                return ExecuteQueryObservable(sql,
                        reader => new SegmentModel
                        {
                            CubeId = reader.GetInt64(reader.GetOrdinal(CubeIdCol)),
                            LedgerId = reader.GetInt64(reader.GetOrdinal(LedgerIdCol)),
                            CoaId = reader.GetInt64(reader.GetOrdinal(CoaIdCol)),
                            SegmentName = reader.GetString(reader.GetOrdinal(SegmentNameCol)),
                            SegmentValueSetId = reader.GetInt64(reader.GetOrdinal(SegmentValueSetIdCol)),
                            SecurityEnabledFlag = reader.GetString(reader.GetOrdinal("securityEnabledFlag")),
                            DefaultType = reader.GetString(reader.GetOrdinal("defaultType")),
                            DefaultValue = reader.GetString(reader.GetOrdinal("defaultValue")),
                            DisplaySize = reader.GetInt32(reader.GetOrdinal("displaySize")),
                            SegmentDelimiter = reader.GetString(reader.GetOrdinal("segmentDelimiter")),
                            ApplicationColumnName = reader.GetString(reader.GetOrdinal(ApplicationColumnNameCol))
                        },
                        new SQLiteParameter(CubeIdParam, cubeId),
                        new SQLiteParameter(LedgerIdParam, ledgerId));
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex, $"Query: {sql} , Parameters: CubeId={cubeId}, LedgerId={ledgerId}");
                return new ObservableCollection<SegmentModel>();
            }
        }

        public static ObservableCollection<SegmentValueModel> GetSegmentValues(SegmentModel segment)
        {
            if (segment == null)
            {
                LogUtility.LogDebug("DataRepository.GetSegmentValues: segment argument is null, returning empty collection.");
                return new ObservableCollection<SegmentValueModel>();
            }

            const string sql = @"
                SELECT cubeId, ledgerId, segmentName, segmentValue, description, summaryFlag, enabledFlag,
                       segmentValueSetId, applicationColumnName
                FROM SEGMENT_VALUES
                WHERE cubeId = @cubeId
                  AND ledgerId = @ledgerId
                  AND segmentName = @segmentName
                  AND segmentValueSetId = @svsid
                  AND (summaryFlag <> 'RG' OR enabledFlag <> 'RG');";

            LogUtility.LogDebug($"DataRepository.GetSegmentValues: CubeId={segment.CubeId}, LedgerId={segment.LedgerId}, SegmentName={segment.SegmentName}, SvsId={segment.SegmentValueSetId}");
            try
            {
                return ExecuteQueryObservable(sql,
                    reader => CreateSegmentValueModel(reader),
                    new SQLiteParameter(CubeIdParam, segment.CubeId),
                    new SQLiteParameter(LedgerIdParam, segment.LedgerId),
                    new SQLiteParameter(SegmentNameParam, segment.SegmentName),
                    new SQLiteParameter(SvsIdParam, segment.SegmentValueSetId));
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex, $"Query: {sql} , Parameters: CubeId={segment?.CubeId}, LedgerId={segment?.LedgerId}, SegmentName={segment?.SegmentName}, SvsId={segment?.SegmentValueSetId}");
                return new ObservableCollection<SegmentValueModel>();
            }
        }

        public static ObservableCollection<SegmentValueModel> GetAllSegmentValues(long cubeId, long ledgerId)
        {
            const string sql = @"
                SELECT cubeId, ledgerId, segmentName, segmentValue, description, summaryFlag, enabledFlag,
                       segmentValueSetId, applicationColumnName
                FROM SEGMENT_VALUES
                WHERE cubeId = @cubeId AND ledgerId = @ledgerId AND (summaryFlag <> 'RG' OR enabledFlag <> 'RG');";

            LogUtility.LogDebug($"DataRepository.GetAllSegmentValues: CubeId={cubeId}, LedgerId={ledgerId}");
            try
            {
                return ExecuteQueryObservable(sql,
                    reader => CreateSegmentValueModel(reader),
                    new SQLiteParameter(CubeIdParam, cubeId),
                    new SQLiteParameter(LedgerIdParam, ledgerId));
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex, $"Query: {sql} , Parameters: CubeId={cubeId}, LedgerId={ledgerId}");
                return new ObservableCollection<SegmentValueModel>();
            }
        }

        public static ObservableCollection<SegmentValueModel> GetSegmentValues_RG(SegmentModel segment)
        {
            if (segment == null)
            {
                LogUtility.LogDebug("DataRepository.GetSegmentValues_RG: segment argument is null, returning empty collection.");
                return new ObservableCollection<SegmentValueModel>();
            }

            const string sql = @"
                WITH RG_Distinct AS (
                    SELECT *,
                           ROW_NUMBER() OVER (PARTITION BY segmentName, description ORDER BY id) AS rn
                    FROM SEGMENT_VALUES
                    WHERE summaryFlag = 'RG'
                      AND cubeId = @cubeId
                      AND ledgerId = @ledgerId
                      AND segmentName = @segmentName
                      AND segmentValueSetId = @svsid
                )
                SELECT id, cubeId, ledgerId, segmentName, segmentValue, description, summaryFlag, enabledFlag,
                       segmentValueSetId, applicationColumnName
                FROM RG_Distinct
                WHERE rn = 1
                UNION ALL
                SELECT id, cubeId, ledgerId, segmentName, segmentValue, description, summaryFlag, enabledFlag,
                       segmentValueSetId, applicationColumnName
                FROM SEGMENT_VALUES
                WHERE summaryFlag <> 'RG'
                  AND cubeId = @cubeId
                  AND ledgerId = @ledgerId
                  AND segmentName = @segmentName
                  AND segmentValueSetId = @svsid
                  AND EXISTS (
                      SELECT 1 FROM SEGMENT_VALUES rg
                      WHERE rg.summaryFlag = 'RG'
                        AND rg.segmentName = SEGMENT_VALUES.segmentName
                        AND rg.segmentValue = SEGMENT_VALUES.segmentValue
                        AND rg.segmentValueSetId = SEGMENT_VALUES.segmentValueSetId
                  )
                ORDER BY segmentName, segmentValue, summaryFlag ASC;";

            LogUtility.LogDebug($"DataRepository.GetSegmentValues_RG: CubeId={segment.CubeId}, LedgerId={segment.LedgerId}, SegmentName={segment.SegmentName}, SvsId={segment.SegmentValueSetId}");
            try
            {
                return ExecuteQueryObservable(sql,
                    reader => CreateSegmentValueModel(reader),
                    new SQLiteParameter(CubeIdParam, segment.CubeId),
                    new SQLiteParameter(LedgerIdParam, segment.LedgerId),
                    new SQLiteParameter(SegmentNameParam, segment.SegmentName),
                    new SQLiteParameter(SvsIdParam, segment.SegmentValueSetId));
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex, $"Query: {sql} , Parameters: CubeId={segment?.CubeId}, LedgerId={segment?.LedgerId}, SegmentName={segment?.SegmentName}, SvsId={segment?.SegmentValueSetId}");
                return new ObservableCollection<SegmentValueModel>();
            }
        }
        private static SegmentValueModel CreateSegmentValueModel(SQLiteDataReader reader)
        {
            try
            {
                var model = new SegmentValueModel
                {
                    CubeId = reader.GetInt64(reader.GetOrdinal(CubeIdCol)),
                    LedgerId = reader.GetInt64(reader.GetOrdinal(LedgerIdCol)),
                    SegmentName = reader.GetString(reader.GetOrdinal(SegmentNameCol)),
                    SegmentValueSetId = reader.GetInt64(reader.GetOrdinal(SegmentValueSetIdCol)),
                    SegmentValue = reader.GetString(reader.GetOrdinal(SegmentValueCol)),
                    Description = reader.GetString(reader.GetOrdinal(DescriptionCol)),
                    SummaryFlag = reader.GetString(reader.GetOrdinal(SummaryFlagCol)),
                    EnabledFlag = reader.GetString(reader.GetOrdinal(EnabledFlagCol)),
                    ApplicationColumnName = reader.GetString(reader.GetOrdinal(ApplicationColumnNameCol)),
                    Parent = string.Empty,
                    Level = 0
                };
                model.MarkLoaded();
                return model;
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex, "Error creating SegmentValueModel from data reader.");
                return new SegmentValueModel();
            }
        }

        public static bool SegmentValuesHierarchyExists(SegmentValueModel segHierarchy)
        {
            const string sql = @"
                SELECT COUNT(*)
                FROM SEGMENT_HIERARCHY_CACHE
                WHERE cubeId = @cubeId
                  AND ledgerId = @ledgerId
                  AND segmentValueSetId = @svsid
                  AND segmentName = @segmentName
                  AND parent = @parent;";

            LogUtility.LogDebug($"DataRepository.SegmentValuesHierarchyExists: CubeId={segHierarchy?.CubeId}, LedgerId={segHierarchy?.LedgerId}, SegmentName={segHierarchy?.SegmentName}, SvsId={segHierarchy?.SegmentValueSetId}, Parent={segHierarchy?.Parent}");
            try
            {
                int count = ExecuteScalar<int>(sql,
                            new SQLiteParameter(CubeIdParam, segHierarchy.CubeId),
                            new SQLiteParameter(LedgerIdParam, segHierarchy.LedgerId),
                            new SQLiteParameter(SvsIdParam, segHierarchy.SegmentValueSetId),
                            new SQLiteParameter(SegmentNameParam, segHierarchy.SegmentName),
                            new SQLiteParameter(ParentParam, segHierarchy.Parent));

                return count > 0;
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex, $"Query: {sql} , Parameters: CubeId={segHierarchy?.CubeId}, LedgerId={segHierarchy?.LedgerId}, SegmentName={segHierarchy?.SegmentName}, SvsId={segHierarchy?.SegmentValueSetId}, Parent={segHierarchy?.Parent}");
                return false;
            }
        }

        public ObservableCollection<SegmentValueModel> GetSegmentValuesHierarchy(SegmentValueModel segValueModel)
        {
            const string sql = @"
                SELECT cubeId, ledgerId, segmentName, segmentValue, description, summaryFlag, enabledFlag,
                       segmentValueSetId, applicationColumnName, parent, lvl
                FROM SEGMENT_HIERARCHY_CACHE
                WHERE cubeId = @cubeId
                  AND ledgerId = @ledgerId
                  AND segmentValueSetId = @svsid
                  AND segmentName = @segmentName
                  AND parent = @parent;";

            LogUtility.LogDebug($"DataRepository.GetSegmentValuesHierarchy: CubeId={segValueModel?.CubeId}, LedgerId={segValueModel?.LedgerId}, SegmentName={segValueModel?.SegmentName}, SvsId={segValueModel?.SegmentValueSetId}, Parent={segValueModel?.SegmentValue}");
            try
                {
                return ExecuteQueryObservable(sql,
                reader => new SegmentValueModel
                {
                    CubeId = reader.GetInt64(reader.GetOrdinal(CubeIdCol)),
                    LedgerId = reader.GetInt64(reader.GetOrdinal(LedgerIdCol)),
                    SegmentName = reader.GetString(reader.GetOrdinal(SegmentNameCol)),
                    SegmentValueSetId = reader.GetInt64(reader.GetOrdinal(SegmentValueSetIdCol)),
                    SegmentValue = reader.GetString(reader.GetOrdinal(SegmentValueCol)),
                    Description = reader.GetString(reader.GetOrdinal(DescriptionCol)),
                    SummaryFlag = reader.GetString(reader.GetOrdinal(SummaryFlagCol)),
                    EnabledFlag = reader.GetString(reader.GetOrdinal(EnabledFlagCol)),
                    ApplicationColumnName = reader.GetString(reader.GetOrdinal(ApplicationColumnNameCol)),
                    Parent = reader.GetString(reader.GetOrdinal(ParentCol)),
                    Level = reader.GetInt32(reader.GetOrdinal(LevelCol))
                }.Also(m => m.MarkLoaded()),
                new SQLiteParameter(CubeIdParam, segValueModel.CubeId),
                new SQLiteParameter(LedgerIdParam, segValueModel.LedgerId),
                new SQLiteParameter(SvsIdParam, segValueModel.SegmentValueSetId),
                new SQLiteParameter(SegmentNameParam, segValueModel.SegmentName),
                new SQLiteParameter(ParentParam, segValueModel.SegmentValue));
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex, $"Query: {sql} , Parameters: CubeId={segValueModel?.CubeId}, LedgerId={segValueModel?.LedgerId}, SegmentName={segValueModel?.SegmentName}, SvsId={segValueModel?.SegmentValueSetId}, Parent={segValueModel?.SegmentValue}");
                return new ObservableCollection<SegmentValueModel>();
            }
        }
        private static bool TryGetRecordsNode(
        JsonElement root,
        out JsonElement recordsNode)
        {
            var recordProp = root.EnumerateObject()
                .FirstOrDefault(prop => string.Equals(prop.Name,
                                                     "records",
                                                     StringComparison.OrdinalIgnoreCase));

            if (recordProp.Value.ValueKind != JsonValueKind.Undefined)
            {
                recordsNode = recordProp.Value;
                return true;
            }

            recordsNode = default;
            return false;
        }
        public static void SaveHierarchyToCache(
                SegmentValueModel selectedHierarchy,
                string jsonData)
        {
            LogUtility.LogDebug($"DataRepository.SaveHierarchyToCache: CubeId={selectedHierarchy?.CubeId}, LedgerId={selectedHierarchy?.LedgerId}, SegmentValue={selectedHierarchy?.SegmentValue}, jsonLength={jsonData?.Length ?? 0}");
            try
            {
                var result =
                    ApiResponseHelper.Parse<JsonElement>(jsonData, JsonGlobals.Options);

                if (!result.IsSuccess)
                {
                    LogUtility.LogWarn("Hierarchy JSON indicates failure.");
                    return;
                }

                JsonElement root = result.Value;

                if (!TryGetRecordsNode(root, out JsonElement recordsElem))
                    recordsElem = root;

                if (recordsElem.ValueKind != JsonValueKind.Array)
                {
                    LogUtility.LogWarn("Hierarchy records not in expected array format.");
                    return;
                }

                var records =
                    JsonSerializer.Deserialize<List<HierarchyRecord>>(
                        recordsElem.GetRawText(),
                        JsonGlobals.Options) ?? new();

                Execute(conn =>
                {
                    using var transaction = conn.BeginTransaction();

                    DeleteExistingHierarchy(conn, transaction, selectedHierarchy);

                    InsertHierarchyRecords(
                        conn,
                        transaction,
                        selectedHierarchy,
                        records);

                    transaction.Commit();
                });
                LogUtility.LogDebug($"DataRepository.SaveHierarchyToCache: cache updated successfully, recordCount={records.Count}");
            }
            catch (JsonException ex)
            {
                LogUtility.LogException(ex, "Error saving hierarchy to cache.");
                LogUtility.LogRawJson("DataRepository.SaveHierarchyToCache", jsonData);
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex, "Error saving hierarchy to cache.");
            }
        }
        private static void DeleteExistingHierarchy(
                    SQLiteConnection conn,
                    SQLiteTransaction transaction,
                    SegmentValueModel model)
        {
            const string deleteSql = @"
        DELETE FROM SEGMENT_HIERARCHY_CACHE
        WHERE cubeId = @cubeId
          AND ledgerId = @ledgerId
          AND segmentValueSetId = @svsid
          AND parent = @parent;";

            using var cmd = conn.CreateCommand();
            cmd.CommandText = deleteSql;
            cmd.Transaction = transaction;

            cmd.Parameters.AddWithValue(CubeIdParam, model.CubeId);
            cmd.Parameters.AddWithValue(LedgerIdParam, model.LedgerId);
            cmd.Parameters.AddWithValue(SvsIdParam, model.SegmentValueSetId);
            cmd.Parameters.AddWithValue(ParentParam, model.SegmentValue);

            cmd.ExecuteNonQuery();
        }
        private static void InsertHierarchyRecords(
    SQLiteConnection conn,
    SQLiteTransaction transaction,
    SegmentValueModel model,
    List<HierarchyRecord> records)
        {
            const string lookupSql = @"
        SELECT summaryFlag, enabledFlag
        FROM SEGMENT_VALUES
        WHERE cubeId = @cubeId
          AND ledgerId = @ledgerId
          AND segmentValueSetId = @svsid
          AND segmentValue = @segVal;";

            const string insertSql = @"
        INSERT INTO SEGMENT_HIERARCHY_CACHE
        (cubeId, ledgerId, segmentValueSetId, segmentName,
         segmentValue, description, parent, lvl,
         summaryFlag, enabledFlag, applicationColumnName)
        VALUES
        (@cubeId, @ledgerId, @svsid, @segmentName,
         @segVal, @desc, @parent, @lvl,
         @sFlag, @eFlag, @appCol);";

            foreach (var rec in records)
            {
                string summaryFlag = "N";
                string enabledFlag = "N";

                using (var lookupCmd = conn.CreateCommand())
                {
                    lookupCmd.CommandText = lookupSql;
                    lookupCmd.Transaction = transaction;

                    lookupCmd.Parameters.AddWithValue(CubeIdParam, model.CubeId);
                    lookupCmd.Parameters.AddWithValue(LedgerIdParam, model.LedgerId);
                    lookupCmd.Parameters.AddWithValue(SvsIdParam, model.SegmentValueSetId);
                    lookupCmd.Parameters.AddWithValue("@segVal", rec.segmentValue);

                    using var reader = lookupCmd.ExecuteReader();
                    if (reader.Read())
                    {
                        summaryFlag = reader.GetString(0);
                        enabledFlag = reader.GetString(1);
                    }
                }

                using var insertCmd = conn.CreateCommand();
                insertCmd.CommandText = insertSql;
                insertCmd.Transaction = transaction;

                insertCmd.Parameters.AddWithValue(CubeIdParam, model.CubeId);
                insertCmd.Parameters.AddWithValue(LedgerIdParam, model.LedgerId);
                insertCmd.Parameters.AddWithValue(SvsIdParam, model.SegmentValueSetId);
                insertCmd.Parameters.AddWithValue(SegmentNameParam, model.SegmentName);
                insertCmd.Parameters.AddWithValue("@parent", model.SegmentValue);
                insertCmd.Parameters.AddWithValue("@appCol", model.ApplicationColumnName);
                insertCmd.Parameters.AddWithValue("@segVal", rec.segmentValue);
                insertCmd.Parameters.AddWithValue("@desc", rec.description ?? (object)DBNull.Value);
                insertCmd.Parameters.AddWithValue("@lvl", rec.lvl);
                insertCmd.Parameters.AddWithValue("@sFlag", summaryFlag);
                insertCmd.Parameters.AddWithValue("@eFlag", enabledFlag);

                insertCmd.ExecuteNonQuery();
            }
        }


        public static int GetSegmentItemsCount(SegmentModel segModel)
        {
            const string sql = @"
                SELECT COUNT(*)
                FROM SEGMENT_VALUES
                WHERE cubeId = @cubeId
                  AND ledgerId = @ledgerId
                  AND segmentName = @segmentName
                  AND segmentValueSetId = @svsid;";

            LogUtility.LogDebug($"DataRepository.GetSegmentItemsCount: CubeId={segModel?.CubeId}, LedgerId={segModel?.LedgerId}, SegmentName={segModel?.SegmentName}, SvsId={segModel?.SegmentValueSetId}");
            try
            {
                return ExecuteScalar<int>(sql,
                    new SQLiteParameter(CubeIdParam, segModel.CubeId),
                    new SQLiteParameter(LedgerIdParam, segModel.LedgerId),
                    new SQLiteParameter(SvsIdParam, segModel.SegmentValueSetId),
                    new SQLiteParameter(SegmentNameParam, segModel.SegmentName));
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex, $"Query: {sql} , Parameters: CubeId={segModel?.CubeId}, LedgerId={segModel?.LedgerId}, SegmentName={segModel?.SegmentName}, SvsId={segModel?.SegmentValueSetId}");
                return 0;
            }
        }

        public static int GetTableItemsCount(long cubeId, long ledgerId, string tableName)
        {
            string sql = $"SELECT COUNT(*) FROM {tableName} WHERE cubeId = @cubeId AND ledgerId = @ledgerId;";

            LogUtility.LogDebug($"DataRepository.GetTableItemsCount: TableName={tableName}, CubeId={cubeId}, LedgerId={ledgerId}");
            try
            {
                return ExecuteScalar<int>(sql,
                    new SQLiteParameter(CubeIdParam, cubeId),
                    new SQLiteParameter(LedgerIdParam, ledgerId));
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex, $"Query: {sql} , Parameters: CubeId={cubeId}, LedgerId={ledgerId}");
                return 0;
            }
        }

        public ObservableCollection<ViewModels.GLConfiguratorViewModel.ActivityModel> GetActivities(long cubeId, long ledgerId)
        {
            const string sql = "SELECT activityType FROM ACTIVITY WHERE cubeId = @cubeId AND ledgerId = @ledgerId;";

            LogUtility.LogDebug($"DataRepository.GetActivities: CubeId={cubeId}, LedgerId={ledgerId}");
            try
            {
                return ExecuteQueryObservable(sql,
                reader => new ViewModels.GLConfiguratorViewModel.ActivityModel
                {
                    CubeId = cubeId,
                    LedgerId = ledgerId,
                    ActivityType = reader.GetString(0)
                },
                new SQLiteParameter(CubeIdParam, cubeId),
                new SQLiteParameter(LedgerIdParam, ledgerId));
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex, $"Query: {sql} , Parameters: CubeId={cubeId}, LedgerId={ledgerId}");
                return new ObservableCollection<ViewModels.GLConfiguratorViewModel.ActivityModel>();
            }
        }

        public ObservableCollection<PeriodModel> GetPeriods(long cubeId, long ledgerId)
        {
            const string sql = @"
                SELECT periodName, periodYear, periodNum, quarterNum, periodSetName, periodType,
                       startDate, endDate, adjustmentPeriodFlag
                FROM PERIODS
                WHERE cubeId = @cubeId AND ledgerId = @ledgerId;";

            LogUtility.LogDebug($"DataRepository.GetPeriods: CubeId={cubeId}, LedgerId={ledgerId}");
            try
            {
                return ExecuteQueryObservable(sql,
                        reader => new PeriodModel
                        {
                            CubeId = cubeId,
                            LedgerId = ledgerId,
                            PeriodName = reader.GetString(0),
                            PeriodYear = reader.GetInt32(1),
                            PeriodNum = reader.GetInt32(2),
                            QuarterNum = reader.GetInt32(3),
                            PeriodSetName = reader.GetString(4),
                            PeriodType = reader.GetString(5),
                            StartDate = DateTimeOffset.FromUnixTimeMilliseconds(reader.GetInt64(6)).LocalDateTime,
                            EndDate = DateTimeOffset.FromUnixTimeMilliseconds(reader.GetInt64(7)).LocalDateTime,
                            AdjustmentPeriodFlag = reader.IsDBNull(8) ? "N" : reader.GetString(8)
                        },
                        new SQLiteParameter(CubeIdParam, cubeId),
                        new SQLiteParameter(LedgerIdParam, ledgerId));
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex, $"Query: {sql} , Parameters: CubeId={cubeId}, LedgerId={ledgerId}");
                return new ObservableCollection<PeriodModel>();
            }
        }

        public ObservableCollection<BudgetModel> GetBudgets(long cubeId, long ledgerId)
        {
            const string sql = "SELECT budgetName FROM BUDGETS WHERE cubeId = @cubeId AND ledgerId = @ledgerId;";

            LogUtility.LogDebug($"DataRepository.GetBudgets: CubeId={cubeId}, LedgerId={ledgerId}");
            try
            {
                return ExecuteQueryObservable(sql,
                reader => new BudgetModel
                {
                    CubeId = cubeId,
                    LedgerId = ledgerId,
                    BudgetName = reader.GetString(0)
                },
                new SQLiteParameter(CubeIdParam, cubeId),
                new SQLiteParameter(LedgerIdParam, ledgerId));
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex, $"Query: {sql} , Parameters: CubeId={cubeId}, LedgerId={ledgerId}");
                return new ObservableCollection<BudgetModel>();
            }
        }

        public ObservableCollection<CurrencyModel> GetCurrencies(long cubeId, long ledgerId)
        {
            const string sql = "SELECT currencyCode FROM CURRENCIES WHERE cubeId = @cubeId AND ledgerId = @ledgerId;";

            LogUtility.LogDebug($"DataRepository.GetCurrencies: CubeId={cubeId}, LedgerId={ledgerId}");
            try
            {
                return ExecuteQueryObservable(sql,
                        reader => new CurrencyModel
                        {
                            CubeId = cubeId,
                            LedgerId = ledgerId,
                            CurrencyCode = reader.GetString(0)
                        },
                        new SQLiteParameter(CubeIdParam, cubeId),
                        new SQLiteParameter(LedgerIdParam, ledgerId));
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex, $"Query: {sql} , Parameters: CubeId={cubeId}, LedgerId={ledgerId}");
                return new ObservableCollection<CurrencyModel>();
            }
        }

        public ObservableCollection<EncumbranceModel> GetEncumbrances(long cubeId, long ledgerId)
        {
            const string sql = "SELECT encumbranceTypeId, encumbranceType FROM ENCUMBRANCES WHERE cubeId = @cubeId AND ledgerId = @ledgerId;";

            LogUtility.LogDebug($"DataRepository.GetEncumbrances: CubeId={cubeId}, LedgerId={ledgerId}");
            try
            {
                return ExecuteQueryObservable(sql,
                        reader =>
                        {
                            // Robustly handle encumbranceTypeId which may be stored as INTEGER, TEXT, FLOAT, or NULL
                            long encumbranceTypeId = 0;
                            int ord = reader.GetOrdinal("encumbranceTypeId");
                            if (!reader.IsDBNull(ord))
                            {
                                var raw = reader.GetValue(ord);
                                try
                                {
                                    switch (raw)
                                    {
                                        case long l:
                                            encumbranceTypeId = l;
                                            break;
                                        case int i:
                                            encumbranceTypeId = i;
                                            break;
                                        case short s:
                                            encumbranceTypeId = s;
                                            break;
                                        case byte b:
                                            encumbranceTypeId = b;
                                            break;
                                        case decimal dec:
                                            encumbranceTypeId = Convert.ToInt64(dec);
                                            break;
                                        case double d:
                                            // truncate fractional part
                                            encumbranceTypeId = Convert.ToInt64(Math.Truncate(d));
                                            break;
                                        case float f:
                                            encumbranceTypeId = Convert.ToInt64(Math.Truncate(f));
                                            break;
                                        case string sraw:
                                            if (!long.TryParse(sraw, NumberStyles.Integer, CultureInfo.InvariantCulture, out var parsed))
                                            {
                                                LogUtility.LogWarn($"Unable to parse encumbranceTypeId value '{sraw}' to Int64 for CubeId={cubeId}, LedgerId={ledgerId}.");
                                            }
                                            else
                                            {
                                                encumbranceTypeId = parsed;
                                            }
                                            break;
                                        default:
                                            try
                                            {
                                                encumbranceTypeId = Convert.ToInt64(raw);
                                            }
                                            catch (Exception ex)
                                            {
                                                LogUtility.LogWarn($"Unexpected encumbranceTypeId raw type {raw?.GetType()} with value '{raw}' - CubeId={cubeId}, LedgerId={ledgerId}. Exception: {ex.Message}");
                                            }
                                            break;
                                    }
                                }
                                catch (Exception ex)
                                {
                                    LogUtility.LogWarn($"Error converting encumbranceTypeId value '{raw}' to Int64 for CubeId={cubeId}, LedgerId={ledgerId}. Exception: {ex.Message}");
                                    encumbranceTypeId = 0;
                                }
                            }

                            var encTypeOrd = reader.GetOrdinal("encumbranceType");
                            return new EncumbranceModel
                            {
                                CubeId = cubeId,
                                LedgerId = ledgerId,
                                EncumbranceTypeId = encumbranceTypeId,
                                EncumbranceType = reader.IsDBNull(encTypeOrd) ? string.Empty : reader.GetString(encTypeOrd)
                            };
                        },
                        new SQLiteParameter(CubeIdParam, cubeId),
                        new SQLiteParameter(LedgerIdParam, ledgerId));
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex, $"Query: {sql} , Parameters: CubeId={cubeId}, LedgerId={ledgerId}");
                return new ObservableCollection<EncumbranceModel>();
            }
        }

        public ObservableCollection<JournalCategoryModel> GetJournalCategories(long cubeId, long ledgerId)
        {
            const string sql = "SELECT jeCategoryName, categoryName FROM JOURNALCATEGORIES WHERE cubeId = @cubeId AND ledgerId = @ledgerId;";

            LogUtility.LogDebug($"DataRepository.GetJournalCategories: CubeId={cubeId}, LedgerId={ledgerId}");
            try
            {
                return ExecuteQueryObservable(sql,
                        reader => new JournalCategoryModel
                        {
                            CubeId = cubeId,
                            LedgerId = ledgerId,
                            JeCategoryName = reader.GetString(0),
                            CategoryName = reader.GetString(1)
                        },
                        new SQLiteParameter(CubeIdParam, cubeId),
                        new SQLiteParameter(LedgerIdParam, ledgerId));
            }
            catch (Exception ex) 
            {
                LogUtility.LogException(ex, $"Query: {sql} , Parameters: CubeId={cubeId}, LedgerId={ledgerId}");
                return new ObservableCollection<JournalCategoryModel>();
            }
        }

        public ObservableCollection<JournalSourceModel> GetJournalSources(long cubeId, long ledgerId)
        {
            const string sql = "SELECT jeSourceName, sourceName FROM JOURNALSOURCES WHERE cubeId = @cubeId AND ledgerId = @ledgerId;";

            LogUtility.LogDebug($"DataRepository.GetJournalSources: CubeId={cubeId}, LedgerId={ledgerId}");
            try
            {
                return ExecuteQueryObservable(sql,
                    reader => new JournalSourceModel
                    {
                        CubeId = cubeId,
                        LedgerId = ledgerId,
                        JeSourceName = reader.GetString(0),
                        SourceName = reader.GetString(1)
                    },
                    new SQLiteParameter(CubeIdParam, cubeId),
                    new SQLiteParameter(LedgerIdParam, ledgerId));
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex, $"Query: {sql} , Parameters: CubeId={cubeId}, LedgerId={ledgerId}");
                return new ObservableCollection<JournalSourceModel>();
            }
        }

        public ObservableCollection<GenericLedgerModel> GetConfiguratorLedgers(long cubeId, long coaId, bool allLedgers)
        {
            string sql = allLedgers
                ? @"SELECT ledgerId, cubeId, ledgerName, coaid, periodSetName, currencyCode
                    FROM ledgers WHERE cubeId = @cubeId ORDER BY ledgerName ASC;"
                : @"SELECT ledgerId, cubeId, ledgerName, coaid, periodSetName, currencyCode
                    FROM ledgers WHERE cubeId = @cubeId AND coaid = @coaid ORDER BY ledgerName ASC;";

            LogUtility.LogDebug($"DataRepository.GetConfiguratorLedgers: CubeId={cubeId}, CoaId={coaId}, AllLedgers={allLedgers}");
            try
            {
                var parameters = new List<SQLiteParameter> { new(CubeIdParam, cubeId) };
                if (!allLedgers)
                    parameters.Add(new SQLiteParameter("@coaid", coaId));

                return ExecuteQueryObservable(sql,
                    reader => new GenericLedgerModel
                    {
                        LedgerId = reader.GetInt64(reader.GetOrdinal(LedgerIdCol)),
                        CubeId = reader.GetInt64(reader.GetOrdinal(CubeIdCol)),
                        LedgerName = reader.GetString(reader.GetOrdinal(LedgerNameCol)),
                        CoaId = reader.GetInt32(reader.GetOrdinal(CoaIdCol)),
                        PeriodSetName = reader.GetString(reader.GetOrdinal(PeriodSetNameCol)),
                        CurrencyCode = reader.GetString(reader.GetOrdinal(CurrencyCodeCol))
                    },
                    parameters.ToArray());
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex, $"Query: {sql} , Parameters: CubeId={cubeId}, CoaId={coaId}, AllLedgers={allLedgers}");
                return new ObservableCollection<GenericLedgerModel>();
            }
        }
        public Dictionary<string, string> GetUserConfigs()
        {
            const string sql = "SELECT PreferenceKey, PreferenceValue FROM USERPREFERENCES;";

            LogUtility.LogDebug("DataRepository.GetUserConfigs: fetching all user preferences.");
            try
            {
                var items = ExecuteQuery(sql,
                            reader => new
                            {
                                Key = reader.GetString(reader.GetOrdinal("PreferenceKey")),
                                Value = reader.GetString(reader.GetOrdinal("PreferenceValue"))
                            });
                return items.ToDictionary(x => x.Key, x => x.Value);
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex, $"Query: {sql}");
                return new Dictionary<string, string>();
            }
        }
        public static int SaveUserConfigs(string key, string value)
        {
            const string insertSql = @"INSERT OR REPLACE INTO USERPREFERENCES (PreferenceKey, PreferenceValue) VALUES (@key, @value);";

            LogUtility.LogDebug($"DataRepository.SaveUserConfigs: Key={key}");
            try
                {
                // First try to update
                int rowsUpdated = ExecuteNonQuery(insertSql,
                    new SQLiteParameter("@key", key),
                    new SQLiteParameter("@value", value));
                if (rowsUpdated > 0)
                {
                    return rowsUpdated; // Successfully updated
                }

                return 0;
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex, $"Query: {insertSql} , Parameters: Key={key}, Value={value}");
                return 0;
            }
        }

    }
    // Small extension helper for fluent MarkLoaded()
    internal static class Extensions
    {
        public static T Also<T>(this T obj, Action<T> action)
        {
            action(obj);
            return obj;
        }
    }
}
