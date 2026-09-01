// DataRepository.cs in GLSense.Addin.Core
// Right-sized carve-out from GLSense\Repositories\DataRepository.cs (FinalWorkingCode) -
// that file is ~860 lines covering segment/period/budget/currency/encumbrance/journal
// queries used by the Balance Configurator pane (Group H) and segment pickers (Group C).
//
// Group B (Cube/Ledger selection - GLCubeDetails + UserConfig) ported:
//   - GetLedgers(cubeId): ledger list for GLCubeDetails' DataGrid.
//   - GetSegments(cubeId, ledgerId): populates the RibSegS ribbon combo in
//     GLCubeDetails.UpdateRibbonForCube right after a cube/ledger is selected.
//   - GetUserConfigs / SaveUserConfigs: backing store for Utilities\UserConfig.cs.
//
// Group C (Segment/Period pickers) added in this pass:
//   - GetTableItemsCount(cubeId, ledgerId, tableName) [static]: DONE here - this is used
//     by all 7 Group C views (to decide whether ledger setup data needs a remote refetch
//     before showing the picker) as well as Services\PeriodDataService.cs. NOTE: an
//     earlier header comment on this file deferred this method to "Group F" - that was
//     stale; Group C needed it first, so it now lives here. Group F should NOT re-add it.
//   - GetPeriods(cubeId, ledgerId): backs GLGetPeriod/GLGetPeriodByYear/GLPeriodByDate/
//     GLPeriodDetails via Services\PeriodDataService.GetPeriodsForLedger.
//   - GetCurrencies(cubeId, ledgerId): backs GLDailyRatesViewModel's currency combos.
//   - GetConfiguratorLedgers(cubeId, coaId, allLedgers): backs the ledger combo shared by
//     all 7 Group C views (they always call it with allLedgers=true).
//   - GetSegmentValues(SegmentModel) [static] / GetAllSegmentValues(cubeId, ledgerId)
//     [static]: backs GLSegmentFuncsViewModel (segment-value combo + Services\
//     SegmentDataService.GetSegmentValues respectively).
//
// Group E (Drilldowns) added in this pass:
//   - GetActivities(cubeId, ledgerId) / GetEncumbrances(cubeId, ledgerId): DONE here -
//     Models\BalanceDtoModel.cs (CreateFromXllParameters) needs both to resolve the
//     activity short/display name and encumbrance-type id list embedded in a
//     GLSense_GetBalance(...) formula. NOTE: an earlier header comment on this file
//     deferred both of these to "Group H" - that was stale; Group E needed them first,
//     so they now live here (same pattern already used when Group C pulled
//     GetTableItemsCount forward from a stale "Group F" deferral, and Group D pulled the
//     hierarchy-cache methods forward). Group H should NOT re-add them. The old
//     GetActivities returned ObservableCollection<ViewModels.GLConfiguratorViewModel.
//     ActivityModel> - that ViewModel is still Group H territory and isn't ported yet, so
//     this pass introduces a standalone Models.ActivityModel (Models\PeriodModels.cs) with
//     the same DisplayName/ShortName parsing instead of reaching into an unported
//     ViewModel. GetEncumbrances returns the already-ported Models.EncumbranceModel
//     (also added to Models\PeriodModels.cs in this pass).
//
// Group H (LOVs/Roller/Account dialogs) resolution - GetSegmentValues_RG/
// GetSegmentItemsCount/GetSegmentValuesHierarchy were previously deferred here as "still
// unowned"; now added near the bottom of this file (GLRollerGroups/GLLOVs/GLSegmentValues
// need them). Group H (Balance Configurator) should NOT re-add them.
//
// Group H (Balance Configurator) additions - port of GLSense\Repositories\
// DataRepository.cs (FinalWorkingCode) GetBudgets/GetJournalSources/GetJournalCategories
// (~lines 588-769). These were the three methods explicitly deferred by the note above
// (now resolved, same as Group C/D/E previously pulled other methods forward out of a
// stale "Group H" deferral). Re-pointed vs. the original: BaseRepository.
// ExecuteQueryObservable -> plain ADO.NET (matching the rest of this file);
// LogUtility.* -> ServiceLocator.Logger?.*. Models.BudgetModel/JournalSourceModel/
// JournalCategoryModel all live in Models\PeriodModels.cs (added alongside this pass).
//
// The original derived from BaseRepository (static SQLiteHelper.GetConnection() +
// ExecuteQueryObservable/ExecuteScalar helpers). This project has no BaseRepository yet -
// matching the CubeDataRepository.cs / LedgerDataRepository.cs pattern already
// established here, this talks to SQLiteHelper.Instance.GetConnection() directly with
// plain ADO.NET reader loops. Re-pointed vs. the original: LogUtility.* ->
// ServiceLocator.Logger.*.
using GLSense.Addin.Core.Helpers;
using GLSense.Addin.Core.Infrastructure;
using GLSense.Addin.Core.Models;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Data.SQLite;
using System.Linq;
using System.Text.Json;

namespace GLSense.Addin.Core.Repositories
{
    public class DataRepository
    {
        private const string CubeIdCol = "cubeId";
        private const string LedgerIdCol = "ledgerId";
        private const string LedgerNameCol = "ledgerName";
        private const string CoaIdCol = "coaid";
        private const string PeriodSetNameCol = "periodSetName";
        private const string CurrencyCodeCol = "currencyCode";
        private const string SegmentNameCol = "segmentName";
        private const string SegmentValueSetIdCol = "segmentValueSetId";
        private const string ApplicationColumnNameCol = "applicationColumnName";

        private const string CubeIdParam = "@cubeId";
        private const string LedgerIdParam = "@ledgerId";
        private const string SvsIdParam = "@svsid";
        private const string SegmentNameParam = "@segmentName";
        private const string ParentParam = "@parent";

        public ObservableCollection<LedgerModel> GetLedgers(long cubeId)
        {
            ServiceLocator.Logger?.LogDebug($"DataRepository.GetLedgers: CubeId={cubeId}");
            const string sql = @"
                SELECT ledgerId, cubeId, ledgerName, coaid, periodSetName, currencyCode
                FROM ledgers
                WHERE cubeId = @cubeId
                ORDER BY ledgerName ASC;";

            var result = new ObservableCollection<LedgerModel>();
            try
            {
                using var conn = SQLiteHelper.Instance.GetConnection();
                using var cmd = conn.CreateCommand();
                cmd.CommandText = sql;
                cmd.Parameters.AddWithValue(CubeIdParam, cubeId);

                using var reader = cmd.ExecuteReader();
                while (reader.Read())
                {
                    result.Add(new LedgerModel
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
                    });
                }
                ServiceLocator.Logger?.LogDebug($"DataRepository.GetLedgers: returning {result.Count} row(s) for CubeId={cubeId}");
                return result;
            }
            catch (Exception ex)
            {
                ServiceLocator.Logger?.LogException(ex, $"Query: {sql} , Parameter: CubeId={cubeId}");
                return result;
            }
        }

        public ObservableCollection<SegmentModel> GetSegments(long cubeId, long ledgerId)
        {
            ServiceLocator.Logger?.LogDebug($"DataRepository.GetSegments: CubeId={cubeId}, LedgerId={ledgerId}");
            const string sql = @"
                SELECT cubeId, ledgerId, coaid, segmentName, segmentValueSetId, securityEnabledFlag,
                       defaultType, defaultValue, displaySize, segmentDelimiter, applicationColumnName
                FROM SEGMENTS
                WHERE cubeId = @cubeId AND ledgerId = @ledgerId;";

            var result = new ObservableCollection<SegmentModel>();
            try
            {
                using var conn = SQLiteHelper.Instance.GetConnection();
                using var cmd = conn.CreateCommand();
                cmd.CommandText = sql;
                cmd.Parameters.AddWithValue(CubeIdParam, cubeId);
                cmd.Parameters.AddWithValue(LedgerIdParam, ledgerId);

                using var reader = cmd.ExecuteReader();
                while (reader.Read())
                {
                    result.Add(new SegmentModel
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
                    });
                }
                ServiceLocator.Logger?.LogDebug($"DataRepository.GetSegments: returning {result.Count} row(s) for CubeId={cubeId}, LedgerId={ledgerId}");
                return result;
            }
            catch (Exception ex)
            {
                ServiceLocator.Logger?.LogException(ex, $"Query: {sql} , Parameters: CubeId={cubeId}, LedgerId={ledgerId}");
                return result;
            }
        }

        public Dictionary<string, string> GetUserConfigs()
        {
            ServiceLocator.Logger?.LogDebug("DataRepository.GetUserConfigs started.");
            const string sql = "SELECT PreferenceKey, PreferenceValue FROM USERPREFERENCES;";

            var result = new Dictionary<string, string>();
            try
            {
                using var conn = SQLiteHelper.Instance.GetConnection();
                using var cmd = conn.CreateCommand();
                cmd.CommandText = sql;

                using var reader = cmd.ExecuteReader();
                while (reader.Read())
                {
                    var key = reader.GetString(reader.GetOrdinal("PreferenceKey"));
                    var value = reader.GetString(reader.GetOrdinal("PreferenceValue"));
                    result[key] = value;
                }
                ServiceLocator.Logger?.LogDebug($"DataRepository.GetUserConfigs: returning {result.Count} preference(s).");
                return result;
            }
            catch (Exception ex)
            {
                ServiceLocator.Logger?.LogException(ex, $"Query: {sql}");
                return result;
            }
        }

        public static int SaveUserConfigs(string key, string value)
        {
            ServiceLocator.Logger?.LogDebug($"DataRepository.SaveUserConfigs: Key={key}");
            const string insertSql = @"INSERT OR REPLACE INTO USERPREFERENCES (PreferenceKey, PreferenceValue) VALUES (@key, @value);";

            try
            {
                using var conn = SQLiteHelper.Instance.GetConnection();
                using var cmd = conn.CreateCommand();
                cmd.CommandText = insertSql;
                cmd.Parameters.AddWithValue("@key", key);
                cmd.Parameters.AddWithValue("@value", value);

                int rows = cmd.ExecuteNonQuery();
                ServiceLocator.Logger?.LogDebug($"DataRepository.SaveUserConfigs: {rows} row(s) affected for Key={key}");
                return rows;
            }
            catch (Exception ex)
            {
                ServiceLocator.Logger?.LogException(ex, $"Query: {insertSql} , Parameters: Key={key}, Value={value}");
                return 0;
            }
        }

        // ---------------------------------------------------------------------
        // Group C (Segment/Period pickers) additions
        // ---------------------------------------------------------------------

        public static int GetTableItemsCount(long cubeId, long ledgerId, string tableName)
        {
            ServiceLocator.Logger?.LogDebug($"DataRepository.GetTableItemsCount: CubeId={cubeId}, LedgerId={ledgerId}, Table={tableName}");
            string sql = $"SELECT COUNT(*) FROM {tableName} WHERE cubeId = @cubeId AND ledgerId = @ledgerId;";

            try
            {
                using var conn = SQLiteHelper.Instance.GetConnection();
                using var cmd = conn.CreateCommand();
                cmd.CommandText = sql;
                cmd.Parameters.AddWithValue(CubeIdParam, cubeId);
                cmd.Parameters.AddWithValue(LedgerIdParam, ledgerId);

                var result = cmd.ExecuteScalar();
                int count = result == null || result == DBNull.Value ? 0 : Convert.ToInt32(result);
                ServiceLocator.Logger?.LogDebug($"DataRepository.GetTableItemsCount: Table={tableName} count={count} for CubeId={cubeId}, LedgerId={ledgerId}");
                return count;
            }
            catch (Exception ex)
            {
                ServiceLocator.Logger?.LogException(ex, $"Query: {sql} , Parameters: CubeId={cubeId}, LedgerId={ledgerId}");
                return 0;
            }
        }

        public ObservableCollection<PeriodModel> GetPeriods(long cubeId, long ledgerId)
        {
            ServiceLocator.Logger?.LogDebug($"DataRepository.GetPeriods: CubeId={cubeId}, LedgerId={ledgerId}");
            const string sql = @"
                SELECT periodName, periodYear, periodNum, quarterNum, periodSetName, periodType,
                       startDate, endDate, adjustmentPeriodFlag
                FROM PERIODS
                WHERE cubeId = @cubeId AND ledgerId = @ledgerId;";

            var result = new ObservableCollection<PeriodModel>();
            try
            {
                using var conn = SQLiteHelper.Instance.GetConnection();
                using var cmd = conn.CreateCommand();
                cmd.CommandText = sql;
                cmd.Parameters.AddWithValue(CubeIdParam, cubeId);
                cmd.Parameters.AddWithValue(LedgerIdParam, ledgerId);

                using var reader = cmd.ExecuteReader();
                while (reader.Read())
                {
                    result.Add(new PeriodModel
                    {
                        CubeId = cubeId,
                        LedgerId = ledgerId,
                        PeriodName = reader.GetString(0),
                        PeriodYear = reader.GetInt32(1),
                        PeriodNum = reader.GetInt32(2),
                        QuarterNum = reader.GetInt32(3),
                        PeriodSetName = reader.GetString(4),
                        PeriodType = reader.GetString(5),
                        // Stored as UTC epoch-ms midnight-of-day timestamps with no real time-of-day
                        // component - .UtcDateTime preserves the exact calendar date. .LocalDateTime
                        // shifted every boundary by the machine's UTC offset (e.g. +5:30 for IST),
                        // which broke any full-DateTime comparison against a date landing exactly on
                        // a period boundary (a date typed as the 1st of the month failed StartDate <=
                        // date and returned #GETTING_DATA, even though the calendar day was correct).
                        StartDate = DateTimeOffset.FromUnixTimeMilliseconds(reader.GetInt64(6)).UtcDateTime,
                        EndDate = DateTimeOffset.FromUnixTimeMilliseconds(reader.GetInt64(7)).UtcDateTime,
                        AdjustmentPeriodFlag = reader.IsDBNull(8) ? "N" : reader.GetString(8)
                    });
                }
                ServiceLocator.Logger?.LogDebug($"DataRepository.GetPeriods: returning {result.Count} row(s) for CubeId={cubeId}, LedgerId={ledgerId}");
                return result;
            }
            catch (Exception ex)
            {
                ServiceLocator.Logger?.LogException(ex, $"Query: {sql} , Parameters: CubeId={cubeId}, LedgerId={ledgerId}");
                return result;
            }
        }

        public ObservableCollection<CurrencyModel> GetCurrencies(long cubeId, long ledgerId)
        {
            ServiceLocator.Logger?.LogDebug($"DataRepository.GetCurrencies: CubeId={cubeId}, LedgerId={ledgerId}");
            const string sql = "SELECT currencyCode FROM CURRENCIES WHERE cubeId = @cubeId AND ledgerId = @ledgerId;";

            var result = new ObservableCollection<CurrencyModel>();
            try
            {
                using var conn = SQLiteHelper.Instance.GetConnection();
                using var cmd = conn.CreateCommand();
                cmd.CommandText = sql;
                cmd.Parameters.AddWithValue(CubeIdParam, cubeId);
                cmd.Parameters.AddWithValue(LedgerIdParam, ledgerId);

                using var reader = cmd.ExecuteReader();
                while (reader.Read())
                {
                    result.Add(new CurrencyModel
                    {
                        CubeId = cubeId,
                        LedgerId = ledgerId,
                        CurrencyCode = reader.GetString(0)
                    });
                }
                ServiceLocator.Logger?.LogDebug($"DataRepository.GetCurrencies: returning {result.Count} row(s) for CubeId={cubeId}, LedgerId={ledgerId}");
                return result;
            }
            catch (Exception ex)
            {
                ServiceLocator.Logger?.LogException(ex, $"Query: {sql} , Parameters: CubeId={cubeId}, LedgerId={ledgerId}");
                return result;
            }
        }

        public ObservableCollection<GenericLedgerModel> GetConfiguratorLedgers(long cubeId, long coaId, bool allLedgers)
        {
            ServiceLocator.Logger?.LogDebug($"DataRepository.GetConfiguratorLedgers: CubeId={cubeId}, CoaId={coaId}, AllLedgers={allLedgers}");
            string sql = allLedgers
                ? @"SELECT ledgerId, cubeId, ledgerName, coaid, periodSetName, currencyCode
                    FROM ledgers WHERE cubeId = @cubeId ORDER BY ledgerName ASC;"
                : @"SELECT ledgerId, cubeId, ledgerName, coaid, periodSetName, currencyCode
                    FROM ledgers WHERE cubeId = @cubeId AND coaid = @coaid ORDER BY ledgerName ASC;";

            var result = new ObservableCollection<GenericLedgerModel>();
            try
            {
                using var conn = SQLiteHelper.Instance.GetConnection();
                using var cmd = conn.CreateCommand();
                cmd.CommandText = sql;
                cmd.Parameters.AddWithValue(CubeIdParam, cubeId);
                if (!allLedgers)
                    cmd.Parameters.AddWithValue("@coaid", coaId);

                using var reader = cmd.ExecuteReader();
                while (reader.Read())
                {
                    result.Add(new GenericLedgerModel
                    {
                        LedgerId = reader.GetInt64(reader.GetOrdinal(LedgerIdCol)),
                        CubeId = reader.GetInt64(reader.GetOrdinal(CubeIdCol)),
                        LedgerName = reader.GetString(reader.GetOrdinal(LedgerNameCol)),
                        CoaId = reader.GetInt32(reader.GetOrdinal(CoaIdCol)),
                        PeriodSetName = reader.GetString(reader.GetOrdinal(PeriodSetNameCol)),
                        CurrencyCode = reader.GetString(reader.GetOrdinal(CurrencyCodeCol))
                    });
                }
                ServiceLocator.Logger?.LogDebug($"DataRepository.GetConfiguratorLedgers: returning {result.Count} row(s) for CubeId={cubeId}, CoaId={coaId}, AllLedgers={allLedgers}");
                return result;
            }
            catch (Exception ex)
            {
                ServiceLocator.Logger?.LogException(ex, $"Query: {sql} , Parameters: CubeId={cubeId}, CoaId={coaId}, AllLedgers={allLedgers}");
                return result;
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
                    SegmentValue = reader.GetString(reader.GetOrdinal("segmentValue")),
                    Description = reader.GetString(reader.GetOrdinal("description")),
                    SummaryFlag = reader.GetString(reader.GetOrdinal("summaryFlag")),
                    EnabledFlag = reader.GetString(reader.GetOrdinal("enabledFlag")),
                    ApplicationColumnName = reader.GetString(reader.GetOrdinal(ApplicationColumnNameCol)),
                    Parent = string.Empty,
                    Level = 0
                };
                model.MarkLoaded();
                return model;
            }
            catch (Exception ex)
            {
                ServiceLocator.Logger?.LogException(ex, "Error creating SegmentValueModel from data reader.");
                return new SegmentValueModel();
            }
        }

        public static ObservableCollection<SegmentValueModel> GetSegmentValues(SegmentModel segment)
        {
            ServiceLocator.Logger?.LogDebug($"DataRepository.GetSegmentValues: CubeId={segment?.CubeId}, LedgerId={segment?.LedgerId}, SegmentName={segment?.SegmentName}, SvsId={segment?.SegmentValueSetId}");
            if (segment == null)
            {
                ServiceLocator.Logger?.LogWarn("DataRepository.GetSegmentValues: segment argument was null.");
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

            var result = new ObservableCollection<SegmentValueModel>();
            try
            {
                using var conn = SQLiteHelper.Instance.GetConnection();
                using var cmd = conn.CreateCommand();
                cmd.CommandText = sql;
                cmd.Parameters.AddWithValue(CubeIdParam, segment.CubeId);
                cmd.Parameters.AddWithValue(LedgerIdParam, segment.LedgerId);
                cmd.Parameters.AddWithValue("@segmentName", segment.SegmentName);
                cmd.Parameters.AddWithValue("@svsid", segment.SegmentValueSetId);

                using var reader = cmd.ExecuteReader();
                while (reader.Read())
                {
                    result.Add(CreateSegmentValueModel(reader));
                }
                ServiceLocator.Logger?.LogDebug($"DataRepository.GetSegmentValues: returning {result.Count} row(s) for SegmentName={segment?.SegmentName}");
                return result;
            }
            catch (Exception ex)
            {
                ServiceLocator.Logger?.LogException(ex, $"Query: {sql} , Parameters: CubeId={segment?.CubeId}, LedgerId={segment?.LedgerId}, SegmentName={segment?.SegmentName}, SvsId={segment?.SegmentValueSetId}");
                return result;
            }
        }

        public static ObservableCollection<SegmentValueModel> GetAllSegmentValues(long cubeId, long ledgerId)
        {
            ServiceLocator.Logger?.LogDebug($"DataRepository.GetAllSegmentValues: CubeId={cubeId}, LedgerId={ledgerId}");
            const string sql = @"
                SELECT cubeId, ledgerId, segmentName, segmentValue, description, summaryFlag, enabledFlag,
                       segmentValueSetId, applicationColumnName
                FROM SEGMENT_VALUES
                WHERE cubeId = @cubeId AND ledgerId = @ledgerId AND (summaryFlag <> 'RG' OR enabledFlag <> 'RG');";

            var result = new ObservableCollection<SegmentValueModel>();
            try
            {
                using var conn = SQLiteHelper.Instance.GetConnection();
                using var cmd = conn.CreateCommand();
                cmd.CommandText = sql;
                cmd.Parameters.AddWithValue(CubeIdParam, cubeId);
                cmd.Parameters.AddWithValue(LedgerIdParam, ledgerId);

                using var reader = cmd.ExecuteReader();
                while (reader.Read())
                {
                    result.Add(CreateSegmentValueModel(reader));
                }
                ServiceLocator.Logger?.LogDebug($"DataRepository.GetAllSegmentValues: returning {result.Count} row(s) for CubeId={cubeId}, LedgerId={ledgerId}");
                return result;
            }
            catch (Exception ex)
            {
                ServiceLocator.Logger?.LogException(ex, $"Query: {sql} , Parameters: CubeId={cubeId}, LedgerId={ledgerId}");
                return result;
            }
        }

        // ---------------------------------------------------------------------
        // Group D (Segment/Period discoverers) additions - port of
        // GLSense\Repositories\DataRepository.cs (FinalWorkingCode) lines ~253-420.
        // Re-pointed vs. the original: BaseRepository.ExecuteScalar/Execute -> plain
        // ADO.NET (SQLiteHelper.Instance.GetConnection() + manual transaction, matching
        // the rest of this file); LogUtility.* -> ServiceLocator.Logger.*;
        // ApiResponseHelper/JsonGlobals resolve via this project's Helpers namespace
        // (both already ported for Group B). TryGetRecordsNode is intentionally a
        // private copy here (not a call into Utilities.CommonFunctions.TryGetRecordsNode,
        // which is private to that class) - same as the original, which also had two
        // independent copies of this helper.
        // ---------------------------------------------------------------------

        public static bool SegmentValuesHierarchyExists(SegmentValueModel segHierarchy)
        {
            ServiceLocator.Logger?.LogDebug($"DataRepository.SegmentValuesHierarchyExists: CubeId={segHierarchy?.CubeId}, LedgerId={segHierarchy?.LedgerId}, SegmentName={segHierarchy?.SegmentName}, Parent={segHierarchy?.SegmentValue}");
            const string sql = @"
                SELECT COUNT(*)
                FROM SEGMENT_HIERARCHY_CACHE
                WHERE cubeId = @cubeId
                  AND ledgerId = @ledgerId
                  AND segmentValueSetId = @svsid
                  AND segmentName = @segmentName
                  AND parent = @parent;";

            try
            {
                using var conn = SQLiteHelper.Instance.GetConnection();
                using var cmd = conn.CreateCommand();
                cmd.CommandText = sql;
                cmd.Parameters.AddWithValue(CubeIdParam, segHierarchy.CubeId);
                cmd.Parameters.AddWithValue(LedgerIdParam, segHierarchy.LedgerId);
                cmd.Parameters.AddWithValue(SvsIdParam, segHierarchy.SegmentValueSetId);
                cmd.Parameters.AddWithValue(SegmentNameParam, segHierarchy.SegmentName);
                cmd.Parameters.AddWithValue(ParentParam, segHierarchy.SegmentValue);

                var result = cmd.ExecuteScalar();
                int count = result == null || result == DBNull.Value ? 0 : Convert.ToInt32(result);
                ServiceLocator.Logger?.LogDebug($"DataRepository.SegmentValuesHierarchyExists: count={count} for SegmentName={segHierarchy?.SegmentName}, Parent={segHierarchy?.SegmentValue}");
                return count > 0;
            }
            catch (Exception ex)
            {
                ServiceLocator.Logger?.LogException(ex, $"Query: {sql} , Parameters: CubeId={segHierarchy?.CubeId}, LedgerId={segHierarchy?.LedgerId}, SegmentName={segHierarchy?.SegmentName}, SvsId={segHierarchy?.SegmentValueSetId}, Parent={segHierarchy?.SegmentValue}");
                return false;
            }
        }

        private static bool TryGetRecordsNode(JsonElement root, out JsonElement recordsNode)
        {
            var recordProp = root.EnumerateObject()
                .FirstOrDefault(prop => string.Equals(prop.Name, "records", StringComparison.OrdinalIgnoreCase));

            if (recordProp.Value.ValueKind != JsonValueKind.Undefined)
            {
                recordsNode = recordProp.Value;
                return true;
            }

            recordsNode = default;
            return false;
        }

        public static void SaveHierarchyToCache(SegmentValueModel selectedHierarchy, string jsonData)
        {
            ServiceLocator.Logger?.LogDebug($"DataRepository.SaveHierarchyToCache: CubeId={selectedHierarchy?.CubeId}, LedgerId={selectedHierarchy?.LedgerId}, SegmentName={selectedHierarchy?.SegmentName}, Parent={selectedHierarchy?.SegmentValue}, JsonLength={jsonData?.Length ?? 0}");
            try
            {
                var result = ApiResponseHelper.Parse<JsonElement>(jsonData, JsonGlobals.Options);

                if (!result.IsSuccess)
                {
                    ServiceLocator.Logger?.LogWarn("Hierarchy JSON indicates failure.");
                    return;
                }

                JsonElement root = result.Value;

                if (!TryGetRecordsNode(root, out JsonElement recordsElem))
                    recordsElem = root;

                if (recordsElem.ValueKind != JsonValueKind.Array)
                {
                    ServiceLocator.Logger?.LogWarn("Hierarchy records not in expected array format.");
                    return;
                }

                var records = JsonSerializer.Deserialize<List<HierarchyRecord>>(
                    recordsElem.GetRawText(),
                    JsonGlobals.Options) ?? new List<HierarchyRecord>();

                using var conn = SQLiteHelper.Instance.GetConnection();
                using var transaction = conn.BeginTransaction();

                DeleteExistingHierarchy(conn, transaction, selectedHierarchy);
                InsertHierarchyRecords(conn, transaction, selectedHierarchy, records);

                transaction.Commit();
                ServiceLocator.Logger?.LogDebug($"DataRepository.SaveHierarchyToCache: committed {records.Count} record(s) for SegmentName={selectedHierarchy?.SegmentName}, Parent={selectedHierarchy?.SegmentValue}");
            }
            catch (JsonException ex)
            {
                ServiceLocator.Logger?.LogException(ex, "Error saving hierarchy to cache.");
                ServiceLocator.Logger?.LogRawJson("DataRepository.SaveHierarchyToCache", jsonData);
            }
            catch (Exception ex)
            {
                ServiceLocator.Logger?.LogException(ex, "Error saving hierarchy to cache.");
            }
        }

        private static void DeleteExistingHierarchy(SQLiteConnection conn, SQLiteTransaction transaction, SegmentValueModel model)
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

        private static void InsertHierarchyRecords(SQLiteConnection conn, SQLiteTransaction transaction, SegmentValueModel model, List<HierarchyRecord> records)
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

        // ---------------------------------------------------------------------
        // Group E (Drilldowns) additions - port of GLSense\Repositories\
        // DataRepository.cs (FinalWorkingCode) lines ~530-721 (GetActivities/
        // GetEncumbrances). Re-pointed vs. the original: BaseRepository.
        // ExecuteQueryObservable -> plain ADO.NET (matching the rest of this file);
        // LogUtility.* -> ServiceLocator.Logger.*; ViewModels.GLConfiguratorViewModel.
        // ActivityModel (unported Group H ViewModel) -> standalone Models.ActivityModel
        // (Models\PeriodModels.cs) with identical DisplayName/ShortName parsing.
        // ---------------------------------------------------------------------

        public ObservableCollection<ActivityModel> GetActivities(long cubeId, long ledgerId)
        {
            ServiceLocator.Logger?.LogDebug($"DataRepository.GetActivities: CubeId={cubeId}, LedgerId={ledgerId}");
            const string sql = "SELECT activityType FROM ACTIVITY WHERE cubeId = @cubeId AND ledgerId = @ledgerId;";

            var result = new ObservableCollection<ActivityModel>();
            try
            {
                using var conn = SQLiteHelper.Instance.GetConnection();
                using var cmd = conn.CreateCommand();
                cmd.CommandText = sql;
                cmd.Parameters.AddWithValue(CubeIdParam, cubeId);
                cmd.Parameters.AddWithValue(LedgerIdParam, ledgerId);

                using var reader = cmd.ExecuteReader();
                while (reader.Read())
                {
                    result.Add(new ActivityModel
                    {
                        CubeId = cubeId,
                        LedgerId = ledgerId,
                        ActivityType = reader.GetString(0)
                    });
                }
                ServiceLocator.Logger?.LogDebug($"DataRepository.GetActivities: returning {result.Count} row(s) for CubeId={cubeId}, LedgerId={ledgerId}");
                return result;
            }
            catch (Exception ex)
            {
                ServiceLocator.Logger?.LogException(ex, $"Query: {sql} , Parameters: CubeId={cubeId}, LedgerId={ledgerId}");
                return result;
            }
        }

        public ObservableCollection<EncumbranceModel> GetEncumbrances(long cubeId, long ledgerId)
        {
            ServiceLocator.Logger?.LogDebug($"DataRepository.GetEncumbrances: CubeId={cubeId}, LedgerId={ledgerId}");
            const string sql = "SELECT encumbranceTypeId, encumbranceType FROM ENCUMBRANCES WHERE cubeId = @cubeId AND ledgerId = @ledgerId;";

            var result = new ObservableCollection<EncumbranceModel>();
            try
            {
                using var conn = SQLiteHelper.Instance.GetConnection();
                using var cmd = conn.CreateCommand();
                cmd.CommandText = sql;
                cmd.Parameters.AddWithValue(CubeIdParam, cubeId);
                cmd.Parameters.AddWithValue(LedgerIdParam, ledgerId);

                using var reader = cmd.ExecuteReader();
                while (reader.Read())
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
                                    if (!long.TryParse(sraw, System.Globalization.NumberStyles.Integer, System.Globalization.CultureInfo.InvariantCulture, out var parsed))
                                    {
                                        ServiceLocator.Logger?.LogWarn($"Unable to parse encumbranceTypeId value '{sraw}' to Int64 for CubeId={cubeId}, LedgerId={ledgerId}.");
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
                                        ServiceLocator.Logger?.LogWarn($"Unexpected encumbranceTypeId raw type {raw?.GetType()} with value '{raw}' - CubeId={cubeId}, LedgerId={ledgerId}. Exception: {ex.Message}");
                                        ServiceLocator.Logger?.LogException(ex, "DataRepository.GetEncumbrances (default type conversion)");
                                    }
                                    break;
                            }
                        }
                        catch (Exception ex)
                        {
                            ServiceLocator.Logger?.LogWarn($"Error converting encumbranceTypeId value '{raw}' to Int64 for CubeId={cubeId}, LedgerId={ledgerId}. Exception: {ex.Message}");
                            ServiceLocator.Logger?.LogException(ex, "DataRepository.GetEncumbrances");
                            encumbranceTypeId = 0;
                        }
                    }

                    var encTypeOrd = reader.GetOrdinal("encumbranceType");
                    result.Add(new EncumbranceModel
                    {
                        CubeId = cubeId,
                        LedgerId = ledgerId,
                        EncumbranceTypeId = encumbranceTypeId,
                        EncumbranceType = reader.IsDBNull(encTypeOrd) ? string.Empty : reader.GetString(encTypeOrd)
                    });
                }
                ServiceLocator.Logger?.LogDebug($"DataRepository.GetEncumbrances: returning {result.Count} row(s) for CubeId={cubeId}, LedgerId={ledgerId}");
                return result;
            }
            catch (Exception ex)
            {
                ServiceLocator.Logger?.LogException(ex, $"Query: {sql} , Parameters: CubeId={cubeId}, LedgerId={ledgerId}");
                return result;
            }
        }

        // ---------------------------------------------------------------------
        // Group H (Balance Configurator) additions - see file header.
        // ---------------------------------------------------------------------

        public ObservableCollection<BudgetModel> GetBudgets(long cubeId, long ledgerId)
        {
            ServiceLocator.Logger?.LogDebug($"DataRepository.GetBudgets: CubeId={cubeId}, LedgerId={ledgerId}");
            const string sql = "SELECT budgetName FROM BUDGETS WHERE cubeId = @cubeId AND ledgerId = @ledgerId;";

            var result = new ObservableCollection<BudgetModel>();
            try
            {
                using var conn = SQLiteHelper.Instance.GetConnection();
                using var cmd = conn.CreateCommand();
                cmd.CommandText = sql;
                cmd.Parameters.AddWithValue(CubeIdParam, cubeId);
                cmd.Parameters.AddWithValue(LedgerIdParam, ledgerId);

                using var reader = cmd.ExecuteReader();
                while (reader.Read())
                {
                    result.Add(new BudgetModel
                    {
                        CubeId = cubeId,
                        LedgerId = ledgerId,
                        BudgetName = reader.GetString(0)
                    });
                }
                ServiceLocator.Logger?.LogDebug($"DataRepository.GetBudgets: returning {result.Count} row(s) for CubeId={cubeId}, LedgerId={ledgerId}");
                return result;
            }
            catch (Exception ex)
            {
                ServiceLocator.Logger?.LogException(ex, $"Query: {sql} , Parameters: CubeId={cubeId}, LedgerId={ledgerId}");
                return result;
            }
        }

        public ObservableCollection<JournalSourceModel> GetJournalSources(long cubeId, long ledgerId)
        {
            ServiceLocator.Logger?.LogDebug($"DataRepository.GetJournalSources: CubeId={cubeId}, LedgerId={ledgerId}");
            const string sql = "SELECT jeSourceName, sourceName FROM JOURNALSOURCES WHERE cubeId = @cubeId AND ledgerId = @ledgerId;";

            var result = new ObservableCollection<JournalSourceModel>();
            try
            {
                using var conn = SQLiteHelper.Instance.GetConnection();
                using var cmd = conn.CreateCommand();
                cmd.CommandText = sql;
                cmd.Parameters.AddWithValue(CubeIdParam, cubeId);
                cmd.Parameters.AddWithValue(LedgerIdParam, ledgerId);

                using var reader = cmd.ExecuteReader();
                while (reader.Read())
                {
                    result.Add(new JournalSourceModel
                    {
                        CubeId = cubeId,
                        LedgerId = ledgerId,
                        JeSourceName = reader.GetString(0),
                        SourceName = reader.GetString(1)
                    });
                }
                ServiceLocator.Logger?.LogDebug($"DataRepository.GetJournalSources: returning {result.Count} row(s) for CubeId={cubeId}, LedgerId={ledgerId}");
                return result;
            }
            catch (Exception ex)
            {
                ServiceLocator.Logger?.LogException(ex, $"Query: {sql} , Parameters: CubeId={cubeId}, LedgerId={ledgerId}");
                return result;
            }
        }

        public ObservableCollection<JournalCategoryModel> GetJournalCategories(long cubeId, long ledgerId)
        {
            ServiceLocator.Logger?.LogDebug($"DataRepository.GetJournalCategories: CubeId={cubeId}, LedgerId={ledgerId}");
            const string sql = "SELECT jeCategoryName, categoryName FROM JOURNALCATEGORIES WHERE cubeId = @cubeId AND ledgerId = @ledgerId;";

            var result = new ObservableCollection<JournalCategoryModel>();
            try
            {
                using var conn = SQLiteHelper.Instance.GetConnection();
                using var cmd = conn.CreateCommand();
                cmd.CommandText = sql;
                cmd.Parameters.AddWithValue(CubeIdParam, cubeId);
                cmd.Parameters.AddWithValue(LedgerIdParam, ledgerId);

                using var reader = cmd.ExecuteReader();
                while (reader.Read())
                {
                    result.Add(new JournalCategoryModel
                    {
                        CubeId = cubeId,
                        LedgerId = ledgerId,
                        JeCategoryName = reader.GetString(0),
                        CategoryName = reader.GetString(1)
                    });
                }
                ServiceLocator.Logger?.LogDebug($"DataRepository.GetJournalCategories: returning {result.Count} row(s) for CubeId={cubeId}, LedgerId={ledgerId}");
                return result;
            }
            catch (Exception ex)
            {
                ServiceLocator.Logger?.LogException(ex, $"Query: {sql} , Parameters: CubeId={cubeId}, LedgerId={ledgerId}");
                return result;
            }
        }

        // ---------------------------------------------------------------------
        // Group H (LOVs/Roller/Account dialogs) additions - port of GLSense\Repositories\
        // DataRepository.cs (FinalWorkingCode) GetSegmentValues_RG/GetSegmentItemsCount/
        // GetSegmentValuesHierarchy. GetSegmentValues_RG/GetSegmentItemsCount were
        // explicitly deferred by this file's header note ("Deliberately NOT ported here -
        // GetSegmentValues_RG (still unowned...)") pending GLRollerGroups/GLLOVs being
        // ported; that's resolved now (same pattern Group C/D/E previously used to pull
        // other methods forward out of a stale deferral). GetSegmentValuesHierarchy reads
        // back the rows SaveHierarchyToCache (already ported above, Group D) wrote into
        // SEGMENT_HIERARCHY_CACHE - it backs SegmentSelectorViewModel's hierarchy-expansion
        // combo (GLSegmentValues). Re-pointed vs. the original the same way as every other
        // method in this file: BaseRepository.ExecuteQueryObservable/ExecuteScalar -> plain
        // ADO.NET; LogUtility.* -> ServiceLocator.Logger?.*.
        // ---------------------------------------------------------------------

        public static ObservableCollection<SegmentValueModel> GetSegmentValues_RG(SegmentModel segment)
        {
            ServiceLocator.Logger?.LogDebug($"DataRepository.GetSegmentValues_RG: CubeId={segment?.CubeId}, LedgerId={segment?.LedgerId}, SegmentName={segment?.SegmentName}, SvsId={segment?.SegmentValueSetId}");
            if (segment == null)
            {
                ServiceLocator.Logger?.LogWarn("DataRepository.GetSegmentValues_RG: segment argument was null.");
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
                SELECT cubeId, ledgerId, segmentName, segmentValue, description, summaryFlag, enabledFlag,
                       segmentValueSetId, applicationColumnName
                FROM RG_Distinct
                WHERE rn = 1
                UNION ALL
                SELECT cubeId, ledgerId, segmentName, segmentValue, description, summaryFlag, enabledFlag,
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

            var result = new ObservableCollection<SegmentValueModel>();
            try
            {
                using var conn = SQLiteHelper.Instance.GetConnection();
                using var cmd = conn.CreateCommand();
                cmd.CommandText = sql;
                cmd.Parameters.AddWithValue(CubeIdParam, segment.CubeId);
                cmd.Parameters.AddWithValue(LedgerIdParam, segment.LedgerId);
                cmd.Parameters.AddWithValue(SegmentNameParam, segment.SegmentName);
                cmd.Parameters.AddWithValue(SvsIdParam, segment.SegmentValueSetId);

                using var reader = cmd.ExecuteReader();
                while (reader.Read())
                {
                    result.Add(CreateSegmentValueModel(reader));
                }
                ServiceLocator.Logger?.LogDebug($"DataRepository.GetSegmentValues_RG: returning {result.Count} row(s) for SegmentName={segment?.SegmentName}");
                return result;
            }
            catch (Exception ex)
            {
                ServiceLocator.Logger?.LogException(ex, $"Query: {sql} , Parameters: CubeId={segment?.CubeId}, LedgerId={segment?.LedgerId}, SegmentName={segment?.SegmentName}, SvsId={segment?.SegmentValueSetId}");
                return result;
            }
        }

        public static int GetSegmentItemsCount(SegmentModel segModel)
        {
            ServiceLocator.Logger?.LogDebug($"DataRepository.GetSegmentItemsCount: CubeId={segModel?.CubeId}, LedgerId={segModel?.LedgerId}, SegmentName={segModel?.SegmentName}, SvsId={segModel?.SegmentValueSetId}");
            const string sql = @"
                SELECT COUNT(*)
                FROM SEGMENT_VALUES
                WHERE cubeId = @cubeId
                  AND ledgerId = @ledgerId
                  AND segmentName = @segmentName
                  AND segmentValueSetId = @svsid;";

            try
            {
                using var conn = SQLiteHelper.Instance.GetConnection();
                using var cmd = conn.CreateCommand();
                cmd.CommandText = sql;
                cmd.Parameters.AddWithValue(CubeIdParam, segModel.CubeId);
                cmd.Parameters.AddWithValue(LedgerIdParam, segModel.LedgerId);
                cmd.Parameters.AddWithValue(SvsIdParam, segModel.SegmentValueSetId);
                cmd.Parameters.AddWithValue(SegmentNameParam, segModel.SegmentName);

                var result = cmd.ExecuteScalar();
                int count = result == null || result == DBNull.Value ? 0 : Convert.ToInt32(result);
                ServiceLocator.Logger?.LogDebug($"DataRepository.GetSegmentItemsCount: count={count} for SegmentName={segModel?.SegmentName}");
                return count;
            }
            catch (Exception ex)
            {
                ServiceLocator.Logger?.LogException(ex, $"Query: {sql} , Parameters: CubeId={segModel?.CubeId}, LedgerId={segModel?.LedgerId}, SegmentName={segModel?.SegmentName}, SvsId={segModel?.SegmentValueSetId}");
                return 0;
            }
        }

        public ObservableCollection<SegmentValueModel> GetSegmentValuesHierarchy(SegmentValueModel segValueModel)
        {
            ServiceLocator.Logger?.LogDebug($"DataRepository.GetSegmentValuesHierarchy: CubeId={segValueModel?.CubeId}, LedgerId={segValueModel?.LedgerId}, SegmentName={segValueModel?.SegmentName}, Parent={segValueModel?.SegmentValue}");
            const string sql = @"
                SELECT cubeId, ledgerId, segmentName, segmentValue, description, summaryFlag, enabledFlag,
                       segmentValueSetId, applicationColumnName, parent, lvl
                FROM SEGMENT_HIERARCHY_CACHE
                WHERE cubeId = @cubeId
                  AND ledgerId = @ledgerId
                  AND segmentValueSetId = @svsid
                  AND segmentName = @segmentName
                  AND parent = @parent;";

            var result = new ObservableCollection<SegmentValueModel>();
            try
            {
                using var conn = SQLiteHelper.Instance.GetConnection();
                using var cmd = conn.CreateCommand();
                cmd.CommandText = sql;
                cmd.Parameters.AddWithValue(CubeIdParam, segValueModel.CubeId);
                cmd.Parameters.AddWithValue(LedgerIdParam, segValueModel.LedgerId);
                cmd.Parameters.AddWithValue(SvsIdParam, segValueModel.SegmentValueSetId);
                cmd.Parameters.AddWithValue(SegmentNameParam, segValueModel.SegmentName);
                cmd.Parameters.AddWithValue(ParentParam, segValueModel.SegmentValue);

                using var reader = cmd.ExecuteReader();
                while (reader.Read())
                {
                    var model = new SegmentValueModel
                    {
                        CubeId = reader.GetInt64(reader.GetOrdinal(CubeIdCol)),
                        LedgerId = reader.GetInt64(reader.GetOrdinal(LedgerIdCol)),
                        SegmentName = reader.GetString(reader.GetOrdinal(SegmentNameCol)),
                        SegmentValueSetId = reader.GetInt64(reader.GetOrdinal(SegmentValueSetIdCol)),
                        SegmentValue = reader.GetString(reader.GetOrdinal("segmentValue")),
                        Description = reader.GetString(reader.GetOrdinal("description")),
                        SummaryFlag = reader.GetString(reader.GetOrdinal("summaryFlag")),
                        EnabledFlag = reader.GetString(reader.GetOrdinal("enabledFlag")),
                        ApplicationColumnName = reader.GetString(reader.GetOrdinal(ApplicationColumnNameCol)),
                        Parent = reader.GetString(reader.GetOrdinal("parent")),
                        Level = reader.GetInt32(reader.GetOrdinal("lvl"))
                    };
                    model.MarkLoaded();
                    result.Add(model);
                }
                ServiceLocator.Logger?.LogDebug($"DataRepository.GetSegmentValuesHierarchy: returning {result.Count} row(s) for SegmentName={segValueModel?.SegmentName}, Parent={segValueModel?.SegmentValue}");
                return result;
            }
            catch (Exception ex)
            {
                ServiceLocator.Logger?.LogException(ex, $"Query: {sql} , Parameters: CubeId={segValueModel?.CubeId}, LedgerId={segValueModel?.LedgerId}, SegmentName={segValueModel?.SegmentName}, SvsId={segValueModel?.SegmentValueSetId}, Parent={segValueModel?.SegmentValue}");
                return result;
            }
        }
    }
}
