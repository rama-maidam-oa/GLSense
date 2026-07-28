using GLSense.Utilities;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SQLite;
using System.Runtime.CompilerServices;
using System.Text;

namespace GLSense.Helpers
{
    /// <summary>
    /// Enhanced database helper with connection pooling, better error handling, and performance optimizations
    /// </summary>
    public static class DatabaseHelper
    {
        /// <summary>
        /// Executes a query and returns results with detailed logging
        /// </summary>
        public static List<T> ExecuteQuery<T>(
            string sql,
            Func<SQLiteDataReader, T> mapper,
            params SQLiteParameter[] parameters)
        {
            using (new LogUtility.LogScope($"ExecuteQuery<{typeof(T).Name}>"))
            {
                var results = new List<T>();

                try
                {
                    LogQueryDetails(sql, parameters);

                    using var connection = SQLiteHelper.GetConnection();
                    using var command = CreateCommand(connection, sql, parameters);
                    using var reader = command.ExecuteReader();

                    int rowCount = 0;
                    while (reader.Read())
                    {
                        try
                        {
                            results.Add(mapper(reader));
                            rowCount++;
                        }
                        catch (Exception ex)
                        {
                            ExceptionHelper.LogDetailedException(ex, $"Error mapping row {rowCount}");
                        }
                    }

                    LogUtility.LogDebug($"Query returned {rowCount} rows");

                    return results;
                }
                catch (Exception ex)
                {
                    ExceptionHelper.LogDetailedException(ex, $"ExecuteQuery failed. SQL: {sql}");
                    throw;
                }
            }
        }

        /// <summary>
        /// Executes a non-query with detailed logging (INSERT/UPDATE/DELETE)
        /// </summary>
        public static int ExecuteNonQuery(
            string sql,
            params SQLiteParameter[] parameters)
        {
            using (new LogUtility.LogScope("ExecuteNonQuery"))
            {
                try
                {
                    LogQueryDetails(sql, parameters);

                    using var connection = SQLiteHelper.GetConnection();
                    using var command = CreateCommand(connection, sql, parameters);

                    int affectedRows = command.ExecuteNonQuery();

                    LogUtility.LogDebug($"Query affected {affectedRows} rows");

                    return affectedRows;
                }
                catch (Exception ex)
                {
                    ExceptionHelper.LogDetailedException(ex, $"ExecuteNonQuery failed. SQL: {sql}");
                    throw;
                }
            }
        }

        /// <summary>
        /// Executes a scalar query with detailed logging
        /// </summary>
        public static T ExecuteScalar<T>(
            string sql,
            params SQLiteParameter[] parameters)
        {
            using (new LogUtility.LogScope($"ExecuteScalar<{typeof(T).Name}>"))
            {
                try
                {
                    LogQueryDetails(sql, parameters);

                    using var connection = SQLiteHelper.GetConnection();
                    using var command = CreateCommand(connection, sql, parameters);

                    var result = command.ExecuteScalar();

                    T typedResult = result is DBNull || result == null
                        ? default
                        : (T)Convert.ChangeType(result, typeof(T));

                    LogUtility.LogDebug($"Scalar result: {typedResult}");

                    return typedResult;
                }
                catch (Exception ex)
                {
                    ExceptionHelper.LogDetailedException(ex, $"ExecuteScalar failed. SQL: {sql}");
                    throw;
                }
            }
        }

        /// <summary>
        /// Executes multiple commands in a transaction with detailed logging
        /// </summary>
        public static void ExecuteInTransaction(Action<SQLiteConnection, SQLiteTransaction> action)
        {
            using (new LogUtility.LogScope("ExecuteInTransaction"))
            {
                SQLiteConnection connection = null;
                SQLiteTransaction transaction = null;

                try
                {
                    LogUtility.LogDebug("Starting database transaction");

                    connection = SQLiteHelper.GetConnection();
                    transaction = connection.BeginTransaction();

                    action(connection, transaction);

                    transaction.Commit();
                    LogUtility.LogDebug("Transaction committed successfully");
                }
                catch (Exception ex)
                {
                    LogUtility.LogError("Transaction failed - attempting rollback");

                    try
                    {
                        transaction?.Rollback();
                        LogUtility.LogDebug("Transaction rolled back successfully");
                    }
                    catch (Exception rollbackEx)
                    {
                        ExceptionHelper.LogDetailedException(rollbackEx, "Transaction rollback failed");
                    }

                    ExceptionHelper.LogDetailedException(ex, "ExecuteInTransaction failed");
                    throw;
                }
                finally
                {
                    transaction?.Dispose();
                    connection?.Dispose();
                }
            }
        }

        /// <summary>
        /// Bulk insert with transaction for better performance
        /// </summary>
        public static int BulkInsert<T>(
            string sql,
            IEnumerable<T> items,
            Action<SQLiteCommand, T> parameterizer)
        {
            using (new LogUtility.LogScope($"BulkInsert<{typeof(T).Name}>"))
            {
                int totalInserted = 0;

                try
                {
                    LogUtility.LogDebug($"Starting bulk insert: {sql}");

                    using var connection = SQLiteHelper.GetConnection();
                    using var transaction = connection.BeginTransaction();
                    using var command = connection.CreateCommand();

                    command.CommandText = sql;
                    command.Transaction = transaction;

                    foreach (var item in items)
                    {
                        try
                        {
                            command.Parameters.Clear();
                            parameterizer(command, item);
                            command.ExecuteNonQuery();
                            totalInserted++;
                        }
                        catch (Exception ex)
                        {
                            LogUtility.LogWarn($"Failed to insert item at index {totalInserted}: {ex.Message}");
                        }
                    }

                    transaction.Commit();
                    LogUtility.LogDebug($"Bulk insert completed: {totalInserted} rows inserted");

                    return totalInserted;
                }
                catch (Exception ex)
                {
                    ExceptionHelper.LogDetailedException(ex, "BulkInsert failed");
                    throw;
                }
            }
        }

        /// <summary>
        /// Creates a command with parameters
        /// </summary>
        private static SQLiteCommand CreateCommand(
            SQLiteConnection connection,
            string sql,
            SQLiteParameter[] parameters)
        {
            var command = connection.CreateCommand();
            command.CommandText = sql;

            if (parameters != null && parameters.Length > 0)
            {
                command.Parameters.AddRange(parameters);
            }

            return command;
        }

        /// <summary>
        /// Logs query details for debugging
        /// </summary>
        private static void LogQueryDetails(string sql, SQLiteParameter[] parameters)
        {
            if (!LogUtility.DebugMode) return;

            var sb = new StringBuilder();
            sb.AppendLine("SQL Query:");
            sb.AppendLine(sql);

            if (parameters != null && parameters.Length > 0)
            {
                sb.AppendLine("Parameters:");
                foreach (var param in parameters)
                {
                    sb.AppendLine($"  {param.ParameterName} = {param.Value ?? "(null)"} ({param.DbType})");
                }
            }

            LogUtility.LogDebug(sb.ToString());
        }

        /// <summary>
        /// Logs query execution error
        /// </summary>
        private static void LogQueryExecutionError(
            string sql,
            SQLiteParameter[] parameters,
            Exception ex)
        {
            var sb = new StringBuilder();
            sb.AppendLine("Query execution failed:");
            sb.AppendLine($"SQL: {sql}");

            if (parameters != null && parameters.Length > 0)
            {
                sb.AppendLine("Parameters:");
                foreach (var param in parameters)
                {
                    sb.AppendLine($"  {param.ParameterName} = {param.Value}");
                }
            }

            sb.AppendLine($"Error: {ex.Message}");

            LogUtility.LogError(sb.ToString());
        }

        /// <summary>
        /// Checks if a table exists in the database
        /// </summary>
        public static bool TableExists(string tableName)
        {
            using (new LogUtility.LogScope($"TableExists: {tableName}"))
            {
                try
                {
                    string sql = "SELECT COUNT(*) FROM sqlite_master WHERE type='table' AND name=@tableName";
                    var param = new SQLiteParameter("@tableName", tableName);

                    int count = ExecuteScalar<int>(sql, param);
                    bool exists = count > 0;

                    LogUtility.LogDebug($"Table '{tableName}' exists: {exists}");

                    return exists;
                }
                catch (Exception ex)
                {
                    ExceptionHelper.LogDetailedException(ex, $"TableExists check failed for table '{tableName}'");
                    return false;
                }
            }
        }

        /// <summary>
        /// Gets the row count of a table
        /// </summary>
        public static long GetRowCount(string tableName, string whereClause = "")
        {
            using (new LogUtility.LogScope($"GetRowCount: {tableName}"))
            {
                try
                {
                    string sql = $"SELECT COUNT(*) FROM {tableName}";

                    if (!string.IsNullOrWhiteSpace(whereClause))
                    {
                        sql += $" WHERE {whereClause}";
                    }

                    long count = ExecuteScalar<long>(sql);

                    LogUtility.LogDebug($"Table '{tableName}' row count: {count}");

                    return count;
                }
                catch (Exception ex)
                {
                    ExceptionHelper.LogDetailedException(ex, $"GetRowCount failed for table '{tableName}'");
                    return 0;
                }
            }
        }
    }
}
