using GLSense.Helpers;
using GLSense.Utilities;
using System.Data.SQLite;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;

namespace GLSense.Repositories
{
    public abstract class BaseRepository
    {
        /// <summary>
        /// Executes a non-query (INSERT/UPDATE/DELETE) and returns affected rows.
        /// </summary>
        protected static int ExecuteNonQuery(string sql, params SQLiteParameter[] parameters)
        {
            LogUtility.LogDebug($"BaseRepository.ExecuteNonQuery: sql=\"{sql}\", paramCount={parameters?.Length ?? 0}");
            return Execute(conn =>
            {
                using var cmd = CreateCommand(conn, sql, parameters);
                int affected = cmd.ExecuteNonQuery();
                LogUtility.LogDebug($"BaseRepository.ExecuteNonQuery: affectedRows={affected}");
                return affected;
            });
        }

        protected static T ExecuteScalar<T>(string sql, params SQLiteParameter[] parameters)
        {
            LogUtility.LogDebug($"BaseRepository.ExecuteScalar<{typeof(T).Name}>: sql=\"{sql}\", paramCount={parameters?.Length ?? 0}");
            return Execute(conn =>
            {
                using var cmd = CreateCommand(conn, sql, parameters);
                var result = cmd.ExecuteScalar();
                return result is DBNull || result == null ? default : (T)Convert.ChangeType(result, typeof(T));
            });
        }

        protected static List<T> ExecuteQuery<T>(string sql, Func<SQLiteDataReader, T> mapper, params SQLiteParameter[] parameters)
        {
            LogUtility.LogDebug($"BaseRepository.ExecuteQuery<{typeof(T).Name}>: sql=\"{sql}\", paramCount={parameters?.Length ?? 0}");
            return Execute(conn =>
            {
                using var cmd = CreateCommand(conn, sql, parameters);
                using var reader = cmd.ExecuteReader();

                var results = new List<T>();
                while (reader.Read())
                {
                    results.Add(mapper(reader));
                }
                LogUtility.LogDebug($"BaseRepository.ExecuteQuery<{typeof(T).Name}>: rowCount={results.Count}");
                return results;
            });
        }

        protected static ObservableCollection<T> ExecuteQueryObservable<T>(string sql, Func<SQLiteDataReader, T> mapper, params SQLiteParameter[] parameters)
        {
            var list = ExecuteQuery(sql, mapper, parameters);
            return new ObservableCollection<T>(list);
        }

        // === CHANGE THESE FROM private TO protected ===
        protected static void Execute(Action<SQLiteConnection> action)
        {
            try
            {
                using var connection = SQLiteHelper.GetConnection();
                action(connection);
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex, "BaseRepository.Execute (Action<SQLiteConnection>)");
                throw;
            }
        }

        protected static TResult Execute<TResult>(Func<SQLiteConnection, TResult> func)
        {
            try
            {
                using var connection = SQLiteHelper.GetConnection();
                return func(connection);
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex, "BaseRepository.Execute<TResult> (Func<SQLiteConnection, TResult>)");
                throw;
            }
        }
        // ===============================================

        private static SQLiteCommand CreateCommand(SQLiteConnection conn, string sql, SQLiteParameter[] parameters)
        {
            var cmd = conn.CreateCommand();
            cmd.CommandText = sql;
            if (parameters != null && parameters.Length > 0)
            {
                cmd.Parameters.AddRange(parameters);
            }
            return cmd;
        }
    }
}
