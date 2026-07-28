using GLSense.Utilities;
using System;
using System.Collections.Generic;
using System.Data.SQLite;
using System.Globalization;
using System.Linq;
using System.Threading;

namespace GLSense.Caching
{
    /// <summary>
    /// Represents a cached formula entry with value and timestamp
    /// </summary>
    public sealed class CachedFormulaEntry
    {
        public string Value { get; set; }
        public DateTime CachedAtUtc { get; set; }

        public CachedFormulaEntry()
        {
            Value = string.Empty;
            CachedAtUtc = DateTime.UtcNow;
        }

        public CachedFormulaEntry(string value, DateTime cachedAtUtc)
        {
            Value = value ?? string.Empty;
            CachedAtUtc = cachedAtUtc.Kind == DateTimeKind.Unspecified
                ? DateTime.SpecifyKind(cachedAtUtc, DateTimeKind.Utc)
                : cachedAtUtc.ToUniversalTime();
        }

        public override string ToString() => Value ?? string.Empty;
    }

    /// <summary>
    /// Application-level cache manager for formula results
    /// Key is already Base64 encoded string
    /// </summary>
    public sealed class FormulaCacheManager : IDisposable
    {
        #region Singleton
        private static readonly Lazy<FormulaCacheManager> _instance =
            new Lazy<FormulaCacheManager>(() => new FormulaCacheManager(), LazyThreadSafetyMode.ExecutionAndPublication);

        public static FormulaCacheManager Instance => _instance.Value;
        #endregion

        #region Fields
        private readonly Dictionary<string, CachedFormulaEntry> _cache = new Dictionary<string, CachedFormulaEntry>();
        private readonly HashSet<string> _dirtyKeys = new HashSet<string>();
        private readonly ReaderWriterLockSlim _lock = new ReaderWriterLockSlim();
        private SQLiteConnection _connection;
        private bool _initialized;
        private bool _disposed;
        private bool _hasChanges; // Flag to track if there are unsaved changes
        #endregion

        #region Properties
        public int Count
        {
            get
            {
                _lock.EnterReadLock();
                try
                {
                    return _cache.Count;
                }
                finally
                {
                    _lock.ExitReadLock();
                }
            }
        }
        public bool HasChanges
        {
            get
            {
                _lock.EnterReadLock();
                try
                {
                    return _hasChanges;
                }
                finally
                {
                    _lock.ExitReadLock();
                }
            }
        }

        public bool IsInitialized => _initialized;
        #endregion

        #region Initialization
        private FormulaCacheManager() { }

        public void Initialize(SQLiteConnection connection)
        {
            if (_initialized) return;
            if (connection == null) throw new ArgumentNullException(nameof(connection));

            _connection = connection;
            EnsureTableExists();
            LoadFromDatabase();
            _initialized = true;
        }

        private void EnsureTableExists()
        {
            const string sql = @"
                CREATE TABLE IF NOT EXISTS GLFORMULAS_CACHE (
                    key TEXT PRIMARY KEY,
                    value TEXT NOT NULL,
                    cached_at_utc TEXT NOT NULL
                );";

            using var cmd = _connection.CreateCommand();
            cmd.CommandText = sql;
            cmd.ExecuteNonQuery();
        }

        private void LoadFromDatabase()
        {
            const string sql = "SELECT key, value, cached_at_utc FROM GLFORMULAS_CACHE;";

            using var cmd = _connection.CreateCommand();
            cmd.CommandText = sql;

            using var reader = cmd.ExecuteReader();
            _lock.EnterWriteLock();
            try
            {
                _cache.Clear();
                while (reader.Read())
                {
                    var key = reader["key"]?.ToString();
                    var value = reader["value"]?.ToString();
                    var timestamp = reader["cached_at_utc"]?.ToString();

                    if (string.IsNullOrEmpty(key)) continue;

                    DateTime cachedAtUtc = DateTime.UtcNow;
                    if (!string.IsNullOrWhiteSpace(timestamp))
                    {
                        DateTime.TryParse(timestamp, CultureInfo.InvariantCulture, DateTimeStyles.RoundtripKind, out cachedAtUtc);
                    }

                    _cache[key] = new CachedFormulaEntry(value, cachedAtUtc);
                }
            }
            finally
            {
                _lock.ExitWriteLock();
            }
        }
        /// <summary>
        /// Saves dirty (changed) entries to database using UPSERT (Insert OR Replace)
        /// </summary>
        private void SaveDirtyEntriesToDatabase()
        {
            if (!_initialized || _connection == null || !_hasChanges) return;

            _lock.EnterReadLock();
            try
            {
                using var transaction = _connection.BeginTransaction();

                const string upsertSql = @"
                INSERT OR REPLACE INTO GLFORMULAS_CACHE (key, value, cached_at_utc) 
                VALUES (@key, @value, @cached_at_utc);";

                const string deleteSql = "DELETE FROM GLFORMULAS_CACHE WHERE key = @key;";

                foreach (var key in _dirtyKeys)
                {
                    if (_cache.TryGetValue(key, out var entry))
                    {
                        // Insert or Update
                        using var cmd = _connection.CreateCommand();
                        cmd.CommandText = upsertSql;
                        cmd.Parameters.AddWithValue("@key", key);
                        cmd.Parameters.AddWithValue("@value", entry.Value ?? string.Empty);
                        cmd.Parameters.AddWithValue("@cached_at_utc", entry.CachedAtUtc.ToString("o", CultureInfo.InvariantCulture));
                        cmd.ExecuteNonQuery();
                    }
                    else
                    {
                        // Delete if key was removed
                        using var cmd = _connection.CreateCommand();
                        cmd.CommandText = deleteSql;
                        cmd.Parameters.AddWithValue("@key", key);
                        cmd.ExecuteNonQuery();
                    }
                }

                transaction.Commit();

                // Clear dirty keys after successful save
                _dirtyKeys.Clear();
                _hasChanges = false;
            }
            finally
            {
                _lock.ExitReadLock();
            }
        }
        /// <summary>
        /// Forces a complete cache save (full replace) - Use sparingly
        /// </summary>
        private void SaveFullCacheToDatabase()
        {
            if (!_initialized || _connection == null) return;

            _lock.EnterReadLock();
            try
            {
                using var transaction = _connection.BeginTransaction();

                // Clear the table
                using (var clearCmd = _connection.CreateCommand())
                {
                    clearCmd.CommandText = "DELETE FROM GLFORMULAS_CACHE;";
                    clearCmd.ExecuteNonQuery();
                }

                // Insert all entries
                const string insertSql = @"
                INSERT INTO GLFORMULAS_CACHE (key, value, cached_at_utc) 
                VALUES (@key, @value, @cached_at_utc);";

                foreach (var kvp in _cache)
                {
                    using var cmd = _connection.CreateCommand();
                    cmd.CommandText = insertSql;
                    cmd.Parameters.AddWithValue("@key", kvp.Key);
                    cmd.Parameters.AddWithValue("@value", kvp.Value.Value ?? string.Empty);
                    cmd.Parameters.AddWithValue("@cached_at_utc", kvp.Value.CachedAtUtc.ToString("o", CultureInfo.InvariantCulture));
                    cmd.ExecuteNonQuery();
                }

                transaction.Commit();
                _dirtyKeys.Clear();
                _hasChanges = false;
            }
            finally
            {
                _lock.ExitReadLock();
            }
        }
        #endregion

        #region Public API
        /// <summary>
        /// Adds or updates a value in the cache
        /// </summary>
        public void AddOrUpdate(string key, object value, DateTime? cachedAtUtc = null)
        {
            if (string.IsNullOrWhiteSpace(key))
                throw new ArgumentException("Key cannot be null or whitespace", nameof(key));

            var entry = value as CachedFormulaEntry ??
                        new CachedFormulaEntry(value?.ToString() ?? string.Empty, cachedAtUtc ?? DateTime.UtcNow);

            _lock.EnterWriteLock();
            try
            {
                bool isNewOrChanged = !_cache.TryGetValue(key, out var existing) ||
                                      existing.Value != entry.Value;

                if (isNewOrChanged)
                {
                    _cache[key] = entry;
                    _dirtyKeys.Add(key);
                    _hasChanges = true;
                }
            }
            finally
            {
                _lock.ExitWriteLock();
            }
        }

        /// <summary>
        /// Tries to get a value from the cache
        /// </summary>
        public bool TryGetValue(string key, out CachedFormulaEntry value)
        {
            if (string.IsNullOrWhiteSpace(key))
            {
                value = null;
                return false;
            }

            _lock.EnterReadLock();
            try
            {
                return _cache.TryGetValue(key, out value);
            }
            finally
            {
                _lock.ExitReadLock();
            }
        }

        /// <summary>
        /// Gets the string value from cache
        /// </summary>
        public string GetValue(string key)
        {
            return TryGetValue(key, out var entry) ? entry?.Value : null;
        }

        /// <summary>
        /// Gets the timestamp when entry was cached
        /// </summary>
        public DateTime? GetTimestamp(string key)
        {
            return TryGetValue(key, out var entry) ? entry?.CachedAtUtc : null;
        }

        /// <summary>
        /// Checks if a key exists
        /// </summary>
        public bool ContainsKey(string key)
        {
            if (string.IsNullOrWhiteSpace(key)) return false;

            _lock.EnterReadLock();
            try
            {
                return _cache.ContainsKey(key);
            }
            finally
            {
                _lock.ExitReadLock();
            }
        }

        /// <summary>
        /// Removes a specific entry
        /// </summary>
        public bool Remove(string key)
        {
            if (string.IsNullOrWhiteSpace(key)) return false;

            _lock.EnterWriteLock();
            try
            {
                if (_cache.Remove(key))
                {
                    _dirtyKeys.Add(key);
                    _hasChanges = true;
                    return true;
                }
                return false;
            }
            finally
            {
                _lock.ExitWriteLock();
            }
        }

        /// <summary>
        /// Removes entries older than specified date
        /// </summary>
        public int RemoveOlderThan(DateTime cutoffUtc)
        {
            var keysToRemove = new List<string>();

            _lock.EnterWriteLock();
            try
            {
                foreach (var kvp in _cache)
                {
                    if (kvp.Value.CachedAtUtc < cutoffUtc)
                    {
                        keysToRemove.Add(kvp.Key);
                    }
                }

                foreach (var key in keysToRemove)
                {
                    _cache.Remove(key);
                    _dirtyKeys.Add(key);
                }

                if (keysToRemove.Count > 0)
                {
                    _hasChanges = true;
                }
            }
            finally
            {
                _lock.ExitWriteLock();
            }

            return keysToRemove.Count;
        }

        /// <summary>
        /// Clears all entries
        /// </summary>
        public void Clear()
        {
            _lock.EnterWriteLock();
            try
            {
                _cache.Clear();
                _dirtyKeys.Clear();
                _hasChanges = true; // Mark for save
            }
            finally
            {
                _lock.ExitWriteLock();
            }
        }
        /// <summary>
        /// Saves cached formulas to database - Call this from Workbook_BeforeSave event
        /// </summary>
        public void PersistToDatabase()
        {
            if (!_initialized || _connection == null) return;

            try
            {
                // Only save if there are changes
                if (_hasChanges)
                {
                    SaveDirtyEntriesToDatabase(); // Incremental save - much faster
                                                  // OR use SaveFullCacheToDatabase() if you prefer full save
                }
            }
            catch (Exception ex)
            {
                // Log but don't crash the workbook save
                ShutdownLogger.LogError("Failed to persist cache to database", ex);
            }
        }
        /// <summary>
        /// Forces an immediate save to database
        /// </summary>
        public void Flush()
        {
            PersistToDatabase();
        }
        #endregion

        #region IDisposable
        public void Dispose()
        {
            if (!_disposed)
            {
                try
                {
                    _lock?.Dispose();
                }
                catch (Exception ex)
                {
                    // Log but don't rethrow during shutdown
                    try { ShutdownLogger.LogError("Error disposing FormulaCacheManager", ex); }
                    catch { /* Silent fail */ }
                }
                finally
                {
                    _disposed = true;
                }
            }
        }
        #endregion
    }

    /// <summary>
    /// Helper class for backward compatibility
    /// </summary>
    public static class CachedFormulaHelper
    {
        public static void Store(string key, object value, DateTime? cachedAtUtc = null)
        {
            FormulaCacheManager.Instance.AddOrUpdate(key, value, cachedAtUtc);
        }

        public static bool TryGetEntry(object value, out CachedFormulaEntry entry)
        {
            if (value is CachedFormulaEntry cachedEntry)
            {
                entry = cachedEntry;
                return true;
            }

            if (value != null)
            {
                entry = new CachedFormulaEntry(value.ToString() ?? string.Empty, DateTime.MinValue);
                return true;
            }

            entry = null;
            return false;
        }

        public static string GetValueText(object value)
        {
            if (value is CachedFormulaEntry cachedEntry)
            {
                return cachedEntry.Value ?? string.Empty;
            }
            return value?.ToString() ?? string.Empty;
        }

        public static DateTime GetCachedAtUtc(object value)
        {
            if (value is CachedFormulaEntry cachedEntry && cachedEntry.CachedAtUtc != default)
            {
                return cachedEntry.CachedAtUtc.Kind == DateTimeKind.Unspecified
                    ? DateTime.SpecifyKind(cachedEntry.CachedAtUtc, DateTimeKind.Utc)
                    : cachedEntry.CachedAtUtc.ToUniversalTime();
            }
            return DateTime.MinValue;
        }

        public static CachedFormulaEntry CreateEntry(object value, DateTime? cachedAtUtc = null)
        {
            if (value is CachedFormulaEntry entry)
            {
                return entry;
            }
            return new CachedFormulaEntry(value?.ToString() ?? string.Empty, cachedAtUtc ?? DateTime.UtcNow);
        }
    }
}