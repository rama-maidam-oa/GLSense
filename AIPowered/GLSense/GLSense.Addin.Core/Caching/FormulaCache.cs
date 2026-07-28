// FormulaCache.cs in GLSense.Addin.Core
// Port of GLSense\Caching\CachedFormulaEntry.cs (FinalWorkingCode) - CachedFormulaEntry
// (simple value+timestamp DTO), FormulaCacheManager (singleton in-memory Dictionary cache
// backed by the SQLite GLFORMULAS_CACHE table, with dirty-key tracking + PersistToDatabase/
// Flush) and CachedFormulaHelper (thin static wrapper). Used by Udf\UdfDispatcher.cs to
// cache the last-known result of every GLSense_* UDF formula so cells can still show a
// value while logged out / mid-batch-calc.
//
// ADAPTATION vs. the original (per PORTING_GUIDE conventions already used elsewhere in this
// project - see Helpers\SQLiteHelper.cs): the old FormulaCacheManager.Initialize(SQLiteConnection)
// took an externally-supplied connection and held it as a field for the add-in's whole
// lifetime. Nothing in this project holds a long-lived SQLite connection - every operation
// opens a fresh GLSense.Addin.Core.Infrastructure.ServiceLocator.Database.GetConnection()
// (already-open, ready-to-use SQLiteConnection) inside a `using` block and disposes it
// immediately. FormulaCacheManager follows that same pattern here: there is no `Initialize`
// method and no `_connection` field; every DB-touching method opens its own connection.
// EnsureTableExists() is gone entirely - the GLFORMULAS_CACHE table is now part of
// SQLiteHelper.CreateTablesIfNotExist's up-front schema (created before any cache code
// runs), so there is nothing left for FormulaCacheManager itself to create. A new
// EnsureLoaded() method (idempotent, thread-safe) replaces the old explicit Initialize call -
// it's invoked lazily from the Instance property getter, so nothing needs to explicitly
// "start" this singleton at add-in startup.
using GLSense.Addin.Core.Infrastructure;
using System;
using System.Collections.Generic;
using System.Data.SQLite;
using System.Globalization;
using System.Linq;
using System.Threading;

namespace GLSense.Addin.Core.Caching
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
    /// Application-level cache manager for formula results.
    /// Key is already a compressed/Base64-encoded string (see Udf\UdfDispatcher.BuildFunctionCacheKey).
    /// </summary>
    public sealed class FormulaCacheManager : IDisposable
    {
        #region Singleton
        private static readonly Lazy<FormulaCacheManager> _instance =
            new Lazy<FormulaCacheManager>(() => new FormulaCacheManager(), LazyThreadSafetyMode.ExecutionAndPublication);

        /// <summary>
        /// Gets the singleton instance, lazily loading its contents from SQLite on first
        /// access (see EnsureLoaded()) - no explicit startup initialization required.
        /// </summary>
        public static FormulaCacheManager Instance
        {
            get
            {
                var instance = _instance.Value;
                instance.EnsureLoaded();
                return instance;
            }
        }
        #endregion

        #region Fields
        private readonly Dictionary<string, CachedFormulaEntry> _cache = new Dictionary<string, CachedFormulaEntry>();
        private readonly HashSet<string> _dirtyKeys = new HashSet<string>();
        private readonly ReaderWriterLockSlim _lock = new ReaderWriterLockSlim();
        private readonly object _initLock = new object();
        private bool _initialized;
        private bool _disposed;
        private bool _hasChanges;
        #endregion

        #region Properties
        public int Count
        {
            get
            {
                _lock.EnterReadLock();
                try { return _cache.Count; }
                finally { _lock.ExitReadLock(); }
            }
        }

        public bool HasChanges
        {
            get
            {
                _lock.EnterReadLock();
                try { return _hasChanges; }
                finally { _lock.ExitReadLock(); }
            }
        }

        public bool IsInitialized => _initialized;
        #endregion

        #region Initialization
        private FormulaCacheManager() { }

        /// <summary>
        /// Idempotent, thread-safe lazy load of the cache contents from SQLite. Safe to call
        /// repeatedly (only the first call does any work). The GLFORMULAS_CACHE table itself
        /// is created up front by SQLiteHelper.CreateTablesIfNotExist, so there is nothing to
        /// create here - only rows to load.
        /// </summary>
        public void EnsureLoaded()
        {
            if (_initialized) return;

            lock (_initLock)
            {
                if (_initialized) return;

                try
                {
                    ServiceLocator.Logger?.LogDebug("FormulaCacheManager.EnsureLoaded: loading formula cache from database...");
                    LoadFromDatabase();
                    ServiceLocator.Logger?.LogDebug($"FormulaCacheManager.EnsureLoaded: loaded {Count} cached formula entries.");
                }
                catch (Exception ex)
                {
                    ServiceLocator.Logger?.LogException(ex, "FormulaCacheManager.EnsureLoaded: failed to load cache from database");
                }
                finally
                {
                    _initialized = true;
                }
            }
        }

        private static void LoadFromDatabase(Dictionary<string, CachedFormulaEntry> target)
        {
            const string sql = "SELECT key, value, cached_at_utc FROM GLFORMULAS_CACHE;";

            using var connection = ServiceLocator.Database.GetConnection();
            using var cmd = connection.CreateCommand();
            cmd.CommandText = sql;

            using var reader = cmd.ExecuteReader();
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

                target[key] = new CachedFormulaEntry(value, cachedAtUtc);
            }
        }

        private void LoadFromDatabase()
        {
            var loaded = new Dictionary<string, CachedFormulaEntry>();
            LoadFromDatabase(loaded);

            _lock.EnterWriteLock();
            try
            {
                _cache.Clear();
                foreach (var kvp in loaded)
                {
                    _cache[kvp.Key] = kvp.Value;
                }
            }
            finally
            {
                _lock.ExitWriteLock();
            }
        }

        /// <summary>
        /// Saves dirty (changed) entries to database using UPSERT (Insert OR Replace).
        /// Opens its own connection (per this project's SQLiteHelper convention) rather than
        /// holding one open for the add-in's lifetime.
        /// </summary>
        private void SaveDirtyEntriesToDatabase()
        {
            _lock.EnterWriteLock();
            try
            {
                if (!_hasChanges) return;

                ServiceLocator.Logger?.LogDebug($"FormulaCacheManager.SaveDirtyEntriesToDatabase: persisting {_dirtyKeys.Count} dirty key(s) to SQLite.");

                using var connection = ServiceLocator.Database.GetConnection();
                using var transaction = connection.BeginTransaction();

                const string upsertSql = @"
                INSERT OR REPLACE INTO GLFORMULAS_CACHE (key, value, cached_at_utc)
                VALUES (@key, @value, @cached_at_utc);";

                const string deleteSql = "DELETE FROM GLFORMULAS_CACHE WHERE key = @key;";

                foreach (var key in _dirtyKeys)
                {
                    if (_cache.TryGetValue(key, out var entry))
                    {
                        using var cmd = connection.CreateCommand();
                        cmd.Transaction = transaction;
                        cmd.CommandText = upsertSql;
                        cmd.Parameters.AddWithValue("@key", key);
                        cmd.Parameters.AddWithValue("@value", entry.Value ?? string.Empty);
                        cmd.Parameters.AddWithValue("@cached_at_utc", entry.CachedAtUtc.ToString("o", CultureInfo.InvariantCulture));
                        cmd.ExecuteNonQuery();
                    }
                    else
                    {
                        using var cmd = connection.CreateCommand();
                        cmd.Transaction = transaction;
                        cmd.CommandText = deleteSql;
                        cmd.Parameters.AddWithValue("@key", key);
                        cmd.ExecuteNonQuery();
                    }
                }

                transaction.Commit();

                ServiceLocator.Logger?.LogDebug($"FormulaCacheManager.SaveDirtyEntriesToDatabase: persisted {_dirtyKeys.Count} dirty key(s) successfully.");

                _dirtyKeys.Clear();
                _hasChanges = false;
            }
            finally
            {
                _lock.ExitWriteLock();
            }
        }

        /// <summary>
        /// Forces a complete cache save (full replace) - use sparingly.
        /// </summary>
        private void SaveFullCacheToDatabase()
        {
            _lock.EnterWriteLock();
            try
            {
                ServiceLocator.Logger?.LogDebug($"FormulaCacheManager.SaveFullCacheToDatabase: performing full replace of {_cache.Count} entries.");

                using var connection = ServiceLocator.Database.GetConnection();
                using var transaction = connection.BeginTransaction();

                using (var clearCmd = connection.CreateCommand())
                {
                    clearCmd.Transaction = transaction;
                    clearCmd.CommandText = "DELETE FROM GLFORMULAS_CACHE;";
                    clearCmd.ExecuteNonQuery();
                }

                const string insertSql = @"
                INSERT INTO GLFORMULAS_CACHE (key, value, cached_at_utc)
                VALUES (@key, @value, @cached_at_utc);";

                foreach (var kvp in _cache)
                {
                    using var cmd = connection.CreateCommand();
                    cmd.Transaction = transaction;
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
                _lock.ExitWriteLock();
            }
        }
        #endregion

        #region Public API
        /// <summary>
        /// Adds or updates a value in the cache
        /// </summary>
        public void AddOrUpdate(string key, object value, DateTime? cachedAtUtc = null)
        {
            EnsureLoaded();

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
                    ServiceLocator.Logger?.LogDebug($"FormulaCacheManager.AddOrUpdate: cached value for key '{key}' -> '{entry.Value}' (dirtyKeys={_dirtyKeys.Count}).");
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
            EnsureLoaded();

            if (string.IsNullOrWhiteSpace(key))
            {
                value = null;
                return false;
            }

            _lock.EnterReadLock();
            try
            {
                bool found = _cache.TryGetValue(key, out value);
                ServiceLocator.Logger?.LogDebug(found
                    ? $"FormulaCacheManager.TryGetValue: cache HIT for key '{key}' (cachedAtUtc={value.CachedAtUtc:o})"
                    : $"FormulaCacheManager.TryGetValue: cache MISS for key '{key}'");
                return found;
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
            return TryGetValue(key, out var entry) ? entry?.CachedAtUtc : (DateTime?)null;
        }

        /// <summary>
        /// Checks if a key exists
        /// </summary>
        public bool ContainsKey(string key)
        {
            EnsureLoaded();

            if (string.IsNullOrWhiteSpace(key)) return false;

            _lock.EnterReadLock();
            try { return _cache.ContainsKey(key); }
            finally { _lock.ExitReadLock(); }
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
                    ServiceLocator.Logger?.LogDebug($"FormulaCacheManager.Remove: removed key '{key}'.");
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

            ServiceLocator.Logger?.LogDebug($"FormulaCacheManager.RemoveOlderThan({cutoffUtc:o}): removed {keysToRemove.Count} entries.");
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
                ServiceLocator.Logger?.LogDebug($"FormulaCacheManager.Clear: clearing all {_cache.Count} cached entries.");
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
        /// Saves cached formulas to database - call this from a workbook-save hook.
        /// </summary>
        public void PersistToDatabase()
        {
            EnsureLoaded();

            try
            {
                if (_hasChanges)
                {
                    ServiceLocator.Logger?.LogDebug("FormulaCacheManager.PersistToDatabase: HasChanges=true, saving dirty entries.");
                    SaveDirtyEntriesToDatabase(); // Incremental save - much faster
                }
                else
                {
                    ServiceLocator.Logger?.LogDebug("FormulaCacheManager.PersistToDatabase: no changes to persist.");
                }
            }
            catch (Exception ex)
            {
                // Log but don't crash the workbook save
                ServiceLocator.Logger?.LogException(ex, "Failed to persist formula cache to database");
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
                    try { ServiceLocator.Logger?.LogException(ex, "Error disposing FormulaCacheManager"); }
                    catch { /* Silent fail during shutdown */ }
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
    /// Helper class for backward compatibility with the old monolith's call sites
    /// (CachedFormulaHelper.Store/TryGetEntry/GetValueText/GetCachedAtUtc/CreateEntry).
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
