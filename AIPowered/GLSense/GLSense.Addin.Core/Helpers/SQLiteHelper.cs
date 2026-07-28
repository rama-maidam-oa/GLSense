// GLSense.Addin.Core/Helpers/SQLiteHelper.cs
using GLSense.Addin.Core.Infrastructure;
using System;
using System.Data.SQLite;
using System.IO;
using System.Runtime.InteropServices;

namespace GLSense.Addin.Core.Helpers
{
    public class SQLiteHelper
    {
        private static SQLiteHelper _instance;
        private static readonly object _instanceLock = new object();
        private bool _initialized = false;
        private readonly object _initLock = new object();
        private string _dbPath;

        // Private constructor to prevent direct instantiation
        private SQLiteHelper()
        {
        }

        /// <summary>
        /// Gets the singleton instance of SQLiteHelper
        /// </summary>
        public static SQLiteHelper Instance
        {
            get
            {
                if (_instance == null)
                {
                    lock (_instanceLock)
                    {
                        if (_instance == null)
                        {
                            _instance = new SQLiteHelper();
                        }
                    }
                }
                return _instance;
            }
        }

        /// <summary>
        /// Indicates whether the database already contains persisted cube data.
        /// Used to avoid clearing tables when opening a new Excel instance.
        /// </summary>
        public bool HasPersistedData()
        {
            ServiceLocator.Logger?.LogDebug("SQLiteHelper.HasPersistedData: checking for existing persisted cube data.");
            try
            {
                EnsureInitialized();

                using var connection = GetConnection();
                using var cmd = connection.CreateCommand();
                cmd.CommandText = "SELECT 1 FROM CUBES LIMIT 1";
                ServiceLocator.Logger?.LogDebug($"SQLiteHelper.HasPersistedData: executing query on connection - {cmd.CommandText}");
                var result = cmd.ExecuteScalar();
                bool hasData = result != null && result != DBNull.Value;
                ServiceLocator.Logger?.LogDebug($"SQLiteHelper.HasPersistedData: result = {hasData}");
                return hasData;
            }
            catch (Exception ex)
            {
                ServiceLocator.Logger?.LogException(ex, "SQLiteHelper.HasPersistedData - Failed to check for persisted SQLite data");
                return false;
            }
        }

        [DllImport("kernel32.dll", SetLastError = true)]
        private static extern bool SetDllDirectory(string lpPathName);

        /// <summary>
        /// Loads the native SQLite DLL from the correct architecture-specific subdirectory.
        /// This must be called before any SQLite operations.
        /// </summary>
        private static void LoadNativeSQLiteDll()
        {
            try
            {
                string platform = Environment.Is64BitProcess ? "x64" : "x86";
                string versionFolderName = $"V{ServiceLocator.Version}";
                string dllPath = Path.Combine(ServiceLocator.Paths.VersionsPath, versionFolderName, platform);

                ServiceLocator.Logger?.LogDebug($"SQLiteHelper.LoadNativeSQLiteDll: attempting to load SQLite native DLL from: {dllPath}");

                if (Directory.Exists(dllPath))
                {
                    SetDllDirectory(dllPath);
                    ServiceLocator.Logger?.LogDebug($"SQLiteHelper.LoadNativeSQLiteDll: SQLite native DLL directory set to: {dllPath}");
                }
                else
                {
                    ServiceLocator.Logger?.LogWarn($"SQLiteHelper.LoadNativeSQLiteDll: SQLite native DLL directory not found: {dllPath}");
                }
            }
            catch (Exception ex)
            {
                ServiceLocator.Logger?.LogException(ex, "SQLiteHelper.LoadNativeSQLiteDll - Error loading native SQLite DLL");
            }
        }

        /// <summary>
        /// Ensures the database is initialized. Called automatically before any operation.
        /// </summary>
        private void EnsureInitialized()
        {
            if (!_initialized)
            {
                InitializeDatabase();
            }
        }

        /// <summary>
        /// Initializes the SQLite database.
        /// - Ensures the database folder exists
        /// - Creates the .db file if it doesn't exist
        /// - Runs schema creation (tables, indexes, etc.)
        /// Call this once at add-in startup.
        /// </summary>
        public void InitializeDatabase()
        {
            lock (_initLock)
            {
                if (_initialized)
                {
                    ServiceLocator.Logger?.LogDebug("SQLiteHelper.InitializeDatabase: SQLite database already initialized");
                    return;
                }

                // LogInfo is not used in this codebase (per project convention) - LogDebug is
                // the safe/cheap equivalent, gated by the ribbon's Debug checkbox.
                ServiceLocator.Logger?.LogDebug("SQLiteHelper.InitializeDatabase: initializing SQLite database...");

                try
                {
                    // Load the native SQLite DLL first
                    LoadNativeSQLiteDll();

                    _dbPath = Path.Combine(ServiceLocator.Paths.Database, "GLSense.sqlite");
                    ServiceLocator.Logger?.LogDebug($"SQLiteHelper.InitializeDatabase: database path resolved to '{_dbPath}'");

                    // Ensure the directory exists
                    string dbFolder = Path.GetDirectoryName(_dbPath);
                    if (!Directory.Exists(dbFolder))
                    {
                        Directory.CreateDirectory(dbFolder);
                        ServiceLocator.Logger?.LogDebug($"SQLiteHelper.InitializeDatabase: created database folder: {dbFolder}");
                    }

                    // Open connection in ReadWriteCreate mode - creates file if missing
                    ServiceLocator.Logger?.LogDebug($"SQLiteHelper.InitializeDatabase: opening connection to '{_dbPath}' (ReadWriteCreate)");
                    using var connection = new SQLiteConnection($"Data Source={_dbPath};Version=3;");
                    connection.Open();

                    // Create tables, indexes, etc.
                    CreateTablesIfNotExist(connection);

                    _initialized = true;
                    ServiceLocator.Logger?.LogDebug($"SQLiteHelper.InitializeDatabase: SQLite database initialized successfully at: {_dbPath}");
                }
                catch (Exception ex)
                {
                    ServiceLocator.Logger?.LogException(ex, "SQLiteHelper.InitializeDatabase - Failed to initialize SQLite database");
                    throw;
                }
            }
        }

        /// <summary>
        /// Returns an open connection to the database.
        /// Automatically initializes if not already done.
        /// </summary>
        public SQLiteConnection GetConnection()
        {
            EnsureInitialized();

            ServiceLocator.Logger?.LogDebug($"SQLiteHelper.GetConnection: opening new connection to '{_dbPath}'");
            var connection = new SQLiteConnection($"Data Source={_dbPath};Version=3;");
            connection.Open();
            return connection;
        }

        private static void CreateTablesIfNotExist(SQLiteConnection connection)
        {
            ServiceLocator.Logger?.LogDebug("SQLiteHelper.CreateTablesIfNotExist: creating tables/indexes if not present (CUBES, LEDGERS, JOURNALSOURCES, ACTIVITY, ENCUMBRANCES, JOURNALCATEGORIES, CURRENCIES, BUDGETS, PERIODS, SEGMENTS, SEGMENT_VALUES, SEGMENT_HIERARCHY_CACHE, USERPREFERENCES, GLFORMULAS_CACHE)");

            string createSql = @"
                PRAGMA foreign_keys = ON;

                CREATE TABLE IF NOT EXISTS CUBES (
                        cubeId INTEGER PRIMARY KEY,
                        cubeName TEXT NOT NULL,
                        userName TEXT,
                        lastRefreshedDate TEXT,
                        blazeEnabled INTEGER DEFAULT 0,
                        erpType TEXT,
                        adaptiveMemoryEnabled INTEGER DEFAULT 0,
                        adaptiveMemoryTableName TEXT,
                        viewBased INTEGER DEFAULT 0
                    );

               CREATE TABLE IF NOT EXISTS LEDGERS (
                        id INTEGER PRIMARY KEY AUTOINCREMENT,
                        cubeId INTEGER NOT NULL,
                        cubeName TEXT,
                        ledgerId INTEGER NOT NULL,
                        ledgerName TEXT NOT NULL,
                        coaid INTEGER,
                        periodSetName TEXT,
                        currencyCode TEXT,
                        FOREIGN KEY (cubeId) REFERENCES CUBES(cubeId) ON DELETE CASCADE,
                        UNIQUE (cubeId, ledgerId)
                    );

               CREATE INDEX IF NOT EXISTS IX_LEDGERS_CubeId ON LEDGERS(cubeId);

               CREATE TABLE IF NOT EXISTS JOURNALSOURCES (
                        id INTEGER PRIMARY KEY AUTOINCREMENT,
                        cubeId INTEGER NOT NULL,
                        ledgerId INTEGER NOT NULL,
                        jeSourceName TEXT,
                        sourceName TEXT,
                        FOREIGN KEY (cubeId, ledgerId) REFERENCES LEDGERS(cubeId, ledgerId) ON DELETE CASCADE
                    );  

               CREATE TABLE IF NOT EXISTS ACTIVITY (
                        id INTEGER PRIMARY KEY AUTOINCREMENT,
                        cubeId INTEGER NOT NULL,
                        ledgerId INTEGER NOT NULL,
                        activityType TEXT,
                        FOREIGN KEY (cubeId, ledgerId) REFERENCES LEDGERS(cubeId, ledgerId) ON DELETE CASCADE
                    );

               CREATE TABLE IF NOT EXISTS ENCUMBRANCES (
                        id INTEGER PRIMARY KEY AUTOINCREMENT,
                        cubeId INTEGER NOT NULL,
                        ledgerId INTEGER NOT NULL,
                        encumbranceTypeId INTEGER,
                        encumbranceType TEXT,
                        FOREIGN KEY (cubeId, ledgerId) REFERENCES LEDGERS(cubeId, ledgerId) ON DELETE CASCADE
                    );

               CREATE TABLE IF NOT EXISTS JOURNALCATEGORIES (
                        id INTEGER PRIMARY KEY AUTOINCREMENT,
                        cubeId INTEGER NOT NULL,
                        ledgerId INTEGER NOT NULL,
                        jeCategoryName TEXT,
                        categoryName TEXT,
                        FOREIGN KEY (cubeId, ledgerId) REFERENCES LEDGERS(cubeId, ledgerId) ON DELETE CASCADE
                    );

              CREATE TABLE IF NOT EXISTS CURRENCIES (
                        id INTEGER PRIMARY KEY AUTOINCREMENT,
                        cubeId INTEGER NOT NULL,
                        ledgerId INTEGER NOT NULL,
                        currencyCode TEXT,
                        FOREIGN KEY (cubeId, ledgerId) REFERENCES LEDGERS(cubeId, ledgerId) ON DELETE CASCADE
                    );

              CREATE TABLE IF NOT EXISTS BUDGETS (
                        id INTEGER PRIMARY KEY AUTOINCREMENT,
                        cubeId INTEGER NOT NULL,
                        ledgerId INTEGER NOT NULL,
                        budgetName TEXT,
                        FOREIGN KEY (cubeId, ledgerId) REFERENCES LEDGERS(cubeId, ledgerId) ON DELETE CASCADE
                    );

              CREATE TABLE IF NOT EXISTS PERIODS (
                        id INTEGER PRIMARY KEY AUTOINCREMENT,
                        cubeId INTEGER NOT NULL,
                        ledgerId INTEGER NOT NULL,
                        periodName TEXT,
                        periodYear INTEGER,
                        periodNum INTEGER,
                        quarterNum INTEGER,
                        periodSetName TEXT,
                        periodType TEXT,
                        startDate INTEGER,
                        endDate INTEGER,
                        adjustmentPeriodFlag TEXT,
                        FOREIGN KEY (cubeId, ledgerId) REFERENCES LEDGERS(cubeId, ledgerId) ON DELETE CASCADE
                    );

              CREATE TABLE IF NOT EXISTS SEGMENTS (
                        id INTEGER PRIMARY KEY AUTOINCREMENT,
                        cubeId INTEGER NOT NULL,
                        ledgerId INTEGER NOT NULL,
                        coaid INTEGER,
                        segmentName TEXT NOT NULL,
                        segmentValueSetId INTEGER NOT NULL,
                        securityEnabledFlag TEXT,
                        defaultType TEXT,
                        defaultValue TEXT,
                        displaySize INTEGER,
                        segmentDelimiter TEXT,
                        applicationColumnName TEXT,
                        FOREIGN KEY (cubeId, ledgerId) REFERENCES LEDGERS(cubeId, ledgerId) ON DELETE CASCADE,
                        UNIQUE (cubeId, ledgerId, applicationColumnName)
                    );

              CREATE INDEX IF NOT EXISTS IX_SEGMENTS_CubeLedger ON SEGMENTS(cubeId, ledgerId);

              CREATE TABLE IF NOT EXISTS SEGMENT_VALUES (
                        id INTEGER PRIMARY KEY AUTOINCREMENT,
                        cubeId INTEGER NOT NULL,
                        ledgerId INTEGER NOT NULL,
                        segmentName TEXT NOT NULL,
                        segmentValue TEXT NOT NULL,
                        description TEXT,
                        summaryFlag TEXT,
                        enabledFlag TEXT,
                        segmentValueSetId INTEGER NOT NULL,
                        applicationColumnName TEXT,
                        FOREIGN KEY (cubeId, ledgerId) REFERENCES LEDGERS(cubeId, ledgerId) ON DELETE CASCADE,
                        UNIQUE (cubeId, ledgerId, segmentValueSetId, segmentValue, summaryFlag, applicationColumnName)
                    );

              CREATE INDEX IF NOT EXISTS IX_SEGMENT_VALUES_Lookup ON SEGMENT_VALUES(cubeId, ledgerId, segmentValueSetId, segmentName);

              CREATE TABLE IF NOT EXISTS SEGMENT_HIERARCHY_CACHE (
                        id INTEGER PRIMARY KEY AUTOINCREMENT,
                        cubeId INTEGER NOT NULL,
                        ledgerId INTEGER NOT NULL,
                        segmentValueSetId INTEGER NOT NULL,
                        segmentName TEXT NOT NULL,
                        segmentValue TEXT NOT NULL,
                        description TEXT,
                        parent TEXT,
                        lvl INTEGER,
                        summaryFlag TEXT,
                        enabledFlag TEXT,
                        applicationColumnName TEXT
                    );

              CREATE INDEX IF NOT EXISTS IX_HIERARCHY_CACHE_Lookup ON SEGMENT_HIERARCHY_CACHE(cubeId, ledgerId, segmentValueSetId, segmentName, parent);

              CREATE TABLE IF NOT EXISTS USERPREFERENCES (
                        PreferenceKey TEXT PRIMARY KEY,
                        PreferenceValue TEXT NOT NULL
                    );

              CREATE TABLE IF NOT EXISTS GLFORMULAS_CACHE (
                        key TEXT PRIMARY KEY,
                        value TEXT NOT NULL,
                        cached_at_utc TEXT NOT NULL
                    );";

            using var cmd = connection.CreateCommand();
            cmd.CommandText = createSql;
            cmd.ExecuteNonQuery();
            ServiceLocator.Logger?.LogDebug("SQLiteHelper.CreateTablesIfNotExist: schema creation/verification completed.");
        }

        /// <summary>
        /// Wipes all session-specific data at Excel startup (after new login).
        /// </summary>
        public void ResetSessionData()
        {
            ServiceLocator.Logger?.LogDebug("SQLiteHelper.ResetSessionData: wiping session-specific tables (new login).");
            try
            {
                EnsureInitialized();

                using var connection = GetConnection();

                var tablesToClear = new[]
                {
                    "SEGMENT_VALUES",
                    "SEGMENT_HIERARCHY_CACHE",
                    "SEGMENTS",
                    "PERIODS",
                    "BUDGETS",
                    "CURRENCIES",
                    "JOURNALCATEGORIES",
                    "JOURNALSOURCES",
                    "ENCUMBRANCES",
                    "ACTIVITY",
                    "LEDGERS",
                    "CUBES"
                };

                foreach (var table in tablesToClear)
                {
                    using var cmd = connection.CreateCommand();
                    cmd.CommandText = $"DELETE FROM [{table}];";
                    ServiceLocator.Logger?.LogDebug($"SQLiteHelper.ResetSessionData: executing '{cmd.CommandText}'");
                    int rowsAffected = cmd.ExecuteNonQuery();
                    ServiceLocator.Logger?.LogDebug($"SQLiteHelper.ResetSessionData: cleared {rowsAffected} rows from {table}");
                }

                // LogInfo is not used in this codebase (per project convention) - LogDebug is
                // the safe/cheap equivalent, gated by the ribbon's Debug checkbox.
                ServiceLocator.Logger?.LogDebug("SQLiteHelper.ResetSessionData: session data reset successfully");
            }
            catch (Exception ex)
            {
                ServiceLocator.Logger?.LogException(ex, "SQLiteHelper.ResetSessionData - Failed to reset session data");
                throw;
            }
        }

        /// <summary>
        /// Checks if the database is initialized
        /// </summary>
        public bool IsInitialized => _initialized;

        /// <summary>
        /// Gets the database file path
        /// </summary>
        public string DatabasePath => _dbPath;

        /// <summary>
        /// Re-initializes the database (use with caution)
        /// </summary>
        public void Reinitialize()
        {
            ServiceLocator.Logger?.LogDebug("SQLiteHelper.Reinitialize: forcing re-initialization of the SQLite database.");
            lock (_initLock)
            {
                _initialized = false;
                InitializeDatabase();
            }
        }
    }
}
