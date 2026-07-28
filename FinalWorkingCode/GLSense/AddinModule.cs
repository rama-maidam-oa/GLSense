using AddinExpress.MSO;
using GLSense.Caching;
using GLSense.Drilldowns;
using GLSense.Helpers;
using GLSense.Models;
using GLSense.Repositories;
using GLSense.Utilities;
using GLSense.Views;
using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SQLite;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Security.Authentication;
using System.Text;
using System.Text.Json;
using System.Text.Json.Nodes;
using System.Threading;
using System.Threading.Tasks;
using System.Web.UI.WebControls;
using System.Windows;
using System.Windows.Forms;
using System.Windows.Threading;
using System.Xml;
using System.Xml.Linq;
using Excel = Microsoft.Office.Interop.Excel;

namespace GLSense
{
    /// <summary>
    ///   Add-in Express Add-in Module
    /// </summary>
    [GuidAttribute("B9FB44F7-6849-4A00-9197-1F8B0D633172"), ProgId("GLSense.AddinModule")]
    public partial class AddinModule : AddinExpress.MSO.ADXAddinModule
    {
        public static RibbonStateHelper RibbonHelper { get; private set; }
        private static RibbonStateHelper _ribbonHelper;
        private bool _isFinalized;
        private static SQLiteConnection _dbConnection;
        // Logging
        public static NLog.Config.LoggingConfiguration LoggerConfiguration { get; set; }
        public static NLog.Logger Logger { get; set; }

        // Flag to enable/disable assembly binding diagnostics (set to false in production)
        private static readonly bool EnableAssemblyDiagnostics = false;

        public AddinModule()
        {
            System.Windows.Forms.Application.EnableVisualStyles();

            InitializeComponent();

            // Hook into assembly binding events for diagnostics (before any other initialization)
            if (EnableAssemblyDiagnostics)
            {
                SetupAssemblyBindingDiagnostics();
            }

            // Please add any initialization code to the AddinInitialize event handler
            this.AddinInitialize += AddinModule_AddinInitialize;
            this.OnError += AddinModule_OnError;
            this.OnRibbonLoaded += AddinModule_OnRibbonLoaded;
            this.AddinBeginShutdown += AddinModule_AddinBeginShutdown;
            this.AddinFinalize += AddinModule_AddinFinalize;
        }

        #region Add-in Express automatic code

        // Required by Add-in Express - do not modify
        // the methods within this region

        public override System.ComponentModel.IContainer GetContainer()
        {
            if (components == null)
                components = new System.ComponentModel.Container();
            return components;
        }

        [ComRegisterFunctionAttribute]
        public static void AddinRegister(Type t)
        {
            AddinExpress.MSO.ADXAddinModule.ADXRegister(t);
        }

        [ComUnregisterFunctionAttribute]
        public static void AddinUnregister(Type t)
        {
            AddinExpress.MSO.ADXAddinModule.ADXUnregister(t);
        }

        public override void UninstallControls()
        {
            base.UninstallControls();
        }

        #endregion

        public static new AddinModule CurrentInstance
        {
            get
            {
                return AddinExpress.MSO.ADXAddinModule.CurrentInstance as AddinModule;
            }
        }

        public object EdgeAddinInstance { get; set; }

        public static object GetEdgeAddinInstance()
        {
            var currentInstance = CurrentInstance;
            if (currentInstance?.EdgeAddinInstance != null)
                return currentInstance.EdgeAddinInstance;

            var hostApplication = currentInstance?.HostApplication ?? AppState.Instance.ExcelApp;
            if (hostApplication == null)
                return null;

            try
            {
                var comAddIns = hostApplication.GetType().InvokeMember("COMAddIns", BindingFlags.GetProperty, null, hostApplication, Array.Empty<object>());
                if (comAddIns is System.Collections.IEnumerable addIns)
                {
                    foreach (var addIn in addIns)
                    {
                        if (addIn == null)
                            continue;

                        var addInType = addIn.GetType();
                        var progId = addInType.InvokeMember("ProgId", BindingFlags.GetProperty, null, addIn, Array.Empty<object>()) as string;
                        var description = addInType.InvokeMember("Description", BindingFlags.GetProperty, null, addIn, Array.Empty<object>()) as string;

                        if (!IsXlEdgeAddin(progId) && !IsXlEdgeAddin(description))
                            continue;

                        var addinObject = addInType.InvokeMember("Object", BindingFlags.GetProperty, null, addIn, Array.Empty<object>());
                        if (currentInstance != null)
                            currentInstance.EdgeAddinInstance = addinObject;

                        return addinObject;
                    }
                }
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex, "Failed to locate XLEdge COM add-in.");
            }

            return null;
        }

        private static bool IsXlEdgeAddin(string value)
        {
            return !string.IsNullOrWhiteSpace(value) && value.IndexOf("XLEdge", StringComparison.OrdinalIgnoreCase) >= 0;
        }

        public Excel._Application ExcelApp
        {
            get
            {
                return (HostApplication as Excel._Application);
            }
        }

        #region Assembly Binding Diagnostics

        /// <summary>
        /// Sets up event handlers to capture and log assembly binding information.
        /// This helps diagnose assembly loading issues in production environments.
        /// </summary>
        private static void SetupAssemblyBindingDiagnostics()
        {
            try
            {
                // Hook into first-chance exceptions to capture assembly binding issues
                AppDomain.CurrentDomain.FirstChanceException += CurrentDomain_FirstChanceException;

                // Hook into assembly resolve events (called when an assembly fails to load)
                AppDomain.CurrentDomain.AssemblyResolve += CurrentDomain_AssemblyResolve;

                // Hook into assembly load events (for logging successful loads)
                AppDomain.CurrentDomain.AssemblyLoad += CurrentDomain_AssemblyLoad;

                // Log initial diagnostic info
                LogAssemblyDiagnosticInfo();
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Failed to setup assembly diagnostics: {ex.Message}");
                LogUtility.LogException(ex, "SetupAssemblyBindingDiagnostics");
            }
        }

        /// <summary>
        /// Logs assembly diagnostic messages directly to the logger, bypassing the ribbon DebugLogs toggle.
        /// </summary>
        private static void LogAssemblyDiag(string level, string message)
        {
            if (!EnableAssemblyDiagnostics) return;

            var logMessage = $"{level} | {DateTime.Now:HH:mm:ss} | {message}";
            var logger = Logger;

            if (logger == null) return;

            var logActions = new Dictionary<string, System.Action>
            {
                ["DEBUG"] = () => logger.Debug(logMessage),
                ["INFO"] = () => logger.Info(logMessage),
                ["WARN"] = () => logger.Warn(logMessage),
                ["ERROR"] = () => logger.Error(logMessage)
            };

            if (logActions.TryGetValue(level, out var logAction))
            {
                logAction();
            }
        }

        private static void CurrentDomain_FirstChanceException(object sender, System.Runtime.ExceptionServices.FirstChanceExceptionEventArgs e)
        {
            if (!EnableAssemblyDiagnostics) return;

            if (IsExpectedInternalException(e.Exception))
            {
                return;
            }

            LogFirstChanceException(e.Exception);
        }

        private static void LogFirstChanceException(Exception exception)
        {
            switch (exception)
            {
                case FileNotFoundException fnfEx:
                    LogFileNotFoundException(fnfEx);
                    break;
                case FileLoadException flEx:
                    LogFileLoadException(flEx);
                    break;
                case BadImageFormatException bifEx:
                    LogBadImageFormatException(bifEx);
                    break;
                case ReflectionTypeLoadException rtlEx:
                    LogReflectionTypeLoadException(rtlEx);
                    break;
                case COMException comEx:
                    LogAssemblyDiag("DEBUG", $"[AssemblyDiag] COMException: HRESULT=0x{comEx.ErrorCode:X8} - {comEx.Message}");
                    break;
                case TargetInvocationException tieEx:
                    LogAssemblyDiag("DEBUG", $"[AssemblyDiag] TargetInvocationException: {tieEx.InnerException?.Message ?? tieEx.Message}");
                    break;
            }
        }

        private static void LogFileNotFoundException(FileNotFoundException fnfEx)
        {
            if (fnfEx.FileName?.Contains(".resources") == true)
            {
                LogAssemblyDiag("DEBUG", $"[AssemblyDiag] Resource probe (expected): {fnfEx.FileName}");
                return;
            }

            LogAssemblyDiag("WARN", $"[AssemblyDiag] FileNotFoundException: {fnfEx.FileName} - {fnfEx.Message}");
            LogFusionLogIfPresent(fnfEx.FusionLog);
        }

        private static void LogFileLoadException(FileLoadException flEx)
        {
            LogAssemblyDiag("WARN", $"[AssemblyDiag] FileLoadException: {flEx.FileName} - {flEx.Message}");
            LogFusionLogIfPresent(flEx.FusionLog);
        }

        private static void LogBadImageFormatException(BadImageFormatException bifEx)
        {
            LogAssemblyDiag("ERROR", $"[AssemblyDiag] BadImageFormatException: {bifEx.FileName} - {bifEx.Message} (possible x86/x64 mismatch)");
            LogFusionLogIfPresent(bifEx.FusionLog);
        }

        private static void LogReflectionTypeLoadException(ReflectionTypeLoadException rtlEx)
        {
            LogAssemblyDiag("ERROR", $"[AssemblyDiag] ReflectionTypeLoadException: {rtlEx.Message}");
            foreach (var loaderEx in rtlEx.LoaderExceptions ?? Array.Empty<Exception>())
            {
                if (loaderEx != null)
                    LogAssemblyDiag("ERROR", $"[AssemblyDiag] Loader Exception: {loaderEx.Message}");
            }
        }

        private static void LogFusionLogIfPresent(string fusionLog)
        {
            if (!string.IsNullOrEmpty(fusionLog))
            {
                LogAssemblyDiag("DEBUG", $"[AssemblyDiag] Fusion Log:\n{fusionLog}");
            }
        }

        /// <summary>
        /// Determines if an exception is an expected internal framework exception that should not be logged.
        /// These are typically caught and handled internally by .NET or third-party libraries.
        /// </summary>
        private static bool IsExpectedInternalException(Exception ex)
        {
            return IsExpectedIOException(ex) ||
                   IsExpectedTargetInvocationException(ex) ||
                   IsExpectedCOMException(ex);
        }

        private static bool IsExpectedIOException(Exception ex)
        {
            if (ex is not IOException ioEx) return false;

            var expectedMessages = new[]
            {
                "being used by another process",
                "cannot find the file",
                "Could not load file"
            };

            if (expectedMessages.Any(msg => ioEx.Message.Contains(msg)))
            {
                LogAssemblyDiag("DEBUG", $"[AssemblyDiag] Expected IOException (assembly probing): {ioEx.Message}");
                return true;
            }

            return false;
        }

        private static bool IsExpectedTargetInvocationException(Exception ex)
        {
            if (ex is TargetInvocationException tie && tie.InnerException is IOException)
            {
                LogAssemblyDiag("DEBUG", $"[AssemblyDiag] Expected TargetInvocationException (reflection/IO): {tie.InnerException.Message}");
                return true;
            }
            return false;
        }

        private static bool IsExpectedCOMException(Exception ex)
        {
            if (ex is not COMException comEx) return false;

            // Common expected Excel COM errors
            var expectedHResults = new[]
            {
                unchecked((int)0x800A03EC), // Command cannot be used on multiple selections
                unchecked((int)0x80010001), // Call was rejected by callee (Excel busy)
                unchecked((int)0x8001010A)  // Message filter indicated that the application is busy
            };

            if (expectedHResults.Contains(comEx.ErrorCode))
            {
                LogAssemblyDiag("DEBUG", $"[AssemblyDiag] Expected COMException (Excel busy/selection): HRESULT=0x{comEx.ErrorCode:X8}");
                return true;
            }

            return false;
        }

        private static Assembly CurrentDomain_AssemblyResolve(object sender, ResolveEventArgs args)
        {
            if (!EnableAssemblyDiagnostics) return null;

            // Parse the assembly name to check if it's a resource assembly
            var assemblyName = new AssemblyName(args.Name);

            // Resource assemblies (satellite assemblies) are expected to fail for non-existent cultures
            // This is normal .NET behavior - it probes for localized resources before falling back to defaults
            if (assemblyName.Name.EndsWith(".resources", StringComparison.OrdinalIgnoreCase))
            {
                LogAssemblyDiag("DEBUG", $"[AssemblyDiag] Resource assembly probe (expected): {assemblyName.Name}, Culture={assemblyName.CultureName}");
                return null; // Let the default resolver handle it (will fall back to neutral resources)
            }

            // Log the assembly resolve attempt - this fires when assembly binding fails
            LogAssemblyDiag("WARN", $"[AssemblyDiag] AssemblyResolve triggered for: {args.Name}");
            LogAssemblyDiag("DEBUG", $"[AssemblyDiag] Requesting assembly: {args.RequestingAssembly?.FullName ?? "Unknown"}");

            // Return null to let the default resolver continue (don't interfere with normal resolution)
            return null;
        }

        private static void CurrentDomain_AssemblyLoad(object sender, AssemblyLoadEventArgs args)
        {
            if (!EnableAssemblyDiagnostics) return;

            // Log assembly loads for diagnostic purposes
            LogAssemblyDiag("DEBUG", $"[AssemblyDiag] Assembly loaded: {args.LoadedAssembly.FullName}");
            LogAssemblyDiag("DEBUG", $"[AssemblyDiag] Location: {(args.LoadedAssembly.IsDynamic ? "Dynamic" : args.LoadedAssembly.Location)}");
        }

        private static void LogAssemblyDiagnosticInfo()
        {
            LogAssemblyDiag("INFO", "[AssemblyDiag] === Assembly Binding Diagnostics Enabled ===");
            LogAssemblyDiag("INFO", $"[AssemblyDiag] AppDomain: {AppDomain.CurrentDomain.FriendlyName}");
            LogAssemblyDiag("INFO", $"[AssemblyDiag] BaseDirectory: {AppDomain.CurrentDomain.BaseDirectory}");
            LogAssemblyDiag("INFO", $"[AssemblyDiag] RelativeSearchPath: {AppDomain.CurrentDomain.RelativeSearchPath ?? "(none)"}");
            LogAssemblyDiag("INFO", $"[AssemblyDiag] CLR Version: {Environment.Version}");
            LogAssemblyDiag("INFO", $"[AssemblyDiag] Process: {(Environment.Is64BitProcess ? "64-bit" : "32-bit")}");
        }

        #endregion

        // Centralized helpers to reduce duplication and improve error handling
        private static void SafeInvokeWpf(System.Action action)
        {
            if (action == null) return;
            try
            {
                WpfAppManager.InvokeOnWpfThread(() =>
                {
                    try
                    {
                        action();
                    }
                    catch (Exception ex)
                    {
                        LogUtility.LogException(ex);
                    }
                });
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex);
            }
        }

        private static void SafeLogException(Exception ex, string context = null)
        {
            if (ex == null) return;
            LogUtility.LogException(ex, context);
        }
        private void AddinModule_AddinBeginShutdown(object sender, EventArgs e)
        {
            try
            {
                // Unsubscribe from all Excel events
                UnsubscribeFromAllExcelEvents();

                // Force hide task panes
                HideTaskPanes();

                // ADXTaskPane mouse-wheel fix (CLAUDE.md section 24.3.5): tears down the
                // dedicated background thread + WH_MOUSE_LL hook so neither lingers if the
                // add-in is disabled/unloaded without Excel itself closing.
                try
                {
                    GLSense.Controls.SuggestAppendComboBox.ShutdownMouseHook();
                }
                catch (Exception ex)
                {
                    ShutdownLogger.LogError("Error in SuggestAppendComboBox.ShutdownMouseHook", ex);
                }

                // Save to database and cleanup (with error handling only)
                try
                {
                    FormulaCacheManager.Instance.Dispose();
                }
                catch (Exception ex)
                {
                    ShutdownLogger.LogError("Error disposing FormulaCacheManager", ex);
                }

                // Close database connection
                try
                {
                    _dbConnection?.Close();
                    _dbConnection?.Dispose();
                }
                catch (Exception ex)
                {
                    ShutdownLogger.LogError("Error closing database connection", ex);
                }

                // Release COM objects
                ReleaseAllComObjectsProperly();
            }
            catch (Exception ex)
            {
                ShutdownLogger.LogError("Critical error in AddinBeginShutdown", ex);
                LogUtility.LogException(ex);
            }
        }

        private void UnsubscribeFromAllExcelEvents()
        {
            try
            {
                if (adxExcelAppEvents1 != null)
                {
                    adxExcelAppEvents1.SheetActivate -= adxExcelAppEvents1_SheetActivate;
                    adxExcelAppEvents1.SheetBeforeDoubleClick -= adxExcelAppEvents1_SheetBeforeDoubleClick;
                    adxExcelAppEvents1.SheetChange -= adxExcelAppEvents1_SheetChange;
                    adxExcelAppEvents1.SheetFollowHyperlink -= adxExcelAppEvents1_SheetFollowHyperlink;
                    adxExcelAppEvents1.SheetSelectionChange -= adxExcelAppEvents1_SheetSelectionChange;
                    adxExcelAppEvents1.WorkbookBeforeSave -= adxExcelAppEvents1_WorkbookBeforeSave;
                    adxExcelAppEvents1.WorkbookActivate -= adxExcelAppEvents1_WorkbookActivate;
                }
            }
            catch (Exception ex)
            {
                ShutdownLogger.LogError("Error unsubscribing from Excel events", ex);
            }
        }

        private void ReleaseAllComObjectsProperly()
        {
            try
            {
                var excelApp = AppState.Instance.ExcelApp;
                if (excelApp != null)
                {
                    // Save original state
                    bool originalEventsState = excelApp.EnableEvents;

                    try
                    {
                        // Disable events during cleanup
                        excelApp.EnableEvents = false;

                        // Release all workbooks
                        if (excelApp.Workbooks != null)
                        {
                            while (excelApp.Workbooks.Count > 0)
                            {
                                Workbook wb = excelApp.Workbooks[1];
                                wb.Close(false);
                                Marshal.FinalReleaseComObject(wb);
                            }
                            Marshal.FinalReleaseComObject(excelApp.Workbooks);
                        }
                    }
                    finally
                    {
                        // RESTORE EVENTS BEFORE releasing the Excel app
                        // This is CRITICAL - but must be done while the COM object is still valid
                        try
                        {
                            excelApp.EnableEvents = originalEventsState;
                        }
                        catch
                        {
                            // If this fails, the COM object might be in a bad state
                            // But we're about to release it anyway
                            ShutdownLogger.LogWarn("Failed to restore Excel events state during shutdown cleanup.");
                        }
                    }

                    // NOW release the Excel application object
                    Marshal.FinalReleaseComObject(excelApp);
                }

                AppState.Instance.ExcelApp = null;
                GC.Collect();
                GC.WaitForPendingFinalizers();
                GC.Collect();
            }
            catch (Exception ex)
            {
                ShutdownLogger.LogError("Error releasing COM objects", ex);
            }
        }

        private void AddinModule_AddinFinalize(object sender, EventArgs e)
        {
            if (_isFinalized)
            {
                return;
            }
            _isFinalized = true;

            try
            {
                // Restore any Excel settings
                ExcelErrorCheckingService.Restore();

                // Dispose ADX events
                if (adxExcelAppEvents1 != null)
                {
                    adxExcelAppEvents1.Dispose();
                    adxExcelAppEvents1 = null;
                }
            }
            catch (Exception ex)
            {
                ShutdownLogger.LogError("Error in AddinFinalize", ex);
                LogUtility.LogException(ex);
            }
        }


        private static void AddinModule_OnRibbonLoaded(object sender, IRibbonUI Ribbon)
        {
            AppState.Instance.ExcelApp = (Excel.Application)AddinModule.CurrentInstance.HostApplication;
            LogHelper.InitializeLogger();

            _ribbonHelper = new RibbonStateHelper(AddinModule.CurrentInstance, Ribbon);
            RibbonHelper = _ribbonHelper; // Expose it globally

            if (!AppState.Instance.IsLoggedIn && !AppState.Instance.IsLoginCompleted)
            {
                _ribbonHelper.ApplyState("LoggedOut");
            }
            else if (AppState.Instance.IsLoggedIn && !AppState.Instance.IsLoginCompleted)
            {
                _ribbonHelper.ApplyState("PartialLoggedIn");
            }
            else
            {
                _ribbonHelper.ApplyState("LoggedIn");
            }

            AddinModule.CurrentInstance.SyncRibbonSelectionWithAppState();
            MahAppsBootstrapper.Init(AppConstants.GLAccentHex, AppConstants.GLTheme);
            MahAppsBootstrapper.PreloadResources();
        }

        private static void AddinModule_AddinInitialize(object sender, EventArgs e)
        {
            try
            {

                // 1. Ensure DB file + tables exist
                SQLiteHelper.InitializeDatabase();
                // 2. Only wipe session data when the database has no cached content yet
                if (!SQLiteHelper.HasPersistedData())
                {
                    SQLiteHelper.ResetSessionData();
                }

                _dbConnection = new SQLiteConnection($"Data Source={AppPaths.DatabasePath};Version=3;");
                _dbConnection.Open();

                // Initialize cache (creates table if not exists, loads data)
                FormulaCacheManager.Instance.Initialize(_dbConnection);

                ExcelErrorCheckingService.Apply(AppState.Instance.ExcelApp);
            }
            catch (Exception ex)
            {
                SafeLogException(ex, "Addin initialization failed");
            }
        }
        private static void AddinModule_OnError(AddinExpress.MSO.ADXErrorEventArgs e)
        {
            e.Handled = true;
            CommonFunctions.GLSenseMessage("Error: " + e.ADXError.ToString(), MessageBoxIcon.Error, MessageBoxButtons.OK);
        }
        private void adxExcelAppEvents1_WorkbookBeforeSave(object sender, ADXHostBeforeSaveEventArgs e)
        {
            try
            {
                LogUtility.LogDebug($"WorkbookBeforeSave fired. CacheInitialized={FormulaCacheManager.Instance.IsInitialized}, HasChanges={FormulaCacheManager.Instance.HasChanges}");

                // Save the formula cache to database when user saves the workbook
                if (FormulaCacheManager.Instance.IsInitialized && FormulaCacheManager.Instance.HasChanges)
                {
                    FormulaCacheManager.Instance.PersistToDatabase();
                }
            }
            catch (Exception ex)
            {
                ShutdownLogger.LogError("Error saving cache during WorkbookBeforeSave", ex);
                // Don't cancel the save - just log the error
            }
        }
        private void adxExcelAppEvents1_SheetActivate(object sender, object hostObj)
        {
            if (ExcelApp == null || !AppState.Instance.IsLoginCompleted) return;

            try
            {
                LogUtility.LogDebug($"SheetActivate fired. Sheet={(hostObj as Excel.Worksheet)?.Name ?? "<unknown>"}");
                _ribbonHelper.ApplyState("ApplySheetActiveState");
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex);
            }
        }

        private void adxExcelAppEvents1_SheetBeforeDoubleClick(object sender, ADXExcelSheetBeforeEventArgs e)
        {
            try
            {
                if ((AppState.Instance.ExcelApp?.Hwnd ?? 0) == 0 || !AppState.Instance.IsLoginCompleted)
                    return;

                var currSheet = (Excel.Worksheet)e.Sheet;
                var currRange = (Excel.Range)e.Range;

                LogUtility.LogDebug($"SheetBeforeDoubleClick fired. Sheet={currSheet?.Name}, Range={currRange?.Address}");

                bool actionTaken = false;
                DrillDownsStart(currRange, currSheet, ref actionTaken);
                e.Cancel = actionTaken;
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex);
            }
        }

        public static void DrillDownsStart(Excel.Range cellRange, Excel.Worksheet currentSheet, ref bool noAction)
        {
            try
            {
                if (HasNoBalanceFormula(cellRange))
                {
                    RunBalancePrecedentDrilldown(cellRange);
                    noAction = true;
                    return;
                }

                string ddType = ResolveDrilldownType(cellRange, currentSheet);
                if (string.IsNullOrWhiteSpace(ddType))
                    return;

                RunDrilldown(cellRange, ddType);
                noAction = true;
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex);
            }
            finally
            {
                try
                {
                    cellRange?.Worksheet?.ClearArrows();
                }
                catch (Exception ex)
                {
                    LogUtility.LogException(ex);
                }
            }
        }

        private static bool TryGetSingleCellFormula(Excel.Range cellRange, out string formulaString)
        {
            formulaString = string.Empty;

            try
            {
                if (cellRange == null)
                    return false;

                if (cellRange.Areas.Count != 1 || cellRange.Rows.Count != 1 || cellRange.Columns.Count != 1)
                    return false;

                if (!(cellRange.HasFormula is bool hf) || !hf)
                    return false;

                formulaString = cellRange.Formula?.ToString() ?? string.Empty;
                return true;
            }
            catch (Exception ex)
            {
                LogUtility.LogDebug($"TryGetSingleCellFormula: could not read formula ({ex.Message}).");
                return false;
            }
        }

        private static bool HasNoBalanceFormula(Excel.Range cellRange)
        {
            return TryGetSingleCellFormula(cellRange, out string formulaString) &&
                   formulaString.IndexOf(AppConstants.glBal, StringComparison.OrdinalIgnoreCase) < 0;
        }

        private static bool HasBalanceFormula(Excel.Range cellRange)
        {
            return TryGetSingleCellFormula(cellRange, out string formulaString) &&
                   formulaString.IndexOf(AppConstants.glBal, StringComparison.OrdinalIgnoreCase) >= 0;
        }

        private static void RunBalancePrecedentDrilldown(Excel.Range cellRange)
        {
            try
            {
                string external = ExcelExternalRef.BuildExternalAddress(cellRange);
                var runProcess = new DrilldownXlPrecedents(AppState.Instance.ExcelApp, external);
                _ = runProcess.ProcessEPDrilldown();
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex);
            }
        }

        private static string ResolveDrilldownType(Excel.Range cellRange, Excel.Worksheet currentSheet)
        {
            if (HasBalanceFormula(cellRange))
                return "BL";

            string a1Text = GetA1Text(currentSheet);
            if (string.IsNullOrWhiteSpace(a1Text))
                return string.Empty;

            string sheetName = currentSheet.Name;
            string markerSheetName = GetDrilldownSheetMarkerValue(currentSheet);

            if (IsBalancesDrilldown(a1Text, sheetName, markerSheetName))
                return "JL";

            if (IsJournalsDrilldown(a1Text, sheetName, markerSheetName))
                return "SL";

            return string.Empty;
        }

        private static string GetA1Text(Excel.Worksheet currentSheet)
        {
            var a1Val = currentSheet?.Range["A1"]?.Value2;
            return a1Val?.ToString().Trim().ToUpperInvariant() ?? string.Empty;
        }

        private static string GetDrilldownSheetMarkerValue(Excel.Worksheet sheet)
        {
            try
            {
                var markerCell = (Excel.Range)sheet.Range[AppConstants.DrilldownSheetMarkerCellAddress];
                var value = markerCell?.Value2 ?? markerCell?.Value;
                return value?.ToString() ?? string.Empty;
            }
            catch (Exception ex)
            {
                LogUtility.LogDebug($"GetDrilldownSheetMarkerValue: could not read marker cell ({ex.Message}).");
                return string.Empty;
            }
        }

        private static bool IsBalancesDrilldown(string a1Text, string sheetName, string markerSheetName)
        {
            if (a1Text.Equals("BALANCES DRILLDOWN", StringComparison.OrdinalIgnoreCase))
                return true;

            return MatchesBalancesPattern(sheetName) || MatchesBalancesPattern(markerSheetName);
        }

        private static bool IsJournalsDrilldown(string a1Text, string sheetName, string markerSheetName)
        {
            if (a1Text.Equals("JOURNALS DRILLDOWN", StringComparison.OrdinalIgnoreCase))
                return true;

            return MatchesJournalsPattern(sheetName) || MatchesJournalsPattern(markerSheetName);
        }

        private static bool MatchesBalancesPattern(string sheetName)
        {
            return !string.IsNullOrWhiteSpace(sheetName) &&
                   sheetName.IndexOf("_BL_", StringComparison.OrdinalIgnoreCase) >= 0 &&
                   sheetName.IndexOf("_JL_", StringComparison.OrdinalIgnoreCase) < 0 &&
                   sheetName.IndexOf("_SL_", StringComparison.OrdinalIgnoreCase) < 0 &&
                   sheetName.IndexOf("_CM_", StringComparison.OrdinalIgnoreCase) >= 0;
        }

        private static bool MatchesJournalsPattern(string sheetName)
        {
            return !string.IsNullOrWhiteSpace(sheetName) &&
                   sheetName.IndexOf("_BL_", StringComparison.OrdinalIgnoreCase) >= 0 &&
                   sheetName.IndexOf("_JL_", StringComparison.OrdinalIgnoreCase) >= 0 &&
                   sheetName.IndexOf("_SL_", StringComparison.OrdinalIgnoreCase) < 0;
        }

        private static void RunDrilldown(Excel.Range cellRange, string ddType)
        {
            try
            {
                string external = ExcelExternalRef.BuildExternalAddress(cellRange);

                switch (ddType)
                {
                    case "BL":
                        {
                            var runProcess = new DrilldownBl(AppState.Instance.ExcelApp, external, "BL");
                            _ = runProcess.ProcessBLDrilldown();
                            break;
                        }

                    case "JL":
                        {
                            var runProcess = new DrilldownJl(AppState.Instance.ExcelApp, external, "JL");
                            _ = runProcess.ProcessJLDrilldown();
                            break;
                        }

                    case "SL":
                        {
                            var runProcess = new DrilldownSl(AppState.Instance.ExcelApp, external);
                            _ = runProcess.ProcessSLDrilldown();
                            break;
                        }

                    default:
                        // Unknown type: optionally log or ignore
                        LogUtility.LogDebug($"Unknown drilldown type requested: {ddType}");
                        break;
                }
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex);
            }
        }

        private async void adxExcelAppEvents1_SheetFollowHyperlink(object sender, object sheet, object hyperlink)
        {
            if (!GuardLoginAndExcel())
                return;

            using var ctsHelper = new CancellationHelper();
            CancellationToken token = ctsHelper.GetToken();

            CommonMethods.DisableExcelSettings();

            try
            {
                Excel.Worksheet sht = sheet as Excel.Worksheet;
                LogUtility.LogDebug($"SheetFollowHyperlink fired. Sheet={sht?.Name}");
                if (!IsValidDrilldownSheet(sht))
                {
                    return;
                }

                Excel.Hyperlink hyprLink = hyperlink as Excel.Hyperlink;

                if (IsCustomDrilldownHyperlink(hyprLink))
                {
                    await HandleCustomDrilldownHyperlink(sht, hyprLink, token);
                }
                else if (hyprLink?.Parent is Excel.Range hyperlinkRange)
                {
                    await HandleJournalAttachmentHyperlink(hyperlinkRange, token);
                }
            }
            catch (OperationCanceledException)
            {
                LogUtility.LogWarn("Sheet Follow Hyperlink operation was canceled by the user.");
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex);
            }
            finally
            {
                CommonMethods.EnableExcelSettings();
            }
        }

        private static bool IsValidDrilldownSheet(Excel.Worksheet sht)
        {
            return sht?.ListObjects.Count > 0 && sht.ListObjects[1].Name.StartsWith("ORB_DD_");
        }

        private static bool IsCustomDrilldownHyperlink(Excel.Hyperlink hyprLink)
        {
            return !string.IsNullOrEmpty(hyprLink?.ScreenTip) &&
                   hyprLink.ScreenTip.IndexOf("CUSTOM DRILLDOWN", StringComparison.OrdinalIgnoreCase) >= 0;
        }

        private async Task HandleCustomDrilldownHyperlink(Excel.Worksheet sht, Excel.Hyperlink hyprLink, CancellationToken token)
        {
            Excel.ListObject tableObj = sht.ListObjects[1];
            Excel.Range rng = sht.Range[hyprLink.SubAddress];

            token.ThrowIfCancellationRequested();

            try
            {
                Excel.Range rngNew = sht.Cells[5, rng.Column] as Excel.Range;
                object cellValue = rngNew.Value2;
                if (cellValue != null)
                {
                    string headerLabel = cellValue.ToString() ?? "";
                    if (!string.IsNullOrWhiteSpace(headerLabel))
                    {
                        await GLSenseCustomDrillDown(tableObj, rng, headerLabel);
                    }
                    else
                    {
                        LogUtility.LogWarn("Exiting the hyperlink sub since the label header is null or empty");
                    }
                }
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex, "Unable to get the header label.");
            }
        }

        private async Task HandleJournalAttachmentHyperlink(Excel.Range hyperlinkRange, CancellationToken token)
        {
            long journalHeaderId = (long)Math.Truncate(Convert.ToDouble(hyperlinkRange.Value2, CultureInfo.InvariantCulture));

            var jsonObj = new JournalAttachments
            {
                cubeId = AppState.Instance.SelectedCube.CubeId,
                journalHeaderId = journalHeaderId
            };

            token.ThrowIfCancellationRequested();

            string apiUrl = $"{AppState.Instance.LoginUrl}{AppConstants.RestSecure}journal-attachment-files";
            string httpPostText = JsonSerializer.Serialize(jsonObj, JsonGlobals.Options);

            string responseData = await ApiHelper.ServerAPI(apiUrl, "JSON", httpPostText, "POST", token);

            token.ThrowIfCancellationRequested();

            var attachmentResult = ProcessAttachmentResponse(responseData);
            if (!attachmentResult.success) return;

            PopulateJournalDictionary(attachmentResult.records, token);

            if (AppState.Instance.JournalDictionary.Count > 0)
            {
                ShowAttachmentsDialog();
            }

            await DownloadSelectedAttachments(token);
        }

        private static (bool success, List<JournalAttachmentRecord> records) ProcessAttachmentResponse(string responseData)
        {
            if (string.IsNullOrWhiteSpace(responseData))
            {
                CommonFunctions.GLSenseMessage("Empty response from server.", MessageBoxIcon.Error);
                return (false, null);
            }

            var result = ApiResponseHelper.Parse<List<JournalAttachmentRecord>>(responseData, JsonGlobals.Options);
            if (!result.IsSuccess)
            {
                CommonFunctions.GLSenseMessage(result.ErrorMessage, MessageBoxIcon.Error);
                return (false, null);
            }

            var records = result.Value;
            if (records == null || records.Count == 0)
            {
                CommonFunctions.GLSenseMessage("No attachment records found.", MessageBoxIcon.Information);
                return (false, null);
            }

            return (true, records);
        }

        private static void PopulateJournalDictionary(List<JournalAttachmentRecord> records, CancellationToken token)
        {
            AppState.Instance.AttachIDs = string.Empty;
            AppState.Instance.JournalDictionary.Clear();

            foreach (var record in records)
            {
                token.ThrowIfCancellationRequested();

                if (string.IsNullOrWhiteSpace(record.FILE_ID))
                    continue;

                if (!AppState.Instance.JournalDictionary.ContainsKey(record.FILE_ID))
                {
                    AppState.Instance.JournalDictionary.Add(record.FILE_ID, record.FILE_NAME ?? string.Empty);
                }
            }
        }

        private static void ShowAttachmentsDialog()
        {
            SafeInvokeWpf(() =>
            {
                var win = new AttachmentsDialog();
                win.ShowDialogWithOwner((IntPtr)AppState.Instance.ExcelApp.Hwnd);
            });
        }

        private async Task DownloadSelectedAttachments(CancellationToken token)
        {
            if (string.IsNullOrWhiteSpace(AppState.Instance.AttachIDs))
                return;

            var jrAttachIDs = AppState.Instance.AttachIDs
                .Split(new[] { ',' }, StringSplitOptions.RemoveEmptyEntries)
                .Select(id => id.Trim())
                .Where(id => long.TryParse(id, out _))
                .Select(long.Parse)
                .ToList();

            if (jrAttachIDs.Count == 0)
                return;

            var downloadRequest = new JrnalAttachRequest
            {
                cubeId = AppState.Instance.SelectedCube.CubeId,
                fileIds = jrAttachIDs.ToArray()
            };

            string downloadUrl = $"{AppState.Instance.LoginUrl}{AppConstants.RestSecure}journal-attachments";
            string downloadPayload = JsonSerializer.Serialize(downloadRequest, JsonGlobals.Options);

            await GLSense_DownloadFile(downloadUrl, downloadPayload, token);
            token.ThrowIfCancellationRequested();
        }

        private async Task GLSenseCustomDrillDown(Excel.ListObject tableObj, Excel.Range rng, string headerLabel)
        {
            GLWaitWindow win = null;
            using var ctsHelper = new CancellationHelper();
            CancellationToken token = ctsHelper.GetToken();

            try
            {
                win = CreateAndShowWaitWindow(ctsHelper);
                await InitializeWaitWindowAsync(win, "Custom Drilldown", "GLSense custom drilldown...");
                LogUtility.LogDebug("Custom Drilldown Process Started.");

                token.ThrowIfCancellationRequested();

                await ProcessCustomDrilldownAsync(tableObj, rng, headerLabel, win, token);
            }
            catch (OperationCanceledException)
            {
                LogUtility.LogWarn("Custom Drilldown operation was canceled by the user.");
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex);
            }
            finally
            {
                await SafelyCloseWaitWindowAsync(win);
            }
        }

        private async Task ProcessCustomDrilldownAsync(Excel.ListObject tableObj, Excel.Range rng, string headerLabel, GLWaitWindow win, CancellationToken token)
        {
            Excel.Workbook wb = AppState.Instance.ExcelApp.ActiveWorkbook;
            if (wb?.CustomXMLParts?.Count <= 0)
                return;

            Excel.Worksheet ws = AppState.Instance.ExcelApp.ActiveSheet as Excel.Worksheet;
            string xmlString = FindDrilldownXmlForSheet(wb, ws.Name);

            token.ThrowIfCancellationRequested();

            if (string.IsNullOrEmpty(xmlString))
            {
                string errorMsg = "Unable to find the drilldown meta " +
                    "object check if sheet renamed! If renamed then regenerate drilldown again to get meta object.";
                await ShowErrorMessageAsync(win, errorMsg);
                return;
            }

            await ExecuteCustomDrilldownAsync(tableObj, rng, headerLabel, xmlString, ws, win, token);
        }

        private static string FindDrilldownXmlForSheet(Excel.Workbook wb, string sheetName)
        {
            var cxps = wb.CustomXMLParts;
            string sheetNameEncoded = System.Net.WebUtility.HtmlEncode(sheetName);

            for (int i = cxps.Count; i >= 1; i--)
            {
                string xml = cxps[i].XML;
                if (string.IsNullOrEmpty(xml) || !xml.Contains("DRILLDOWNSHEET"))
                    continue;

                string shtNameInXml = System.Net.WebUtility.HtmlEncode(ExtractDrilldownSheet_XPath(xml));
                if (string.IsNullOrEmpty(shtNameInXml) || !shtNameInXml.Equals(sheetNameEncoded, StringComparison.OrdinalIgnoreCase))
                    continue;

                LogUtility.LogDebug($"Found XML in internal memory location: {xml}");
                return System.Net.WebUtility.HtmlDecode(xml);
            }

            return null;
        }

        private async Task ExecuteCustomDrilldownAsync(Excel.ListObject tableObj, Excel.Range rng, string headerLabel, string xmlString, Excel.Worksheet ws, GLWaitWindow win, CancellationToken token)
        {
            try
            {
                XDocument doc = XDocument.Parse(xmlString);
                XElement columnNameElement = doc.Descendants("COLUMNNAME")
                    .FirstOrDefault(x => x.Attribute("Name")?.Value == headerLabel);

                if (columnNameElement == null)
                    return;

                token.ThrowIfCancellationRequested();

                string jsonData = columnNameElement.Value.Trim();
                LogUtility.LogDebug($"JSON to Parse: {jsonData}");

                JsonNode jsonObject = JsonNode.Parse(jsonData);
                UpdateJsonParameters(jsonObject, tableObj, rng, ws, token);

                string updatedJsonString = JsonSerializer.Serialize(jsonObject, JsonGlobals.Options);
                LogUtility.LogDebug($"Request Sent to Server with JSON Body: {updatedJsonString}");

                string responseString = await SendCustomDrilldownRequestAsync(updatedJsonString, token);
                token.ThrowIfCancellationRequested();

                LogUtility.LogDebug($"Response Received: {responseString}");

                var result = ApiResponseHelper.Parse<JsonElement>(responseString, JsonGlobals.Options);

                if (!result.IsSuccess)
                {
                    string errorMsg = responseString?.Contains("A task was canceled.") == true
                        ? "Task has been canceled."
                        : result.ErrorMessage;

                    await ShowErrorMessageAsync(win, errorMsg);
                    return;
                }

                await WriteCustomDrilldownToWorksheet(ws, rng, jsonData, updatedJsonString, responseString, win, token);
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex);
            }
        }

        private static void UpdateJsonParameters(JsonNode jsonObject, Excel.ListObject tableObj, Excel.Range rng, Excel.Worksheet ws, CancellationToken token)
        {
            JsonArray parameters = jsonObject["parameters"].AsArray();

            foreach (JsonObject param in parameters.Cast<JsonObject>())
            {
                token.ThrowIfCancellationRequested();

                string paramType = param["type"]?.GetValue<string>();

                if (paramType == "COLUMN")
                {
                    UpdateColumnParameter(param, tableObj, rng, ws, token);
                }

                param.Remove("id");
            }
        }

        private static void UpdateColumnParameter(JsonObject param, Excel.ListObject tableObj, Excel.Range rng, Excel.Worksheet ws, CancellationToken token)
        {
            string colName = param[AppConstants.value]?.GetValue<string>();

            try
            {
                int colNumber = FindColumnNumber(tableObj, colName, token);

                if (colNumber > 0)
                {
                    Excel.Range rng1 = ws.Cells[rng.Row, colNumber] as Excel.Range;
                    object cellValue = rng1.Value2;
                    string val = cellValue?.ToString() ?? "";
                    param[AppConstants.value] = val;
                }
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex);
            }
        }

        private static int FindColumnNumber(Excel.ListObject tableObj, string colName, CancellationToken token)
        {
            foreach (Excel.Range hdrRange in tableObj.HeaderRowRange)
            {
                token.ThrowIfCancellationRequested();
                object offsetValue = hdrRange.Offset[-1, 0].Value;
                if (offsetValue != null && offsetValue.ToString().Trim() == colName?.Trim())
                {
                    return hdrRange.Column;
                }
            }
            return 0;
        }

        private static async Task<string> SendCustomDrilldownRequestAsync(string jsonPayload, CancellationToken token)
        {
            string url = $"{AppState.Instance.LoginUrl}{AppConstants.RestSecure}custom-drilldown?cubeId={AppState.Instance.SelectedCube.CubeId}";
            return await ApiHelper.ServerAPI(url, "JSON", jsonPayload, "POST", token);
        }

        private static async Task WriteCustomDrilldownToWorksheet(Excel.Worksheet ws, Excel.Range rng, string jsonData, string updatedJsonString, string responseString, GLWaitWindow win, CancellationToken token)
        {
            string obj1 = $"{ws.Name}_CM_{rng.Address[false, false, Excel.XlReferenceStyle.xlA1]}";
            string obj2 = rng.Address[true, true, Excel.XlReferenceStyle.xlA1, true];
            string obj3 = DateTime.Now.ToString("dd-MMM-yyyy hh:mm:ss tt");
            string obj4 = "CM";
            string obj5 = MergeJsonAndGenerateString(jsonData, updatedJsonString);
            string finalData = $"{obj1}*{obj2}*{obj3}*{obj4}*{obj5}";

            var dataToSheet = new DDDatatoWorksheet(AppState.Instance.ExcelApp, responseString, "CM", finalData, token, win);
            await dataToSheet.DD_DatetoWorksheet();

            token.ThrowIfCancellationRequested();
        }
        private async void RibHighlight_OnClick(object sender, IRibbonControl control, bool pressed)
        {
            LogUtility.LogDebug("RibHighlight_OnClick clicked.");
            if (!GuardLoginAndExcel())
                return;

            Excel.Range adaptiveBalanceRange = null;
            CommonMethods.DisableExcelSettings();
            GLWaitWindow win = null;
            using var ctsHelper = new CancellationHelper();
            CancellationToken token = ctsHelper.GetToken();

            try
            {
                Excel.Worksheet wrkSheet = AppState.Instance.ExcelApp.ActiveSheet as Excel.Worksheet;

                LogUtility.LogDebug("Selecting balance cells whose values are from adaptive memory.");
                LogUtility.LogDebug($"Worksheet Name : {wrkSheet?.Name}");

                if (!ValidateHighlightPreconditions(wrkSheet))
                    return;

                win = CreateAndShowWaitWindow(ctsHelper);
                await InitializeWaitWindowAsync(win, "Highlighting Adaptive Memory", "Searching adaptive memory cells...");

                adaptiveBalanceRange = await FindAdaptiveMemoryCellsFast(wrkSheet, win, token);
            }
            catch (OperationCanceledException)
            {
                LogUtility.LogWarn("Highlight Adaptive Memory operation was canceled by the user.");
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex);
            }
            finally
            {
                await SafelyCloseWaitWindowAsync(win);
                SelectAdaptiveBalanceRange(adaptiveBalanceRange);
                CommonMethods.EnableExcelSettings();
            }
        }
        private static bool ValidateHighlightPreconditions(Excel.Worksheet wrkSheet)
        {
            bool balancesExists = CommonFunctions.BalanceFormulaExists(wrkSheet?.Name);

            if (!balancesExists)
            {
                CommonFunctions.GLSenseMessage("No balance formulas exists in the current worksheet.", MessageBoxIcon.Exclamation, MessageBoxButtons.OK);
                return false;
            }

            System.Data.DataTable dt = AppState.Instance.CalculatedBalances;

            if (dt == null || !dt.Columns.Contains("cache") || dt.Rows.Count == 0)
            {
                LogUtility.LogDebug("No calculated values or 'cache' column found in balance refresh memory, or no rows present.");
                CommonFunctions.GLSenseMessage("Worksheet has to be refreshed first.", MessageBoxIcon.Exclamation, MessageBoxButtons.OK);
                return false;
            }

            return true;
        }

        private static async Task<Excel.Range> FindAdaptiveMemoryCellsFast(Excel.Worksheet wrkSheet, GLWaitWindow win, CancellationToken token)
        {
            System.Data.DataTable dt = AppState.Instance.CalculatedBalances;
            string sheetNameEscaped = wrkSheet.Name.Replace("'", "''");
            string dataTableFilter = $"[excelSheet]='{sheetNameEscaped}' AND [cache] = True";

            token.ThrowIfCancellationRequested();

            if (win != null)
            {
                await MessageWaitWindowAsync(win, "Filtering data...");
                await Task.Delay(1, token);
            }

            DataRow[] sheetRows = dt.Select(dataTableFilter);

            if (sheetRows == null || sheetRows.Length == 0)
            {
                await ShowErrorMessageAsync(win, "No balances from adaptive memory");
                return null;
            }

            if (win != null)
            {
                await MessageWaitWindowAsync(win, $"Parsing {sheetRows.Length} cell addresses...");
                await Task.Delay(1, token);
            }

            // STEP 1: Parse and deduplicate in one pass
            var cellMap = new Dictionary<string, (int col, int row)>(StringComparer.OrdinalIgnoreCase);
            var addressOrder = new List<string>(sheetRows.Length);

            for (int i = 0; i < sheetRows.Length; i++)
            {
                token.ThrowIfCancellationRequested();

                string cellAddress = sheetRows[i]["excelCell"]?.ToString();
                if (string.IsNullOrWhiteSpace(cellAddress)) continue;

                if (!cellMap.ContainsKey(cellAddress))
                {
                    var (col, row) = ParseExcelCell(cellAddress);
                    cellMap[cellAddress] = (col, row);
                    addressOrder.Add(cellAddress);
                }
            }

            if (cellMap.Count == 0)
            {
                await ShowErrorMessageAsync(win, "No valid cell addresses found");
                return null;
            }

            if (win != null)
            {
                await MessageWaitWindowAsync(win, $"Sorting {cellMap.Count} unique cells...");
                await Task.Delay(1, token);
            }

            // STEP 2: Sort unique addresses by column then row

            var sortedAddresses = addressOrder
                .OrderBy(addr => cellMap[addr].col)
                .ThenBy(addr => cellMap[addr].row)
                .ToList();

            if (win != null)
            {
                await MessageWaitWindowAsync(win, $"Creating range from {sortedAddresses.Count} cells...");
                await Task.Delay(1, token);
            }

            // STEP 3: Build range efficiently
            Excel.Range result = await BuildRangeEfficiently(wrkSheet, sortedAddresses, win, token);

            if (win != null)
            {
                await MessageWaitWindowAsync(win, "Ready");
                await Task.Delay(1, token);
            }

            return result;
        }

        private static async Task<Excel.Range> BuildRangeEfficiently(Excel.Worksheet wrkSheet, List<string> addresses, GLWaitWindow win, CancellationToken token)
        {
            try
            {
                Excel.Range finalRange = null;
                int currentIndex = 0;
                int totalBatches = 0;

                while (currentIndex < addresses.Count)
                {
                    token.ThrowIfCancellationRequested();

                    // Build batch based on character count (max 200 chars for safety)
                    var batch = new List<string>();
                    int currentLength = 0;

                    while (currentIndex < addresses.Count)
                    {
                        string address = addresses[currentIndex];
                        int addedLength = address.Length;

                        // Add 1 for comma separator (except first item)
                        if (batch.Count > 0)
                        {
                            addedLength += 1;
                        }

                        if (currentLength + addedLength > 200)
                        {
                            break;
                        }

                        batch.Add(address);
                        currentLength += addedLength;
                        currentIndex++;
                    }

                    // Ensure we process at least one address
                    if (batch.Count == 0 && currentIndex < addresses.Count)
                    {
                        batch.Add(addresses[currentIndex]);
                        currentIndex++;
                    }

                    totalBatches++;

                    if (win != null)
                    {
                        await MessageWaitWindowAsync(win, $"Building range: batch {totalBatches} ({batch.Count} cells)");
                        await Task.Delay(1, token);
                    }

                    string batchAddress = string.Join(",", batch);

                    try
                    {
                        Excel.Range batchRange = wrkSheet.Range[batchAddress];

                        if (finalRange == null)
                        {
                            finalRange = batchRange;
                        }
                        else
                        {
                            finalRange = wrkSheet.Application.Union(finalRange, batchRange);
                        }
                    }
                    catch (Exception ex)
                    {
                        LogUtility.LogException(ex);

                        // Fallback: Process each cell individually for this batch
                        foreach (string cell in batch)
                        {
                            try
                            {
                                Excel.Range cellRange = wrkSheet.Range[cell];
                                if (finalRange == null)
                                {
                                    finalRange = cellRange;
                                }
                                else
                                {
                                    finalRange = wrkSheet.Application.Union(finalRange, cellRange);
                                }
                            }
                            catch (Exception cellEx)
                            {
                                LogUtility.LogException(cellEx);
                            }
                        }
                    }
                }
                return finalRange;
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex);
                return await BuildRangeCellByCell(wrkSheet, addresses, win, token);
            }
        }

        private static async Task<Excel.Range> BuildRangeCellByCell(Excel.Worksheet wrkSheet, List<string> addresses, GLWaitWindow win, CancellationToken token)
        {
            Excel.Range finalRange = null;
            int total = addresses.Count;

            for (int i = 0; i < total; i++)
            {
                token.ThrowIfCancellationRequested();

                if (win != null && i % 100 == 0)
                {
                    await MessageWaitWindowAsync(win, $"Processing cell {i}/{total}");
                    await Task.Delay(1, token);
                }

                try
                {
                    Excel.Range cellRange = wrkSheet.Range[addresses[i]];

                    if (finalRange == null)
                    {
                        finalRange = cellRange;
                    }
                    else
                    {
                        finalRange = wrkSheet.Application.Union(finalRange, cellRange);
                    }
                }
                catch (Exception ex)
                {
                    LogUtility.LogException(ex);
                }
            }

            return finalRange;
        }

        private static (int col, int row) ParseExcelCell(string cellAddress)
        {
            int col = 0;
            int row = 0;
            int idx = 0;
            int len = cellAddress.Length;

            // Skip leading $
            if (len > 0 && cellAddress[idx] == '$') idx++;

            // Parse column letters (A-Z)
            while (idx < len)
            {
                char c = cellAddress[idx];
                if (c >= 'A' && c <= 'Z')
                {
                    col = col * 26 + (c - 'A' + 1);
                    idx++;
                }
                else if (c >= 'a' && c <= 'z')
                {
                    col = col * 26 + (c - 'a' + 1);
                    idx++;
                }
                else
                {
                    break;
                }
            }

            // Skip $ before row
            if (idx < len && cellAddress[idx] == '$') idx++;

            // Parse row digits
            while (idx < len)
            {
                char c = cellAddress[idx];
                if (c >= '0' && c <= '9')
                {
                    row = row * 10 + (c - '0');
                    idx++;
                }
                else
                {
                    break;
                }
            }

            return (col, row);
        }
        private static void SelectAdaptiveBalanceRange(Excel.Range adaptiveBalanceRange)
        {
            if (adaptiveBalanceRange == null)
                return;

            LogUtility.LogDebug($"Balance cells found with adaptive memory : {adaptiveBalanceRange?.Address}");
            try
            {
                AppState.Instance.ExcelApp.ActiveWorkbook?.Activate();
                Excel.Worksheet ws = AppState.Instance.ExcelApp.ActiveSheet as Excel.Worksheet;
                ws?.Activate();

                adaptiveBalanceRange.Select();
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex);
            }
        }

        private async void RibRefreshRange_OnClick(object sender, IRibbonControl control, bool pressed)
        {
            LogUtility.LogDebug("RibRefreshRange_OnClick clicked.");
            if (!GuardLoginAndExcel())
            {
                CommonFunctions.GLSenseMessage("Please log in to the instance.", MessageBoxIcon.Exclamation, MessageBoxButtons.OK);
                return;
            }

            CommonMethods.DisableExcelSettings();
            GLWaitWindow win = null;
            using var ctsHelper = new CancellationHelper();
            CancellationToken token = ctsHelper.GetToken();

            try
            {
                Excel.Range formulaCells = AppState.Instance.ExcelApp.Selection as Excel.Range;
                int glBalCount = CountBalanceFormulas(formulaCells, token);

                if (!ValidateRefreshRange(glBalCount))
                    return;

                token.ThrowIfCancellationRequested();

                win = CreateAndShowWaitWindow(ctsHelper);

                await InitializeWaitWindowAsync(win, "Range Refresh", "Refreshing selected range...");
                await Task.Yield();

                await RefreshFormulaCells(formulaCells, win, token);
            }
            catch (OperationCanceledException)
            {
                LogUtility.LogWarn("Range refresh operation was canceled by the user.");
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex);
            }
            finally
            {
                CommonMethods.EnableExcelSettings();
                await SafelyCloseWaitWindowAsync(win);
            }
        }

        private static int CountBalanceFormulas(Excel.Range formulaCells, CancellationToken token)
        {
            int count = 0;
            foreach (Excel.Range cell in formulaCells)
            {
                token.ThrowIfCancellationRequested();

                if (cell.HasFormula is true &&
                    cell.Formula is string formula &&
                    formula.IndexOf(AppConstants.glBal, StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    count++;
                }
            }
            return count;
        }

        private static bool ValidateRefreshRange(int glBalCount)
        {
            if (glBalCount == 0)
            {
                CommonFunctions.GLSenseMessage("No balance formulas exists in the selected range.", MessageBoxIcon.Exclamation, MessageBoxButtons.OK);
                return false;
            }


            if (glBalCount > UserConfig.RefreshCells)
            {
                string msg = $"The selected range contains {glBalCount} balance formulas.\nThe configured refresh range is {UserConfig.RefreshCells}.\nChange the configuration and try again! Max refresh range limit is 100.";
                CommonFunctions.GLSenseMessage(msg, MessageBoxIcon.Warning, MessageBoxButtons.OK);
                return false;
            }

            return true;
        }

        private static async Task RefreshFormulaCells(Excel.Range formulaCells, GLWaitWindow win, CancellationToken token)
        {
            try
            {
                AppState.Instance.SingleRefresh = true;
                foreach (Excel.Range cell in formulaCells)
                {
                    token.ThrowIfCancellationRequested();
                    await MessageWaitWindowAsync(win, $"Refreshing range {cell.Address}.");
                    await Task.Yield();
                    cell.Dirty();
                    cell.Calculate();
                }
                await MessageWaitWindowAsync(win, "Completed refreshing the range.");
            }
            catch (Exception ex)
            {
                AppState.Instance.SingleRefresh = false;
                LogUtility.LogException(ex);
            }
        }

        private static string ExtractDrilldownSheet_XPath(string xml)
        {
            if (string.IsNullOrWhiteSpace(xml)) return null;

            var doc = new XmlDocument { XmlResolver = null };

            try
            {
                doc.LoadXml(xml);
            }
            catch (XmlException ex)
            {
                LogUtility.LogWarn($"AddinModule.ExtractDrilldownSheet_XPath: malformed XML (ignored): {ex.Message}");
                return null;
            }

            try
            {
                var nav = doc.CreateNavigator();
                var node = nav.SelectSingleNode("//DRILLDOWNSHEET/text()[1]");
                var value = node?.Value?.Trim();
                return string.IsNullOrEmpty(value) ? null : value;
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex);
                return null;
            }
        }

        private static string MergeJsonAndGenerateString(string jsonNames, string jsonValues)
        {
            JsonNode namesNode = JsonNode.Parse(jsonNames);
            JsonNode valuesNode = JsonNode.Parse(jsonValues);

            string identifierName = namesNode["identifierName"]?.GetValue<string>();

            var paramMap = new Dictionary<string, string>();
            foreach (JsonObject param in namesNode["parameters"].AsArray().Cast<JsonObject>())
            {
                string name = param["name"]?.GetValue<string>();
                string value = param[AppConstants.value]?.GetValue<string>();
                if (name != null)
                    paramMap[name] = value;
            }

            var result = new StringBuilder();
            result.Append($"Report Name:= {identifierName}");

            foreach (JsonObject param in valuesNode["parameters"].AsArray().Cast<JsonObject>())
            {
                string paramName = param["name"]?.GetValue<string>();
                string paramValue = param[AppConstants.value]?.GetValue<string>();
                string paramType = param["type"]?.GetValue<string>();

                if (paramMap.TryGetValue(paramName, out string mappedName))
                {
                    if (paramType == "CONSTANT")
                        result.Append($", Constant:= {paramValue}");
                    else
                        result.Append($", {mappedName}:= {paramValue}");
                }
            }

            return result.ToString();
        }

        private async Task GLSense_DownloadFile(string strURL, string postData, CancellationToken token)
        {
            try
            {
                LogUtility.LogDebug($"Downloading file from {strURL}");

                token.ThrowIfCancellationRequested();

                var handler = new HttpClientHandler
                {
                    SslProtocols = SslProtocols.Tls12 | SslProtocols.Tls13,
                    ServerCertificateCustomValidationCallback = StrictCertificateValidator.Validate
                };

                token.ThrowIfCancellationRequested();

                using var client = new HttpClient(handler);
                client.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", AppState.Instance.LoginToken);
                client.Timeout = Timeout.InfiniteTimeSpan;

                HttpContent content = string.IsNullOrWhiteSpace(postData) ? null : new StringContent(postData, Encoding.UTF8, "application/json");

                HttpResponseMessage response = await client.PostAsync(strURL, content, token);

                token.ThrowIfCancellationRequested();

                if (response.IsSuccessStatusCode)
                {
                    string fileName = response.Content.Headers.ContentDisposition?.FileName?.Trim('"') ?? "DownloadedFile.zip";

                    string downloadPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.UserProfile), "Downloads", fileName);

                    using var fileStream = new FileStream(downloadPath, FileMode.Create, FileAccess.Write, FileShare.None);
                    await response.Content.CopyToAsync(fileStream);

                    CommonFunctions.GLSenseMessage($"Attachment saved to downloads folder as \"{fileName}\"", MessageBoxIcon.Information);
                }
                else
                {
                    LogUtility.LogError($"Download failed: {response.StatusCode} - {response.ReasonPhrase}");
                }
            }
            catch (HttpRequestException ex)
            {
                LogUtility.LogException(ex);
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex);
            }
        }

        private void adxExcelAppEvents1_SheetChange(object sender, object sheet, object range)
        {

            if (!GuardLoginAndExcel())
            {
                return;
            }

            try
            {
                if (range is not Excel.Range rng) return;

                LogUtility.LogDebug($"SheetChange fired. Sheet={(sheet as Excel.Worksheet)?.Name}, Range={rng.Address}");

                if (TryGetSingleCellFormula(rng, out string formula) &&
                    formula.IndexOf(AppConstants.glBal, StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    _ribbonHelper.ApplyState("LoggedIn");
                }
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex, "adxExcelAppEvents1_SheetChange");
            }
        }

        private void adxExcelAppEvents1_SheetSelectionChange(object sender, object sheet, object range)
        {
            try
            {
                if (!AppState.Instance.IsLoginCompleted) return;
                if (range is not Excel.Range rng) return;
                if (rng.Rows.Count != 1 || rng.Columns.Count != 1) return;

                LogUtility.LogDebug($"SheetSelectionChange fired. Sheet={(sheet as Excel.Worksheet)?.Name}, Cell={rng.Address}");

                AppState.Instance.BalancePane = GetPaneInstance();
                if (AppState.Instance.BalancePane != null && AppState.Instance.BalancePane.Visible)
                {
                    if (TryGetSingleCellFormula(rng, out string formula) &&
                        formula.IndexOf(AppConstants.glBal, StringComparison.OrdinalIgnoreCase) >= 0)
                    {
                        _ = AppState.Instance.BalancePane.RelaunchPane();
                    }
                    else
                    {
                        _ = AppState.Instance.BalancePane.ResetPaneReference();
                    }
                }
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex, "adxExcelAppEvents1_SheetSelectionChange");
            }
        }

        public GLConfiguratorPane GetPaneInstance()
        {
            try
            {
                return (GLConfiguratorPane)adxExcelTaskPanesCollectionItem1.TaskPaneInstance;
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex);
                return null;
            }
        }

        private void SyncRibbonSelectionWithAppState()
        {
            if (!AppState.Instance.IsLoginCompleted)
                return;

            try
            {
                var cube = AppState.Instance.SelectedCube;
                var ledger = AppState.Instance.SelectedLedger;

                if (cube == null)
                    return;

                RibGetCube.Caption = $"Cube: {cube.CubeName}";

                Ribledger.Items.Clear();
                if (cube.Ledgers != null)
                {
                    foreach (var l in cube.Ledgers.OrderBy(l => l.LedgerName))
                    {
                        Ribledger.Items.Add(new AddinExpress.MSO.ADXRibbonItem { Caption = l.LedgerName });
                    }
                }
                Ribledger.Text = ledger?.LedgerName ?? string.Empty;

                RibSegS.Items.Clear();
                if (ledger != null)
                {
                    var repository = new DataRepository();
                    foreach (var s in repository.GetSegments(cube.CubeId, ledger.LedgerId))
                    {
                        RibSegS.Items.Add(new AddinExpress.MSO.ADXRibbonItem { Caption = s.SegmentName });
                    }

                    if (!string.IsNullOrWhiteSpace(AppState.Instance.DefaultSegment))
                    {
                        RibSegS.Text = AppState.Instance.DefaultSegment;
                    }
                }
                else
                {
                    RibSegS.Text = string.Empty;
                }
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex, "Failed to sync ribbon selection with app state.");
            }
        }

        private void adxExcelAppEvents1_WorkbookActivate(object sender, object hostObj)
        {
            if (AppState.Instance.ExcelApp == null || !AppState.Instance.IsLoginCompleted || RibLogin.Visible) return;

            try
            {
                LogUtility.LogDebug($"WorkbookActivate fired. Workbook={AppState.Instance.ExcelApp?.ActiveWorkbook?.Name}");
                _ribbonHelper.ApplyState("LoggedIn");
                SyncRibbonSelectionWithAppState();
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex, "adxExcelAppEvents1_WorkbookActivate");
            }
        }

        private void RibHelp_OnClick(object sender, IRibbonControl control, bool pressed)
        {
            try
            {
                LogUtility.LogDebug("RibHelp_OnClick clicked.");
                if (!AppState.Instance.IsLoginCompleted)
                {
                    CommonFunctions.GLSenseMessage("Please log in to the instance.", MessageBoxIcon.Exclamation, MessageBoxButtons.OK);
                    return;
                }

                if (!string.IsNullOrEmpty(AppState.Instance.LoginToken))
                {
                    string helpUrl = AppState.Instance.LoginUrl + "/web/public/redirect-help/Excel_GLSense.htm?jwtParam=" + AppState.Instance.LoginToken;
                    LogUtility.LogDebug(helpUrl);
                    Process.Start(helpUrl);
                }
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex);
            }
        }

        private void RibAbout_OnClick(object sender, IRibbonControl control, bool pressed)
        {
            LogUtility.LogDebug("RibAbout_OnClick clicked.");
            SafeInvokeWpf(() =>
            {
                var win = new GLAbout();
                win.ShowDialogWithOwner((IntPtr)AppState.Instance.ExcelApp.Hwnd);
            });
        }

        private void RibVersionCheck_OnClick(object sender, IRibbonControl control, bool pressed)
        {
            LogUtility.LogDebug($"RibVersionCheck_OnClick clicked. Pressed={pressed}");
            AppState.Instance.VersionCheck = pressed;
        }
        private void RibDebug_OnClick(object sender, IRibbonControl control, bool pressed)
        {
            AppState.Instance.DebugLogs = pressed;

            if (pressed)
            {
                LogUtility.LogDebug("Debug session started. Detailed traces will be written until disabled.");
            }
            else
            {
                LogUtility.FlushDebugLogs("Debug session ended");
            }
        }

        private void RibUserConfig_OnClick(object sender, IRibbonControl control, bool pressed)
        {
            LogUtility.LogDebug("RibUserConfig_OnClick clicked.");
            SafeInvokeWpf(() =>
            {
                var win = new GLUserConfig();
                win.ShowDialogWithOwner((IntPtr)AppState.Instance.ExcelApp.Hwnd);
            });
        }

        private void Riburl_OnClick(object sender, IRibbonControl control, bool pressed)
        {
            LogUtility.LogDebug("Riburl_OnClick clicked.");
            SafeInvokeWpf(() =>
            {
                var win = new GLServerConfiguration();
                win.ShowDialogWithOwner((IntPtr)AppState.Instance.ExcelApp.Hwnd);
            });
        }

        private void RibDailyRate_OnClick(object sender, IRibbonControl control, bool pressed)
        {
            LogUtility.LogDebug("RibDailyRate_OnClick clicked.");
            SafeInvokeWpf(() =>
            {
                var win = new GLDailyRates();
                win.ShowDialogWithOwner((IntPtr)AppState.Instance.ExcelApp.Hwnd);
            });
        }

        private void LaunchPeriodDetails(string FuncName)
        {
            LogUtility.LogDebug($"LaunchPeriodDetails invoked. FuncName={FuncName}");
            SafeInvokeWpf(() =>
            {
                var win = new GLGetPeriodDetails(FuncName);
                win.ShowDialogWithOwner((IntPtr)AppState.Instance.ExcelApp.Hwnd);
            });
        }

        private void LaunchPeriodStarEnd(string FuncName)
        {
            LogUtility.LogDebug($"LaunchPeriodStarEnd invoked. FuncName={FuncName}");
            SafeInvokeWpf(() =>
            {
                var win = new GLGetPeriodStartEnd(FuncName);
                win.ShowDialogWithOwner((IntPtr)AppState.Instance.ExcelApp.Hwnd);
            });
        }

        private void RibPeriodEnd_OnClick(object sender, IRibbonControl control, bool pressed) => LaunchPeriodStarEnd("END");
        private void RibPeriodStart_OnClick(object sender, IRibbonControl control, bool pressed) => LaunchPeriodStarEnd("START");
        private void RibPeriodYear_OnClick(object sender, IRibbonControl control, bool pressed) => LaunchPeriodDetails("YEAR");
        private void RibPeriodQtr_OnClick(object sender, IRibbonControl control, bool pressed) => LaunchPeriodDetails("QTR");
        private void RibPeriodNum_OnClick(object sender, IRibbonControl control, bool pressed) => LaunchPeriodDetails("NUM");

        private void RibPeriodbyYear_OnClick(object sender, IRibbonControl control, bool pressed)
        {
            LogUtility.LogDebug("RibPeriodbyYear_OnClick clicked.");
            SafeInvokeWpf(() =>
            {
                var win = new GLGetPeriodByYear();
                win.ShowDialogWithOwner((IntPtr)AppState.Instance.ExcelApp.Hwnd);
            });
        }

        private void RibPeriodbyDate_OnClick(object sender, IRibbonControl control, bool pressed)
        {
            LogUtility.LogDebug("RibPeriodbyDate_OnClick clicked.");
            SafeInvokeWpf(() =>
            {
                var win = new GLGetPeriodByDate();
                win.ShowDialogWithOwner((IntPtr)AppState.Instance.ExcelApp.Hwnd);
            });
        }

        private void RibPeriod_OnClick(object sender, IRibbonControl control, bool pressed)
        {
            LogUtility.LogDebug("RibPeriod_OnClick clicked.");
            SafeInvokeWpf(() =>
            {
                var win = new GLGetPeriod();
                win.ShowDialogWithOwner((IntPtr)AppState.Instance.ExcelApp.Hwnd);
            });
        }

        private void LaunchSegmentWindow(string FuncName)
        {
            LogUtility.LogDebug($"LaunchSegmentWindow invoked. FuncName={FuncName}");
            SafeInvokeWpf(() =>
            {
                var win = new GLSegmentFunctions(FuncName);
                win.ShowDialogWithOwner((IntPtr)AppState.Instance.ExcelApp.Hwnd);
            });
        }

        private void RibSegmentEnabledFlag_OnClick(object sender, IRibbonControl control, bool pressed) => LaunchSegmentWindow("ENABLEDFLAG");
        private void RibSummaryFlag_OnClick(object sender, IRibbonControl control, bool pressed) => LaunchSegmentWindow("SUMMARYFLAG");
        private void RibSegment_OnClick(object sender, IRibbonControl control, bool pressed) => LaunchSegmentWindow("DESCRIPTION");
        private void RibNextSegment_OnClick(object sender, IRibbonControl control, bool pressed) => LaunchSegmentWindow("NEXTSEGMENT");
        private void RibPreviousSegment_OnClick(object sender, IRibbonControl control, bool pressed) => LaunchSegmentWindow("PREVIOUSSEGMENT");
        private void RibSegmentDFF_OnClick(object sender, IRibbonControl control, bool pressed) => LaunchSegmentWindow("DFF");
        private void RibSegmentAccountType_OnClick(object sender, IRibbonControl control, bool pressed) => LaunchSegmentWindow("ACCOUNTTYPE");

        private void RibDrillJobs_OnClick(object sender, IRibbonControl control, bool pressed)
        {
            LogUtility.LogDebug("RibDrillJobs_OnClick clicked.");
            SafeInvokeWpf(() =>
            {
                var win = new GLJobsMonitor();
                win.ShowDialogWithOwner((IntPtr)AppState.Instance.ExcelApp.Hwnd);
            });
        }

        private void RibDDConfiguration_OnClick(object sender, IRibbonControl control, bool pressed)
        {
            LogUtility.LogDebug("RibDDConfiguration_OnClick clicked.");
            SafeInvokeWpf(() =>
            {
                var win = new GLDrilldownCustomization();
                win.ShowDialogWithOwner((IntPtr)AppState.Instance.ExcelApp.Hwnd);
            });
        }

        private static async Task RunBalanceDrilldownAsync(string ddType)
        {
            try
            {
                LogUtility.LogDebug($"RunBalanceDrilldownAsync invoked. ddType={ddType}");
                Excel.Range rng = (Excel.Range)AppState.Instance.ExcelApp.Selection;
                string external = ExcelExternalRef.BuildExternalAddress(rng);
                var runProcess = new DrilldownBl(AppState.Instance.ExcelApp, external, ddType);
                await runProcess.ProcessBLDrilldown();
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex);
            }
        }

        private async void RibBalanceDD_OnClick(object sender, IRibbonControl control, bool pressed) => await RunBalanceDrilldownAsync("BL");
        private async void RibBalanceJournalDD_OnClick(object sender, IRibbonControl control, bool pressed) => await RunBalanceDrilldownAsync("BL_JL");
        private async void RibBalanceSubLedgerDD_OnClick(object sender, IRibbonControl control, bool pressed) => await RunBalanceDrilldownAsync("BL_SL");
        private async void RibTotaDD_OnClick(object sender, IRibbonControl control, bool pressed) => await RunBalanceDrilldownAsync("UF");

        private async void RibJournalDD_OnClick(object sender, IRibbonControl control, bool pressed)
        {
            try
            {
                await RunDrilldownAsync("JL");
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex);
            }
        }
        private async void RibBalancesDDToSubLedger_OnClick(object sender, IRibbonControl control, bool pressed)
        {
            try
            {
                await RunDrilldownAsync("BLDD_SL");
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex);
            }
        }

        private async void RibBalancesDDToUnified_OnClick(object sender, IRibbonControl control, bool pressed)
        {
            try
            {
                await RunDrilldownAsync("BLDD_UF");
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex);
            }
        }
        private async Task RunDrilldownAsync(string ddType)
        {
            try
            {
                LogUtility.LogDebug($"RunDrilldownAsync invoked. ddType={ddType}");
                Excel.Range rng = (Excel.Range)AppState.Instance.ExcelApp.Selection;
                string external = ExcelExternalRef.BuildExternalAddress(rng);

                var runProcess = new DrilldownJl(AppState.Instance.ExcelApp, external, ddType);
                await runProcess.ProcessJLDrilldown();
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex);
            }
        }
        private async void RibSubledgerDD_OnClick(object sender, IRibbonControl control, bool pressed)
        {
            try
            {
                LogUtility.LogDebug("RibSubledgerDD_OnClick clicked.");
                Excel.Range rng = (Excel.Range)AppState.Instance.ExcelApp.Selection;
                string external = ExcelExternalRef.BuildExternalAddress(rng);
                var runProcess = new DrilldownSl(AppState.Instance.ExcelApp, external);
                await runProcess.ProcessSLDrilldown();
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex);
            }
        }

        private async void RibCellHighlight_OnClick(object sender, IRibbonControl control, bool pressed)
        {
            LogUtility.LogDebug("RibCellHighlight_OnClick clicked.");
            await DrillCellHighlighter.RibCellHighlight_OnClick();
        }

        private void RibRefreshAll_OnClick(object sender, IRibbonControl control, bool pressed)
        {
            LogUtility.LogDebug("RibRefreshAll_OnClick clicked.");
            _ = BalanceRefresh.RefreshingBalancesAsync("Refresh", "Sheet");
        }
        private void RibRefreshBook_OnClick(object sender, IRibbonControl control, bool pressed)
        {
            LogUtility.LogDebug("RibRefreshBook_OnClick clicked.");
            _ = BalanceRefresh.RefreshingBalancesAsync("Refresh", "Book");
        }
        private void RibSnapSubmit_OnClick(object sender, IRibbonControl control, bool pressed)
        {
            LogUtility.LogDebug($"RibSnapSubmit_OnClick clicked. Pressed={RibSnapSubmit.Pressed}");
            AppState.Instance.SnapshotJob = RibSnapSubmit.Pressed;
        }
        private void RibSnapShot_OnClick(object sender, IRibbonControl control, bool pressed)
        {
            string mode = RibSnapWorksheet.Pressed ? "Sheet" : "Book";
            LogUtility.LogDebug($"RibSnapShot_OnClick clicked. Mode={mode}, Pressed={pressed}");

            ADXRibbonCheckBox ribbonCheckBox = RibSnapSubmit.AsRibbonCheckBox;

            AppState.Instance.SnapshotJob = ribbonCheckBox != null && ribbonCheckBox.Pressed;

            if (AppState.Instance.SnapshotJob)
            {
                _ = BalanceRefresh.SubmitSnapAsync(mode);
            }
            else
            {
                _ = BalanceRefresh.RefreshingBalancesAsync("Snapshot", mode);
            }
        }
        private static async Task ResetBalances(string ResetType)
        {
            GLWaitWindow win = null;
            using var ctsHelper = new CancellationHelper();
            CancellationToken token = ctsHelper.GetToken();

            try
            {
                LogUtility.LogDebug($"ResetBalances invoked. ResetType={ResetType}");

                Excel.Worksheet wrksheet = AppState.Instance.ExcelApp.ActiveSheet as Excel.Worksheet;

                // Regression fix: check whether the target sheet/workbook has any balance
                // formulas at all FIRST, before doing anything else (no DisableExcelSettings,
                // no wait window) - previously this check (see CommonFunctions.
                // BalanceFormulaExists, the same helper ValidateHighlightPreconditions uses
                // for the Highlight feature) ran only after the wait window was already
                // shown, so the user would see a "Reset Balance Formulas" window flash up
                // right before immediately being told there was nothing to reset. Now it's
                // the very first thing that happens, so the window only ever appears once
                // there's actually something to do. Calls CommonFunctions.GLSenseMessage
                // directly (not the ShowErrorMessageAsync(win, ...) helper used below) since
                // win doesn't exist yet at this point and there's nothing to close first.
                bool balancesExist = ResetType == "Sheet"
                    ? wrksheet != null && CommonFunctions.BalanceFormulaExists(wrksheet.Name)
                    : AppState.Instance.ExcelApp.ActiveWorkbook.Worksheets
                          .Cast<Excel.Worksheet>()
                          .Any(sht => CommonFunctions.BalanceFormulaExists(sht.Name));

                if (!balancesExist)
                {
                    string scope = ResetType == "Sheet"
                        ? $"worksheet \"{wrksheet?.Name}\""
                        : $"workbook \"{AppState.Instance.ExcelApp.ActiveWorkbook?.Name}\"";
                    LogUtility.LogWarn($"ResetBalances: no balance formulas found in {scope} - nothing to reset.");
                    CommonFunctions.GLSenseMessage($"No balance formulas found in {scope}.", MessageBoxIcon.Warning, MessageBoxButtons.OK);
                    return;
                }

                AppState.Instance.ResetFormulas = true;
                CommonMethods.DisableExcelSettings();

                win = CreateAndShowWaitWindow(ctsHelper);
                await InitializeWaitWindowAsync(win, "Reset Balance Formulas", "Resetting balances...");

                await MessageWaitWindowAsync(win, "Checking if there are any broken links in the workbook.");
                string brokenLinks = CommonFunctions.WorkbookBrokenLinks();
                if (!string.IsNullOrWhiteSpace(brokenLinks))
                {
                    var linksText = brokenLinks.Any() ? $"\"{string.Join("\", \"", brokenLinks)}\"" : "unknown links";
                    await ShowErrorMessageAsync(win, $"The workbook has broken links: {Environment.NewLine}{linksText}{Environment.NewLine}.Please fix them.");
                    return;
                }

                if (ResetType == "Sheet")
                {
                    if (wrksheet != null)
                    {
                        BalancesReset(wrksheet.Name);
                        token.ThrowIfCancellationRequested();
                    }
                }
                else
                {
                    foreach (Excel.Worksheet sht in AppState.Instance.ExcelApp.ActiveWorkbook.Worksheets)
                    {
                        token.ThrowIfCancellationRequested();
                        sht.Activate();
                        BalancesReset(sht.Name);
                    }
                    wrksheet?.Activate();
                }
            }
            catch (OperationCanceledException)
            {
                LogUtility.LogWarn("Reset Balances operation was canceled by the user.");
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex);
            }
            finally
            {
                CommonMethods.EnableExcelSettings();
                AppState.Instance.ResetFormulas = false;
                await SafelyCloseWaitWindowAsync(win);
            }
        }

        private static void BalancesReset(string shtname)
        {
            Excel.Worksheet ws = AppState.Instance.ExcelApp.ActiveWorkbook.Worksheets[shtname] as Excel.Worksheet;
            try
            {
                //Forcing the excel to recalculate all formulas in the sheet by toggling the calculation mode. This is to reset any cached balance values without having to loop through all cells.
                ws.EnableCalculation = false;
                ws.EnableCalculation = true;

                string cleanSheetName = shtname.Replace("'", "");

                AppState.Instance.CalculatedBalances?
                    .AsEnumerable()?
                    .Where(row => string.Equals(
                        row.Field<string>("excelSheet").Replace("'", ""),
                        cleanSheetName,
                        StringComparison.OrdinalIgnoreCase))
                    .ToList()
                    .ForEach(row => row["cache"] = false);
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex, $"Exception encountered by resetting balances in worksheet \"{shtname}\"");
            }
        }

        private void RibClearSheet_OnClick(object sender, IRibbonControl control, bool pressed) { LogUtility.LogDebug("RibClearSheet_OnClick clicked."); _ = ResetBalances("Sheet"); }
        private void RibClear_OnClick(object sender, IRibbonControl control, bool pressed) { LogUtility.LogDebug("RibClear_OnClick clicked."); _ = ResetBalances("Book"); }
        private void RibSegProperty_OnClick(object sender, IRibbonControl control, bool pressed) { LogUtility.LogDebug("RibSegProperty_OnClick clicked."); _ = SegmentDiscoverer.SegmentAction("Property"); }
        private void RibSegmentExpand_OnClick(object sender, IRibbonControl control, bool pressed)
        {
            LogUtility.LogDebug("RibSegmentExpand_OnClick clicked.");
            // Was a "Hierarchy" menu with two items (RibExpandAll/RibbonExpand1Level ->
            // SegmentDiscoverer.SegmentAction("HierarchyAll"/"Hierarchy1Level") directly).
            // Now a single button that opens GLExpandOptions, where the user picks the
            // level (All/1 Level) and fill direction (Rows/Columns) before
            // SegmentDiscoverer.SegmentAction gets called from that dialog. Only a
            // lightweight DefaultSegment/SegmentPickedIndex guard happens here -
            // SegmentAction (invoked from the dialog's Expand button) already does the
            // full active-cell/segment-value validation, same as it did when these were
            // two separate ribbon buttons.
            if (AppState.Instance.DefaultSegment == null || string.IsNullOrEmpty(AppState.Instance.DefaultSegment) || AppState.Instance.SegmentPickedIndex < 0)
            {
                CommonFunctions.GLSenseMessage("Please select a segment from the dropdown.", MessageBoxIcon.Exclamation, MessageBoxButtons.OK);
                return;
            }

            SafeInvokeWpf(() =>
            {
                var win = new GLExpandOptions();
                win.ShowDialogWithOwner((IntPtr)AppState.Instance.ExcelApp.Hwnd);
            });
        }
        private void RibExpodeAll_OnClick(object sender, IRibbonControl control, bool pressed) { LogUtility.LogDebug("RibExpodeAll_OnClick clicked."); _ = SegmentDiscoverer.SegmentAction("ExplodeAll"); }
        private void RibbonExplode1Level_OnClick(object sender, IRibbonControl control, bool pressed) { LogUtility.LogDebug("RibbonExplode1Level_OnClick clicked."); _ = SegmentDiscoverer.SegmentAction("Explode1Level"); }
        private void RibDiscoverPeriod_OnClick(object sender, IRibbonControl control, bool pressed) { LogUtility.LogDebug("RibDiscoverPeriod_OnClick clicked."); _ = PeriodsDiscoverer.FillPeriods(); }

        private void RigSegDiscover_OnClick(object sender, IRibbonControl control, bool pressed)
        {
            LogUtility.LogDebug("RigSegDiscover_OnClick clicked.");
            if (AppState.Instance.DefaultSegment == null || string.IsNullOrEmpty(AppState.Instance.DefaultSegment) || AppState.Instance.SegmentPickedIndex < 0)
            {
                CommonFunctions.GLSenseMessage("Please select a segment from the dropdown.", MessageBoxIcon.Exclamation, MessageBoxButtons.OK);
                return;
            }

            var activeCell = AppState.Instance.ExcelApp?.ActiveCell;
            var activeCellValue = activeCell?.Value2;
            string activeCellText = activeCellValue?.ToString().Trim();

            if (string.IsNullOrEmpty(activeCellText))
            {
                CommonFunctions.GLSenseMessage("Active cell is empty. Please select a cell with a segment value.", MessageBoxIcon.Exclamation, MessageBoxButtons.OK);
                return;
            }

            var repository = new DataRepository();
            var segments = repository.GetSegments(AppState.Instance.SelectedCube.CubeId, AppState.Instance.SelectedLedger.LedgerId);
            var segment = segments?.FirstOrDefault(s => s.SegmentName == AppState.Instance.DefaultSegment);
            if (segment != null)
            {
                var segmentValues = DataRepository.GetSegmentValues(segment);

                if (segmentValues != null)
                {
                    bool isValidValue = segmentValues.Any(v => string.Equals(v.SegmentValue?.Trim(), activeCellText, StringComparison.OrdinalIgnoreCase));
                    if (isValidValue)
                    {
                        SafeInvokeWpf(() =>
                        {
                            var win = new GLSegmentDiscovery();
                            win.ShowDialogWithOwner((IntPtr)AppState.Instance.ExcelApp.Hwnd);
                        });
                    }
                    else
                    {
                        CommonFunctions.GLSenseMessage($"The value in the active cell \"{activeCellText}\" does not match any of the values for the selected segment \"{segment.SegmentName}\" . Please select a cell with a valid segment value.", MessageBoxIcon.Exclamation, MessageBoxButtons.OK);
                    }
                }
                else
                {
                    CommonFunctions.GLSenseMessage($"Failed in fetching segment values for the selected segment \"{segment.SegmentName}\" .", MessageBoxIcon.Exclamation, MessageBoxButtons.OK);
                }
            }
            else
            {
                CommonFunctions.GLSenseMessage("Selected segment not found. Please re-select the segment from the dropdown.", MessageBoxIcon.Exclamation, MessageBoxButtons.OK);
            }
        }

        private void RibSegS_OnChange(object sender, IRibbonControl Control, string text)
        {
            LogUtility.LogDebug($"RibSegS_OnChange fired. Text={text}");
            AppState.Instance.DefaultSegment = string.Empty;
            AppState.Instance.SegmentPickedIndex = -1;

            if (string.IsNullOrEmpty(text))
                return;

            var item = RibSegS.Items.Cast<AddinExpress.MSO.ADXRibbonItem>()
                            .Select((itm, index) => new { Item = itm, Index = index })
                            .FirstOrDefault(x => x.Item.Caption == text);

            if (item != null)
            {
                AppState.Instance.DefaultSegment = text;
                AppState.Instance.SegmentPickedIndex = item.Index;
            }
        }

        private void RibLiveCalc_OnClick(object sender, IRibbonControl control, bool pressed)
        {
            LogUtility.LogDebug($"RibLiveCalc_OnClick clicked. Pressed={RibLiveCalc.Pressed}");
            AppState.Instance.SingleRefresh = RibLiveCalc.Pressed;
        }

        private void RibFSG_OnClick(object sender, IRibbonControl control, bool pressed)
        {
            try
            {
                LogUtility.LogDebug("RibFSG_OnClick clicked.");
                AppState.Instance.displayConfigurator = true;

                AppState.Instance.BalancePane = GetPaneInstance();
                GLConfiguratorPane blpane = AppState.Instance.BalancePane;

                if (blpane != null)
                {
                    blpane.Visible = !blpane.Visible;
                    if (blpane.Visible)
                    {
                        _ = blpane.RelaunchPane();
                    }
                }
                else
                {
                    adxExcelTaskPanesCollectionItem1.Position = AddinExpress.XL.ADXExcelTaskPanePosition.Right;
                    blpane = (GLConfiguratorPane)adxExcelTaskPanesCollectionItem1.CreateTaskPaneInstance();
                    if (blpane != null)
                    {
                        blpane.Show();
                        blpane.Visible = true;
                    }
                    AppState.Instance.BalancePane = blpane;
                }
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex);
            }
            finally
            {
                AppState.Instance.displayConfigurator = false;
            }
        }

        private void RibLOVs_OnClick(object sender, IRibbonControl control, bool pressed)
        {
            LogUtility.LogDebug("RibLOVs_OnClick clicked.");
            SafeInvokeWpf(() =>
            {
                var win = new GLLOVs();
                win.ShowDialogWithOwner((IntPtr)AppState.Instance.ExcelApp.Hwnd);
            });
        }

        private void RibRollerGroup_OnClick(object sender, IRibbonControl control, bool pressed)
        {
            LogUtility.LogDebug("RibRollerGroup_OnClick clicked.");
            SafeInvokeWpf(() =>
            {
                var win = new GLRollerGroups();
                win.ShowDialogWithOwner((IntPtr)AppState.Instance.ExcelApp.Hwnd);
            });
        }

        private void RibAccount_OnClick(object sender, IRibbonControl control, bool pressed)
        {
            LogUtility.LogDebug("RibAccount_OnClick clicked.");
            SafeInvokeWpf(() =>
            {
                var win = new GLSegmentValues();
                win.ShowDialogWithOwner((IntPtr)AppState.Instance.ExcelApp.Hwnd);
            });
        }

        private static bool ValidatePreconditions()
        {
            if (!AppState.Instance.IsLoginCompleted || AppState.Instance.SelectedCube == null)
            {
                CommonFunctions.GLSenseMessage("Please log in first.", MessageBoxIcon.Exclamation);
                return false;
            }
            return true;
        }

        private bool ValidateLedgerChange(string text, out LedgerRecord ledger, out bool shouldContinue, out bool sheetClear)
        {
            ledger = AppState.Instance.SelectedCube.Ledgers.FirstOrDefault(l => l.LedgerName == text);
            if (ledger == null)
            {
                CommonFunctions.GLSenseMessage("Ledger not found.", MessageBoxIcon.Exclamation);
                shouldContinue = false; sheetClear = false;
                return false;
            }

            shouldContinue = true; sheetClear = false;
            if (AppState.Instance.SelectedLedger?.CoaId != ledger.Coaid)
            {
                var result = CommonFunctions.GLSenseMessage("Different chart of account detected. Clear sheet?", MessageBoxIcon.Question, MessageBoxButtons.YesNoCancel);
                shouldContinue = result != MessageBoxResult.Cancel;
                sheetClear = result == MessageBoxResult.Yes;
                if (!shouldContinue) Ribledger.Text = AppState.Instance.SelectedLedger?.LedgerName ?? "";
            }
            return true;
        }

        private async Task PerformLedgerChangeAsync(LedgerRecord ledger, bool sheetClear, CancellationToken token, GLWaitWindow win)
        {
            await Task.Run(async () =>
            {
                try
                {
                    if (sheetClear)
                    {
                        _ = win.Dispatcher.InvokeAsync(() => win.SetProcessMessage("Clearing sheet..."));
                        token.ThrowIfCancellationRequested();
                        if (ExcelApp.ActiveSheet is Excel.Worksheet sheet) sheet.Cells.Clear();
                    }

                    _ = win.Dispatcher.InvokeAsync(() => win.SetProcessMessage("Loading ledger data..."));
                    token.ThrowIfCancellationRequested();

                    await CommonFunctions.FillResponsibilitiesAsync(ledger.LedgerId, AppState.Instance.SelectedCube.CubeId, token);

                    token.ThrowIfCancellationRequested();
                    UpdateRibbonUI(ledger);
                }
                catch (OperationCanceledException)
                {
                    LogUtility.LogDebug("PerformLedgerChangeAsync: operation canceled.");
                    throw;
                }
                catch (Exception ex)
                {
                    LogUtility.LogDebug($"PerformLedgerChangeAsync: rethrowing exception to caller ({ex.Message}).");
                    throw;
                }
            }, token);
        }

        private void UpdateRibbonUI(LedgerRecord ledger)
        {
            var rib = AddinModule.CurrentInstance;
            rib.RibSegS.Items.Clear();

            AppState.Instance.DefaultSegment = string.Empty;
            AppState.Instance.SegmentPickedIndex = -1;

            var repository = new DataRepository();
            foreach (var s in repository.GetSegments(AppState.Instance.SelectedCube.CubeId, ledger.LedgerId))
            {
                rib.RibSegS.Items.Add(new ADXRibbonItem { Caption = s.SegmentName });
            }

            AppState.Instance.SelectedLedger = repository.GetLedgers(AppState.Instance.SelectedCube.CubeId)?.FirstOrDefault(l => l.LedgerId == ledger.LedgerId);
        }

        private async void Ribledger_OnChange(object sender, IRibbonControl Control, string text)
        {
            LogUtility.LogDebug($"Ribledger_OnChange fired. Text={text}");
            if (!ValidatePreconditions()) return;

            using var ctsHelper = new CancellationHelper();
            var token = ctsHelper.GetToken();
            GLWaitWindow win = null;

            try
            {
                if (!ValidateLedgerChange(text, out var ledger, out var shouldContinue, out var sheetClear)) return;

                win = CreateAndShowWaitWindow(ctsHelper);

                if (shouldContinue)
                {
                    await PerformLedgerChangeAsync(ledger, sheetClear, token, win);
                    AppState.Instance.BalancePane = GetPaneInstance();
                    if (AppState.Instance.BalancePane != null && AppState.Instance.BalancePane.Visible)
                    {
                        _ = AppState.Instance.BalancePane.RelaunchPane();
                    }
                }
            }
            catch (OperationCanceledException)
            {
                LogUtility.LogWarn("Ledger change operation cancelled by user!");
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex);
                if (win != null) await win.Dispatcher.InvokeAsync(() => win.RequestClose());
                CommonFunctions.GLSenseMessage(ex.Message, MessageBoxIcon.Error, MessageBoxButtons.OK);
            }
            finally
            {
                if (win != null) await win.Dispatcher.InvokeAsync(() => win.RequestClose());
            }
        }

        private void RibGetCube_OnClick(object sender, IRibbonControl control, bool pressed)
        {
            LogUtility.LogDebug("RibGetCube_OnClick clicked.");
            SafeInvokeWpf(() =>
            {
                var win = new GLCubeDetails();
                win.ShowDialogWithOwner((IntPtr)AppState.Instance.ExcelApp.Hwnd);
            });
        }

        private void RibDBL1_OnAction(object sender, IRibbonControl control, bool pressed)
        {
            LogUtility.LogDebug("RibDBL1_OnAction clicked.");
            SafeInvokeWpf(() =>
            {
                var win = new GLLoginDetails();
                win.ShowDialogWithOwner((IntPtr)AppState.Instance.ExcelApp.Hwnd);
            });
        }

        private void RibLogin_OnClick(object sender, IRibbonControl control, bool pressed)
        {
            LogUtility.LogDebug("RibLogin_OnClick clicked.");
            SafeInvokeWpf(() =>
            {
                AppState.Instance.LoginToken = string.Empty;

                // Original login UI
                var win = new GLLogin();
                win.ShowDialogWithOwner((IntPtr)AppState.Instance.ExcelApp.Hwnd);
            });
        }

        private void RibLogout_OnClick(object sender, IRibbonControl control, bool pressed)
        {
            LogUtility.LogDebug("RibLogout_OnClick clicked.");
            try
            {
                object edgeAddin = AddinModule.GetEdgeAddinInstance();
                edgeAddin?.GetType().InvokeMember("LogoffFromAddin", BindingFlags.InvokeMethod, null, edgeAddin, new object[] { });
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex, "Exception while logging out from XLEdge via GLSense!");
            }
            GLSenseLogout();
        }

        private void GLSenseLogout()
        {
            using var ctsHelper = new CancellationHelper();
            var token = ctsHelper.GetToken();

            string message = string.Empty;
            MessageBoxIcon icon = MessageBoxIcon.None;
            try
            {
                HideTaskPanes();
                (message, icon) = Task.Run(() => GetLogoutResponseAsync(token)).GetAwaiter().GetResult();
            }
            catch (OperationCanceledException)
            {
                LogUtility.LogWarn("Logout operation was canceled by the user.");
                message = "Logout was canceled.";
                icon = MessageBoxIcon.Warning;
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex);
                message = "An unexpected error occurred during logout.";
                icon = MessageBoxIcon.Error;
            }
            finally
            {
                LoggOff();
                CommonFunctions.GLSenseMessage(message, icon);
            }
        }

        private static async Task<(string Message, MessageBoxIcon Icon)> GetLogoutResponseAsync(CancellationToken token)
        {
            var apiUrl = $"{AppState.Instance.LoginUrl}{AppConstants.WebSecure}applogout";

            LogUtility.LogDebug(apiUrl);

            var result = await ApiHelper.ServerAPI(apiUrl, "Form", "", "GET", token);

            LogUtility.LogDebug($"Logout API response: {result}");
            token.ThrowIfCancellationRequested();
            var logout = ApiResponseHelper.Parse<JsonElement>(result, JsonGlobals.Options);
            if (!logout.IsSuccess) return (logout.ErrorMessage, MessageBoxIcon.Error);
            string message = logout.Value.TryGetProperty("message", out var msgEl) ? msgEl.GetString() : "Successfully logged out.";
            return (message ?? "Successfully logged out.", MessageBoxIcon.Information);
        }

        private void LoggOff()
        {
            AppState.Instance.Reset();
            Ribledger.Items.Clear();
            Ribledger.Text = string.Empty;
            RibGetCube.Caption = "Cube:  Select cube";
            RibbonHelper.ApplyState("LoggedOut");
        }

        private void HideTaskPanes()
        {
            try
            {
                if (adxExcelTaskPanesCollectionItem1?.TaskPaneInstances == null) return;
                if (adxExcelTaskPanesCollectionItem1.TaskPaneInstances.Count > 0)
                {
                    foreach (GLConfiguratorPane xlTaskpane in adxExcelTaskPanesCollectionItem1.TaskPaneInstances)
                    {
                        xlTaskpane.Visible = false;
                    }
                }
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex);
            }
        }
        

        /// <summary>
        /// The below method is executed from XLEdge Add-in
        /// </summary>
        public void LogoutSession()
        {
            try
            {
                LoggOff();
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex, "Exception in logout method called from XLEdge!");
            }
        }
        /// <summary>
        /// The below method is executed from XLEdge Add-in
        /// </summary>
        public void GetGLCubeInformation(string token, string url, string userName)
        {


            if (string.IsNullOrWhiteSpace(token) || string.IsNullOrWhiteSpace(url))
                return;

            try
            {
                AppState.Instance.LoginToken = token;
                AppState.Instance.LoginUrl = url;

                string apiUrl = $"{AppState.Instance.LoginUrl}{AppConstants.RestSecure}finance-cubes";

                string apiResponse = ExecuteApiCallWithTimeout(apiUrl);

                if (string.IsNullOrWhiteSpace(apiResponse))
                {
                    LogUtility.LogWarn("Received an empty response when extracting cubes.");
                    return;
                }

                var result = ApiResponseHelper.Parse<List<CubeRecord>>(
                        apiResponse,
                        JsonGlobals.Options);

                if (!result.IsSuccess)
                {
                    AddinModule.RibbonHelper.ApplyState("NoCubes");
                    LogUtility.LogWarn(apiResponse);
                    return;
                }

                CubeCache.AllCubes =
                    result.Value!.OrderBy(c => c.CubeName).ToList();

                if (CubeCache.AllCubes.Count == 0)
                {
                    LogUtility.LogWarn($"No cubes found for the user \"{userName}\" logged in from XLEdge.");
                    AddinModule.RibbonHelper.ApplyState("NoCubes");
                    return;
                }

                CubeDataRepository.InsertCubeDataAsync().GetAwaiter().GetResult();

                AppState.Instance.LoginUserName = userName;

                AddinModule.RibbonHelper.ApplyState("PartialLoggedIn");

                LogExcelState("After processing");
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex, "Exception in GetGLCubeInformation");
            }
        }
        private void LogExcelState(string stage)
        {
            try
            {
                var excel = this.ExcelApp;

                if (excel != null)
                {
                    try
                    {
                        // Try to access a property to see if it's alive
                        var name = excel.Name;

                        // Check active workbook/sheet
                        if (excel.ActiveWorkbook != null)
                        {
                            LogUtility.LogWarn($"{stage} - ActiveWorkbook: {excel.ActiveWorkbook.Name}");
                        }
                    }
                    catch (Exception ex)
                    {
                        LogUtility.LogWarn($"{stage} - ExcelApp accessible: false, error: {ex.Message}");
                    }
                }
            }
            catch (Exception ex)
            {
                LogUtility.LogWarn($"{stage} - Error logging Excel state: {ex.Message}");
            }
        }
        private static string ExecuteApiCallWithTimeout(string apiUrl)
        {
            using var cts = new CancellationTokenSource(TimeSpan.FromSeconds(1800));
            try
            {
                Task<string> apiTask = Task.Run(async () => await ApiHelper.ServerAPI(apiUrl, "Form", "", "POST", cts.Token));
                if (apiTask.Wait(TimeSpan.FromSeconds(1800))) return apiTask.Result;

                LogUtility.LogWarn("API timeout - Wait limit reached.");
                return null;
            }
            catch (AggregateException ex) when (ex.InnerException is OperationCanceledException)
            {
                LogUtility.LogError($"API call timed out: {ex.InnerException?.Message}");
                return null;
            }
            catch (AggregateException ex) when (ex.InnerException is HttpRequestException)
            {
                LogUtility.LogError($"HTTP error: {ex.InnerException?.Message}");
                return null;
            }
            catch (AggregateException ex)
            {
                LogUtility.LogError($"API error: {ex.InnerException?.Message}");
                return null;
            }
        }
        private async void RibHideRows_OnClick(object sender, IRibbonControl control, bool pressed)
        {
            LogUtility.LogDebug("RibHideRows_OnClick clicked.");
            var processor = new HideRowProcessor();
            await processor.ExecuteAsync("Hiding Rows");
        }

        private async void RibUnHideRows_OnClick(object sender, IRibbonControl control, bool pressed)
        {
            LogUtility.LogDebug("RibUnHideRows_OnClick clicked.");
            var processor = new UnhideRowProcessor();
            await processor.ExecuteAsync("Unhiding Rows");
        }

        public abstract class RowProcessor
        {
            public async Task ExecuteAsync(string operationName)
            {
                LogUtility.LogDebug($"RowProcessor.ExecuteAsync started. Operation={operationName}");
                CommonMethods.DisableExcelSettings();
                GLWaitWindow win = null;
                using var ctsHelper = new CancellationHelper();
                CancellationToken token = ctsHelper.GetToken();
                try
                {
                    if (!GuardLoginAndExcel()) return;
                    win = CreateAndShowWaitWindow(ctsHelper);
                    await InitializeWaitWindowAsync(win, operationName, operationName == "Hiding Rows" ? "Hiding rows for 0 balances…" : "Unhiding rows please wait...");
                    await Task.Yield();
                    Excel.Range selection = GetSelection();
                    Excel.Worksheet sheet = selection.Worksheet;
                    Excel.Range balanceRange = await BalancesRangeAsync(selection);
                    if (balanceRange == null)
                    {
                        await SafelyCloseWaitWindowAsync(win);
                        CommonFunctions.GLSenseMessage("No balance formula's in the selection!", MessageBoxIcon.Exclamation, MessageBoxButtons.OK);
                        return;
                    }
                    await ProcessRowsCoreAsync(sheet, balanceRange, win, token);
                }
                catch (OperationCanceledException) { LogUtility.LogError("Operation cancelled by user."); }
                catch (Exception ex) { LogUtility.LogException(ex); }
                finally
                {
                    await SafelyCloseWaitWindowAsync(win);
                    CommonMethods.EnableExcelSettings();
                }
            }
            protected abstract Task ProcessRowsCoreAsync(Excel.Worksheet sheet, Excel.Range selection, GLWaitWindow win, CancellationToken token);
        }

        public sealed class HideRowProcessor : RowProcessor
        {
            protected override async Task ProcessRowsCoreAsync(Excel.Worksheet sheet, Excel.Range selection, GLWaitWindow win, CancellationToken token)
            {
                token.ThrowIfCancellationRequested();
                var (formulas, values) = await GetFormulaAndValueArraysAsync(selection);
                if (formulas == null || values == null) return;
                token.ThrowIfCancellationRequested();
                var hideRows = FindHideRows(formulas, values, selection.Row);
                await ProcessHideRowsAsync(sheet, hideRows, win, token);
            }
        }

        public sealed class UnhideRowProcessor : RowProcessor
        {
            protected override async Task ProcessRowsCoreAsync(Excel.Worksheet sheet, Excel.Range selection, GLWaitWindow win, CancellationToken token)
            {
                token.ThrowIfCancellationRequested();
                var (formulas, values) = await GetFormulaAndValueArraysAsync(selection);
                if (formulas == null || values == null) return;
                token.ThrowIfCancellationRequested();
                var rowsToUnhide = FindHideRows(formulas, values, selection.Row);
                if (rowsToUnhide.Count == 0) { await MessageWaitWindowAsync(win, "Nothing to unhide in the current selection."); return; }
                double standardHeight = sheet.StandardHeight;
                await ProcessUnhideRowsByBatchesAsync(sheet, rowsToUnhide, standardHeight, win, token);
            }
        }

        private static List<int> FindHideRows(object[,] formulas, object[,] values, int startRow)
        {
            var hideRows = new List<int>();
            int rLo = formulas.GetLowerBound(0), rHi = formulas.GetUpperBound(0), cLo = formulas.GetLowerBound(1), cHi = formulas.GetUpperBound(1);
            if (values.GetLength(0) != formulas.GetLength(0) || values.GetLength(1) != formulas.GetLength(1))
                throw new InvalidOperationException("Formulas and values arrays have different shapes.");
            for (int r = rLo; r <= rHi; r++)
            {
                if (ShouldHideRow(formulas, values, r, cLo, cHi)) hideRows.Add(startRow + (r - rLo));
            }
            return hideRows;
        }

        private static bool ShouldHideRow(object[,] formulas, object[,] values, int r, int cLo, int cHi)
        {
            for (int c = cLo; c <= cHi; c++)
            {
                if (IsGetBalanceFormula(formulas[r, c]) && IsZero(values[r, c])) return true;
            }
            return false;
        }

        private static bool IsGetBalanceFormula(object f)
        {
            if (f == null) return false;
            var s = f as string ?? f.ToString();
            if (string.IsNullOrWhiteSpace(s)) return false;
            return s.TrimStart('=', '@').IndexOf(AppConstants.glBal, StringComparison.OrdinalIgnoreCase) >= 0;
        }

        private static bool IsZero(object value) => value switch
        {
            null => true,
            double d => Math.Abs(d) < 1e-9,
            int i => i == 0,
            decimal m => m == 0m,
            string s => double.TryParse(s, out double parsed) && Math.Abs(parsed) < 1e-9,
            _ => false,
        };

        private static bool GuardLoginAndExcel() => AppState.Instance.IsLoginCompleted && AppState.Instance.ExcelApp != null;

        private static Excel.Range GetSelection() => AppState.Instance.ExcelApp.Selection as Excel.Range ?? throw new InvalidOperationException("No selection available");

        private static async Task<(object[,] formulas, object[,] values)> GetFormulaAndValueArraysAsync(Excel.Range selection)
        {
            await Task.Yield();
            return (CoerceTo2D(selection.Formula), CoerceTo2D(selection.Value2));
        }

        private static object[,] CoerceTo2D(object value) => value switch
        {
            object[,] array2d => array2d,
            null => new object[1, 1] { { null! } },
            _ => new object[1, 1] { { value } }
        };

        private static async Task<Excel.Range> BalancesRangeAsync(Excel.Range selection)
        {
            string rngAddress = ExcelExternalRef.BuildExternalAddress(selection);
            Excel.Range totalRange = CommonFunctions.GetBalanceTotalRange(rngAddress);
            if (totalRange != null) return totalRange;
            await Task.Yield();
            return null;
        }

        private static async Task ProcessHideRowsAsync(Excel.Worksheet sheet, List<int> hideRows, GLWaitWindow win, CancellationToken token)
        {
            if (hideRows == null || hideRows.Count == 0) return;
            hideRows.Sort();
            int i = 0;
            while (i < hideRows.Count)
            {
                token.ThrowIfCancellationRequested();
                int start = hideRows[i], end = start;
                i++;
                while (i < hideRows.Count && hideRows[i] == end + 1) { end = hideRows[i]; i++; }
                sheet.Range[$"{start}:{end}"].RowHeight = 0.1;
                await MessageWaitWindowAsync(win, $"Hid rows {start}:{end}…");
                await Task.Yield();
            }
        }

        private static async Task ProcessUnhideRowsByBatchesAsync(Excel.Worksheet sheet, List<int> unhideRows, double standardHeight, GLWaitWindow win, CancellationToken token)
        {
            if (unhideRows == null || unhideRows.Count == 0) return;
            unhideRows.Sort();
            int i = 0;
            while (i < unhideRows.Count)
            {
                token.ThrowIfCancellationRequested();
                int start = unhideRows[i], end = start;
                i++;
                while (i < unhideRows.Count && unhideRows[i] == end + 1) { end = unhideRows[i]; i++; }
                var rng = sheet.Rows[$"{start}:{end}"] as Excel.Range;
                if (rng != null) rng.RowHeight = standardHeight;
                await MessageWaitWindowAsync(win, $"Unhid rows {start}:{end}…");
                await Task.Yield();
            }
            await MessageWaitWindowAsync(win, "Completed successfully.");
        }

        private static GLWaitWindow CreateAndShowWaitWindow(CancellationHelper cts)
        {
            try
            {
                return WpfAppManager.InvokeOnWpfThread(() =>
                {
                    try
                    {
                        var win = new GLWaitWindow(cts);
                        win.ShowWithOwner((IntPtr)AppState.Instance.ExcelApp.Hwnd);
                        win.StartMonitoring();
                        return win;
                    }
                    catch (Exception ex) { LogUtility.LogException(ex); return null; }
                });
            }
            catch (Exception ex) { LogUtility.LogException(ex); return null; }
        }

        private static Task InitializeWaitWindowAsync(GLWaitWindow _win, string title, string message)
        {
            if (_win == null || _win.Dispatcher == null) return Task.CompletedTask;
            try
            {
                return _win.Dispatcher.InvokeAsync(() => { _win.SetProcessTitle(title); _win.SetProcessMessage(message); }, DispatcherPriority.Normal).Task;
            }
            catch (TaskCanceledException) { return Task.CompletedTask; }
            catch (Exception ex) { LogUtility.LogException(ex); return Task.CompletedTask; }
        }

        private static async Task SafelyCloseWaitWindowAsync(GLWaitWindow _win)
        {
            if (_win == null) return;
            try
            {
                if (_win.Dispatcher.CheckAccess()) _win.RequestClose();
                await _win.Dispatcher.InvokeAsync(() => _win.RequestClose());
            }
            catch (Exception ex) { LogUtility.LogError($"Error closing wait window: {ex.Message}"); }
        }

        private static Task MessageWaitWindowAsync(GLWaitWindow _win, string message)
        {
            if (_win == null || _win.Dispatcher == null) return Task.CompletedTask;
            try
            {
                return _win.Dispatcher.InvokeAsync(() => _win.SetProcessMessage(message), DispatcherPriority.Normal).Task;
            }
            catch (TaskCanceledException) { return Task.CompletedTask; }
            catch (Exception ex) { LogUtility.LogException(ex); return Task.CompletedTask; }
        }

        private static async Task ShowErrorMessageAsync(GLWaitWindow _win, string message)
        {
            await SafelyCloseWaitWindowAsync(_win);
            CommonFunctions.GLSenseMessage(message, MessageBoxIcon.Error, MessageBoxButtons.OK);
        }
        private void RibSnapWorksheet_OnClick(object sender, IRibbonControl control, bool pressed)
        {
            LogUtility.LogDebug($"RibSnapWorksheet_OnClick clicked. Pressed={pressed}");
            if (!AppState.Instance.IsLoginCompleted || AppState.Instance.SelectedCube == null)
            {
                return;
            }

            RibSnapWorkbook.Pressed = !pressed;
        }

        private void RibSnapWorkbook_OnClick(object sender, IRibbonControl control, bool pressed)
        {
            LogUtility.LogDebug($"RibSnapWorkbook_OnClick clicked. Pressed={pressed}");
            if (!AppState.Instance.IsLoginCompleted || AppState.Instance.SelectedCube == null)
            {
                return;
            }

            RibSnapWorksheet.Pressed = !pressed;

        }
    }
}

