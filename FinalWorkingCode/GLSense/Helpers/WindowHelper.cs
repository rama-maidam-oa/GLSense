using GLSense.Utilities;
using System;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Threading;
using Excel = Microsoft.Office.Interop.Excel;

namespace GLSense.Helpers
{
    /// <summary>
    /// Centralized helper for common window operations to reduce repeated code
    /// </summary>
    public static class WindowHelper
    {
        /// <summary>
        /// Gets the active cell address in A1 notation with sheet name
        /// </summary>
        public static string GetActiveCellAddress(Excel.Application excelApp)
        {
            using (new LogUtility.LogScope("GetActiveCellAddress"))
            {
                try
                {
                    LogUtility.LogDebug("Getting active cell address");

                    var rng = excelApp?.ActiveCell;
                    if (rng == null)
                    {
                        LogUtility.LogWarn("Active cell is null");
                        return string.Empty;
                    }

                    string sheetName = ((Excel.Worksheet)rng.Parent).Name;
                    string cellAddress = rng.Address[true, true, Excel.XlReferenceStyle.xlA1, false];
                    string fullAddress = $"'{sheetName}'!{cellAddress}";

                    LogUtility.LogDebug($"Active cell address: {fullAddress}");

                    return fullAddress;
                }
                catch (Exception ex)
                {
                    ExceptionHelper.LogDetailedException(ex, "GetActiveCellAddress");
                    return string.Empty;
                }
            }
        }

        /// <summary>
        /// Sets Excel window as owner of a WPF window
        /// </summary>
        public static void SetExcelAsOwner(Window window, Excel.Application excelApp)
        {
            try
            {
                if (window == null || excelApp == null)
                {
                    LogUtility.LogWarn("Cannot set Excel as owner - null parameters");
                    return;
                }

                LogUtility.LogDebug($"Setting Excel as owner for window: {window.GetType().Name}");

                var excelHandle = new IntPtr(excelApp.Hwnd);
                var helper = new System.Windows.Interop.WindowInteropHelper(window);
                helper.Owner = excelHandle;
                window.WindowStartupLocation = WindowStartupLocation.CenterOwner;

                LogUtility.LogDebug("Excel set as window owner successfully");
            }
            catch (Exception ex)
            {
                ExceptionHelper.LogDetailedException(ex, "SetExcelAsOwner");
                window.WindowStartupLocation = WindowStartupLocation.CenterScreen;
            }
        }

        /// <summary>
        /// Validates Excel application and login status
        /// </summary>
        public static bool ValidateExcelAndLogin()
        {
            using (new LogUtility.LogScope("ValidateExcelAndLogin"))
            {
                bool excelValid = AppState.Instance.ExcelApp != null;
                bool loginValid = AppState.Instance.IsLoginCompleted;

                LogUtility.LogDebug($"Excel available: {excelValid}, Login completed: {loginValid}");

                if (!excelValid)
                {
                    LogUtility.LogWarn("Excel application is not available");
                }

                if (!loginValid)
                {
                    LogUtility.LogWarn("User is not logged in");
                }

                return excelValid && loginValid;
            }
        }

        /// <summary>
        /// Invokes action on dispatcher with error handling
        /// </summary>
        public static async Task InvokeAsync(
            Dispatcher dispatcher,
            Action action,
            DispatcherPriority priority = DispatcherPriority.Normal)
        {
            try
            {
                if (dispatcher == null || action == null)
                {
                    LogUtility.LogWarn("InvokeAsync called with null dispatcher or action");
                    return;
                }

                await dispatcher.InvokeAsync(action, priority);
            }
            catch (Exception ex)
            {
                ExceptionHelper.LogDetailedException(ex, "InvokeAsync");
            }
        }

        /// <summary>
        /// Invokes async function on dispatcher with error handling
        /// </summary>
        public static async Task InvokeAsync(
            Dispatcher dispatcher,
            Func<Task> action,
            DispatcherPriority priority = DispatcherPriority.Normal)
        {
            try
            {
                if (dispatcher == null || action == null)
                {
                    LogUtility.LogWarn("InvokeAsync called with null dispatcher or action");
                    return;
                }

                await dispatcher.InvokeAsync(action, priority);
            }
            catch (Exception ex)
            {
                ExceptionHelper.LogDetailedException(ex, "InvokeAsync");
            }
        }

        /// <summary>
        /// Loads window with standard initialization pattern
        /// </summary>
        public static async Task StandardWindowLoad(
            Dispatcher dispatcher,
            Views.AppOverlay overlayControl,
            Func<Task> loadAction,
            string operationName = "Loading")
        {
            using (new LogUtility.LogScope($"StandardWindowLoad: {operationName}"))
            {
                var helper = new CancellationHelper();

                try
                {
                    await OverlayHelper.ShowBusyOverlayAsync(dispatcher, overlayControl, helper, operationName);

                    LogUtility.LogDebug($"Executing load action: {operationName}");

                    await loadAction();

                    LogUtility.LogDebug($"Load action completed: {operationName}");
                }
                catch (OperationCanceledException)
                {
                    LogUtility.LogWarn($"Load operation cancelled: {operationName}");
                }
                catch (Exception ex)
                {
                    ExceptionHelper.LogDetailedException(ex, $"StandardWindowLoad failed: {operationName}");
                    await OverlayHelper.ShowErrorAsync(dispatcher, overlayControl, ExceptionHelper.GetFriendlyErrorMessage(ex));
                }
                finally
                {
                    await OverlayHelper.HideBusyOverlayAsync(dispatcher, overlayControl);
                }
            }
        }
    }
}
