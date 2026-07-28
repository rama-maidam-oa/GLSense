using GLSense.Utilities;
using System;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Threading;

namespace GLSense.Helpers
{
    /// <summary>
    /// Centralized UI overlay helper to reduce repeated code across windows
    /// </summary>
    public static class OverlayHelper
    {
        /// <summary>
        /// Shows busy overlay with cancellation support and detailed logging
        /// </summary>
        public static async Task ShowBusyOverlayAsync(
            Dispatcher dispatcher,
            Views.AppOverlay overlayControl,
            CancellationHelper helper,
            string message)
        {
            using (new LogUtility.LogScope($"ShowBusyOverlay: {message}"))
            {
                try
                {
                    LogUtility.LogDebug($"Showing busy overlay: {message}");

                    await dispatcher.InvokeAsync(() =>
                    {
                        overlayControl.ShowBusyasyn(
                            message: message + " (Click cancel to stop)",
                            cancelAction: async () =>
                            {
                                if (!helper.IsCancellationRequested)
                                {
                                    helper.Cancel();
                                }
                                await Task.Delay(80);
                            }
                        );
                    }, DispatcherPriority.Background);
                }
                catch (Exception ex)
                {
                    ExceptionHelper.LogDetailedException(ex, $"ShowBusyOverlayAsync: {message}");
                }
            }
        }

        /// <summary>
        /// Executes an action with busy overlay and automatic cleanup
        /// </summary>
        public static async Task ExecuteWithBusyOverlay(
            Dispatcher dispatcher,
            Views.AppOverlay overlayControl,
            string message,
            Func<CancellationHelper, Task> action)
        {
            using (new LogUtility.LogScope($"ExecuteWithBusyOverlay: {message}"))
            {
                var helper = new CancellationHelper();

                try
                {
                    await ShowBusyOverlayAsync(dispatcher, overlayControl, helper, message);
                    LogUtility.LogDebug($"Starting action: {message}");

                    await action(helper);

                    LogUtility.LogDebug($"Action completed: {message}");
                }
                catch (OperationCanceledException)
                {
                    LogUtility.LogWarn($"Operation cancelled: {message}");
                    throw;
                }
                catch (Exception ex)
                {
                    ExceptionHelper.LogDetailedException(ex, $"ExecuteWithBusyOverlay: {message}");
                    throw;
                }
                finally
                {
                    await HideBusyOverlayAsync(dispatcher, overlayControl);
                }
            }
        }

        /// <summary>
        /// Hides busy overlay with logging
        /// </summary>
        public static async Task HideBusyOverlayAsync(
            Dispatcher dispatcher,
            Views.AppOverlay overlayControl)
        {
            try
            {
                LogUtility.LogDebug("Hiding busy overlay");

                await dispatcher.InvokeAsync(async () =>
                {
                    await overlayControl.HideBusyAsync();
                }, DispatcherPriority.Background);
            }
            catch (Exception ex)
            {
                ExceptionHelper.LogDetailedException(ex, "HideBusyOverlayAsync");
            }
        }

        /// <summary>
        /// Shows warning message with logging
        /// </summary>
        public static void ShowWarning(
            Dispatcher dispatcher,
            Views.AppOverlay overlayControl,
            string message)
        {
            try
            {
                LogUtility.LogWarn($"Warning shown to user: {message}");

                dispatcher.Invoke(() =>
                {
                    overlayControl.ShowWarning(message);
                });
            }
            catch (Exception ex)
            {
                ExceptionHelper.LogDetailedException(ex, "ShowWarning");
            }
        }

        /// <summary>
        /// Shows warning message asynchronously with logging
        /// </summary>
        public static async Task ShowWarningAsync(
            Dispatcher dispatcher,
            Views.AppOverlay overlayControl,
            string message)
        {
            try
            {
                LogUtility.LogWarn($"Warning shown to user: {message}");

                await dispatcher.InvokeAsync(() =>
                {
                    overlayControl.ShowWarning(message);
                }, DispatcherPriority.Background);
            }
            catch (Exception ex)
            {
                ExceptionHelper.LogDetailedException(ex, "ShowWarningAsync");
            }
        }

        /// <summary>
        /// Shows info message with logging
        /// </summary>
        public static void ShowInfo(
            Dispatcher dispatcher,
            Views.AppOverlay overlayControl,
            string message)
        {
            try
            {

                dispatcher.Invoke(() =>
                {
                    overlayControl.ShowInfo(message);
                });
            }
            catch (Exception ex)
            {
                ExceptionHelper.LogDetailedException(ex, "ShowInfo");
            }
        }

        /// <summary>
        /// Shows info message asynchronously with logging
        /// </summary>
        public static async Task ShowInfoAsync(
            Dispatcher dispatcher,
            Views.AppOverlay overlayControl,
            string message)
        {
            try
            {

                await dispatcher.InvokeAsync(() =>
                {
                    overlayControl.ShowInfo(message);
                }, DispatcherPriority.Background);
            }
            catch (Exception ex)
            {
                ExceptionHelper.LogDetailedException(ex, "ShowInfoAsync");
            }
        }

        /// <summary>
        /// Shows error message with logging
        /// </summary>
        public static void ShowError(
            Dispatcher dispatcher,
            Views.AppOverlay overlayControl,
            string message)
        {
            try
            {
                LogUtility.LogError($"Error shown to user: {message}");

                dispatcher.Invoke(() =>
                {
                    overlayControl.ShowError(message);
                });
            }
            catch (Exception ex)
            {
                ExceptionHelper.LogDetailedException(ex, "ShowError");
            }
        }

        /// <summary>
        /// Shows error message asynchronously with logging
        /// </summary>
        public static async Task ShowErrorAsync(
            Dispatcher dispatcher,
            Views.AppOverlay overlayControl,
            string message)
        {
            try
            {
                LogUtility.LogError($"Error shown to user: {message}");

                await dispatcher.InvokeAsync(async () =>
                {
                    await overlayControl.ShowErrorAsync(message);
                }, DispatcherPriority.Background);
            }
            catch (Exception ex)
            {
                ExceptionHelper.LogDetailedException(ex, "ShowErrorAsync");
            }
        }
    }
}
