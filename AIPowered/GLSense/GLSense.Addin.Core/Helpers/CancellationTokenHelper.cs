// CancellationTokenHelper.cs in GLSense.Addin.Core
// Port of GLSense\Helpers\CancellationTokenHelper.cs (FinalWorkingCode).
// Changes: LogUtility.* -> ServiceLocator.Logger.*. Not to be confused with
// CancellationHelper.cs (also in this folder) - that's the CancellationTokenSource
// wrapper instance; this is a set of static utility functions over a plain
// CancellationToken.
using GLSense.Addin.Core.Infrastructure;
using System;
using System.Runtime.CompilerServices;
using System.Threading;
using System.Threading.Tasks;

namespace GLSense.Addin.Core.Helpers
{
    /// <summary>
    /// Helper for cancellation token operations with detailed logging
    /// </summary>
    public static class CancellationTokenHelper
    {
        /// <summary>
        /// Throws if cancellation requested with detailed logging
        /// </summary>
        public static void ThrowIfCancelled(
            CancellationToken token,
            [CallerMemberName] string operationName = "")
        {
            if (token.IsCancellationRequested)
            {
                ServiceLocator.Logger?.LogDebug($"Cancellation requested, throwing in: {operationName}");
                token.ThrowIfCancellationRequested();
            }
        }

        /// <summary>
        /// Checks if cancellation is requested and logs
        /// </summary>
        public static bool IsCancellationRequested(
            CancellationToken token,
            [CallerMemberName] string operationName = "")
        {
            if (token.IsCancellationRequested)
            {
                ServiceLocator.Logger?.LogWarn($"Cancellation detected in: {operationName}");
                return true;
            }

            return false;
        }

        /// <summary>
        /// Waits asynchronously with cancellation support and logging
        /// </summary>
        public static async Task DelayWithLogging(
            int milliseconds,
            CancellationToken token,
            string reason = "")
        {
            try
            {
                ServiceLocator.Logger?.LogDebug($"Waiting {milliseconds}ms{(string.IsNullOrEmpty(reason) ? "" : $" - {reason}")}");

                await Task.Delay(milliseconds, token);
            }
            catch (OperationCanceledException)
            {
                ServiceLocator.Logger?.LogWarn($"Wait cancelled after delay{(string.IsNullOrEmpty(reason) ? "" : $": {reason}")}");
                throw;
            }
        }
    }
}
