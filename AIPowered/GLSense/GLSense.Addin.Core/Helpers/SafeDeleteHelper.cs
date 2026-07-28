// SafeDeleteHelper.cs in GLSense.Addin.Core
// Ported from GLSense\Helpers\SafeDeleteHelper.cs (FinalWorkingCode), including the
// QuarantineDeleteHelper class that lives in the same file in the original project.
// Group F (transitive dependency of BalanceRefresh.ClearFiles/TryDeleteFileWithFallbackAsync,
// used to best-effort delete the working refresh/zip files after a bulk refresh).
// Re-pointed vs. the original: LogUtility.* -> ServiceLocator.Logger?.*.
using GLSense.Addin.Core.Infrastructure;
using System;
using System.IO;
using System.Threading;
using System.Threading.Tasks;

namespace GLSense.Addin.Core.Helpers
{
    public static class SafeDeleteHelper
    {
        /// <summary>
        /// Attempts to delete the file. If locked, it retries until timeout elapses.
        /// Returns true if deleted or file doesn't exist; false if not deleted within timeout.
        /// Throws on invalid arguments or if cancellation requested.
        /// </summary>
        public static async Task<bool> TryDeleteFileAsync(
            string path,
            TimeSpan timeout,
            CancellationToken ct = default,
            int retryIntervalMs = 300)
        {
            ServiceLocator.Logger?.LogDebug($"SafeDeleteHelper.TryDeleteFileAsync: attempting to delete '{path}' (timeout: {timeout})");
            ValidatePath(path);

            var startTime = DateTime.UtcNow;

            while (!ShouldStopRetrying(startTime, timeout))
            {
                ct.ThrowIfCancellationRequested();

                if (TryDeleteFileOnce(path, ct))
                {
                    ServiceLocator.Logger?.LogDebug($"SafeDeleteHelper.TryDeleteFileAsync: '{path}' deleted (or already absent).");
                    return true;
                }

                await WaitForRetryAsync(retryIntervalMs, ct);
            }

            ServiceLocator.Logger?.LogWarn($"SafeDeleteHelper.TryDeleteFileAsync: timed out deleting '{path}' after {timeout}.");
            return false;
        }
        private static void ValidatePath(string path)
        {
            if (string.IsNullOrWhiteSpace(path))
                throw new ArgumentException("Path is required.", nameof(path));
        }

        private static bool ShouldStopRetrying(DateTime startTime, TimeSpan timeout)
        {
            return DateTime.UtcNow - startTime >= timeout;
        }

        private static bool TryDeleteFileOnce(string path, CancellationToken ct)
        {
            ct.ThrowIfCancellationRequested();

            if (!File.Exists(path))
                return true; // File doesn't exist, consider it deleted

            if (FileLockHelper.IsFileLocked(path))
                return false; // File is locked, can't delete now

            return AttemptDelete(path, ct);
        }

        private static bool AttemptDelete(string path, CancellationToken ct)
        {
            try
            {
                AttemptDirectDelete(path);
                return true;
            }
            catch (IOException ex)
            {
                // File may still be in use; will retry
                ServiceLocator.Logger?.LogWarn($"AttemptDelete: IO exception deleting '{path}' (will retry): {ex.Message}");
                return false;
            }
            catch (UnauthorizedAccessException ex)
            {
                ServiceLocator.Logger?.LogWarn($"AttemptDelete: unauthorized deleting '{path}', attempting to remove read-only attribute: {ex.Message}");
                return HandleUnauthorizedAccess(path, ct);
            }
            catch (Exception ex)
            {
                // Other exceptions - retry on next iteration
                ServiceLocator.Logger?.LogWarn($"AttemptDelete: unexpected exception deleting '{path}' (will retry): {ex.Message}");
                return false;
            }
        }

        private static void AttemptDirectDelete(string path)
        {
            File.Delete(path);
        }

        private static bool HandleUnauthorizedAccess(string path, CancellationToken ct)
        {
            TryRemoveReadOnlyAttribute(path);

            try
            {
                ct.ThrowIfCancellationRequested();
                File.Delete(path);
                return true;
            }
            catch (Exception ex)
            {
                // Still failing after removing read-only attribute
                ServiceLocator.Logger?.LogWarn($"HandleUnauthorizedAccess: failed to delete '{path}' after removing read-only attribute: {ex.Message}");
                return false;
            }
        }

        private static void TryRemoveReadOnlyAttribute(string path)
        {
            try
            {
                var fileInfo = new FileInfo(path);
                if (fileInfo.Exists && fileInfo.IsReadOnly)
                {
                    fileInfo.IsReadOnly = false;
                }
            }
            catch (Exception ex)
            {
                ServiceLocator.Logger?.LogWarn($"TryRemoveReadOnlyAttribute: failed to modify attributes for '{path}' (non-fatal): {ex.Message}");
            }
        }

        private static async Task WaitForRetryAsync(int retryIntervalMs, CancellationToken ct)
        {
            await Task.Delay(retryIntervalMs, ct).ConfigureAwait(false);
        }
        private static class FileLockHelper
        {
            /// <summary>
            /// Returns true if the file exists and cannot be opened for exclusive read (i.e., locked by another process).
            /// </summary>
            public static bool IsFileLocked(string path)
            {
                if (string.IsNullOrWhiteSpace(path)) return false;
                if (!File.Exists(path)) return false;

                FileStream stream = null;
                try
                {
                    // Try exclusive open (no sharing). If this fails, the file is locked.
                    stream = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.None);
                    return false; // opened successfully => not locked
                }
                catch (IOException ex)
                {
                    // Typical exception when file is in use
                    ServiceLocator.Logger?.LogDebug($"IsFileLocked: IOException while testing lock for '{path}': {ex.Message}");
                    return true;
                }
                catch (Exception ex)
                {
                    // For other exceptions, conservatively treat as locked
                    ServiceLocator.Logger?.LogWarn($"IsFileLocked: unexpected exception while testing lock for '{path}' (treating as locked): {ex.Message}");
                    return true;
                }
                finally
                {
                    stream?.Dispose();
                }
            }
        }
    }

    public static class QuarantineDeleteHelper
    {
        /// <summary>
        /// Attempts to rename a locked file to a temporary name so the original path becomes free.
        /// Returns the new path if rename succeeded; null otherwise.
        /// </summary>
        public static string TryQuarantine(string path)
        {
            ServiceLocator.Logger?.LogDebug($"QuarantineDeleteHelper.TryQuarantine: attempting to quarantine locked file '{path}'");

            if (string.IsNullOrWhiteSpace(path) || !File.Exists(path))
                return null;

            var dir = Path.GetDirectoryName(path);
            var name = Path.GetFileName(path);
            var newName = $"{name}.del_{DateTime.UtcNow:yyyyMMddHHmmssfff}";
            var newPath = Path.Combine(dir ?? "", newName);

            try
            {
                File.Move(path, newPath);
                ServiceLocator.Logger?.LogDebug($"QuarantineDeleteHelper.TryQuarantine: moved '{path}' to '{newPath}'");
                return newPath;
            }
            catch (Exception ex)
            {
                ServiceLocator.Logger?.LogWarn($"TryQuarantine: failed to move '{path}' to quarantine (non-fatal): {ex.Message}");
                return null;
            }
        }
    }
}
