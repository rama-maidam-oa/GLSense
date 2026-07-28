// ApiOperationHelper.cs in GLSense.Addin.Core
// Right-sized port of GLSense\Helpers\ApiOperationHelper.cs (FinalWorkingCode).
// Ported: ExecuteWithRetry, ValidateApiResponse, ParseJsonResponse, EnsureValidApiResponse.
// Changes vs. the original: LogUtility.* -> ServiceLocator.Logger.*.
// Deliberately NOT ported: the NotifyApiError(string, Dispatcher, AppOverlay, string)
// overload. The original showed the error either on a passed-in AppOverlay/Dispatcher
// pair, or (if neither was given) via CommonFunctions.GLSenseMessage using the old
// WinForms MessageBoxIcon/MessageBoxButtons enums - a signature this project's
// GLSenseMessage no longer has (it takes WPF MessageBoxImage/MessageBoxButton
// directly, see Utilities/CommonFunctions.cs). EnsureValidApiResponse below is
// simplified to validate + throw only; callers already show errors explicitly at the
// UI layer (see GLLogin.xaml.cs's AppOverlayControl.ShowErrorAsync calls) rather than
// through this helper, so nothing is lost - just don't reintroduce a
// Dispatcher/AppOverlay parameter pair here without picking a real call site first.
using GLSense.Addin.Core.Infrastructure;
using System;
using System.Text.Json;
using System.Threading;
using System.Threading.Tasks;

namespace GLSense.Addin.Core.Helpers
{
    /// <summary>
    /// Centralized helper for API operations with retry logic and better error handling
    /// </summary>
    public static class ApiOperationHelper
    {
        private const int MaxRetries = 3;
        private const int RetryDelayMs = 1000;

        /// <summary>
        /// Executes an API call with retry logic for transient failures
        /// </summary>
        public static async Task<string> ExecuteWithRetry(
            Func<CancellationToken, Task<string>> apiCall,
            CancellationToken cancellationToken,
            string operationName = "API call")
        {
            int retryCount = 0;
            Exception lastException = null;

            while (retryCount < MaxRetries)
            {
                try
                {
                    CancellationTokenHelper.ThrowIfCancelled(cancellationToken, operationName);

                    if (retryCount > 0)
                    {
                        ServiceLocator.Logger?.LogDebug($"Retry attempt {retryCount} of {MaxRetries - 1}");
                    }

                    var result = await apiCall(cancellationToken);

                    if (retryCount > 0)
                    {
                        // LogInfo is not used in this codebase (per project convention) - LogDebug
                        // is the safe/cheap equivalent, gated by the ribbon's Debug checkbox.
                        ServiceLocator.Logger?.LogDebug($"{operationName} succeeded after {retryCount} retries");
                    }

                    return result;
                }
                catch (OperationCanceledException)
                {
                    ServiceLocator.Logger?.LogWarn($"{operationName} was cancelled");
                    throw;
                }
                catch (Exception ex) when (IsTransientError(ex))
                {
                    lastException = ex;
                    retryCount++;

                    if (retryCount >= MaxRetries)
                    {
                        ExceptionHelper.LogDetailedException(ex, $"{operationName} - Max retries exceeded");
                        throw;
                    }

                    ServiceLocator.Logger?.LogWarn($"{operationName} failed (transient error) - retrying in {RetryDelayMs}ms: {ex.Message}");

                    await CancellationTokenHelper.DelayWithLogging(
                        RetryDelayMs * retryCount, // Exponential backoff
                        cancellationToken,
                        $"Retry delay for {operationName}");
                }
                catch (Exception ex)
                {
                    // Non-transient error - don't retry
                    ExceptionHelper.LogDetailedException(ex, operationName);
                    throw;
                }
            }

            // Should never reach here, but for safety
            if (lastException != null)
            {
                throw lastException;
            }

            return string.Empty;
        }

        /// <summary>
        /// Determines if an error is transient and worth retrying
        /// </summary>
        private static bool IsTransientError(Exception ex)
        {
            return ex is System.Net.Http.HttpRequestException ||
                   ex is TimeoutException ||
                   ex is System.Net.WebException ||
                   (ex is System.Net.Sockets.SocketException socketEx &&
                    IsTransientSocketError(socketEx));
        }

        private static bool IsTransientSocketError(System.Net.Sockets.SocketException socketEx)
        {
            // Connection timeout, connection refused, network unreachable, etc.
            return socketEx.SocketErrorCode == System.Net.Sockets.SocketError.TimedOut ||
                   socketEx.SocketErrorCode == System.Net.Sockets.SocketError.ConnectionRefused ||
                   socketEx.SocketErrorCode == System.Net.Sockets.SocketError.NetworkUnreachable ||
                   socketEx.SocketErrorCode == System.Net.Sockets.SocketError.HostDown;
        }

        /// <summary>
        /// Validates API response for common error patterns
        /// </summary>
        public static bool ValidateApiResponse(string response, out string errorMessage)
        {
            errorMessage = string.Empty;

            try
            {
                if (string.IsNullOrWhiteSpace(response))
                {
                    errorMessage = "Empty response from server";
                    ServiceLocator.Logger?.LogWarn(errorMessage);
                    return false;
                }

                if (response.IndexOf("(401) Unauthorized", StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    errorMessage = "Session expired! Please re-login.";
                    ServiceLocator.Logger?.LogError(errorMessage);
                    return false;
                }

                if (response.StartsWith("Error", StringComparison.OrdinalIgnoreCase))
                {
                    errorMessage = response;
                    ServiceLocator.Logger?.LogError($"API returned error: {response}");
                    return false;
                }

                if (response.IndexOf("<!DOCTYPE", StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    errorMessage = "Received HTML response instead of expected data";
                    ServiceLocator.Logger?.LogError(errorMessage);
                    return false;
                }

                ServiceLocator.Logger?.LogDebug("API response validation passed");
                return true;
            }
            catch (Exception ex)
            {
                ExceptionHelper.LogDetailedException(ex, "ValidateApiResponse");
                errorMessage = "Error validating API response";
                return false;
            }
        }

        /// <summary>
        /// Parses JSON response with detailed error logging
        /// </summary>
        public static T ParseJsonResponse<T>(string jsonResponse, string operationName = "")
        {
            try
            {
                ServiceLocator.Logger?.LogDebug($"Parsing JSON response (Length: {jsonResponse?.Length ?? 0})");

                var result = JsonSerializer.Deserialize<T>(jsonResponse, JsonGlobals.Options);

                ServiceLocator.Logger?.LogDebug("JSON parsed successfully");
                return result;
            }
            catch (JsonException ex)
            {
                ServiceLocator.Logger?.LogRawJson($"ParseJsonResponse<{typeof(T).Name}>", jsonResponse ?? string.Empty);
                ExceptionHelper.LogDetailedException(ex, $"ParseJsonResponse: {operationName}");
                throw;
            }
        }

        /// <summary>
        /// Validates the API response and throws when invalid. Callers are responsible
        /// for showing the error to the user (e.g. via AppOverlayControl.ShowErrorAsync)
        /// - this helper only validates and logs.
        /// </summary>
        public static void EnsureValidApiResponse(string response, string operationName)
        {
            if (ValidateApiResponse(response, out var errorMessage))
                return;

            throw new InvalidOperationException(string.IsNullOrWhiteSpace(errorMessage)
                ? "Invalid API response."
                : errorMessage);
        }
    }
}
