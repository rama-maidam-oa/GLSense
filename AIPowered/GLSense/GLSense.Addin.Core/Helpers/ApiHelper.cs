// ApiHelper.cs in GLSense.Addin.Core
// Port of GLSense\Helpers\APIHelper.cs (FinalWorkingCode).
// Changes vs. the original:
//   - LogUtility.* -> ServiceLocator.Logger.*
//   - AppState.Instance.LoginToken -> still AppState.Instance.LoginToken, but now valid
//     because AppState.cs in this project has been expanded with LoginToken (see
//     AppState.cs - added alongside LoginUrl/LoginUserName/IsLoggedIn for the Login
//     ribbon group).
//   - AppPaths.TempFilesPath -> ServiceLocator.Paths.Temp (this project's IPathProvider
//     equivalent).
//   - AppConstants.RefreshZipFileName resolves via the enclosing GLSense.Addin.Core
//     namespace (already present there, no change needed).
//   - PerformanceHelper.MeasureExecutionTime / ApiOperationHelper.ExecuteWithRetry /
//     StrictCertificateValidator.Validate all resolve via this project's Helpers
//     namespace (all already ported, see this same folder).
using GLSense.Addin.Core.Infrastructure;
using System;
using System.IO;
using System.Net;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Security.Authentication;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace GLSense.Addin.Core.Helpers
{
    public static class ApiHelper
    {
        // Remove shared client - create new clients per request to avoid connection issues
        private static readonly object _lock = new object();

        static ApiHelper()
        {
            ServiceLocator.Logger?.LogDebug("Initializing ApiHelper");

            // Configure ServicePointManager for better connection handling
            ServicePointManager.DefaultConnectionLimit = 100;
            ServicePointManager.Expect100Continue = false;
            ServicePointManager.MaxServicePointIdleTime = 5000; // 5 seconds idle timeout
            ServicePointManager.SetTcpKeepAlive(false, 0, 0); // Disable TCP keep-alive
            ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12 | SecurityProtocolType.Tls13;

            ServiceLocator.Logger?.LogDebug("ApiHelper initialized successfully");
        }

        private static HttpClient CreateHttpClient()
        {
            var handler = new HttpClientHandler
            {
                SslProtocols = SslProtocols.Tls12 | SslProtocols.Tls13,
                ServerCertificateCustomValidationCallback = StrictCertificateValidator.Validate,
                MaxConnectionsPerServer = 10,
                UseProxy = false,
                AutomaticDecompression = DecompressionMethods.GZip | DecompressionMethods.Deflate
            };

            var client = new HttpClient(handler)
            {
                Timeout = TimeSpan.FromMinutes(5) // Set reasonable timeout instead of Infinite
            };

            client.DefaultRequestHeaders.ExpectContinue = false;
            client.DefaultRequestHeaders.ConnectionClose = true; // Force connection close after request

            return client;
        }

        public static async Task<string> ServerAPI(string sendURL, string StrContentType, string PostData = "", string MethodType = "POST", CancellationToken cancellationToken = default)
        {
            using (var perfScope = PerformanceHelper.MeasureExecutionTime("API Call"))
            {
                // Use retry logic for transient failures
                return await ApiOperationHelper.ExecuteWithRetry(
                    async (token) =>
                    {
                        using (var client = CreateHttpClient())
                        {
                            return await ExecuteApiCall(client, sendURL, StrContentType, PostData, MethodType, token);
                        }
                    },
                    cancellationToken,
                    $"API call to {sendURL}"
                ).ConfigureAwait(false);
            }
        }

        private static async Task<string> ExecuteApiCall(HttpClient client, string sendURL, string StrContentType, string PostData, string MethodType, CancellationToken cancellationToken)
        {
            try
            {
                ServiceLocator.Logger?.LogDebug($"Starting API call to: {sendURL}");
                ServiceLocator.Logger?.LogDebug($"Method: {MethodType}, ContentType: {StrContentType}");

                var content = CreateHttpContent(PostData, StrContentType);
                LogRequestDetails(sendURL, MethodType, StrContentType, PostData, client.Timeout);

                cancellationToken.ThrowIfCancellationRequested();

                // Set authorization header for this request
                lock (_lock)
                {
                    client.DefaultRequestHeaders.Authorization =
                        new AuthenticationHeaderValue("Bearer", AppState.Instance.LoginToken);
                }

                var response = await SendRequestAsync(client, sendURL, MethodType, content, cancellationToken);

                // Read response as stream in chunks to avoid large single-buffer reads
                var responseBody = await ReadResponseStreamAsStringAsync(response, cancellationToken).ConfigureAwait(false);

                LogResponseDetails(response, responseBody);

                var result = ProcessResponse(response, responseBody);
                ApiOperationHelper.EnsureValidApiResponse(result, $"API call to {sendURL}");

                ServiceLocator.Logger?.LogDebug($"API call completed successfully, response length: {result?.Length ?? 0}");

                return result;
            }
            catch (OperationCanceledException)
            {
                ServiceLocator.Logger?.LogWarn($"API call cancelled for {sendURL}");
                throw;
            }
            catch (HttpRequestException ex) when (ex.InnerException is WebException webEx &&
                   (webEx.Status == WebExceptionStatus.KeepAliveFailure ||
                    webEx.Status == WebExceptionStatus.ConnectionClosed))
            {
                // These are transient errors that should be retried
                ServiceLocator.Logger?.LogWarn($"Connection closed by server for {sendURL}: {ex.Message}");
                throw new HttpRequestException("Connection was closed by server (transient)", ex);
            }
            catch (Exception ex)
            {
                ServiceLocator.Logger?.LogException(ex, $"ApiHelper.ExecuteApiCall - unhandled error calling {sendURL}");
                throw;
            }
        }

        private static HttpContent CreateHttpContent(string postData, string contentType)
        {
            if (string.IsNullOrWhiteSpace(postData))
                return null;

            var mediaType = contentType == "JSON" ? "application/json" : "application/x-www-form-urlencoded";
            return new StringContent(postData, Encoding.UTF8, mediaType);
        }

        private static void LogRequestDetails(string url, string method, string contentType, string payload, TimeSpan timeout)
        {
            ServiceLocator.Logger?.LogDebug($"API Request - URL: {url}");
            ServiceLocator.Logger?.LogDebug($"Method: {method}, ContentType: {contentType}, Timeout: {timeout}");
            bool hasAuthToken = !string.IsNullOrWhiteSpace(AppState.Instance.LoginToken);
            ServiceLocator.Logger?.LogDebug($"Authorization: Bearer present={hasAuthToken}");

            if (!string.IsNullOrWhiteSpace(payload))
            {
                LogPayloadChunks(payload);
            }
            else
            {
                ServiceLocator.Logger?.LogDebug("Payload: (empty)");
            }
        }

        private static void LogPayloadChunks(string payload)
        {
            const int chunkSize = 1000;

            for (int i = 0; i < payload.Length; i += chunkSize)
            {
                var length = Math.Min(chunkSize, payload.Length - i);
                var chunk = payload.Substring(i, length);
                ServiceLocator.Logger?.LogDebug($"Payload: {chunk}");
            }
        }

        private static async Task<HttpResponseMessage> SendRequestAsync(HttpClient client, string url, string method, HttpContent content, CancellationToken cancellationToken)
        {
            var httpMethod = new HttpMethod(method.ToUpperInvariant());
            using var request = new HttpRequestMessage(httpMethod, url)
            {
                Content = content
            };

            // Add Connection: close header explicitly
            request.Headers.ConnectionClose = true;

            return await client.SendAsync(request, HttpCompletionOption.ResponseHeadersRead, cancellationToken).ConfigureAwait(false);
        }

        private static async Task<string> ReadResponseStreamAsStringAsync(HttpResponseMessage response, CancellationToken cancellationToken)
        {
            if (response?.Content == null)
                return string.Empty;

            // Try to determine encoding from response headers, default to UTF8
            Encoding encoding = Encoding.UTF8;
            try
            {
                var charset = response.Content.Headers.ContentType?.CharSet;
                if (!string.IsNullOrWhiteSpace(charset))
                {
                    charset = charset.Trim('"');
                    encoding = Encoding.GetEncoding(charset);
                }
            }
            catch (Exception ex)
            {
                ServiceLocator.Logger?.LogWarn($"ApiHelper.ReadResponseStreamAsStringAsync: failed to resolve response charset, defaulting to UTF8. {ex.Message}");
                encoding = Encoding.UTF8;
            }

            // Read the response stream in chunks
            using var stream = await response.Content.ReadAsStreamAsync().ConfigureAwait(false);
            using var reader = new StreamReader(stream, encoding, detectEncodingFromByteOrderMarks: true, bufferSize: 65536, leaveOpen: false);
            var sb = new StringBuilder();
            var buffer = new char[65536];
            int read;
            while ((read = await reader.ReadAsync(buffer, 0, buffer.Length).ConfigureAwait(false)) > 0)
            {
                cancellationToken.ThrowIfCancellationRequested();
                sb.Append(buffer, 0, read);
            }

            return sb.ToString();
        }

        private static void LogResponseDetails(HttpResponseMessage response, string responseBody)
        {
            ServiceLocator.Logger?.LogDebug($"API Response - Status: {(int)response.StatusCode} {response.StatusCode}, ContentType: {response.Content?.Headers?.ContentType?.MediaType ?? "N/A"}, Length: {response.Content?.Headers?.ContentLength ?? 0}");

            if (!string.IsNullOrWhiteSpace(responseBody))
            {
                // Always log the full body, even when huge: a partial/truncated response
                // mixed in with other log lines is worse than a long one - it hides
                // exactly the data (often malformed/unexpected JSON) that debugging a
                // real issue needs to see in full.
                ServiceLocator.Logger?.LogDebug($"Response ({responseBody.Length} chars): {responseBody}");
            }
            else
            {
                ServiceLocator.Logger?.LogDebug("Response: (empty)");
            }
        }

        private static string ProcessResponse(HttpResponseMessage response, string responseBody)
        {
            if (response.IsSuccessStatusCode)
                return CleanResponse(responseBody);

            return !string.IsNullOrWhiteSpace(responseBody) && responseBody.Contains("status")
                ? CleanResponse(responseBody)
                : response.StatusCode.ToString();
        }

        private static string CleanResponse(string response)
        {
            if (string.IsNullOrEmpty(response))
                return response;

            var cleaned = response.Replace("null", string.Empty);

            var sb = new StringBuilder(cleaned.Length + 8);
            bool inString = false;
            bool escape = false;

            for (int i = 0; i < cleaned.Length; i++)
            {
                char c = cleaned[i];

                if (inString)
                {
                    sb.Append(c);

                    if (escape)
                    {
                        escape = false;
                        continue;
                    }

                    if (c == '\\')
                    {
                        escape = true;
                    }
                    else if (c == '"')
                    {
                        inString = false;
                    }

                    continue;
                }

                if (c == '"')
                {
                    inString = true;
                    sb.Append(c);
                    continue;
                }

                if (c == ':')
                {
                    int j = i + 1;
                    while (j < cleaned.Length && char.IsWhiteSpace(cleaned[j]))
                    {
                        j++;
                    }

                    if (j < cleaned.Length && (cleaned[j] == ',' || cleaned[j] == '}' || cleaned[j] == ']'))
                    {
                        sb.Append(':');
                        sb.Append("\"\"");
                        i = j - 1;
                        continue;
                    }
                }

                sb.Append(c);
            }

            return sb.ToString();
        }

        public static async Task<string> HttpUploadFileAsync(string uri, CancellationToken token)
        {
            using (var perfScope = PerformanceHelper.MeasureExecutionTime("File Upload"))
            {
                return await ApiOperationHelper.ExecuteWithRetry(
                    async (cancellationToken) =>
                    {
                        try
                        {
                            ServiceLocator.Logger?.LogDebug($"Starting file upload to: {uri}");
                            cancellationToken.ThrowIfCancellationRequested();

                            var path = ServiceLocator.Paths.Temp ?? string.Empty;
                            var zipPath = Path.Combine(path, AppConstants.RefreshZipFileName);

                            ServiceLocator.Logger?.LogDebug($"File path: {zipPath}");

                            cancellationToken.ThrowIfCancellationRequested();

                            if (!File.Exists(zipPath))
                            {
                                ServiceLocator.Logger?.LogError($"File not found: {zipPath}");
                                return "Error: File not found.";
                            }

                            var fileInfo = new FileInfo(zipPath);
                            ServiceLocator.Logger?.LogDebug($"File size: {fileInfo.Length:N0} bytes");

                            var handler = new HttpClientHandler
                            {
                                SslProtocols = SslProtocols.Tls12 | SslProtocols.Tls13,
                                ServerCertificateCustomValidationCallback = StrictCertificateValidator.Validate,
                                MaxConnectionsPerServer = 10,
                                UseProxy = false,
                                AutomaticDecompression = DecompressionMethods.GZip | DecompressionMethods.Deflate
                            };

                            using var client = new HttpClient(handler);
                            bool hasAuthToken = !string.IsNullOrWhiteSpace(AppState.Instance.LoginToken);
                            ServiceLocator.Logger?.LogDebug($"Authorization: Bearer present={hasAuthToken}");
                            client.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", AppState.Instance.LoginToken);
                            client.Timeout = TimeSpan.FromMinutes(10);
                            client.DefaultRequestHeaders.ConnectionClose = true;

                            cancellationToken.ThrowIfCancellationRequested();

                            ServiceLocator.Logger?.LogDebug("Opening file stream");
                            using var fileStream = new FileStream(zipPath, FileMode.Open, FileAccess.Read, FileShare.Read);
                            using var content = new StreamContent(fileStream);
                            content.Headers.ContentType = new MediaTypeHeaderValue("application/zip");

                            ServiceLocator.Logger?.LogDebug("Sending POST request with file content");
                            using var response = await client.PostAsync(uri, content, cancellationToken).ConfigureAwait(false);

                            cancellationToken.ThrowIfCancellationRequested();

                            var responseText = await response.Content.ReadAsStringAsync().ConfigureAwait(false);

                            ServiceLocator.Logger?.LogDebug($"Upload response: Status={response.StatusCode}, Response length={responseText?.Length ?? 0}");

                            if (response.IsSuccessStatusCode)
                            {
                                return responseText;
                            }
                            else
                            {
                                ServiceLocator.Logger?.LogError($"Upload failed: {response.StatusCode} - {responseText}");
                                return $"Error: {response.StatusCode} - {responseText}";
                            }
                        }
                        catch (TaskCanceledException)
                        {
                            ServiceLocator.Logger?.LogWarn("File upload was cancelled by user");
                            return "Error: Upload was canceled.";
                        }
                        catch (HttpRequestException ex)
                        {
                            ServiceLocator.Logger?.LogException(ex, $"HTTP request failed during file upload to {uri}");
                            throw; // Let retry logic handle it
                        }
                        catch (Exception ex)
                        {
                            ServiceLocator.Logger?.LogException(ex, $"Unexpected error during file upload to {uri}");
                            throw;
                        }
                    },
                    token,
                    $"File upload to {uri}"
                ).ConfigureAwait(false);
            }
        }
    }
}
