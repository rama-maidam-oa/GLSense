using GLSense.Utilities;
using System;
using System.IO;
using System.Net;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Security.Authentication;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace GLSense.Helpers
{
    public static class ApiHelper
    {
        // Remove shared client - create new clients per request to avoid connection issues
        private static readonly object _lock = new object();

        static ApiHelper()
        {
            LogUtility.LogDebug("Initializing ApiHelper");

            // Configure ServicePointManager for better connection handling
            ServicePointManager.DefaultConnectionLimit = 100;
            ServicePointManager.Expect100Continue = false;
            ServicePointManager.MaxServicePointIdleTime = 5000; // 5 seconds idle timeout
            ServicePointManager.SetTcpKeepAlive(false, 0, 0); // Disable TCP keep-alive
            ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12 | SecurityProtocolType.Tls13;

            LogUtility.LogDebug("ApiHelper initialized successfully");
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
                            return await ExecuteApiCall(client, sendURL, StrContentType, PostData, MethodType, token).ConfigureAwait(false);
                        }
                    },
                    cancellationToken,
                    $"API call to {sendURL}"
                ).ConfigureAwait(false);
            }
        }

        private static async Task<string> ExecuteApiCall(HttpClient client, string sendURL, string StrContentType, string PostData, string MethodType, CancellationToken cancellationToken)
        {
            string what = DescribeEndpoint(sendURL);
            ServerWait wait = null;
            try
            {
                LogUtility.LogDebug($"Starting API call to: {sendURL}");
                LogUtility.LogDebug($"Method: {MethodType}, ContentType: {StrContentType}");

                var content = CreateHttpContent(PostData, StrContentType);
                LogRequestDetails(sendURL, MethodType, StrContentType, PostData, client.Timeout);

                cancellationToken.ThrowIfCancellationRequested();

                // Set authorization header for this request
                lock (_lock)
                {
                    client.DefaultRequestHeaders.Authorization =
                        new AuthenticationHeaderValue("Bearer", AppState.Instance.LoginToken);
                }

                wait = new ServerWait(what, sendURL);
                // OISR-23022: ConfigureAwait(false) - a caller blocking on ServerAPI from Excel's thread (ledger setup data in
                // BalanceDto / PeriodDataService) deadlocked when this continuation was posted back to that thread.
                var response = await SendRequestAsync(client, sendURL, MethodType, content, cancellationToken).ConfigureAwait(false);
                wait.ResponseStarted(response.StatusCode);

                // Read response as stream in chunks to avoid large single-buffer reads
                var responseBody = await ReadResponseStreamAsStringAsync(response, cancellationToken).ConfigureAwait(false);
                wait.Completed(response.StatusCode, responseBody?.Length ?? 0);

                LogResponseDetails(response, responseBody);

                var result = ProcessResponse(response, responseBody);
                ApiOperationHelper.EnsureValidApiResponse(result, $"API call to {sendURL}");

                LogUtility.LogDebug($"API call completed successfully, response length: {result?.Length ?? 0}");

                return result;
            }
            catch (OperationCanceledException) when (!cancellationToken.IsCancellationRequested)
            {
                // HttpClient.Timeout ends the call with the same exception as Cancel - say which it was.
                LogUtility.LogWarn($"The server did not answer {what} within {client.Timeout.TotalMinutes:0.#} minutes " +
                    $"(timed out after {wait?.ElapsedSeconds ?? 0:0} s): {sendURL}");
                throw;
            }
            catch (OperationCanceledException)
            {
                LogUtility.LogWarn($"API call cancelled for {sendURL} after {wait?.ElapsedSeconds ?? 0:0} s.");
                throw;
            }
            catch (HttpRequestException ex) when (ex.InnerException is WebException webEx &&
                   (webEx.Status == WebExceptionStatus.KeepAliveFailure ||
                    webEx.Status == WebExceptionStatus.ConnectionClosed))
            {
                // These are transient errors that should be retried
                LogUtility.LogWarn($"Connection closed by server for {sendURL}: {ex.Message}");
                throw new HttpRequestException("Connection was closed by server (transient)", ex);
            }
            catch (Exception)
            {
                throw;
            }
            finally
            {
                wait?.Dispose();
            }
        }

        /// <summary>
        /// OISR-23022: writes a server request's waiting time to the log, so a slow or silent server is visible there:
        /// "Still waiting for the server" every 30 s while the request is open (WARN, also without debug mode), then one
        /// timing line - time to the server's first answer (response headers) and to the end of the response. The caller
        /// marks the first answer with <see cref="ResponseStarted"/> on a ConfigureAwait(false) continuation, so the log
        /// reflects the network even when the calling thread is busy.
        /// </summary>
        internal sealed class ServerWait : IDisposable
        {
            private const int HeartbeatMs = 30000;
            private const double SlowSeconds = 10;
            private readonly System.Diagnostics.Stopwatch _sw = System.Diagnostics.Stopwatch.StartNew();
            private readonly Timer _heartbeat;
            private readonly string _what, _url;
            private volatile string _stage = "no response from the server yet";
            private double _firstAnswerSeconds = -1;

            internal ServerWait(string what, string url)
            {
                _what = what;
                _url = url;
                _heartbeat = new Timer(_ => LogUtility.LogWarn(
                    $"Still waiting for the server: {_what} - {_stage} after {_sw.Elapsed.TotalSeconds:0} s ({_url})."), null, HeartbeatMs, HeartbeatMs);
            }

            internal double ElapsedSeconds => _sw.Elapsed.TotalSeconds;

            internal void ResponseStarted(HttpStatusCode status)
            {
                _firstAnswerSeconds = _sw.Elapsed.TotalSeconds;
                _stage = $"the server answered HTTP {(int)status} after {_firstAnswerSeconds:0.0} s, still receiving its response";
            }

            /// <summary>The timing line: INFO when the request took longer than 10 s, else DEBUG.</summary>
            internal void Completed(HttpStatusCode status, long length)
            {
                _heartbeat.Dispose();
                string line = $"Server timing: {_what} - HTTP {(int)status}, server's first answer after {_firstAnswerSeconds:0.0} s, " +
                              $"response complete after {_sw.Elapsed.TotalSeconds:0.0} s, {length:N0} chars ({_url}).";
                if (_sw.Elapsed.TotalSeconds > SlowSeconds) LogUtility.LogInfo(line);
                else LogUtility.LogDebug(line);
            }

            public void Dispose() => _heartbeat.Dispose();
        }

        // ".../finance/ledger-setup-data?ledgerId=1" -> "ledger setup data"
        internal static string DescribeEndpoint(string url)
        {
            try
            {
                var segments = new Uri(url).Segments;
                string last = segments.Length == 0 ? string.Empty : segments[segments.Length - 1].Trim('/');
                return last.Length == 0 ? "request" : last.Replace('-', ' ').Replace('_', ' ');
            }
            catch (Exception)
            {
                return "request";
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
            // Bearer token itself must never be logged - only whether one will be sent.
            bool hasAuthToken = !string.IsNullOrWhiteSpace(AppState.Instance.LoginToken);

            LogUtility.LogDebug($"┌─ API Request ─────────────────");
            LogUtility.LogDebug($"│ URL: {url}");
            LogUtility.LogDebug($"│ Method: {method}");
            LogUtility.LogDebug($"│ ContentType: {contentType}");
            LogUtility.LogDebug($"│ Authorization: Bearer present={hasAuthToken}");
            LogUtility.LogDebug($"│ Timeout: {timeout}");

            if (!string.IsNullOrWhiteSpace(payload))
            {
                LogPayloadChunks(payload);
            }
            else
            {
                LogUtility.LogDebug($"│ Payload: (empty)");
            }
            LogUtility.LogDebug($"└───────────────────────────────");
        }

        private static void LogPayloadChunks(string payload)
        {
            const int chunkSize = 1000;

            for (int i = 0; i < payload.Length; i += chunkSize)
            {
                var length = Math.Min(chunkSize, payload.Length - i);
                var chunk = payload.Substring(i, length);
                LogUtility.LogDebug($"│ Payload: {chunk}");
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
            catch
            {
                // Missing/invalid charset header - can be ignored as expected, already
                // defaulted to UTF8 above.
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
            LogUtility.LogDebug($"┌─ API Response ────────────────");
            LogUtility.LogDebug($"│ Status Code: {(int)response.StatusCode} {response.StatusCode}");
            LogUtility.LogDebug($"│ Content Type: {response.Content?.Headers?.ContentType?.MediaType ?? "N/A"}");
            LogUtility.LogDebug($"│ Content Length: {response.Content?.Headers?.ContentLength ?? 0}");

            if (!string.IsNullOrWhiteSpace(responseBody))
            {
                // Always log the full body, even when huge: a partial/truncated response
                // mixed in with other log lines is worse than a long one - it hides
                // exactly the data (often malformed/unexpected JSON) that debugging a
                // real issue needs to see in full.
                LogUtility.LogDebug($"│ Response ({responseBody.Length} chars): {responseBody}");
            }
            else
            {
                LogUtility.LogDebug($"│ Response: (empty)");
            }
            LogUtility.LogDebug($"└───────────────────────────────");
        }

        private static string ProcessResponse(HttpResponseMessage response, string responseBody)
        {
            if (response.IsSuccessStatusCode)
                return CleanResponse(responseBody);

            return !string.IsNullOrWhiteSpace(responseBody) && responseBody.Contains("status")
                ? CleanResponse(responseBody)
                : response.StatusCode.ToString();
        }

        // Removes JSON null literals, but only OUTSIDE string values. The old blanket
        // Replace("null", "") also ran inside strings, so real data containing "null"
        // was silently corrupted (e.g. "annulled" -> "aned"). Outside a string, "null"
        // can only be the JSON literal, so removing it there keeps the previous
        // behavior exactly ("x":null -> "x": -> "x":"" via the colon pass below,
        // [null] -> []).
        private static string RemoveNullLiterals(string json)
        {
            var sb = new StringBuilder(json.Length);
            bool inString = false;
            bool escape = false;

            for (int i = 0; i < json.Length; i++)
            {
                char c = json[i];

                if (inString)
                {
                    sb.Append(c);

                    if (escape)
                        escape = false;
                    else if (c == '\\')
                        escape = true;
                    else if (c == '"')
                        inString = false;

                    continue;
                }

                if (c == '"')
                {
                    inString = true;
                    sb.Append(c);
                    continue;
                }

                if (c == 'n' && string.CompareOrdinal(json, i, "null", 0, 4) == 0)
                {
                    i += 3;
                    continue;
                }

                sb.Append(c);
            }

            return sb.ToString();
        }

        private static string CleanResponse(string response)
        {
            if (string.IsNullOrEmpty(response))
                return response;

            var cleaned = RemoveNullLiterals(response);

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
                            LogUtility.LogDebug($"Starting file upload to: {uri}");
                            cancellationToken.ThrowIfCancellationRequested();

                            var path = AppPaths.TempFilesPath ?? string.Empty;
                            var zipPath = Path.Combine(path, AppConstants.RefreshZipFileName);

                            LogUtility.LogDebug($"File path: {zipPath}");

                            cancellationToken.ThrowIfCancellationRequested();

                            if (!File.Exists(zipPath))
                            {
                                LogUtility.LogError($"File not found: {zipPath}");
                                return "Error: File not found.";
                            }

                            var fileInfo = new FileInfo(zipPath);
                            LogUtility.LogDebug($"File size: {fileInfo.Length:N0} bytes");

                            var handler = new HttpClientHandler
                            {
                                SslProtocols = SslProtocols.Tls12 | SslProtocols.Tls13,
                                ServerCertificateCustomValidationCallback = StrictCertificateValidator.Validate,
                                MaxConnectionsPerServer = 10,
                                UseProxy = false,
                                AutomaticDecompression = DecompressionMethods.GZip | DecompressionMethods.Deflate
                            };

                            using var client = new HttpClient(handler);
                            client.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", AppState.Instance.LoginToken);
                            client.Timeout = TimeSpan.FromMinutes(10);
                            client.DefaultRequestHeaders.ConnectionClose = true;

                            LogUtility.LogDebug($"Method: POST, ContentType: application/zip, Authorization: Bearer present={!string.IsNullOrWhiteSpace(AppState.Instance.LoginToken)}");

                            cancellationToken.ThrowIfCancellationRequested();

                            LogUtility.LogDebug("Opening file stream");
                            using var fileStream = new FileStream(zipPath, FileMode.Open, FileAccess.Read, FileShare.Read);
                            using var content = new StreamContent(fileStream);
                            content.Headers.ContentType = new MediaTypeHeaderValue("application/zip");

                            LogUtility.LogDebug("Sending POST request with file content");
                            using var wait = new ServerWait(DescribeEndpoint(uri), uri);
                            using var response = await client.PostAsync(uri, content, cancellationToken).ConfigureAwait(false);
                            wait.ResponseStarted(response.StatusCode);

                            cancellationToken.ThrowIfCancellationRequested();

                            var responseText = await response.Content.ReadAsStringAsync().ConfigureAwait(false);
                            wait.Completed(response.StatusCode, responseText?.Length ?? 0);

                            LogUtility.LogDebug($"Upload response: Status={response.StatusCode}, Response length={responseText?.Length ?? 0}");

                            if (response.IsSuccessStatusCode)
                            {
                                return responseText;
                            }
                            else
                            {
                                LogUtility.LogError($"Upload failed: {response.StatusCode} - {responseText}");
                                return $"Error: {response.StatusCode} - {responseText}";
                            }
                        }
                        catch (TaskCanceledException)
                        {
                            LogUtility.LogWarn("File upload was cancelled by user");
                            return "Error: Upload was canceled.";
                        }
                        catch (HttpRequestException ex)
                        {
                            LogUtility.LogException(ex, $"HTTP request failed during file upload to {uri}");
                            throw; // Let retry logic handle it
                        }
                        catch (Exception ex)
                        {
                            LogUtility.LogException(ex, $"Unexpected error during file upload to {uri}");
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