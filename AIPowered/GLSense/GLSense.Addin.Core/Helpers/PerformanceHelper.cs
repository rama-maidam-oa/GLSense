// PerformanceHelper.cs in GLSense.Addin.Core
// Port of GLSense\Helpers\PerformanceHelper.cs (FinalWorkingCode).
// Changes: LogUtility.* -> ServiceLocator.Logger.*.
using GLSense.Addin.Core.Infrastructure;
using System;
using System.Diagnostics;
using System.Runtime.CompilerServices;
using System.Threading.Tasks;

namespace GLSense.Addin.Core.Helpers
{
    /// <summary>
    /// Performance monitoring helper to track method execution times and identify bottlenecks
    /// </summary>
    public static class PerformanceHelper
    {
        /// <summary>
        /// Executes an async action and measures its execution time
        /// </summary>
        public static async Task MeasureExecutionTimeAsync(
            Func<Task> action,
            [CallerMemberName] string operationName = "")
        {
            using (new PerformanceScope(operationName))
            {
                if (action != null)
                {
                    await action();
                }
            }
        }

        /// <summary>
        /// Executes an async function and measures its execution time, returning the result
        /// </summary>
        public static async Task<T> MeasureExecutionTimeAsync<T>(
            Func<Task<T>> func,
            [CallerMemberName] string operationName = "")
        {
            using (new PerformanceScope(operationName))
            {
                return await func();
            }
        }

        /// <summary>
        /// Executes an action and measures its execution time
        /// </summary>
        public static void MeasureExecutionTime(
            Action action,
            [CallerMemberName] string operationName = "")
        {
            using (new PerformanceScope(operationName))
            {
                action?.Invoke();
            }
        }

        /// <summary>
        /// Executes a function and measures its execution time, returning the result
        /// </summary>
        public static T MeasureExecutionTime<T>(
            Func<T> func,
            [CallerMemberName] string operationName = "")
        {
            using (new PerformanceScope(operationName))
            {
                return func();
            }
        }

        /// <summary>
        /// Executes an action and measures its execution time
        /// </summary>
        /// <param name="operationName">Name of the operation being measured</param>
        /// <param name="url">Optional API URL or navigating address</param>
        public static PerformanceScope MeasureExecutionTime(string operationName, string url = null)
        {
            return new PerformanceScope(operationName, url);
        }

        /// <summary>
        /// Disposable scope for performance measurement
        /// </summary>
        public sealed class PerformanceScope : IDisposable
        {
            private readonly string _operationName;
            private readonly string _url;
            private readonly Stopwatch _stopwatch;
            private bool _disposed;

            public PerformanceScope(string operationName, string url = null)
            {
                _operationName = operationName;
                _url = url;
                _stopwatch = Stopwatch.StartNew();

                var message = string.IsNullOrEmpty(_url)
                    ? $"Performance measurement started: {_operationName}"
                    : $"Performance measurement started: {_operationName} | URL: {_url}";

                ServiceLocator.Logger?.LogDebug(message);
            }

            public void Dispose()
            {
                if (_disposed) return;

                _stopwatch.Stop();

                var elapsed = _stopwatch.Elapsed;
                string timeString = FormatElapsedTime(elapsed);

                var message = string.IsNullOrEmpty(_url)
                    ? $"Performance measurement completed: {_operationName} - {timeString}"
                    : $"Performance measurement completed: {_operationName} - {timeString} | URL: {_url}";

                ServiceLocator.Logger?.LogDebug(message);

                if (elapsed.TotalSeconds > 45)
                {
                    var warnMessage = string.IsNullOrEmpty(_url)
                        ? $"Slow operation detected: {_operationName} took {timeString}"
                        : $"Slow operation detected: {_operationName} took {timeString} | URL: {_url}";

                    ServiceLocator.Logger?.LogWarn(warnMessage);
                }

                _disposed = true;
            }

            private static string FormatElapsedTime(TimeSpan elapsed)
            {
                if (elapsed.TotalMilliseconds < 1000)
                    return $"{elapsed.TotalMilliseconds:F2}ms";

                if (elapsed.TotalSeconds < 60)
                    return $"{elapsed.TotalSeconds:F2}s";

                return $"{elapsed.TotalMinutes:F2}m";
            }
        }
    }
}
