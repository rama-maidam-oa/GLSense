using System;
using System.Runtime.InteropServices;

namespace GLSense.Models
#nullable enable
{
    [ComVisible(false)]
    public sealed class ApiResult<T>
    {
        public T? Value { get; }
        public Exception? Exception { get; }
        public int? StatusCode { get; }

        public bool IsSuccess => Exception == null;
        public string ErrorMessage => Exception?.Message ?? string.Empty;

        private ApiResult(T? value, Exception? exception, int? statusCode = null)
        {
            Value = value;
            Exception = exception;
            StatusCode = statusCode;
        }

        // ---------------------------
        // Factory Methods
        // ---------------------------

        public static ApiResult<T> Success(T value, int? statusCode = null)
            => new(value, null, statusCode);

        public static ApiResult<T> Failure(Exception exception, int? statusCode = null)
        {
            if (exception == null)
                throw new ArgumentNullException(nameof(exception));

            return new(default, exception, statusCode);
        }

        public static ApiResult<T> Failure(string message, int? statusCode = null)
        {
            if (string.IsNullOrWhiteSpace(message))
                message = "Unknown error occurred.";

            return new(default, new InvalidOperationException(message), statusCode);
        }
    }
#nullable restore
}
