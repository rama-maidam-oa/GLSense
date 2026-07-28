// ExceptionHelper.cs in GLSense.Addin.Core
// Verbatim port of GLSense\Helpers\ExceptionHelper.cs (FinalWorkingCode).
// Only change: LogUtility.LogException (static) -> ServiceLocator.Logger.LogException.
using GLSense.Addin.Core.Infrastructure;
using System;
using System.Text;

namespace GLSense.Addin.Core.Helpers
{
    public static class ExceptionHelper
    {
        /// <summary>
        /// Logs an exception with additional context.
        /// </summary>
        public static void LogDetailedException(Exception ex, string context)
        {
            try
            {
                var message = new StringBuilder();
                if (!string.IsNullOrWhiteSpace(context))
                {
                    message.Append(context);
                    message.Append(" | ");
                }
                message.Append(ex.GetType().FullName);
                message.Append(": ");
                message.Append(ex.Message);

                if (ex.InnerException != null)
                {
                    message.Append(" | Inner: ");
                    message.Append(ex.InnerException.GetType().FullName);
                    message.Append(": ");
                    message.Append(ex.InnerException.Message);
                }

                ServiceLocator.Logger?.LogException(ex, message.ToString());
            }
            catch (Exception loggingEx)
            {
                // Deliberately do not call back into ServiceLocator.Logger here - that's the
                // component that just failed, and retrying it risks an infinite loop if it's
                // the cause. Fall back to Debug.WriteLine so this isn't completely silent.
                System.Diagnostics.Debug.WriteLine($"ExceptionHelper.LogDetailedException: logging itself failed - {loggingEx.Message}");
            }
        }

        /// <summary>
        /// Returns a simplified friendly error message for UI display.
        /// </summary>
        public static string GetFriendlyErrorMessage(Exception ex)
        {
            if (ex == null) return "An unexpected error occurred.";

            // Prefer inner exception message if present
            return !string.IsNullOrWhiteSpace(ex.InnerException?.Message)
                ? ex.InnerException.Message
                : ex.Message;
        }
    }
}
