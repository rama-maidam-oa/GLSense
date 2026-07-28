using GLSense.Utilities;
using System;
using System.Text;

namespace GLSense.Helpers
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

                LogUtility.LogException(ex, message.ToString());
            }
            catch
            {
                // Swallow any logging exceptions to avoid secondary failures
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
