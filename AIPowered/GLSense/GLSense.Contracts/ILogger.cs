// ILogger.cs in GLSense.Contracts
using System;

namespace GLSense.Contracts
{
    public interface ILogger
    {
        void LogInfo(string msg);
        void LogWarn(string msg);
        void LogError(string msg, Exception ex = null);
        void LogDebug(string msg);
        void LogException(Exception ex, string context = "");
        void LogRawJson(string context, string rawJson);
        void FlushDebugLogs(string section = "Buffered Logs");
        void LogMethodEntry([System.Runtime.CompilerServices.CallerMemberName] string methodName = "");
        void LogMethodExit([System.Runtime.CompilerServices.CallerMemberName] string methodName = "");

        // Opens a per-action buffered log scope (see Logger.cs's ActionBuffer design) -
        // callers dispose it (typically via `using`) when the action completes, which
        // flushes everything logged inside as one batched write instead of one file
        // write per line. Returns an IDisposable rather than the concrete LogScope type
        // so Addin.Core code never needs a direct Logger reference, only this interface.
        IDisposable BeginLogScope(string scopeName);
    }
}
