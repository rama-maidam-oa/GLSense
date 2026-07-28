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
    }
}
