// Utilities/ComMessageFilter.cs in GLSense.Addin.Core
// Port of GLSense\Utilities\ComMessageFilter.cs (FinalWorkingCode) - identical logic,
// no ServiceLocator/GLSense.Shared dependency needed (this class has no side effects
// besides the raw Win32 P/Invoke), so it ported verbatim other than the namespace.
using System;
using System.Runtime.InteropServices;

namespace GLSense.Addin.Core.Utilities
{
    // Registers the classic OLE "busy" message filter for this process. Without one
    // registered, Excel rejecting an incoming automation call while it's busy (e.g. the
    // user clicking into a sheet while GLSense is mid-operation) surfaces immediately as
    // COMException 0x800AC472 (VBA_E_IGNORE) instead of being retried - this is what
    // caused the row hide/unhide hang (GetBalanceTotalRange/RowHeight COM errors, then
    // EnableExcelSettings itself failing the same way while trying to restore
    // ScreenUpdating). Registering this filter lets such rejections retry transparently.
    internal static class ComMessageFilter
    {
        private const int SERVERCALL_ISHANDLED = 0;
        private const int SERVERCALL_RETRYLATER = 2;
        private const int PENDINGMSG_WAITDEFPROCESS = 2;
        private const int RETRY_DELAY_MS = 250;
        private const int RETRY_GIVE_UP = -1;

        private static Filter _filter;

        public static void Register()
        {
            try
            {
                _filter = new Filter();
                IOleMessageFilter oldFilter;
                CoRegisterMessageFilter(_filter, out oldFilter);
                Infrastructure.ServiceLocator.Logger?.LogDebug("ComMessageFilter registered.");
            }
            catch (Exception ex)
            {
                Infrastructure.ServiceLocator.Logger?.LogException(ex, "Failed to register ComMessageFilter");
            }
        }

        public static void Revoke()
        {
            try
            {
                if (_filter == null) return;
                IOleMessageFilter oldFilter;
                CoRegisterMessageFilter(null, out oldFilter);
                _filter = null;
                Infrastructure.ServiceLocator.Logger?.LogDebug("ComMessageFilter revoked.");
            }
            catch (Exception ex)
            {
                Infrastructure.ServiceLocator.Logger?.LogException(ex, "Failed to revoke ComMessageFilter");
            }
        }

        [DllImport("Ole32.dll")]
        private static extern int CoRegisterMessageFilter(IOleMessageFilter newFilter, out IOleMessageFilter oldFilter);

        [ComImport, Guid("00000016-0000-0000-C000-000000000046"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        private interface IOleMessageFilter
        {
            [PreserveSig]
            int HandleInComingCall(int dwCallType, IntPtr hTaskCaller, int dwTickCount, IntPtr lpInterfaceInfo);

            [PreserveSig]
            int RetryRejectedCall(IntPtr hTaskCallee, int dwTickCount, int dwRejectType);

            [PreserveSig]
            int MessagePending(IntPtr hTaskCallee, int dwTickCount, int dwPendingType);
        }

        private sealed class Filter : IOleMessageFilter
        {
            int IOleMessageFilter.HandleInComingCall(int dwCallType, IntPtr hTaskCaller, int dwTickCount, IntPtr lpInterfaceInfo)
                => SERVERCALL_ISHANDLED;

            int IOleMessageFilter.RetryRejectedCall(IntPtr hTaskCallee, int dwTickCount, int dwRejectType)
                => dwRejectType == SERVERCALL_RETRYLATER ? RETRY_DELAY_MS : RETRY_GIVE_UP;

            int IOleMessageFilter.MessagePending(IntPtr hTaskCallee, int dwTickCount, int dwPendingType)
                => PENDINGMSG_WAITDEFPROCESS;
        }
    }
}
