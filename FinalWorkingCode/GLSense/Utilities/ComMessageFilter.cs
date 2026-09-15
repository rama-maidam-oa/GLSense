using System;
using System.Runtime.InteropServices;

namespace GLSense.Utilities
{
#nullable enable
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
        private const int PENDINGMSG_WAITNOPROCESS = 1;
        private const int RETRY_DELAY_MS = 250;
        private const int RETRY_GIVE_UP = -1;

        private static Filter? _filter;

        public static void Register()
        {
            try
            {
                _filter = new Filter();
                CoRegisterMessageFilter(_filter, out _);
                LogUtility.LogDebug("ComMessageFilter registered.");
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex, "Failed to register ComMessageFilter");
            }
        }

        public static void Revoke()
        {
            try
            {
                if (_filter == null) return;
                CoRegisterMessageFilter(null, out _);
                _filter = null;
                LogUtility.LogDebug("ComMessageFilter revoked.");
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex, "Failed to revoke ComMessageFilter");
            }
        }

        [DllImport("Ole32.dll")]
        private static extern int CoRegisterMessageFilter(IOleMessageFilter? newFilter, out IOleMessageFilter? oldFilter);

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

            // WAITNOPROCESS (not WAITDEFPROCESS): confirmed via a controlled A/B test that
            // simply registering ANY IOleMessageFilter at all - even one whose
            // MessagePending returns the ostensibly "default" WAITDEFPROCESS value -
            // changes COM's internal handling of window-activation-related messages while
            // an outgoing call to Excel is pending (which is constantly, since nearly
            // every Excel property read is an outgoing COM call). That let Excel's own
            // window silently reclaim OS activation from the floating Balance
            // Configurator's window moments before the user's next keystroke, regardless
            // of the floating window's own implementation (reproduced identically across
            // a WPF Window, a WinForms Form hosting WPF via ElementHost, and a WinForms
            // Form with zero WPF content at all - see GLBalanceConfiguratorForm.cs's own
            // history). WAITNOPROCESS tells COM not to pump/forward messages at all while
            // the call is pending, which fixed it in that same A/B test. This does not
            // affect this filter's actual documented purpose (retrying Excel's "busy"
            // rejections) - that lives entirely in RetryRejectedCall above, a separate
            // callback untouched by this change.
            int IOleMessageFilter.MessagePending(IntPtr hTaskCallee, int dwTickCount, int dwPendingType)
                => PENDINGMSG_WAITNOPROCESS;
        }
    }
}
