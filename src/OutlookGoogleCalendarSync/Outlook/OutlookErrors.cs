using Ogcs = OutlookGoogleCalendarSync;
using log4net;
using System;

namespace OutlookGoogleCalendarSync.Outlook {
    public class Errors {
        private static readonly ILog log = LogManager.GetLogger(typeof(Errors));

        public enum ErrorType {
            InvokedObjectDisconnectedFromClients,
            ObjectSeparatedFromRcw,
            OperationFailed,
            PermissionFailure,
            RpcFailed,
            RpcRejected,
            RpcServerUnavailable,
            Unavailable,
            WrongThread,
            Unhandled
        }

        public static ErrorType HandleComError(System.Exception ex) {
            return HandleComError(ex, out _);
        }

        public static ErrorType HandleComError(System.Exception ex, out String hResult) {
            ErrorType retVal = ErrorType.Unhandled;
            try {
                hResult = ex.GetErrorCode();
                log.Warn($"[{hResult}] >> {ex.Message}");

                if (hResult == "0x800401E3" && ex.Message.Contains("MK_E_UNAVAILABLE")) {
                    return retVal = ErrorType.Unavailable;
                }
                
                if (ex is System.InvalidCastException && hResult == "0x80004002") {
                    if (ex.Message.Contains("0x8001010E (RPC_E_WRONG_THREAD")) return retVal = ErrorType.WrongThread;
                    if (ex.Message.Contains("0x800706BA")) return retVal = ErrorType.RpcServerUnavailable;
                }

                if (ex is System.Runtime.InteropServices.COMException) {
                    if (hResult == "0x80004002" && ex.Message.Contains("0x8001010E (RPC_E_WRONG_THREAD")) return retVal = ErrorType.WrongThread;

                    if (hResult == "0x80010001" && ex.Message.Contains("RPC_E_CALL_REJECTED")) return retVal = ErrorType.RpcRejected;
                    if (hResult == "0x80040201") return retVal = ErrorType.OperationFailed; //The messaging interfaces have returned an unknown error. If the problem persists, restart Outlook.
                    if (ex.Message.Contains("0x80010108(RPC_E_DISCONNECTED)")) return retVal = ErrorType.InvokedObjectDisconnectedFromClients;
                    if (hResult == "0x800706BA") return retVal = ErrorType.RpcServerUnavailable;
                    if (hResult == "0x800706BE") return retVal = ErrorType.RpcFailed;
                    if (hResult == "0x80080005" && ex.Message.Contains("CO_E_SERVER_EXEC_FAILURE")) return retVal = ErrorType.PermissionFailure;
                }

                if (ex is System.Runtime.InteropServices.InvalidComObjectException) {
                    if (hResult == "0x80131527") return retVal = ErrorType.ObjectSeparatedFromRcw;
                }
            } finally {
                if (retVal != ErrorType.Unhandled)
                    Ogcs.Exception.Analyse(Ogcs.Exception.LogAsFail(ex));
            }
            
            return retVal;
        }
    }
}
