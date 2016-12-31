using log4net;
using System;
using System.ComponentModel;

namespace OutlookGoogleCalendarSync {
    class UserCancelledSyncException : Exception {
        public UserCancelledSyncException() { }
        public UserCancelledSyncException(string message) : base(message) { }
        public UserCancelledSyncException(string message, Exception inner) : base(message, inner) { }
    }

    class OGCSexception {
        private static readonly ILog log = LogManager.GetLogger(typeof(OGCSexception));

        public static void Analyse(System.Exception ex, Boolean includeStackTrace = false) {
            log.Error(ex.GetType().FullName +": "+ ex.Message);
            log.Error("Code: " + getErrorCode(ex).ToString());
            if (includeStackTrace) log.Error(ex.StackTrace);
        }

        private static int getErrorCode(System.Exception ex) {
            try {
                var w32ex = ex as Win32Exception;
                if (w32ex == null) {
                    w32ex = ex.InnerException as Win32Exception;
                }
                if (w32ex != null) {
                    return w32ex.ErrorCode;
                }
            } catch {
                log.Error("Failed to obtain Win32Exception.");
            }
            try {
                return System.Runtime.InteropServices.Marshal.GetHRForException(ex);
            } catch {
                log.Error("Failed to get HResult.");
            }
            return -1;
        }
    }
}
