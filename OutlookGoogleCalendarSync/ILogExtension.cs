using System;
using log4net;

namespace OutlookGoogleCalendarSync {
    public static class ILogExtentions {
        
        private static void Fine(this ILog log, string message, Exception exception) {
            log.Logger.Log(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType,
                Program.MyFineLevel, message, exception);
        }

        public static void Fine(this ILog log, string message) {
            log.Fine(message, null);
        }
    }
}
