using System;
using log4net;

namespace OutlookGoogleCalendarSync {
    public static class ILogExtentions {
        
        private static void Fine(this ILog log, string message, System.Exception exception) {
            log.Logger.Log(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType,
                Program.MyFineLevel, message, exception);
        }
        public static void Fine(this ILog log, string message) {
            log.Fine(message, exception:null);
        }
        public static void Fine(this ILog log, string message, String containsEmail) {
            if (Settings.Instance.LoggingLevel != "ULTRA-FINE") {
                message = message.Replace(containsEmail, EmailAddress.maskAddress(containsEmail));
            }
            log.Fine(message);
        }
        public static Boolean IsFineEnabled(this ILog log) {
            return log.Logger.IsEnabledFor(Program.MyFineLevel);
        }

        private static void UltraFine(this ILog log, string message, System.Exception exception) {
            log.Logger.Log(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType,
                Program.MyUltraFineLevel, message, exception);
        }
        public static void UltraFine(this ILog log, string message) {
            log.UltraFine(message, null);
        }
    }
}
