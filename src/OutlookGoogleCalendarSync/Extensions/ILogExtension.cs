using System;
using System.Windows.Forms;
using log4net;
using log4net.Core;

namespace OutlookGoogleCalendarSync {

    public static class ILogExtensions {

        private static void Fine(this ILog log, string message, System.Exception exception) {
            log.Logger.Log(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType,
                Program.MyFineLevel, message, exception);
        }
        public static void Fine(this ILog log, string message) {
            log.Fine(message, exception:null);
        }
        public static void Fine(this ILog log, string message, String containsEmail) {
            if (Settings.Instance.LoggingLevel != "ULTRA-FINE" && !string.IsNullOrEmpty(containsEmail)) {
                message = message.Replace(containsEmail, EmailAddress.MaskAddress(containsEmail));
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

    public class ErrorFlagAppender : log4net.Appender.AppenderSkeleton {
        private Boolean errorOccurred = false;

        /// <summary>
        /// When an error is logged, check if user has chosen to upload logs or not
        /// </summary>
        protected override void Append(LoggingEvent loggingEvent) {
            if (errorOccurred) return;
            errorOccurred = true;
            if (Settings.Instance.IsLoaded && Settings.Instance.CloudLogging != null) return;
            if (!Settings.Instance.IsLoaded && XMLManager.ImportElement("CloudLogging", Settings.ConfigFile) != "") return;

            Forms.CloudLogging frm = new Forms.CloudLogging();
            DialogResult dr = frm.ShowDialog();
            if (dr == DialogResult.Cancel) {
                errorOccurred = false;
                return;
            }
            Settings.Instance.CloudLogging = dr == DialogResult.Yes;
            if (!Settings.Instance.IsLoaded) XMLManager.ExportElement("CloudLogging", Settings.Instance.CloudLogging, Settings.ConfigFile);
            Analytics.Send(Analytics.Category.ogcs, Analytics.Action.setting, "CloudLogging=" + Settings.Instance.CloudLogging.ToString());

            Forms.Main.Instance.SetControlPropertyThreadSafe(Forms.Main.Instance.cbCloudLogging, "Checked", Settings.Instance.CloudLogging);
        }
    }
}
