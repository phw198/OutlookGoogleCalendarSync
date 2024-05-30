using Ogcs = OutlookGoogleCalendarSync;
using log4net;
using log4net.Core;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Windows.Forms;

namespace OutlookGoogleCalendarSync {

    public static class ILogExtensions {

        #region Fail
        private static void Fail(this ILog log, string message, System.Exception exception) {
            log.Logger.Log(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType,
                Program.MyFailLevel, message, exception);
        }
        public static void Fail(this ILog log, string message) {
            log.Fail(message, null);
        }
        #endregion

        #region Fine
        private static void Fine(this ILog log, string message, System.Exception exception) {
            log.Logger.Log(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType,
                Program.MyFineLevel, message, exception);
        }
        public static void Fine(this ILog log, string message) {
            log.Fine(message, exception: null);
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
        #endregion

        #region UltraFine
        private static void UltraFine(this ILog log, string message, System.Exception exception) {
            log.Logger.Log(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType,
                Program.MyUltraFineLevel, message, exception);
        }
        public static void UltraFine(this ILog log, string message) {
            log.UltraFine(message, null);
        }
        public static Boolean IsUltraFineEnabled(this ILog log) {
            return log.Logger.IsEnabledFor(Program.MyUltraFineLevel);
        }
        #endregion

        /// <summary>
        /// Log a message at either of these levels
        /// </summary>
        /// <param name="log"></param>
        /// <param name="message"></param>
        /// <param name="level">The level to log the message at</param>
        public static void ErrorOrFail(this ILog log, String message, log4net.Core.Level level) {
            if (level == Program.MyFailLevel) log.Fail(message);
            else log.Error(message);
        }
    }

    public class ErrorFlagAppender : log4net.Appender.AppenderSkeleton {
        private static readonly ILog log = LogManager.GetLogger(typeof(ErrorFlagAppender));

        /// <summary>
        /// When an error is logged, check if user has chosen to upload logs or not
        /// </summary>
        protected override void Append(LoggingEvent loggingEvent) {
            if (!Ogcs.Google.ErrorReporting.Initialised || Ogcs.Google.ErrorReporting.ErrorOccurred) return;
            Ogcs.Google.ErrorReporting.ErrorOccurred = true;
            String configSetting = null;

            if (Settings.AreLoaded) configSetting = Settings.Instance.CloudLogging.ToString();
            else configSetting = XMLManager.ImportElement("CloudLogging", Settings.ConfigFile);

            if (!string.IsNullOrEmpty(configSetting)) {
                if (Convert.ToBoolean(configSetting) && Ogcs.Google.ErrorReporting.GetThreshold().ToString().ToUpper() != "ALL") {
                    Ogcs.Google.ErrorReporting.SetThreshold(true);
                    replayLogs();
                } else if (!Convert.ToBoolean(configSetting) && Ogcs.Google.ErrorReporting.GetThreshold().ToString().ToUpper() != "OFF") {
                    Ogcs.Google.ErrorReporting.SetThreshold(false);
                }
                return;
            }

            //Cloud logging value not set yet - let's ask the user
            Forms.ErrorReporting frm = Forms.ErrorReporting.Instance;
            DialogResult dr = frm.ShowDialog();
            if (dr == DialogResult.Cancel) {
                Ogcs.Google.ErrorReporting.ErrorOccurred = false;
                return;
            }
            Boolean confirmative = dr == DialogResult.Yes;
            if (Settings.AreLoaded) Settings.Instance.CloudLogging = confirmative;
            Telemetry.Send(Analytics.Category.ogcs, Analytics.Action.setting, "CloudLogging=" + confirmative.ToString());
            Telemetry.GA4Event.Event settingGa4Ev = new(Telemetry.GA4Event.Event.Name.setting);
            settingGa4Ev.AddParameter("cloud_logging", confirmative);
            settingGa4Ev.Send();

            try {
                Forms.Main.Instance.SetControlPropertyThreadSafe(Forms.Main.Instance.cbCloudLogging, "CheckState", confirmative ? CheckState.Checked : CheckState.Unchecked);
            } catch { }

            if (confirmative) replayLogs();
        }

        /// <summary>
        /// Replay the logs that the CloudLogger appender did not buffer (because it was off)
        /// </summary>
        private void replayLogs() {
            try {
                String logFile = Path.Combine(log4net.GlobalContext.Properties["LogPath"].ToString(), log4net.GlobalContext.Properties["LogFilename"].ToString());
                List<String> lines = new List<String>();
                using (FileStream logFileStream = new FileStream(logFile, FileMode.Open, FileAccess.Read, FileShare.ReadWrite)) {
                    StreamReader logFileReader = new StreamReader(logFileStream);
                    while (!logFileReader.EndOfStream) {
                        lines.Add(logFileReader.ReadLine());
                    }
                }
                //"2018-07-14 17:22:41,740 DEBUG  10 OutlookGoogleCalendarSync.XMLManager [59] -  Retrieved setting 'CloudLogging' with value 'true'"
                //We want the logging level and the message strings
                Regex rgx = new Regex(@"^\d{4}-\d{2}-\d{2}\s[\d:,]+\s(\w+)\s+\d+\s[\w\.]+\s\[\d+\]\s-\s+(.*?)$", RegexOptions.IgnoreCase);
                foreach (String line in lines.Skip(lines.Count - 50).ToList()) {
                    MatchCollection matches = rgx.Matches(line);
                    if (matches.Count > 0) {
                        switch (matches[0].Groups[1].ToString()) {
                            case "FINE": log.Fine(matches[0].Groups[2].ToString()); break;
                            case "DEBUG": log.Debug(matches[0].Groups[2]); break;
                            case "INFO": log.Info(matches[0].Groups[2]); break;
                            case "WARN": log.Warn(matches[0].Groups[2]); break;
                            case "ERROR": log.Error(matches[0].Groups[2]); break;
                            default: log.Debug(matches[0].Groups[2]); break;
                        }

                    } else log.Debug(line);

                }
            } catch (System.Exception ex) {
                log.Warn("Failed to replay logs.");
                Ogcs.Exception.Analyse(ex);
            }
        }
    }
}
