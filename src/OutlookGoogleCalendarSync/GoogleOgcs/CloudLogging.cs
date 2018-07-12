using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Management;
using log4net;

namespace OutlookGoogleCalendarSync.GoogleOgcs {
    class CloudLogging {
        private static readonly ILog log = LogManager.GetLogger(typeof(CloudLogging));

        public static String LogId {
            set {
                log4net.GlobalContext.Properties["CloudLogId"] = value;
            }
        }

        public static String logUuid {
            set {
                log4net.GlobalContext.Properties["CloudLogUuid"] = value;
                ((log4net.Repository.Hierarchy.Hierarchy)LogManager.GetRepository()).RaiseConfigurationChanged(EventArgs.Empty);
            }
        }
        public static String LogUuid {
            get {
                return log4net.GlobalContext.Properties["CloudLogUuid"].ToString();
            }            
        }

        /// <summary>
        /// Set cloud logger to include unique ID in each log line
        /// </summary>
        /// <returns>MD5 hash to correlate log entries to distinct, anonymous user</returns>
        public static void UpdateLogUuId() {
            String uuid = null;
            try {
                //Check if Settings have been loaded yet and has Gmail account set
                if (Settings.Instance.IsLoaded && !string.IsNullOrEmpty(Settings.Instance.GaccountEmail)) {
                    logUuid = GoogleOgcs.Authenticator.GetMd5(Settings.Instance.GaccountEmail, true);

                } else { //Check if the raw settings file has Gmail account set
                    String gmailAccount = null;
                    try {
                        gmailAccount = XMLManager.ImportElement("GaccountEmail", Settings.ConfigFile, false);
                    } catch { }

                    if (!string.IsNullOrEmpty(gmailAccount)) {
                        logUuid =  GoogleOgcs.Authenticator.GetMd5(gmailAccount, true);
                    } else {
                        //Make a "unique" string based on:
                        //ComputerName;Processor;C-driveSerial
                        ManagementClass mc = new ManagementClass("win32_processor");
                        ManagementObjectCollection moc = mc.GetInstances();
                        foreach (ManagementObject mo in moc) {
                            uuid = mo.Properties["SystemName"].Value.ToString();
                            uuid += ";" + mo.Properties["Name"].Value.ToString();
                            break;
                        }
                        String drive = "C";
                        ManagementObject dsk = new ManagementObject(@"win32_logicaldisk.deviceid=""" + drive + @":""");
                        dsk.Get();
                        String volumeSerial = dsk["VolumeSerialNumber"].ToString();
                        uuid += ";" + volumeSerial;

                        logUuid = GoogleOgcs.Authenticator.GetMd5(uuid);
                    }
                }

            } catch {
                Random random = new Random();
                logUuid = random.Next().ToString();
            }
        }

        public static void SetThreshold(Boolean cloudLoggingEnabled) {
            try {
                log4net.Appender.BufferingForwardingAppender cloudLogger = (log4net.Appender.BufferingForwardingAppender)LogManager.GetRepository().GetAppenders().Where(a => a.Name == "CloudLogger").FirstOrDefault();
                if (cloudLogger == null) {
                    log.Warn("Could not find CloudLogger appender.");
                }

                if (cloudLoggingEnabled)
                    cloudLogger.Threshold = log4net.Core.Level.All;
                else
                    cloudLogger.Threshold = log4net.Core.Level.Off;

                log.Info("Turned cloud logging " + (cloudLoggingEnabled ? "ON" : "OFF"));

            } catch (System.Exception ex) {
                log.Error("Failed to configure cloud logging appender.");
                OGCSexception.Analyse(ex);
            }
        }
    }
}
