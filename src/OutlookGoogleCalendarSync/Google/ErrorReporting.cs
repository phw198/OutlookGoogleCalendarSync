using Ogcs = OutlookGoogleCalendarSync;
using log4net;
using log4net.Appender;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;

namespace OutlookGoogleCalendarSync.Google {
    class ErrorReporting {
        private static readonly ILog log = LogManager.GetLogger(typeof(ErrorReporting));

        public static Boolean Initialised = true;
        public static Boolean ErrorOccurred = false;
        private static String templateCredFile = Path.Combine(System.Windows.Forms.Application.StartupPath, "ErrorReportingTemplate.json");
        private static String credFile = Path.Combine(System.Windows.Forms.Application.StartupPath, "ErrorReporting.json");

        public static void Initialise() {
            if (Program.StartedWithSquirrelArgs && !(Environment.GetCommandLineArgs()[1].ToLower().Equals("--squirrel-firstrun"))) return;
            if (Program.InDeveloperMode) return;

            //Note, logging isn't actually initialised yet, so log4net won't log any lines within this function

            String cloudCredsURL = "https://raw.githubusercontent.com/phw198/OutlookGoogleCalendarSync/master/docs/keyring.md";
            String html = null;
            String line = null;
            String placeHolder = "###";
            String cloudID = null;
            String cloudKey = null;

            log.Debug("Getting credential attributes");
            try {
                try {
                    html = new Extensions.OgcsWebClient().DownloadString(cloudCredsURL);
                    html = html.Replace("\n", "");
                } catch (System.Exception ex) {
                    log.Error("Failed to retrieve data: " + ex.Message);
                }

                if (string.IsNullOrEmpty(html)) {
                    throw new ApplicationException("Not able to retrieve error reporting credentials.");
                }

                Regex rgx = new Regex(@"### Error Reporting.*\|ID\|(.*)\|\|Key\|(.*?)\|", RegexOptions.IgnoreCase);
                MatchCollection keyRecords = rgx.Matches(html);
                if (keyRecords.Count == 1) {
                    cloudID = keyRecords[0].Groups[1].ToString();
                    cloudKey = keyRecords[0].Groups[2].ToString();
                } else
                    throw new ApplicationException("Unexpected parse of error reporting credentials.");

                List<String> newLines = new List<string>();
                StreamReader sr = new StreamReader(templateCredFile);
                while ((line = sr.ReadLine()) != null) {
                    if (line.IndexOf(placeHolder) > 0) {
                        if (line.IndexOf("private_key_id") > 0) {
                            line = line.Replace(placeHolder, cloudID);

                        } else if (line.IndexOf("private_key") > 0) {
                            line = line.Replace(placeHolder, cloudKey);
                        }
                    }
                    newLines.Add(line);
                }
                try {
                    File.WriteAllLines(credFile, newLines.ToArray());
                } catch (System.IO.IOException ex) {
                    if (ex.GetErrorCode() == "0x80070020")
                        log.Warn("ErrorReporting.json is being used by another process (perhaps multiple instances of OGCS are being started on system startup?)");
                    else
                        throw;
                }
                Environment.SetEnvironmentVariable("GOOGLE_APPLICATION_CREDENTIALS", credFile);

            } catch (ApplicationException ex) {
                log.Warn(ex.Message);
                Initialised = false;

            //} catch (System.Exception ex) {
                //Logging isn't initialised yet, so don't catch this error - let it crash out so user is aware and hopefully reports it!
                //Ogcs.Extensions.MessageBox.Show(ex.Message);
                //log.Debug("Failed to initialise error reporting.");
                //Ogcs.Exception.Analyse(ex);
            }
        }            

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
        public static void UpdateLogUuId() {
            logUuid = Telemetry.Instance.UpdateAnonymousUniqueUserId();
        }

        private static log4net.Appender.BufferingForwardingAppender getAppender() {
            log4net.Appender.BufferingForwardingAppender cloudLogger = (log4net.Appender.BufferingForwardingAppender)LogManager.GetRepository().GetAppenders().Where(a => a.Name == "CloudLogger").FirstOrDefault();
            if (cloudLogger == null) {
                log.Warn("Could not find CloudLogger appender.");
                return null;
            }
            return cloudLogger;
        }

        public static log4net.Core.Level GetThreshold() {
            log4net.Appender.BufferingForwardingAppender cloudLogger = getAppender();
            if (cloudLogger == null) return null;
            else return cloudLogger.Threshold;
        }

        public static void SetThreshold(Boolean cloudLoggingEnabled) {
            try {
                log4net.Appender.BufferingForwardingAppender cloudLogger = getAppender();
                if (cloudLogger == null) return;

                if (cloudLoggingEnabled) {
                    if (cloudLogger.Threshold != log4net.Core.Level.All) {
                        cloudLogger.Threshold = log4net.Core.Level.All;
                        log.Info("Turned error reporting ON. Anonymous ID: "+ LogUuid);
                    }
                } else {
                    if (cloudLogger.Threshold != log4net.Core.Level.Off) {
                        if (cloudLogger.Threshold == null) log.Info("Initialising error reporting to OFF");
                        else log.Info("Turned error reporting OFF");
                        cloudLogger.Threshold = log4net.Core.Level.Off;
                    }
                }

            } catch (System.Exception ex) {
                log.Error("Failed to configure error reporting appender.");
                Ogcs.Exception.Analyse(ex);
            }
        }

        public class LevelRewritingAppender : ForwardingAppender {
            
            //Map unusual levels to Google Cloud Logging native levels
            protected Dictionary<String, log4net.Core.Level> levelMap = new Dictionary<String, log4net.Core.Level> { 
                { "FAIL", log4net.Core.Level.Warn },
                { "FINE", log4net.Core.Level.Debug },
                { "ULTRA-FINE", log4net.Core.Level.Debug } //This doesn't seem to get passed, even with no filter defined
            };

            // This method is called for every event that passes the parent filters, by default up to DEBUG
            protected override void Append(log4net.Core.LoggingEvent loggingEvent) {

                if (levelMap.Keys.Contains(loggingEvent.Level.Name)) {
                    // Create a new LoggingEvent with a standard, mapped level.
                    log4net.Core.LoggingEventData transposedEventData = loggingEvent.GetLoggingEventData();
                    transposedEventData.Message = loggingEvent.Level.Name + ": " + transposedEventData.Message;
                    transposedEventData.Level = levelMap[loggingEvent.Level.Name];
                    
                    log4net.Core.LoggingEvent transposedEvent= new log4net.Core.LoggingEvent(transposedEventData);
                    base.Append(transposedEvent);

                } else {
                    // For all other levels, pass the original event
                    base.Append(loggingEvent);
                }
            }
        }
    }
}
