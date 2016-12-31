using log4net;
using NodaTime.TimeZones;
using System;
using System.IO;
using System.Text.RegularExpressions;

namespace OutlookGoogleCalendarSync {
    public class TimezoneDB {
        private static TimezoneDB instance;
        private static readonly ILog log = LogManager.GetLogger(typeof(NotificationTray));
        private TzdbDateTimeZoneSource source;
        private String tzdbFilename = "tzdb.nzd";

        public static TimezoneDB Instance {
            get {   
                if (instance == null) instance = new TimezoneDB();
                return instance;
            }
        }

        public TzdbDateTimeZoneSource Source {
            get { return source; }
        }
        public String Version {
            get { return source.TzdbVersion; }
        }
        
        private TimezoneDB() {
            try {
                using (Stream stream = File.OpenRead(tzdbFilename)) {
                    source = TzdbDateTimeZoneSource.FromStream(stream);
                }
            } catch {
                log.Warn("Custom TZDB source failed. Falling back to NodaTime.dll");
                source = TzdbDateTimeZoneSource.Default;
            }
            log.Info("Using NodaTime "+ source.VersionId);
        }

        public void CheckForUpdate() {
            System.Threading.Thread updateDBthread = new System.Threading.Thread(x => checkForUpdate(source.TzdbVersion));
            updateDBthread.Start();
        }        
        private void checkForUpdate(String localVersion) {
            if (System.Diagnostics.Debugger.IsAttached && File.Exists(tzdbFilename)) return;

            log.Debug("Checking for new timezone database...");
            String nodatimeURL = "http://nodatime.org/tzdb/latest.txt";
            String html = "";
            try {
                html = new System.Net.WebClient().DownloadString(nodatimeURL);
            } catch (System.Exception ex) {
                log.Error("Failed to get latest NodaTime db version.");
                OGCSexception.Analyse(ex);
                return;
            }

            if (string.IsNullOrEmpty(html)) {
                log.Warn("Empty response from " + nodatimeURL);
            } else {
                html = html.TrimEnd('\r', '\n');
                if (html.EndsWith(localVersion + ".nzd")) {
                    log.Debug("Already have latest version.");
                } else {
                    Regex rgx = new Regex(@"https*:.*/tzdb(.*)\.nzd$", RegexOptions.IgnoreCase);
                    MatchCollection matches = rgx.Matches(html);
                    if (matches.Count > 0) {
                        String remoteVersion = matches[0].Result("$1");
                        if (string.Compare(localVersion, remoteVersion, System.StringComparison.InvariantCultureIgnoreCase) < 0) {
                            log.Debug("There is a new version " + remoteVersion);
                            try {
                                new System.Net.WebClient().DownloadFile(html, tzdbFilename);
                                log.Debug("New version downloaded - disposing of reference to old db data.");
                                instance = null;
                            } catch (System.Exception ex) {
                                log.Error("Failed to download new database from " + html);
                                OGCSexception.Analyse(ex);
                            }
                        }
                    } else {
                        log.Warn("Regex to extract latest version is no longer working!");
                    }
                }
            }
        }
    }
}
