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
            System.Net.WebClient wc = new System.Net.WebClient();
            wc.Headers.Add("user-agent", Settings.Instance.Proxy.BrowserUserAgent);
            try {
                html = wc.DownloadString(nodatimeURL);
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
                    log.Debug("Already have latest TZDB version.");
                } else {
                    Regex rgx = new Regex(@"https*:.*/tzdb(.*)\.nzd$", RegexOptions.IgnoreCase);
                    MatchCollection matches = rgx.Matches(html);
                    if (matches.Count > 0) {
                        String remoteVersion = matches[0].Result("$1");
                        if (string.Compare(localVersion, remoteVersion, System.StringComparison.InvariantCultureIgnoreCase) < 0) {
                            log.Debug("There is a new version " + remoteVersion);
                            try {
                                wc.DownloadFile(html, tzdbFilename);
                                log.Debug("New TZDB version downloaded - disposing of reference to old db data.");
                                instance = null;
                            } catch (System.Exception ex) {
                                log.Error("Failed to download new TZDB database from " + html);
                                OGCSexception.Analyse(ex);
                            }
                        }
                    } else {
                        log.Warn("Regex to extract latest version is no longer working!");
                    }
                }
            }
        }

        public static String FixAlexa(String timezone) {
            //Alexa (Amazon Echo) is a bit dumb - she creates Google Events with a GMT offset "timezone". Eg GMT-5
            //This isn't actually a timezone at all, but an area, and not a legal IANA value.
            //So to workaround this, we'll turn it into something valid at least, by inverting the offset sign and prefixing "Etc\"
            //Issues:- 
            // * As it's an area, Microsoft will just guess at the zone - so GMT-5 for CST may end up as Bogata/Lima.
            // * Not sure what happens with half hour offset, such as in India with GMT+4:30
            // * Not sure what happens with Daylight Saving, as zones in the same area may or may not follow DST.

            try {
                System.Text.RegularExpressions.Regex rgx = new System.Text.RegularExpressions.Regex(@"^GMT([+-])(\d{1,2})(:\d\d)*$");
                System.Text.RegularExpressions.MatchCollection matches = rgx.Matches(timezone);
                if (matches.Count > 0) {
                    log.Debug("Found an Alexa \"timezone\" of " + timezone);
                    String fixedTimezone = "Etc/GMT" + (matches[0].Groups[1].Value == "+" ? "-" : "+") + Convert.ToInt16(matches[0].Groups[2].Value).ToString();
                    log.Debug("Translated to " + fixedTimezone);
                    return fixedTimezone;
                }
            } catch (System.Exception ex) {
                log.Error("Failed to detect and translate Alexa timezone.");
                OGCSexception.Analyse(ex);
            }
            return timezone;
        }
    }
}
