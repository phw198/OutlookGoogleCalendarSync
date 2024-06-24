﻿using Ogcs = OutlookGoogleCalendarSync;
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
        private const String tzdbFilename = "tzdb.nzd";
        private String tzdbFile {
            get { return Path.Combine(Path.GetDirectoryName(System.Windows.Forms.Application.ExecutablePath), tzdbFilename); }
        }

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
        
        public static DateTime LastSystemResume = DateTime.Now.AddDays(-1);

        private TimezoneDB() {
            try {
                using (Stream stream = File.OpenRead(tzdbFile)) {
                    source = TzdbDateTimeZoneSource.FromStream(stream);
                }
            } catch {
                log.Warn("Custom TZDB source failed. Falling back to NodaTime.dll");
                source = TzdbDateTimeZoneSource.Default;
            }
            log.Info("Using NodaTime " + source.VersionId);
            this.RevertKyiv = false;

            Microsoft.Win32.SystemEvents.TimeChanged += SystemEvents_TimeChanged;
        }

        private static void SystemEvents_TimeChanged(object sender, EventArgs e) {
            log.Info("Detected system timezone change.");
            System.Globalization.CultureInfo.CurrentCulture.ClearCachedData();
            LastSystemResume = DateTime.Now;
        }

        public void CheckForUpdate() {
            System.Threading.Thread updateDBthread = new System.Threading.Thread(x => checkForUpdate(source.TzdbVersion));
            updateDBthread.Start();
        }
        private void checkForUpdate(String localVersion) {
            try {
                if (Program.InDeveloperMode && File.Exists(tzdbFile)) return;

                log.Debug("Checking for new timezone database...");
                String nodatimeURL = "http://nodatime.org/tzdb/latest.txt";
                String html = "";
                Extensions.OgcsWebClient wc = new Extensions.OgcsWebClient();
                try {
                    html = wc.DownloadString(nodatimeURL);
                } catch (System.Exception ex) {
                    log.Error("Failed to get latest NodaTime db version.");
                    Ogcs.Exception.Analyse(ex);
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
                                    wc.DownloadFile(html, tzdbFile);
                                    log.Debug("New TZDB version downloaded - disposing of reference to old db data.");
                                    instance = null;
                                } catch (System.Exception ex) {
                                    log.Error("Failed to download new TZDB database from " + html);
                                    Ogcs.Exception.Analyse(ex);
                                }
                            }
                        } else {
                            log.Warn("Regex to extract latest version is no longer working!");
                        }
                    }
                }
            } catch (System.Exception ex) {
                ex.Analyse("Could not check for timezone data update.");
            }
        }

        /// <summary>
        /// Alexa (Amazon Echo) is a bit dumb - she creates Google Events with a GMT offset "timezone". Eg GMT-5
        /// This isn't actually a timezone at all, but an area, and not a legal IANA value.
        /// So to workaround this, we'll turn it into something valid at least, by inverting the offset sign and prefixing "Etc\"
        /// </summary>
        public static String FixAlexa(String timezone) {
            //Issues:- 
            // * As it's an area, Microsoft will just guess at the zone - so GMT-5 for CST may end up as Bogata/Lima.
            // * Not sure what happens with half hour offset, such as in India with GMT+4:30
            // * Not sure what happens with Daylight Saving, as zones in the same area may or may not follow DST.

            try {
                Regex rgx = new Regex(@"^GMT([+-])(\d{1,2})(:\d\d)*$");
                MatchCollection matches = rgx.Matches(timezone);
                if (matches.Count > 0) {
                    log.Debug("Found an Alexa \"timezone\" of " + timezone);
                    String fixedTimezone = "Etc/GMT" + (matches[0].Groups[1].Value == "+" ? "-" : "+") + Convert.ToInt16(matches[0].Groups[2].Value).ToString();
                    log.Debug("Translated to " + fixedTimezone);
                    return fixedTimezone;
                }
            } catch (System.Exception ex) {
                ex.Analyse("Failed to detect and translate Alexa timezone: " + timezone);
            }
            return timezone;
        }

        /// <summary>
        /// Sometime an Outlook timezone name contains a GMT offset, which isn't valid.
        /// </summary>
        /// <param name="outlookTimezone">The string stored in Outlook by appointment organiser, eg "(GMT+10:00) AUS Eastern Standard Time"</param>
        /// <returns>Offset, if present</returns>
        public static Int16? GetTimezoneOffset(String outlookTimezone) {
            //timezone = "(GMT+10:00) AUS Eastern Standard Time"; <== WebEx is known to store timezones like this.
            try {
                Regex rgx = new Regex(@"^\((GMT|UTC)([+-]\d{1,2})*:*\d{0,2}\)\s.*$");
                MatchCollection matches = rgx.Matches(outlookTimezone);
                if (matches != null && matches.Count > 0) {
                    String gmtOffset_str = matches[0].Groups[2].Value.Trim();
                    if (string.IsNullOrEmpty(gmtOffset_str)) return 0;
                    Int16 gmtOffset = Convert.ToInt16(gmtOffset_str);
                    log.Debug("Found a " + matches[0].Groups[1].Value.ToString() + " timezone offset of " + gmtOffset);
                    return gmtOffset;
                }
            } catch (System.Exception ex) {
                ex.Analyse("Failed to detect any timezone offset for: " + outlookTimezone);
            }
            return null;
        }

        /// <summary>
        /// Convert IANA timezone to a UTC offset.
        /// </summary>
        /// <param name="IanaTimezone">eg "America/Vancouver"</param>
        /// <returns>The offset eg -7</returns>
        public static Int16 GetUtcOffset(String IanaTimezone) {
            Int16 utcOffset = 0;
            try {
                NodaTime.IDateTimeZoneProvider tzProvider = NodaTime.DateTimeZoneProviders.Tzdb;
                if (!tzProvider.Ids.Contains(IanaTimezone)) {
                    log.Warn("Could not map IANA timezone '" + IanaTimezone + "' to UTC offset.");
                } else {
                    NodaTime.DateTimeZone tz = tzProvider[IanaTimezone];
                    NodaTime.Instant instant = NodaTime.Instant.FromDateTimeUtc(DateTime.UtcNow);
                    NodaTime.Offset offset = tz.GetUtcOffset(instant);
                    utcOffset = Convert.ToInt16(offset.Seconds / 3600);
                }
            } catch (System.Exception ex) {
                ex.Analyse("Not able to convert IANA timezone '" + IanaTimezone + "' to UTC offset.");
            }
            return utcOffset;
        }

        /// <summary>
        /// IANA updated in Aug-2022 to Europe/Kyiv, but Google is still on Europe/Kiev.
        /// Once Google catches up, this can be removed.
        /// </summary>
        public Boolean RevertKyiv { get; set; }
    }
}
