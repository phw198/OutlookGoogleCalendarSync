using Google.Apis.Calendar.v3.Data;
using log4net;
using Microsoft.Office.Interop.Outlook;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

namespace OutlookGoogleCalendarSync.OutlookOgcs {
    class CustomProperty {
        private static readonly ILog log = LogManager.GetLogger(typeof(CustomProperty));

        private static String calendarKeyName = metadataIdKeyName(MetadataId.gCalendarId);

        /// <summary>
        /// These properties can be stored multiple times against a single calendar item.
        /// The first default set is NOT appended with a number
        /// Subsequent sets are appended with "-<2-digit-sequence>" - eg "googleCalendarID-02"
        /// </summary>
        public enum MetadataId {
            gEventID,
            gCalendarId,
            ogcsModified,
            forceSave,
            locallyCopied
        }

        /// <summary>
        /// The name of the keys as held in the custom attribute.
        /// Names can be stored with numbers appended to support syncing the same object between multiple calendars.
        /// CalendarID is the master keyname to determine an ID set number.
        /// </summary>
        private static String metadataIdKeyName(MetadataId Id) {
            switch (Id) {
                case MetadataId.gEventID: return "googleEventID";
                case MetadataId.gCalendarId: return "googleCalendarID";
                case MetadataId.ogcsModified: return "OGCSmodified";
                case MetadataId.forceSave: return "forceSave";
                default: return Id.ToString();
            }
        }

        /// <summary>
        /// Return number appended to key name for current calendar key
        /// </summary>
        /// <param name="maxSet">The set number of the last contiguous run of ID sets (to aid defragmentation).</param>
        /// <returns>The set number, if it exists</returns>
        private static int? getKeySet(AppointmentItem ai, out int maxSet) {
            String returnSet = "";
            maxSet = 0;
            Dictionary<String, String> calendarKeys = new Dictionary<string, string>();
            UserProperties ups = null;
            try {
                ups = ai.UserProperties;
                for (int p = 1; p <= ups.Count; p++) {
                    UserProperty up = null;
                    try {
                        up = ups[p];
                        if (up.Name.StartsWith(calendarKeyName))
                            calendarKeys.Add(up.Name, up.Value.ToString());
                    } finally {
                        up = (UserProperty)OutlookOgcs.Calendar.ReleaseObject(up);
                    }
                }
            } finally {
                ups = (UserProperties)OutlookOgcs.Calendar.ReleaseObject(ups);
            }

            //For backward compatibility, always default to key names with no set number appended
            if (!calendarKeys.ContainsKey(calendarKeyName) ||
                (calendarKeys.Count == 1 && calendarKeys.ContainsKey(calendarKeyName) && calendarKeys[calendarKeyName] == Settings.Instance.UseGoogleCalendar.Id))
            {
                maxSet = -1;
                return null;
            }

            foreach (KeyValuePair<String, String> kvp in calendarKeys.OrderBy(k => k.Key)) {
                Regex rgx = new Regex("^" + calendarKeyName + "-*(\\d{0,2})", RegexOptions.IgnoreCase);
                MatchCollection matches = rgx.Matches(kvp.Key);

                if (matches.Count > 0) {
                    int appendedNos = 0;
                    if (matches[0].Groups[1].Value != "")
                        appendedNos = Convert.ToInt16(matches[0].Groups[1].Value);
                    if (appendedNos - maxSet == 1) maxSet = appendedNos;
                    if (kvp.Value == Settings.Instance.UseGoogleCalendar.Id)
                        returnSet = matches[0].Groups[1].Value;
                }
            }

            if (string.IsNullOrEmpty(returnSet)) return null;
            else return Convert.ToInt16(returnSet);
        }

        public static Boolean GoogleIdMissing(AppointmentItem ai) {
            //Make sure Outlook appointment has all Google IDs stored
            String missingIds = "";
            if (!Exists(ai, MetadataId.gEventID)) missingIds += metadataIdKeyName(MetadataId.gEventID) + "|";
            if (!Exists(ai, MetadataId.gCalendarId)) missingIds += metadataIdKeyName(MetadataId.gCalendarId) + "|";
            if (!string.IsNullOrEmpty(missingIds))
                log.Warn("Found Outlook item missing Google IDs (" + missingIds.TrimEnd('|') + "). " + Calendar.GetEventSummary(ai));
            return !string.IsNullOrEmpty(missingIds);
        }

        public static Boolean Exists(AppointmentItem ai, MetadataId searchId) {
            String throwAway;
            return Exists(ai, searchId, out throwAway);
        }
        public static Boolean Exists(AppointmentItem ai, MetadataId searchId, out String searchKey) {
            searchKey = metadataIdKeyName(searchId);

            int maxSet;
            int? keySet = getKeySet(ai, out maxSet);
            if (keySet.HasValue) searchKey += "-" + keySet.Value.ToString("D2");

            UserProperties ups = null;
            UserProperty prop = null;
            try {
                ups = ai.UserProperties;
                prop = ups.Find(searchKey);
                if (searchId == MetadataId.gCalendarId)
                    return (prop != null && prop.Value.ToString() == Settings.Instance.UseGoogleCalendar.Id);
                else {
                    return (prop != null && Get(ai, MetadataId.gCalendarId) == Settings.Instance.UseGoogleCalendar.Id);
                }
            } catch {
                return false;
            } finally {
                prop = (UserProperty)OutlookOgcs.Calendar.ReleaseObject(prop);
                ups = (UserProperties)OutlookOgcs.Calendar.ReleaseObject(ups);
            }
        }

        public static Boolean ExistsAny(AppointmentItem ai) {
            if (Exists(ai, MetadataId.gEventID)) return true;
            if (Exists(ai, MetadataId.gCalendarId)) return true;
            return false;
        }

        /// <summary>
        /// Add the Google event IDs into Outlook appointment.
        /// </summary>
        public static void AddGoogleIDs(ref AppointmentItem ai, Event ev) {
            Add(ref ai, MetadataId.gEventID, ev.Id);
            Add(ref ai, MetadataId.gCalendarId, Settings.Instance.UseGoogleCalendar.Id);
            LogProperties(ai, log4net.Core.Level.Debug);
        }

        public static void Add(ref AppointmentItem ai, MetadataId key, String value) {
            add(ref ai, key, OlUserPropertyType.olText, value);
        }
        public static void Add(ref AppointmentItem ai, MetadataId key, DateTime value) {
            add(ref ai, key, OlUserPropertyType.olDateTime, value);
        }
        private static void add(ref AppointmentItem ai, MetadataId key, OlUserPropertyType keyType, object keyValue) {
            String addkeyName = metadataIdKeyName(key);

            UserProperties ups = null;
            try {
                if (!Exists(ai, key)) {
                    int newSet;
                    int? keySet = getKeySet(ai, out newSet);
                    keySet = keySet ?? newSet + 1;
                    if (keySet.HasValue && keySet.Value != 0) addkeyName += "-" + keySet.Value.ToString("D2");

                    try {
                        ups = ai.UserProperties;
                        ups.Add(addkeyName, keyType);
                    } catch (System.Exception ex) {
                        OGCSexception.Analyse(ex);
                        ups.Add(addkeyName, keyType, false);
                    } finally {
                        ups = (UserProperties)Calendar.ReleaseObject(ups);
                    }
                }
                ups = ai.UserProperties;
                ups[addkeyName].Value = keyValue;
                log.Fine("Set userproperty " + addkeyName + "=" + keyValue.ToString());

            } finally {
                ups = (UserProperties)Calendar.ReleaseObject(ups);
            }
        }

        public static String Get(AppointmentItem ai, MetadataId key) {
            String retVal = null;
            String searchKey;
            if (Exists(ai, key, out searchKey)) {
                UserProperties ups = null;
                UserProperty prop = null;
                try {
                    ups = ai.UserProperties;
                    prop = ups.Find(searchKey);
                    if (prop != null) {
                        if (prop.Type != OlUserPropertyType.olText) log.Warn("Non-string property " + searchKey + " being retrieved as String.");
                        retVal = prop.Value.ToString();
                    }
                } finally {
                    prop = (UserProperty)OutlookOgcs.Calendar.ReleaseObject(prop);
                    ups = (UserProperties)OutlookOgcs.Calendar.ReleaseObject(ups);
                }
            }
            return retVal;
        }
        private static DateTime get_datetime(AppointmentItem ai, MetadataId key) {
            DateTime retVal = new DateTime();
            String searchKey;
            if (Exists(ai, key, out searchKey)) {
                UserProperties ups = null;
                UserProperty prop = null;
                try {
                    ups = ai.UserProperties;
                    prop = ups.Find(searchKey);
                    if (prop != null) {
                        try {
                            if (prop.Type != OlUserPropertyType.olDateTime) {
                                log.Warn("Non-datetime property " + searchKey + " being retrieved as DateTime.");
                                retVal = DateTime.Parse(prop.Value.ToString());
                            } else
                                retVal = (DateTime)prop.Value;
                        } catch (System.Exception ex) {
                            log.Error("Failed to retrieve DateTime value for property " + searchKey);
                            OGCSexception.Analyse(ex);
                        }
                    }
                } finally {
                    prop = (UserProperty)Calendar.ReleaseObject(prop);
                    ups = (UserProperties)Calendar.ReleaseObject(ups);
                }
            }
            return retVal;
        }

        public static void RemoveAll(ref AppointmentItem ai) {
            Remove(ref ai, MetadataId.gEventID);
            Remove(ref ai, MetadataId.gCalendarId);
            Remove(ref ai, MetadataId.forceSave);
            Remove(ref ai, MetadataId.locallyCopied);
            Remove(ref ai, MetadataId.ogcsModified);
        }
        public static void Remove(ref AppointmentItem ai, MetadataId key) {
            String searchKey;
            if (Exists(ai, key, out searchKey)) {
                UserProperties ups = null;
                UserProperty prop = null;
                try {
                    ups = ai.UserProperties;
                    prop = ups.Find(searchKey);
                    prop.Delete();
                    log.Debug("Removed " + searchKey + " property.");
                } finally {
                    prop = (UserProperty)Calendar.ReleaseObject(prop);
                    ups = (UserProperties)Calendar.ReleaseObject(ups);
                }
            }
        }

        public static DateTime GetOGCSlastModified(AppointmentItem ai) {
            return get_datetime(ai, MetadataId.ogcsModified);
        }
        public static void SetOGCSlastModified(ref AppointmentItem ai) {
            Add(ref ai, MetadataId.ogcsModified, DateTime.Now);
        }

        /// <summary>
        /// Log the various User Properties.
        /// </summary>
        /// <param name="ai">The Appointment item.</param>
        /// <param name="thresholdLevel">Only log if logging configured at this level or higher.</param>
        public static void LogProperties(AppointmentItem ai, log4net.Core.Level thresholdLevel) {
            if (((log4net.Repository.Hierarchy.Hierarchy)LogManager.GetRepository()).Root.Level.Value > thresholdLevel.Value) return;

            UserProperties ups = null;
            UserProperty up = null;
            try {
                log.Debug(OutlookOgcs.Calendar.GetEventSummary(ai));
                ups = ai.UserProperties;
                for (int p = 1; p <= ups.Count; p++) {
                    try {
                        up = ups[p];
                        log.Debug(up.Name + "=" + up.Value.ToString());
                    } finally {
                        up = (UserProperty)OutlookOgcs.Calendar.ReleaseObject(up);
                    }
                }
            } catch (System.Exception ex) {
                OGCSexception.Analyse("Failed to log Appointment UserProperties", ex);
            } finally {
                ups = (UserProperties)OutlookOgcs.Calendar.ReleaseObject(ups);
            }
        }
    }
}
