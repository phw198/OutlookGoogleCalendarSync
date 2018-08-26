using Google.Apis.Calendar.v3.Data;
using log4net;
using Microsoft.Office.Interop.Outlook;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

namespace OutlookGoogleCalendarSync.GoogleOgcs {
    class CustomProperty {
        private static readonly ILog log = LogManager.GetLogger(typeof(CustomProperty));

        private static String calendarKeyName = metadataIdKeyName(MetadataId.oCalendarId);

        /// <summary>
        /// These properties can be stored multiple times against a single calendar item.
        /// The first default set is NOT appended with a number
        /// Subsequent sets are appended with "-<2-digit-sequence>" - eg "outlook_CalendarID-02"
        /// </summary>
        public enum MetadataId {
            oEntryId,
            oGlobalApptId,
            oCalendarId,
            ogcsModified,
            apiLimitHit,
            forceSave
        }

        /// <summary>
        /// The name of the keys as held in the custom attribute.
        /// Names can be stored with numbers appended to support syncing the same object between multiple calendars.
        /// CalendarID is the master keyname to determine an ID set number.
        /// </summary>
        private static String metadataIdKeyName(MetadataId id) {
            switch (id) {
                case MetadataId.oEntryId: return "outlook_EntryID";
                case MetadataId.oGlobalApptId: return "outlook_GlobalApptID";
                case MetadataId.oCalendarId: return "outlook_CalendarID";
                case MetadataId.ogcsModified: return "OGCSmodified";
                case MetadataId.apiLimitHit: return "APIlimitHit";
                case MetadataId.forceSave: return "forceSave";
                default: return "outlook_EntryID";
            }
        }

        /// <summary>
        /// Return number appended to key name for current calendar key
        /// </summary>
        /// <param name="maxSet">The set number of the last contiguous run of ID sets (to aid defragmentation).</param>
        /// <returns>The set number, if it exists</returns>
        private static int? getKeySet(Event ev, out int maxSet) {
            String returnSet = "";
            maxSet = 0;
            Dictionary<String, String> calendarKeys = ev.ExtendedProperties.Private__.Where(k => k.Key.StartsWith(calendarKeyName)).OrderBy(k => k.Key).ToDictionary(k => k.Key, k => k.Value);

            //For backward compatibility, always default to key names with no set number appended
            if (!calendarKeys.ContainsKey(calendarKeyName) ||
                (calendarKeys.Count == 1 && calendarKeys.ContainsKey(calendarKeyName)) && calendarKeys[calendarKeyName] == OutlookOgcs.Calendar.Instance.UseOutlookCalendar.EntryID)
            {
                maxSet = -1;
                return null;
            }

            foreach (KeyValuePair<String, String> kvp in calendarKeys) {
                Regex rgx = new Regex("^" + calendarKeyName + "-*(\\d{0,2})", RegexOptions.IgnoreCase);
                MatchCollection matches = rgx.Matches(kvp.Key);

                if (matches.Count > 0) {
                    int appendedNos = 0;
                    if (matches[0].Groups[1].Value != "")
                        appendedNos = Convert.ToInt16(matches[0].Groups[1].Value);
                    if (appendedNos - maxSet == 1) maxSet = appendedNos;
                    if (kvp.Value == OutlookOgcs.Calendar.Instance.UseOutlookCalendar.EntryID)
                        returnSet = matches[0].Groups[1].Value;
                }
            }

            if (string.IsNullOrEmpty(returnSet)) return null;
            else return Convert.ToInt16(returnSet);
        }

        /// <summary>
        /// Check if any Outlook IDs are missing.
        /// </summary>
        public static Boolean OutlookIdMissing(Event ev) {
            //Make sure Google event has all Outlook IDs stored
            String missingIds = "";
            if (!Exists(ev, MetadataId.oGlobalApptId)) missingIds += metadataIdKeyName(MetadataId.oGlobalApptId) + "|";
            if (!Exists(ev, MetadataId.oCalendarId)) missingIds += metadataIdKeyName(MetadataId.oCalendarId) + "|";
            if (!Exists(ev, MetadataId.oEntryId)) missingIds += metadataIdKeyName(MetadataId.oEntryId) + "|";
            if (!string.IsNullOrEmpty(missingIds))
                log.Warn("Found Google item missing Outlook IDs (" + missingIds.TrimEnd('|') + "). " + GoogleOgcs.Calendar.GetEventSummary(ev));
            return !string.IsNullOrEmpty(missingIds);
        }

        public static Boolean Exists(Event ev, MetadataId searchId) {
            String throwAway;
            return Exists(ev, searchId, out throwAway);
        }
        public static Boolean Exists(Event ev, MetadataId searchId, out String searchKey) {
            searchKey = null;
            if (ev.ExtendedProperties == null || ev.ExtendedProperties.Private__ == null) return false;

            searchKey = metadataIdKeyName(searchId);

            int maxSet;
            int? keySet = getKeySet(ev, out maxSet);
            if (keySet.HasValue) searchKey += "-" + keySet.Value.ToString("D2");
            if (searchId == MetadataId.oCalendarId)
                return ev.ExtendedProperties.Private__.ContainsKey(searchKey) && ev.ExtendedProperties.Private__[searchKey] == OutlookOgcs.Calendar.Instance.UseOutlookCalendar.EntryID;
            else
                return ev.ExtendedProperties.Private__.ContainsKey(searchKey) && Get(ev, MetadataId.oCalendarId) == OutlookOgcs.Calendar.Instance.UseOutlookCalendar.EntryID;
        }

        public static Boolean ExistsAny(Event ev) {
            if (Exists(ev, MetadataId.oEntryId)) return true;
            if (Exists(ev, MetadataId.oGlobalApptId)) return true;
            if (Exists(ev, MetadataId.oCalendarId)) return true;
            return false;
        }

        /// <summary>
        /// Add the Outlook appointment IDs into Google event.
        /// </summary>
        public static void AddOutlookIDs(ref Event ev, AppointmentItem ai) {
            Add(ref ev, MetadataId.oCalendarId, OutlookOgcs.Calendar.Instance.UseOutlookCalendar.EntryID);
            Add(ref ev, MetadataId.oEntryId, ai.EntryID);
            Add(ref ev, MetadataId.oGlobalApptId, OutlookOgcs.Calendar.Instance.IOutlook.GetGlobalApptID(ai));
            CustomProperty.LogProperties(ev, log4net.Core.Level.Debug);
        }

        public static void Add(ref Event ev, MetadataId key, String value) {
            String addkeyName = metadataIdKeyName(key);
            if (ev.ExtendedProperties == null) ev.ExtendedProperties = new Event.ExtendedPropertiesData();
            if (ev.ExtendedProperties.Private__ == null) ev.ExtendedProperties.Private__ = new Dictionary<String, String>();

            int newSet;
            int? keySet = getKeySet(ev, out newSet);
            add(ref ev, addkeyName, value, keySet ?? newSet + 1);
        }
        private static void Add(ref Event ev, MetadataId key, DateTime value) {
            Add(ref ev, key, value.ToString("yyyyMMddHHmmss", System.Globalization.CultureInfo.InvariantCulture));
        }
        private static void add(ref Event ev, String keyName, String keyValue, int? keySet) {
            if (keySet.HasValue && keySet.Value != 0) keyName += "-" + keySet.Value.ToString("D2");
            if (ev.ExtendedProperties.Private__.ContainsKey(keyName))
                ev.ExtendedProperties.Private__[keyName] = keyValue;
            else
                ev.ExtendedProperties.Private__.Add(keyName, keyValue);

            log.Fine("Set extendedproperty " + keyName + "=" + keyValue);
        }

        public static String Get(Event ev, MetadataId id) {
            String key;
            if (Exists(ev, id, out key)) {
                return ev.ExtendedProperties.Private__[key];
            } else
                return null;
        }

        public static void Remove(ref Event ev, MetadataId id) {
            String key;
            if (Exists(ev, id, out key))
                ev.ExtendedProperties.Private__.Remove(key);
        }

        public static DateTime GetOGCSlastModified(Event ev) {
            if (Exists(ev, MetadataId.ogcsModified)) {
                String lastModded = Get(ev, MetadataId.ogcsModified);
                try {
                    return DateTime.ParseExact(lastModded, "yyyyMMddHHmmss", System.Globalization.CultureInfo.InvariantCulture);
                } catch (System.FormatException) {
                    //Bugfix <= v2.2, 
                    log.Fine("Date wasn't stored as invariant culture.");
                    DateTime retDate;
                    if (DateTime.TryParse(lastModded, out retDate)) {
                        log.Fine("Fall back to current culture successful.");
                        return retDate;
                    } else {
                        log.Debug("Fall back to current culture for date failed. Last resort: setting to a month ago.");
                        return DateTime.Now.AddMonths(-1);
                    }
                }
            } else {
                return new DateTime();
            }
        }
        public static void SetOGCSlastModified(ref Event ev) {
            Add(ref ev, MetadataId.ogcsModified, DateTime.Now);
        }

        /// <summary>
        /// Log the various User Properties.
        /// </summary>
        /// <param name="ev">The Event.</param>
        /// <param name="thresholdLevel">Only log if logging configured at this level or higher.</param>
        public static void LogProperties(Event ev, log4net.Core.Level thresholdLevel) {
            if (((log4net.Repository.Hierarchy.Hierarchy)LogManager.GetRepository()).Root.Level.Value > thresholdLevel.Value) return;

            try {
                if (ev.ExtendedProperties != null && ev.ExtendedProperties.Private__ != null) {
                    log.Debug(GoogleOgcs.Calendar.GetEventSummary(ev));
                    foreach (KeyValuePair<String, String> prop in ev.ExtendedProperties.Private__.OrderBy(k => k.Key)) {
                        log.Debug(prop.Key + "=" + prop.Value);
                    }
                }
            } catch (System.Exception ex) {
                OGCSexception.Analyse("Failed to log Event ExtendedProperties", ex);
            }
        }
    }
}
