using Google.Apis.Calendar.v3.Data;
using log4net;
using Microsoft.Office.Interop.Outlook;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

namespace OutlookGoogleCalendarSync.GoogleOgcs {
    public class EphemeralProperties {

        private Dictionary<Event, Dictionary<EphemeralProperty.PropertyName, Object>> ephemeralProperties;

        public EphemeralProperties() {
            ephemeralProperties = new Dictionary<Event, Dictionary<EphemeralProperty.PropertyName, Object>>();
        }

        public void Clear() {
            ephemeralProperties = new Dictionary<Event, Dictionary<EphemeralProperty.PropertyName, object>>();
        }

        public void Add(Event ev, EphemeralProperty property) {
            if (!ExistAny(ev)) {
                ephemeralProperties.Add(ev, new Dictionary<EphemeralProperty.PropertyName, object> { { property.Name, property.Value } });
            } else {
                if (PropertyExists(ev, property.Name)) ephemeralProperties[ev][property.Name] = property.Value;
                else ephemeralProperties[ev].Add(property.Name, property.Value);
            }
        }

        /// <summary>
        /// Is the Event already registered with any ephemeral properties?
        /// </summary>
        /// <param name="ev">The Event to check</param>
        public Boolean ExistAny(Event ev) {
            return ephemeralProperties.ContainsKey(ev);
        }
        /// <summary>
        /// Does a specific ephemeral property exist for an Event?
        /// </summary>
        /// <param name="ev">The Event to check</param>
        /// <param name="propertyName">The property to check</param>
        public Boolean PropertyExists(Event ev, EphemeralProperty.PropertyName propertyName) {
            if (!ExistAny(ev)) return false;
            return ephemeralProperties[ev].ContainsKey(propertyName);
        }

        public Object GetProperty(Event ev, EphemeralProperty.PropertyName propertyName) {
            if (this.ExistAny(ev)) {
                if (PropertyExists(ev, propertyName)) {
                    Object ep = ephemeralProperties[ev][propertyName];
                    switch (propertyName) {
                        case EphemeralProperty.PropertyName.KeySet:
                            if (ep is int && ep != null) return Convert.ToInt16(ep);
                            else return null;
                        case EphemeralProperty.PropertyName.MaxSet:
                            if (ep is int && ep != null) return Convert.ToInt16(ep);
                            else return null;
                        default:
                            return ep;
                    }
                }
            }
            return null;
        }
    }

    public class EphemeralProperty {
        //These keys are only stored in memory against the Event, not saved anwhere.
        public enum PropertyName {
            KeySet, //Current set for calendar being synced
            MaxSet  //Last set in continquous sequence
        }
        public PropertyName Name { get; private set; }
        public Object Value { get; private set; }

        public EphemeralProperty(PropertyName propertyName, Object value) {
            Name = propertyName;
            Value = value;
        }
    }

    public class CustomProperty {
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
            int? returnVal = null;
            maxSet = 0;

            if (GoogleOgcs.Calendar.Instance.EphemeralProperties.PropertyExists(ev, EphemeralProperty.PropertyName.KeySet) &&
                GoogleOgcs.Calendar.Instance.EphemeralProperties.PropertyExists(ev, EphemeralProperty.PropertyName.MaxSet)) 
            {
                Object ep_keySet = GoogleOgcs.Calendar.Instance.EphemeralProperties.GetProperty(ev, EphemeralProperty.PropertyName.KeySet);
                Object ep_maxSet = GoogleOgcs.Calendar.Instance.EphemeralProperties.GetProperty(ev, EphemeralProperty.PropertyName.MaxSet);
                maxSet = Convert.ToInt16(ep_maxSet ?? ep_keySet);
                if (ep_keySet != null) returnVal = Convert.ToInt16(ep_keySet);
                return returnVal;
            }

            try {
                Dictionary<String, String> calendarKeys = ev.ExtendedProperties.Private__.Where(k => k.Key.StartsWith(calendarKeyName)).OrderBy(k => k.Key).ToDictionary(k => k.Key, k => k.Value);

                //For backward compatibility, always default to key names with no set number appended
                if (calendarKeys.Count == 0||
                    (calendarKeys.Count == 1 && calendarKeys.ContainsKey(calendarKeyName)) && calendarKeys[calendarKeyName] == Sync.Engine.Calendar.Instance.Profile.UseOutlookCalendar.Id)
                {
                    maxSet = -1;
                    return returnVal;
                }

                foreach (KeyValuePair<String, String> kvp in calendarKeys) {
                    Regex rgx = new Regex("^" + calendarKeyName + "-*(\\d{0,2})", RegexOptions.IgnoreCase);
                    MatchCollection matches = rgx.Matches(kvp.Key);

                    if (matches.Count > 0) {
                        int appendedNos = 0;
                        if (matches[0].Groups[1].Value != "")
                            appendedNos = Convert.ToInt16(matches[0].Groups[1].Value);
                        if (appendedNos - maxSet == 1) maxSet = appendedNos;
                        if (kvp.Value == Sync.Engine.Calendar.Instance.Profile.UseOutlookCalendar.Id)
                            returnSet = (matches[0].Groups[1].Value == "") ? "0" : matches[0].Groups[1].Value;
                    }
                }

                if (!string.IsNullOrEmpty(returnSet)) returnVal = Convert.ToInt16(returnSet);

            } finally {
                GoogleOgcs.Calendar.Instance.EphemeralProperties.Add(ev, new EphemeralProperty(EphemeralProperty.PropertyName.KeySet, returnVal));
                GoogleOgcs.Calendar.Instance.EphemeralProperties.Add(ev, new EphemeralProperty(EphemeralProperty.PropertyName.MaxSet, maxSet));
            }
            return returnVal;
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
            if (keySet.HasValue && keySet.Value != 0) searchKey += "-" + keySet.Value.ToString("D2");
            if (searchId == MetadataId.oCalendarId)
                return ev.ExtendedProperties.Private__.ContainsKey(searchKey) && ev.ExtendedProperties.Private__[searchKey] == Sync.Engine.Calendar.Instance.Profile.UseOutlookCalendar.Id;
            else
                return ev.ExtendedProperties.Private__.ContainsKey(searchKey) && Get(ev, MetadataId.oCalendarId) == Sync.Engine.Calendar.Instance.Profile.UseOutlookCalendar.Id;
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
            Add(ref ev, MetadataId.oCalendarId, Sync.Engine.Calendar.Instance.Profile.UseOutlookCalendar.Id);
            Add(ref ev, MetadataId.oEntryId, ai.EntryID);
            Add(ref ev, MetadataId.oGlobalApptId, OutlookOgcs.Calendar.Instance.IOutlook.GetGlobalApptID(ai));
            CustomProperty.LogProperties(ev, log4net.Core.Level.Debug);
        }

        /// <summary>
        /// Remove the Outlook appointment IDs from a Google event.
        /// </summary>
        public static void RemoveOutlookIDs(ref Event ev) {
            Remove(ref ev, MetadataId.oEntryId);
            Remove(ref ev, MetadataId.oGlobalApptId);
            Remove(ref ev, MetadataId.oCalendarId);
        }

        public static void Add(ref Event ev, MetadataId key, String value) {
            String addkeyName = metadataIdKeyName(key);
            if (ev.ExtendedProperties == null) ev.ExtendedProperties = new Event.ExtendedPropertiesData();
            if (ev.ExtendedProperties.Private__ == null) ev.ExtendedProperties.Private__ = new Dictionary<String, String>();

            int maxSet;
            int? keySet = getKeySet(ev, out maxSet);
            if (key == MetadataId.oCalendarId && (keySet ?? 0) == 0) //Couldn't find key set for calendar
                keySet = maxSet + 1; //So start a new one
            else if (key != MetadataId.oCalendarId && keySet == null) //Couldn't find non-calendar key in the current set
                keySet = 0; //Add them in to the default key set

            add(ref ev, addkeyName, value, keySet);
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

            GoogleOgcs.Calendar.Instance.EphemeralProperties.Add(ev, new EphemeralProperty(EphemeralProperty.PropertyName.KeySet, keySet));
            GoogleOgcs.Calendar.Instance.EphemeralProperties.Add(ev, new EphemeralProperty(EphemeralProperty.PropertyName.MaxSet, keySet));
            log.Fine("Set extendedproperty " + keyName + "=" + keyValue);
        }

        public static String Get(Event ev, MetadataId id) {
            String key;
            if (Exists(ev, id, out key)) {
                return ev.ExtendedProperties.Private__[key];
            } else
                return null;
        }

        public static void RemoveAll(ref Event ev) {
            Remove(ref ev, MetadataId.apiLimitHit);
            Remove(ref ev, MetadataId.forceSave);
            Remove(ref ev, MetadataId.oEntryId);
            Remove(ref ev, MetadataId.ogcsModified);
            Remove(ref ev, MetadataId.oGlobalApptId);
            Remove(ref ev, MetadataId.oCalendarId); //This one must be removed last
        }
        public static void Remove(ref Event ev, MetadataId id) {
            String key;
            if (Exists(ev, id, out key))
                ev.ExtendedProperties.Private__.Remove(key);
        }
        /// <summary>
        /// Completely remove all OGCS custom properties
        /// </summary>
        /// <param name="ev">The Event to strip properties from</param>
        /// <returns>Whether any properties were removed</returns>
        public static Boolean Extirpate(ref Event ev) {
            Boolean removedProperty = false;
            List<String> keyNames = new List<String>() {
                metadataIdKeyName(MetadataId.apiLimitHit),
                metadataIdKeyName(MetadataId.forceSave),
                metadataIdKeyName(MetadataId.oCalendarId),
                metadataIdKeyName(MetadataId.oEntryId),
                metadataIdKeyName(MetadataId.ogcsModified),
                metadataIdKeyName(MetadataId.oGlobalApptId)
            };
            if (ev.ExtendedProperties != null && ev.ExtendedProperties.Private__ != null) {
                for (int i = ev.ExtendedProperties.Private__.Count - 1; i >= 0; i--) {
                    String ep = ev.ExtendedProperties.Private__.Keys.ToArray()[i];
                    if (keyNames.Exists(k => ep.StartsWith(k))) {
                        ev.ExtendedProperties.Private__.Remove(ep);
                        log.Fine("Removed " + ep);
                        removedProperty = true;
                    }
                }
            }
            return removedProperty;
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
