using Ogcs = OutlookGoogleCalendarSync;
using Google.Apis.Calendar.v3.Data;
using log4net;
using Microsoft.Office.Interop.Outlook;
using OutlookGoogleCalendarSync.Extensions;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

namespace OutlookGoogleCalendarSync.Outlook {
    public class EphemeralProperties {

        private Dictionary<AppointmentItem, Dictionary<EphemeralProperty.PropertyName, Object>> ephemeralProperties;

        public EphemeralProperties() {
            ephemeralProperties = new Dictionary<AppointmentItem, Dictionary<EphemeralProperty.PropertyName, Object>>();
        }

        public void Clear() {
            ephemeralProperties = new Dictionary<AppointmentItem, Dictionary<EphemeralProperty.PropertyName, object>>();
        }

        public void Add(AppointmentItem ai, EphemeralProperty property) {
            if (!ExistAny(ai)) {
                ephemeralProperties.Add(ai, new Dictionary<EphemeralProperty.PropertyName, object> { { property.Name, property.Value } });
            } else {
                if (PropertyExists(ai, property.Name)) ephemeralProperties[ai][property.Name] = property.Value;
                else ephemeralProperties[ai].Add(property.Name, property.Value);
            }
        }

        /// <summary>
        /// Is the AppointmentItem already registered with any ephemeral properties?
        /// </summary>
        /// <param name="ai">The AppointmentItem to check</param>
        public Boolean ExistAny(AppointmentItem ai) {
            return ephemeralProperties.ContainsKey(ai);
        }
        /// <summary>
        /// Does a specific ephemeral property exist for an AppointmentItem?
        /// </summary>
        /// <param name="ai">The AppointmentItem to check</param>
        /// <param name="propertyName">The property to check</param>
        public Boolean PropertyExists(AppointmentItem ai, EphemeralProperty.PropertyName propertyName) {
            if (!ExistAny(ai)) return false;
            return ephemeralProperties[ai].ContainsKey(propertyName);
        }

        public Object GetProperty(AppointmentItem ai, EphemeralProperty.PropertyName propertyName) {
            if (this.ExistAny(ai)) {
                if (PropertyExists(ai, propertyName)) {
                    Object ep = ephemeralProperties[ai][propertyName];
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
        //These keys are only stored in memory against the AppointmentItem, not saved anwhere.
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
            ogcsModified,       //Associated with olDateTime type - deprecated due to truncation of seconds
            ogcsModifiedText,   //Associated with olText type - format of yyyyMMddHHmmss
            forceSave,
            gMeetUrl,
            locallyCopied,
            originalStartDate
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
                case MetadataId.ogcsModifiedText: return "OGCSmodifiedText";
                case MetadataId.forceSave: return "forceSave";
                default: return Id.ToString();
            }
        }

        /// <summary>
        /// Return number appended to key name for current calendar key
        /// </summary>
        /// <param name="maxSet">The set number of the last contiguous run of ID sets (to aid defragmentation).</param>
        /// <returns>The set number, if it exists</returns>
        private static int? getKeySet(AppointmentItem ai, out int maxSet, SettingsStore.Calendar profile = null) {
            String returnSet = "";
            int? returnVal = null;
            maxSet = 0;
            profile ??= Sync.Engine.Calendar.Instance.Profile;

            if (Calendar.Instance.EphemeralProperties.PropertyExists(ai, EphemeralProperty.PropertyName.KeySet) &&
                Calendar.Instance.EphemeralProperties.PropertyExists(ai, EphemeralProperty.PropertyName.MaxSet))
            {
                Object ep_keySet = Calendar.Instance.EphemeralProperties.GetProperty(ai, EphemeralProperty.PropertyName.KeySet);
                Object ep_maxSet = Calendar.Instance.EphemeralProperties.GetProperty(ai, EphemeralProperty.PropertyName.MaxSet);
                maxSet = Convert.ToInt16(ep_maxSet ?? ep_keySet);
                if (ep_keySet != null) returnVal = Convert.ToInt16(ep_keySet);
                return returnVal;
            }

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
                        up = (UserProperty)Calendar.ReleaseObject(up);
                    }
                }
            } finally {
                ups = (UserProperties)Calendar.ReleaseObject(ups);
            }

            try {
                //For backward compatibility, always default to key names with no set number appended
                if (calendarKeys.Count == 0 ||
                    (calendarKeys.Count == 1 && calendarKeys.ContainsKey(calendarKeyName) && calendarKeys[calendarKeyName] == profile.UseGoogleCalendar.Id)) //
                {
                    maxSet = -1;
                    return returnVal;
                }

                foreach (KeyValuePair<String, String> kvp in calendarKeys.OrderBy(k => k.Key)) {
                    Regex rgx = new Regex("^" + calendarKeyName + "-*(\\d{0,2})", RegexOptions.IgnoreCase);
                    MatchCollection matches = rgx.Matches(kvp.Key);

                    if (matches.Count > 0) {
                        int appendedNos = 0;
                        if (matches[0].Groups[1].Value != "")
                            appendedNos = Convert.ToInt16(matches[0].Groups[1].Value);
                        if (appendedNos - maxSet == 1) maxSet = appendedNos;
                        if (kvp.Value == profile.UseGoogleCalendar.Id)
                            returnSet = (matches[0].Groups[1].Value == "") ? "0" : matches[0].Groups[1].Value;
                    }
                }

                if (!string.IsNullOrEmpty(returnSet)) returnVal = Convert.ToInt16(returnSet);

            } catch (System.Exception ex) {
                ex.Analyse(true);
                throw;
            } finally {
                Calendar.Instance.EphemeralProperties.Add(ai, new EphemeralProperty(EphemeralProperty.PropertyName.KeySet, returnVal));
                Calendar.Instance.EphemeralProperties.Add(ai, new EphemeralProperty(EphemeralProperty.PropertyName.MaxSet, maxSet));
            }
            return returnVal;
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

        public static Boolean Exists(AppointmentItem ai, MetadataId searchId, SettingsStore.Calendar profile = null) {
            String throwAway;
            return Exists(ai, searchId, out throwAway, profile);
        }
        public static Boolean Exists(AppointmentItem ai, MetadataId searchId, out String searchKey, SettingsStore.Calendar profile = null) {
            profile ??= Sync.Engine.Calendar.Instance.Profile;
            searchKey = metadataIdKeyName(searchId);

            int maxSet;
            int? keySet = getKeySet(ai, out maxSet, profile);
            if (keySet.HasValue && keySet.Value != 0) searchKey += "-" + keySet.Value.ToString("D2");

            UserProperties ups = null;
            UserProperty prop = null;
            try {
                ups = ai.UserProperties;
                prop = ups.Find(searchKey);
                if (searchId == MetadataId.gCalendarId)
                    return (prop != null && prop.Value.ToString() == profile.UseGoogleCalendar.Id);
                else {
                    return (prop != null && Get(ai, MetadataId.gCalendarId, profile) == profile.UseGoogleCalendar.Id);
                }
            } catch {
                return false;
            } finally {
                prop = (UserProperty)Calendar.ReleaseObject(prop);
                ups = (UserProperties)Calendar.ReleaseObject(ups);
            }
        }

        public static Boolean ExistAnyGoogleIDs(AppointmentItem ai, SettingsStore.Calendar profile = null) {
            if (Exists(ai, MetadataId.gEventID, profile)) return true;
            if (Exists(ai, MetadataId.gCalendarId, profile)) return true;
            return false;
        }

        /// <summary>
        /// Are there any properties that start with key name (irrespective of key set value)
        /// </summary>
        public static Boolean AnyStartsWith(AppointmentItem ai, MetadataId key) {
            String keyName = metadataIdKeyName(key);

            UserProperties ups = null;
            try {
                ups = ai.UserProperties;
                for (int p = ups.Count; p > 0; p--) {
                    UserProperty prop = null;
                    try {
                        prop = ups[p];
                        if (prop.Name.StartsWith(keyName)) {
                            return true;
                        }
                    } finally {
                        prop = (UserProperty)Calendar.ReleaseObject(prop);
                    }
                }
            } finally {
                ups = (UserProperties)Calendar.ReleaseObject(ups);
            }
            return false;
        }

        /// <summary>
        /// Add the Google event IDs into Outlook appointment.
        /// </summary>
        public static void AddGoogleIDs(ref AppointmentItem ai, Event ev) {
            Add(ref ai, MetadataId.gCalendarId, Sync.Engine.Calendar.Instance.Profile.UseGoogleCalendar.Id);
            Add(ref ai, MetadataId.gEventID, ev.Id);
            LogProperties(ai, log4net.Core.Level.Debug);
        }

        /// <summary>
        /// Remove the Google event IDs from an Outlook appointment.
        /// </summary>
        public static void RemoveGoogleIDs(ref AppointmentItem ai, SettingsStore.Calendar profile = null) {
            Remove(ref ai, MetadataId.gEventID, profile);
            Remove(ref ai, MetadataId.gCalendarId, profile);
        }

        public static void Add(ref AppointmentItem ai, MetadataId key, String value) {
            add(ref ai, key, OlUserPropertyType.olText, value);
        }
        public static void Add(ref AppointmentItem ai, MetadataId key, DateTimeOffset value) {
            if (key == MetadataId.ogcsModifiedText || key == MetadataId.ogcsModified /* Store deprecated key properly */)
                //We can't use OlUserPropertyType.olDateTime, because the stupid OOM silently truncates it to the nearest minute!!
                add(ref ai, MetadataId.ogcsModifiedText, OlUserPropertyType.olText, value.ToPreciseString());
            else
                //Only store in this way for date values not requiring accuracy greater than minutes
                add(ref ai, key, OlUserPropertyType.olDateTime, value.UtcDateTime);
        }
        private static void add(ref AppointmentItem ai, MetadataId key, OlUserPropertyType keyType, object keyValue) {
            String addkeyName = metadataIdKeyName(key);

            UserProperties ups = null;
            try {
                int maxSet;
                int? keySet = null;
                String currentKeyName = null;
                if (!Exists(ai, key, out currentKeyName)) {
                    keySet = getKeySet(ai, out maxSet);
                    if (key == MetadataId.gCalendarId && (keySet ?? 0) == 0) //Couldn't find key set for calendar
                        keySet = maxSet + 1; //So start a new one
                    else if (key != MetadataId.gCalendarId && keySet == null) //Couldn't find non-calendar key in the current set
                        keySet = 0; //Add them in to the default key set

                    if (keySet.HasValue && keySet.Value != 0) addkeyName += "-" + keySet.Value.ToString("D2");

                    try {
                        ups = ai.UserProperties;
                        ups.Add(addkeyName, keyType);
                    } catch (System.Exception ex) {
                        Ogcs.Exception.Analyse(ex);
                        ups.Add(addkeyName, keyType, false);
                    } finally {
                        ups = (UserProperties)Calendar.ReleaseObject(ups);
                    }
                } else
                    addkeyName = currentKeyName; //Might be suffixed with "-01"
                ups = ai.UserProperties;
                ups[addkeyName].Value = keyValue;
                Calendar.Instance.EphemeralProperties.Add(ai, new EphemeralProperty(EphemeralProperty.PropertyName.KeySet, keySet));
                Calendar.Instance.EphemeralProperties.Add(ai, new EphemeralProperty(EphemeralProperty.PropertyName.MaxSet, keySet));
                log.Fine("Set userproperty " + addkeyName + "=" + keyValue.ToString());

            } finally {
                ups = (UserProperties)Calendar.ReleaseObject(ups);
            }
        }

        public static String Get(AppointmentItem ai, MetadataId key, SettingsStore.Calendar profile = null) {
            String retVal = null;
            String searchKey;
            if (Exists(ai, key, out searchKey, profile)) {
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
                    prop = (UserProperty)Calendar.ReleaseObject(prop);
                    ups = (UserProperties)Calendar.ReleaseObject(ups);
                }
            }
            return retVal;
        }
        private static DateTimeOffset get_datetime(ref AppointmentItem ai, MetadataId key) {
            DateTimeOffset minVal = new System.DateTime(1, 1, 1, 0, 0, 0, 0, DateTimeKind.Utc);
            DateTimeOffset retVal = minVal;
            String searchKey;
            if (Exists(ai, key, out searchKey)) {
                UserProperties ups = null;
                UserProperty prop = null;
                try {
                    ups = ai.UserProperties;
                    prop = ups.Find(searchKey);
                    try {
                        if (key == MetadataId.ogcsModifiedText && prop.Type == OlUserPropertyType.olText) {
                            try {
                                //Without explicit cast to String, follow error:-
                                //Microsoft.CSharp.RuntimeBinder.RuntimeBinderException: 'string' does not contain a definition for 'GetPreciseDate'
                                retVal = ((String)prop.Value.ToString()).GetPreciseDate();
                            } catch (FormatException) {
                                retVal = minVal;
                                throw;
                            }

                        } else if (key == MetadataId.ogcsModified && prop.Type == OlUserPropertyType.olDateTime) {
                            log.Info("Migrating away from deprecated olDateTime user property for ogcsModified.");
                            retVal = new System.DateTime(((System.DateTime)prop.Value).Ticks, DateTimeKind.Local);
                            Add(ref ai, MetadataId.ogcsModifiedText, retVal);
                            prop.Delete();
                            Add(ref ai, MetadataId.forceSave, true.ToString());

                        } else {
                            log.Warn($"Property '{searchKey}' of type {prop.Type.ToString()} being retrieved as DateTime.");
                            retVal = System.DateTime.Parse(prop.Value.ToString());
                        }

                    } catch (System.Exception ex) {
                        ex.Analyse("Failed to retrieve DateTime value for property " + searchKey);
                    }

                } finally {
                    prop = (UserProperty)Calendar.ReleaseObject(prop);
                    ups = (UserProperties)Calendar.ReleaseObject(ups);
                }

            } else if (key == MetadataId.ogcsModifiedText && Exists(ai, MetadataId.ogcsModified)) {
                retVal = get_datetime(ref ai, MetadataId.ogcsModified);
            }
            return retVal;
        }

        public static void RemoveAll(ref AppointmentItem ai) {
            Remove(ref ai, MetadataId.gEventID);
            Remove(ref ai, MetadataId.gCalendarId);
            Remove(ref ai, MetadataId.forceSave);
            Remove(ref ai, MetadataId.locallyCopied);
            Remove(ref ai, MetadataId.ogcsModified);
            Remove(ref ai, MetadataId.ogcsModifiedText);
        }
        public static void Remove(ref AppointmentItem ai, MetadataId key, SettingsStore.Calendar profile = null) {
            String searchKey;
            if (Exists(ai, key, out searchKey, profile)) {
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
        /// <summary>
        /// Completely remove all OGCS custom properties
        /// </summary>
        /// <param name="ai">The AppointmentItem to strip attributes from</param>
        /// <returns>Whether any properties were removed</returns>
        public static Boolean Extirpate(ref AppointmentItem ai) {
            List<String> keyNames = new List<String>() {
                metadataIdKeyName(MetadataId.forceSave),
                metadataIdKeyName(MetadataId.gCalendarId),
                metadataIdKeyName(MetadataId.gEventID),
                metadataIdKeyName(MetadataId.gMeetUrl),
                metadataIdKeyName(MetadataId.locallyCopied),
                metadataIdKeyName(MetadataId.ogcsModified),
                metadataIdKeyName(MetadataId.ogcsModifiedText),
                metadataIdKeyName(MetadataId.originalStartDate)
            };
            Boolean removedProperty = false;
            UserProperties ups = null;
            try {
                ups = ai.UserProperties;
                for (int p = ups.Count; p > 0; p--) {
                    UserProperty prop = null;
                    try {
                        prop = ups[p];
                        if (keyNames.Exists(k => prop.Name.StartsWith(k))) {
                            log.Fine("Removed " + prop.Name);
                            prop.Delete();
                            removedProperty = true;
                        }
                    } finally {
                        prop = (UserProperty)Calendar.ReleaseObject(prop);
                    }
                }
            } finally {
                ups = (UserProperties)Calendar.ReleaseObject(ups);
            }
            return removedProperty;
        }

        public static DateTimeOffset GetOGCSlastModified(AppointmentItem ai) {
            return get_datetime(ref ai, MetadataId.ogcsModifiedText);
        }
        public static void SetOGCSlastModified(ref AppointmentItem ai) {
            Add(ref ai, MetadataId.ogcsModifiedText, System.DateTimeOffset.UtcNow);
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
                log.Debug(Calendar.GetEventSummary(ai));
                ups = ai.UserProperties;
                for (int p = 1; p <= ups.Count; p++) {
                    try {
                        up = ups[p];
                        if (up.Name == metadataIdKeyName(MetadataId.gCalendarId))
                            log.Debug(up.Name + "=" + EmailAddress.MaskAddress(up.Value.ToString()));
                        else
                            log.Debug(up.Name + "=" + up.Value.ToString());
                    } finally {
                        up = (UserProperty)Calendar.ReleaseObject(up);
                    }
                }
            } catch (System.Exception ex) {
                ex.Analyse("Failed to log Appointment UserProperties");
            } finally {
                ups = (UserProperties)Calendar.ReleaseObject(ups);
            }
        }
    }

    public static class ReflectionProperties {
        private static readonly ILog log = LogManager.GetLogger(typeof(ReflectionProperties));

        public static OlBodyFormat BodyFormat(this AppointmentItem ai) {
            OlBodyFormat format = OlBodyFormat.olFormatUnspecified;
            try {
                format = (OlBodyFormat)ai.GetType().InvokeMember("BodyFormat", System.Reflection.BindingFlags.GetProperty, null, ai, null);
            } catch (System.Exception ex) {
                ex.Analyse("Unable to determine AppointmentItem body format.");
            }
            return format;
        }

        public static String RTFBodyAsString(this AppointmentItem ai) {
#if DEVELOP_AGAINST_2007
            return "";
#else
            return System.Text.Encoding.ASCII.GetString(ai.RTFBody as byte[]);
#endif
        }
        private static Boolean RTFIsHtml(this AppointmentItem ai) {
            //RTF Specification: https://learn.microsoft.com/en-us/openspecs/exchange_server_protocols/ms-oxrtfex/4e5f466b-068a-42b2-b3d5-c9b3d5872438
            String bodyCode = ai.RTFBodyAsString();
            Regex rgx = new Regex(@"\\rtf1.*?\\fromhtml1.*?\\fonttbl", RegexOptions.IgnoreCase);
            return rgx.IsMatch(bodyCode);
        }
    }
}
