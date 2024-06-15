using Google.Apis.Calendar.v3.Data;
using log4net;
using Microsoft.Office.Interop.Outlook;
using OutlookGoogleCalendarSync.Extensions;
using OutlookGoogleCalendarSync.GraphExtension;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

namespace OutlookGoogleCalendarSync.Outlook.Graph {
    public class EphemeralProperties {

        private Dictionary<Microsoft.Graph.Event, Dictionary<EphemeralProperty.PropertyName, Object>> ephemeralProperties;

        public EphemeralProperties() {
            ephemeralProperties = new Dictionary<Microsoft.Graph.Event, Dictionary<EphemeralProperty.PropertyName, Object>>();
        }

        public void Clear() {
            ephemeralProperties = new Dictionary<Microsoft.Graph.Event, Dictionary<EphemeralProperty.PropertyName, Object>>();
        }

        public void Add(Microsoft.Graph.Event ai, EphemeralProperty property) {
            if (!ExistAny(ai)) {
                ephemeralProperties.Add(ai, new Dictionary<EphemeralProperty.PropertyName, object> { { property.Name, property.Value } });
            } else {
                if (PropertyExists(ai, property.Name)) ephemeralProperties[ai][property.Name] = property.Value;
                else ephemeralProperties[ai].Add(property.Name, property.Value);
            }
        }

        /// <summary>
        /// Is the Graph Event already registered with any ephemeral properties?
        /// </summary>
        /// <param name="ai">The Graph Event to check</param>
        public Boolean ExistAny(Microsoft.Graph.Event ai) {
            return ephemeralProperties.ContainsKey(ai);
        }
        /// <summary>
        /// Does a specific ephemeral property exist for a Graph Event?
        /// </summary>
        /// <param name="ai">The Graph Event to check</param>
        /// <param name="propertyName">The property to check</param>
        public Boolean PropertyExists(Microsoft.Graph.Event ai, EphemeralProperty.PropertyName propertyName) {
            if (!ExistAny(ai)) return false;
            return ephemeralProperties[ai].ContainsKey(propertyName);
        }

        public Object GetProperty(Microsoft.Graph.Event ai, EphemeralProperty.PropertyName propertyName) {
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
        //These keys are only stored in memory against the Graph Event, not saved anwhere.
        public enum PropertyName {
            KeySet, //Current set for calendar being synced
            MaxSet  //Last set in contiguous sequence
        }
        public PropertyName Name { get; private set; }
        public Object Value { get; private set; }

        public EphemeralProperty(PropertyName propertyName, Object value) {
            Name = propertyName;
            Value = value;
        }
    }
    
    class O365CustomProperty {
        private static readonly ILog log = LogManager.GetLogger(typeof(CustomProperty));

        private static String calendarKeyName = metadataIdKeyName(MetadataId.gCalendarId);

        private const String extensionName = "Ogcs.Properties";
        /// <summary>
        /// The name of the OGCS extension property bag
        /// </summary>
        /// <param name="prefixWithMsType">Once the Graph Event is saved, Microsoft add a prefix</param>
        public static String ExtensionName(Boolean prefixWithMsType = false) {
            return extensionName.Prepend(prefixWithMsType ? "Microsoft.OutlookServices.OpenTypeExtension." : "");
        }

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
            gMeetUrl/*,
            locallyCopied,
            originalStartDate*/
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
        private static int? getKeySet(Microsoft.Graph.Event ai, out int maxSet) {
            String returnSet = "";
            int? returnVal = null;
            maxSet = 0;

            if (Calendar.Instance.EphemeralProperties.PropertyExists(ai, EphemeralProperty.PropertyName.KeySet) &&
                Calendar.Instance.EphemeralProperties.PropertyExists(ai, EphemeralProperty.PropertyName.MaxSet)) {
                Object ep_keySet = Calendar.Instance.EphemeralProperties.GetProperty(ai, EphemeralProperty.PropertyName.KeySet);
                Object ep_maxSet = Calendar.Instance.EphemeralProperties.GetProperty(ai, EphemeralProperty.PropertyName.MaxSet);
                maxSet = Convert.ToInt16(ep_maxSet ?? ep_keySet);
                if (ep_keySet != null) returnVal = Convert.ToInt16(ep_keySet);
                return returnVal;
            }

            Dictionary<String, String> calendarKeys = new Dictionary<string, string>();
            /*UserProperties ups = null;
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
            }*/

            try {
                //For backward compatibility, always default to key names with no set number appended
                if (calendarKeys.Count == 0 ||
                    (calendarKeys.Count == 1 && calendarKeys.ContainsKey(calendarKeyName) && calendarKeys[calendarKeyName] == Sync.Engine.Calendar.Instance.Profile.UseGoogleCalendar.Id)) {
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
                        if (kvp.Value == Sync.Engine.Calendar.Instance.Profile.UseGoogleCalendar.Id)
                            returnSet = (matches[0].Groups[1].Value == "") ? "0" : matches[0].Groups[1].Value;
                    }
                }

                if (!string.IsNullOrEmpty(returnSet)) returnVal = Convert.ToInt16(returnSet);

            } finally {
                Calendar.Instance.EphemeralProperties.Add(ai, new EphemeralProperty(EphemeralProperty.PropertyName.KeySet, returnVal));
                Calendar.Instance.EphemeralProperties.Add(ai, new EphemeralProperty(EphemeralProperty.PropertyName.MaxSet, maxSet));
            }
            return returnVal;
        }

        public static Boolean GoogleIdMissing(Microsoft.Graph.Event ai) {
            //Make sure Outlook appointment has all Google IDs stored
            String missingIds = "";
            if (!Exists(ai, MetadataId.gEventID)) missingIds += metadataIdKeyName(MetadataId.gEventID) + "|";
            if (!Exists(ai, MetadataId.gCalendarId)) missingIds += metadataIdKeyName(MetadataId.gCalendarId) + "|";
            if (!string.IsNullOrEmpty(missingIds))
                log.Warn("Found Outlook item missing Google IDs (" + missingIds.TrimEnd('|') + "). " + Calendar.GetEventSummary(ai));
            return !string.IsNullOrEmpty(missingIds);
        }

        public static Boolean Exists(Microsoft.Graph.Event ai, MetadataId searchId) {
            String throwAway;
            return Exists(ai, searchId, out throwAway);
        }
        public static Boolean Exists(Microsoft.Graph.Event ai, MetadataId searchId, out String searchKey) {
            searchKey = metadataIdKeyName(searchId);

            int maxSet;
            int? keySet = getKeySet(ai, out maxSet);
            if (keySet.HasValue && keySet.Value != 0) searchKey += "-" + keySet.Value.ToString("D2");

            UserProperties ups = null;
            UserProperty prop = null;
            /*try {
                ups = ai.UserProperties;
                prop = ups.Find(searchKey);
                if (searchId == MetadataId.gCalendarId)
                    return (prop != null && prop.Value.ToString() == Sync.Engine.Calendar.Instance.Profile.UseGoogleCalendar.Id);
                else {
                    return (prop != null && Get(ai, MetadataId.gCalendarId) == Sync.Engine.Calendar.Instance.Profile.UseGoogleCalendar.Id);
                }
            } catch {
                return false;
            } finally {
                prop = (UserProperty)Calendar.ReleaseObject(prop);
                ups = (UserProperties)Calendar.ReleaseObject(ups);
            }*/
            return false; //Temporary hack
        }

        public static Boolean ExistAnyGoogleIDs(Microsoft.Graph.Event ai) {
            if (Exists(ai, MetadataId.gEventID)) return true;
            if (Exists(ai, MetadataId.gCalendarId)) return true;
            return false;
        }

        /// <summary>
        /// Are there any properties that start with key name (irrespective of key set value)
        /// </summary>
        public static Boolean AnyStartsWith(Microsoft.Graph.Event ai, MetadataId key) {
            String keyName = metadataIdKeyName(key);

            /*UserProperties ups = null;
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
            }*/
            return false;
        }

        /// <summary>
        /// Add the Google Event IDs into Outlook Event.
        /// </summary>
        public static void AddGoogleIDs(ref Microsoft.Graph.Event ai, Event ev) {
            Add(ref ai, MetadataId.gCalendarId, Sync.Engine.Calendar.Instance.Profile.UseGoogleCalendar.Id);
            Add(ref ai, MetadataId.gEventID, ev.Id);
            LogProperties(ai, log4net.Core.Level.Debug);
        }

        /// <summary>
        /// Remove the Google Event IDs from an Outlook Event.
        /// </summary>
        public static void RemoveGoogleIDs(ref Microsoft.Graph.Event ai) {
            Remove(ref ai, MetadataId.gEventID);
            Remove(ref ai, MetadataId.gCalendarId);
        }

        public static void Add(ref Microsoft.Graph.Event ai, MetadataId key, String value) {
            add(ref ai, key, OlUserPropertyType.olText, value);
        }
        public static void Add(ref Microsoft.Graph.Event ai, MetadataId key, System.DateTime value) {
            add(ref ai, key, OlUserPropertyType.olDateTime, value);
        }
        private static void add(ref Microsoft.Graph.Event ai, MetadataId key, OlUserPropertyType keyType, object keyValue) {
            String addkeyName = metadataIdKeyName(key);

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
            } else
                addkeyName = currentKeyName; //Might be suffixed with "-01"

            if (ai.Extensions == null)
                ai.Extensions = new Microsoft.Graph.EventExtensionsCollectionPage();

            if (ai.Extensions.Count == 0)
                ai.Extensions.Add(new Microsoft.Graph.OpenTypeExtension {
                    ExtensionName = extensionName,
                    Id = extensionName,
                    AdditionalData = new Dictionary<String, Object>()
                });
            
            ai.Extensions.Where(e => e.Id == extensionName).First().AdditionalData[addkeyName] = keyValue.ToString();
            Calendar.Instance.EphemeralProperties.Add(ai, new EphemeralProperty(EphemeralProperty.PropertyName.KeySet, keySet));
            Calendar.Instance.EphemeralProperties.Add(ai, new EphemeralProperty(EphemeralProperty.PropertyName.MaxSet, keySet));
            log.Fine("Set userproperty " + addkeyName + "=" + keyValue.ToString());
        }

        public static String Get(Microsoft.Graph.Event ai, MetadataId key) {
            String retVal = null;
            String searchKey;
            if (Exists(ai, key, out searchKey)) {
                /*UserProperties ups = null;
                UserProperty prop = null;
                *try {
                    ups = ai.UserProperties;
                    prop = ups.Find(searchKey);
                    if (prop != null) {
                        if (prop.Type != OlUserPropertyType.olText) log.Warn("Non-string property " + searchKey + " being retrieved as String.");
                        retVal = prop.Value.ToString();
            }
                } finally {
                    prop = (UserProperty)Calendar.ReleaseObject(prop);
                    ups = (UserProperties)Calendar.ReleaseObject(ups);
                }*/
            }
            return retVal;
        }
        private static System.DateTime get_datetime(Microsoft.Graph.Event ai, MetadataId key) {
            System.DateTime retVal = new System.DateTime();
            String searchKey;
            if (Exists(ai, key, out searchKey)) {
                UserProperties ups = null;
                UserProperty prop = null;
                /*try {
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
                            Ogcs.Exception.Analyse(ex);
                        }
                    }
                } finally {
                    prop = (UserProperty)Calendar.ReleaseObject(prop);
                    ups = (UserProperties)Calendar.ReleaseObject(ups);
                }*/
            }
            return retVal;
        }

        public static void RemoveAll(ref Microsoft.Graph.Event ai) {
            Remove(ref ai, MetadataId.gEventID);
            Remove(ref ai, MetadataId.gCalendarId);
            Remove(ref ai, MetadataId.forceSave);
            //Remove(ref ai, MetadataId.locallyCopied);
            Remove(ref ai, MetadataId.ogcsModified);
        }
        public static void Remove(ref Microsoft.Graph.Event ai, MetadataId key) {
            String searchKey;
            if (Exists(ai, key, out searchKey)) {
                UserProperties ups = null;
                UserProperty prop = null;
                /*try {
                    ups = ai.UserProperties;
                    prop = ups.Find(searchKey);
                    prop.Delete();
                    log.Debug("Removed " + searchKey + " property.");
                } finally {
                    prop = (UserProperty)Calendar.ReleaseObject(prop);
                    ups = (UserProperties)Calendar.ReleaseObject(ups);
                }*/
            }
        }
        /// <summary>
        /// Completely remove all OGCS custom properties
        /// </summary>
        /// <param name="ai">The Graph Event to strip attributes from</param>
        /// <returns>Whether any properties were removed</returns>
        public static Boolean Extirpate(ref Microsoft.Graph.Event ai) {
            List<String> keyNames = new List<String>() {
                metadataIdKeyName(MetadataId.forceSave),
                metadataIdKeyName(MetadataId.gCalendarId),
                metadataIdKeyName(MetadataId.gEventID),
                metadataIdKeyName(MetadataId.gMeetUrl),
                //metadataIdKeyName(MetadataId.locallyCopied),
                metadataIdKeyName(MetadataId.ogcsModified),
                //metadataIdKeyName(MetadataId.originalStartDate)
            };
            Boolean removedProperty = false;
            UserProperties ups = null;
            /*try {
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
            }*/
            return removedProperty;
        }

        public static System.DateTime GetOGCSlastModified(Microsoft.Graph.Event ai) {
            return get_datetime(ai, MetadataId.ogcsModified);
        }
        public static void SetOGCSlastModified(ref Microsoft.Graph.Event ai) {
            Add(ref ai, MetadataId.ogcsModified, System.DateTime.Now);
        }

        /// <summary>
        /// Log the various User Properties.
        /// </summary>
        /// <param name="ai">The Graph Event item.</param>
        /// <param name="thresholdLevel">Only log if logging configured at this level or higher.</param>
        public static void LogProperties(Microsoft.Graph.Event ai, log4net.Core.Level thresholdLevel) {
            if (((log4net.Repository.Hierarchy.Hierarchy)LogManager.GetRepository()).Root.Level.Value > thresholdLevel.Value) return;

            try {
                log.Debug(Calendar.GetEventSummary(ai));
                Microsoft.Graph.Extension ext = ai.Extensions.Where(e => e.Id == extensionName).First();
                foreach (KeyValuePair<String, Object> prop in ext.AdditionalData) {
                    if (prop.Key == metadataIdKeyName(MetadataId.gCalendarId))
                        log.Debug(prop.Key + "=" + EmailAddress.MaskAddress(prop.Value.ToString()));
                    else
                        log.Debug(prop.Key + "=" + prop.Value.ToString());
                }
            } catch (System.Exception ex) {
                ex.Analyse("Failed to log Appointment UserProperties");
            }
        }
    }

    /*public static class ReflectionProperties {
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
    }*/
}
