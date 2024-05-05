using Google.Apis.Calendar.v3.Data;
using Microsoft.Office.Interop.Outlook;
using System;
using System.Collections.Generic;
using System.Runtime.Serialization;

namespace OutlookGoogleCalendarSync {
    /// <summary>
    /// Description of CalendarListEntry.
    /// </summary>

    [DataContract]
    public class GoogleCalendarListEntry {
        //Default values before loading from xml and attribute not yet serialized
        [OnDeserializing]
        void OnDeserializing(StreamingContext context) {
            setDefaults();
        }

        private void setDefaults() {
            this.colourId = "0";
        }

        [DataMember]
        private String Name { get; set; }
        [DataMember]
        public String Id { get; internal set; }
        [DataMember]
        private String AccessRole { get; set; }
        private String colourId = "0";
        [DataMember]
        public String ColourId {
            get { return colourId; }
            internal set { colourId = value; }
        }

        private Boolean primary { get; set; }
        public Boolean Hidden { get; protected set; }

        private Boolean readOnly {
            get {
                if (AccessRole == null) return false;
                return AccessRole.ToLower().Contains("reader");
            }
        }

        public GoogleCalendarListEntry() {
        }

        public GoogleCalendarListEntry(CalendarListEntry init) {
            AccessRole = init.AccessRole;
            Id = init.Id;
            Name = init.SummaryOverride ?? init.Summary;
            primary = init.Primary ?? false;
            Hidden = init.Hidden ?? false;
            ColourId = init.ColorId;
        }

        public override String ToString() {
            List<String> prefix = new List<String>();
            if (readOnly) prefix.Add("Read Only");
            if (Hidden) prefix.Add("Hidden");

            if (prefix.Count > 0)
                return "[" + string.Join(",", prefix) + "] " + Name;

            return Name;
        }

        public string ToString(bool withId) {
            return EmailAddress.MaskAddressWithinText(this.ToString()) + (Id != null ? " (ID: " + EmailAddress.MaskAddressWithinText(Id) + ")" : "");
        }

        public string Sorted() {
            switch (AccessRole.ToLower()) {
                case "owner": return (primary ? "0-" : "1-") + Name;
                case "writer": return "1-" + Name;
                case "reader": return "2-" + Name;
                case "freebusyreader": return "2-" + Name;
                default: return "1-" + Name;
            }
        }
    }

    [DataContract]
    public class OutlookCalendarListEntry {
        [DataMember]
        public String Name { get; internal set; }
        [DataMember]
        public String Id { get; internal set; }

        public OutlookCalendarListEntry() {
        }

        public OutlookCalendarListEntry(MAPIFolder calendarFolder) {
            Id = calendarFolder.EntryID;
            Name = calendarFolder.Name;
        }

        public OutlookCalendarListEntry(Microsoft.Graph.Calendar calendarFolder) {
            Id = calendarFolder.Id;
            Name = calendarFolder.Name;
        }

        public override String ToString() {
            return EmailAddress.MaskAddressWithinText(Name) + " (ID: " + Id + ")";
        }
    }
}
