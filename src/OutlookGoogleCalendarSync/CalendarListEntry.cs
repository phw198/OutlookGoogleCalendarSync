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
        public String Name { get; protected set; }
        [DataMember]
        public String Id { get; protected set; }
        [DataMember]
        private String AccessRole { get; set; }
        private String colourId = "0";
        [DataMember]
        public String ColourId {
            get { return colourId; }
            internal set { colourId = value; }
        }

        public Boolean Primary { get; protected set; }
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
            Primary = init.Primary ?? false;
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
                case "owner": return (Primary ? "0-" : "1-") + Name;
                case "writer": return "1-" + Name;
                case "reader": return "2-" + Name;
                case "freebusyreader": return "2-" + Name;
                default: return "1-" + Name;
            }
        }
    }

    [DataContract]
    public class OutlookCalendarListEntry {
        #region Common attributes
        [DataMember]
        public String Name { get; protected set; }
        [DataMember]
        public String Id { get; protected set; }
        #endregion

        #region Graph-specific attributes
        [DataMember]
        public Boolean ReadOnly {get; protected set; }
        #endregion

        private Boolean primary { get; set; }

        public OutlookCalendarListEntry() {
        }

        public OutlookCalendarListEntry(MAPIFolder calendarFolder) {
            Id = calendarFolder.EntryID;
            Name = calendarFolder.Name;
            
            //Set default for irrelevant Graph attributes
            ReadOnly = false;
        }

        public OutlookCalendarListEntry(Outlook.Graph.CustomClient.Models.Calendar calendarFolder) {
            Id = calendarFolder.Id;
            Name = calendarFolder.Name;
            primary = calendarFolder.IsDefaultCalendar ?? false;
            ReadOnly = !calendarFolder.CanEdit ?? false;
        }

        public override String ToString() {
            List<String> prefix = new List<String>();
            if (ReadOnly) prefix.Add("Read Only");

            if (prefix.Count > 0)
                return "[" + string.Join(",", prefix) + "] " + Name;

            return Name;
        }

        public String ToString(bool withId) {
            return EmailAddress.MaskAddressWithinText(Name) + " (ID: " + Id + ")";
        }

        public string Sorted(String key) {
            if (this.primary) return "0-" + key;
            if (this.ReadOnly) return "2-" + key;
            return "1-" + key;
        }
    }
}
