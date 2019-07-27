using Google.Apis.Calendar.v3.Data;
using Microsoft.Office.Interop.Outlook;
using System.Runtime.Serialization;

namespace OutlookGoogleCalendarSync {
    /// <summary>
    /// Description of CalendarListEntry.
    /// </summary>
    
    [DataContract]
    public class GoogleCalendarListEntry {
        [DataMember]
        private string Name { get; set; }
        [DataMember]
        public string Id { get; internal set; }
        [DataMember]
        private string AccessRole { get; set; }

        private bool primary { get; set; }

        private bool readOnly {
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
        }

        public override string ToString() {
            return (readOnly ? "[Read Only] " : "") + Name;
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
        public string Name { get; internal set; }
        [DataMember]
        public string Id { get; internal set; }

        public OutlookCalendarListEntry() {
        }

        public OutlookCalendarListEntry(MAPIFolder calendarFolder) {
            Id = calendarFolder.EntryID;
            Name = calendarFolder.Name;
        }

        public override string ToString() {
            return Name;
        }
    }
}
