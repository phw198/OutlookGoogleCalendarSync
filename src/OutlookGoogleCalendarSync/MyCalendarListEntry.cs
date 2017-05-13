using Google.Apis.Calendar.v3.Data;
using Microsoft.Office.Interop.Outlook;
using System.Runtime.Serialization;

namespace OutlookGoogleCalendarSync {
    /// <summary>
    /// Description of MyCalendarListEntry.
    /// </summary>
    
    [DataContract]
    public class MyGoogleCalendarListEntry {
        [DataMember]
        public string Name { get; internal set; }
        [DataMember]
        public string Id { get; internal set; }

        public MyGoogleCalendarListEntry() {
        }

        public MyGoogleCalendarListEntry(CalendarListEntry init) {
            Id = init.Id;
            Name = init.Summary;
        }

        public override string ToString() {
            return Name;
        }
    }

    [DataContract]
    public class MyOutlookCalendarListEntry {
        [DataMember]
        public string Name { get; internal set; }
        [DataMember]
        public string Id { get; internal set; }

        public MyOutlookCalendarListEntry() {
        }

        public MyOutlookCalendarListEntry(MAPIFolder calendarFolder) {
            Id = calendarFolder.EntryID;
            Name = calendarFolder.Name;
        }

        public override string ToString() {
            return Name;
        }
    }
}
