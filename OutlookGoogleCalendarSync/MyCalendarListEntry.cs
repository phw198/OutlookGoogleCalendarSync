using Google.Apis.Calendar.v3.Data;
using System.Runtime.Serialization;

namespace OutlookGoogleCalendarSync {
    /// <summary>
    /// Description of MyCalendarListEntry.
    /// </summary>
    
    [DataContract]
    public class MyCalendarListEntry {
        [DataMember]
        public string Name { get; internal set; }
        [DataMember]
        public string Id { get; internal set; }

        public MyCalendarListEntry() {
        }

        public MyCalendarListEntry(CalendarListEntry init) {
            Id = init.Id;
            Name = init.Summary;
        }

        public override string ToString() {
            return Name;
        }
    }
}
