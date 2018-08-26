using System;
using System.Runtime.Serialization;

namespace OutlookGoogleCalendarSync.Sync {

    [DataContract(Namespace="http://schemas.datacontract.org/2004/07/OutlookGoogleCalendarSync")]
    public sealed class Direction {

        [DataMember]
        public String Name { get; internal set; }
        [DataMember]
        public int Id { get; internal set; }

        public static readonly Direction OutlookToGoogle = new Direction(1, "Outlook → Google");
        public static readonly Direction GoogleToOutlook = new Direction(2, "Outlook ← Google");
        public static readonly Direction Bidirectional = new Direction(3,   "Outlook ↔ Google");

        private Direction(int id, String name) {
            this.Name = name;
            this.Id = id;
        }

        public override String ToString() {
            return Name;
        }
    }
}
