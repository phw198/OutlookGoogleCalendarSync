using System;
using System.Runtime.Serialization;

namespace OutlookGoogleCalendarSync.Sync {

    [DataContract]
    public sealed class Direction {

        [DataMember]
        public readonly String Name;
        [DataMember]
        public readonly int Id;

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
