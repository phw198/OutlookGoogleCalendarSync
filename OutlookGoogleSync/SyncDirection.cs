using System;
using System.Runtime.Serialization;

namespace OutlookGoogleSync {

    [DataContract]
    public sealed class SyncDirection {

        [DataMember]
        public readonly String Name;
        [DataMember]
        public readonly int Id;

        public static readonly SyncDirection OutlookToGoogle = new SyncDirection(1, "Outlook ====> Google");
        public static readonly SyncDirection GoogleToOutlook = new SyncDirection(2, "Outlook <==== Google");
        public static readonly SyncDirection Bidirectional = new SyncDirection(3,   "Outlook <===> Google");

        public SyncDirection() { }

        private SyncDirection(int id, String name) {
            this.Name = name;
            this.Id = id;
        }

        public override String ToString() {
            return Name;
        }
    }
}
