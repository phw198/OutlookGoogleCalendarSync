using log4net;
using GcalData = Google.Apis.Calendar.v3.Data;
using Ogcs = OutlookGoogleCalendarSync;

namespace OutlookGoogleCalendarSync.Google.Graph {
    class CustomProperty {
        private static readonly ILog log = LogManager.GetLogger(typeof(CustomProperty));

        /// <summary>
        /// Add the Outlook Graph Event IDs into Google event.
        /// </summary>
        public static void AddOutlookIDs(ref GcalData.Event ev, Microsoft.Graph.Event ai) {
            Ogcs.Google.CustomProperty.Add(ref ev, Ogcs.Google.CustomProperty.MetadataId.oCalendarId, Sync.Engine.Calendar.Instance.Profile.UseOutlookCalendar.Id);
            Ogcs.Google.CustomProperty.Add(ref ev, Ogcs.Google.CustomProperty.MetadataId.oEntryId, ai.Id);
            Ogcs.Google.CustomProperty.Add(ref ev, Ogcs.Google.CustomProperty.MetadataId.oGlobalApptId, ai.ICalUId);
            Ogcs.Google.CustomProperty.LogProperties(ev, log4net.Core.Level.Debug);
        }
    }
}
