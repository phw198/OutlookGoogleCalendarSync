using Ogcs = OutlookGoogleCalendarSync;
using Google.Apis.Calendar.v3.Data;
using log4net;
using System;
using System.Collections.Generic;
using System.Linq;

namespace OutlookGoogleCalendarSync.Google {
    static class GMeet {
        private static readonly ILog log = LogManager.GetLogger(typeof(GMeet));

        /// <summary>
        /// Add/update/delete Google Meet conference data in Google event
        /// </summary>
        /// <param name="ev">The Event to update</param>
        /// <param name="outlookGMeetUrl">The GMeet URL detected in the Outlook appointment body</param>
        public static void GoogleMeet(this Event ev, String outlookGMeetUrl) {
            try {
                if (String.IsNullOrEmpty(ev.HangoutLink)) {
                    log.Debug("Adding Google Meet conference data.");
                    log.Debug("Conference ID: " + outlookGMeetUrl);
                    ev.ConferenceData = new ConferenceData() {
                        ConferenceSolution = new ConferenceSolution() {
                            Key = new ConferenceSolutionKey { Type = "hangoutsMeet" }
                        },
                        EntryPoints = new List<EntryPoint>() { new EntryPoint() { EntryPointType = "video", Uri = outlookGMeetUrl } }
                    };

                } else if (String.IsNullOrEmpty(outlookGMeetUrl)) {
                    log.Debug("Removing Google Meet conference data.");
                    if (ev.ConferenceData.ConferenceSolution.Key.Type != "hangoutsMeet") {
                        log.Warn("Unexpected conference solution type '" + ev.ConferenceData.ConferenceSolution.Key.Type + "'. Remove abandoned.");
                    } else {
                        EntryPoint ep = ev.ConferenceData.EntryPoints.Where(ep => ep.EntryPointType == "video").FirstOrDefault();
                        log.Fine("Removing the 'video' conference entry point.");
                        ev.ConferenceData.EntryPoints.Remove(ep);
                        ev.ConferenceData.ConferenceSolution.Name = null;
                    }

                } else {
                    log.Debug("Updating Google Meet conference data.");
                    if (ev.ConferenceData.ConferenceSolution.Key.Type != "hangoutsMeet") {
                        log.Warn("Unexpected conference solution type '" + ev.ConferenceData.ConferenceSolution.Key.Type + "'. Update abandoned.");
                    } else {
                        EntryPoint ep = ev.ConferenceData.EntryPoints.Where(ep => ep.EntryPointType == "video").FirstOrDefault();
                        log.Fine("Replacing the 'video' conference entry point.");
                        ev.ConferenceData.EntryPoints.Remove(ep);
                        ep.Uri = outlookGMeetUrl;
                        ev.ConferenceData.EntryPoints.Add(ep);
                        ev.ConferenceData.ConferenceSolution.Name = null;
                    }
                }

            } catch (System.Exception ex) {
                Ogcs.Exception.Analyse("Could not alter Event conference data.", ex);
            }
        }
    }
}
