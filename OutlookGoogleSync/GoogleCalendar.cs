using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Windows.Forms;
using DotNetOpenAuth.OAuth2;
using Google.Apis.Authentication.OAuth2;
using Google.Apis.Authentication.OAuth2.DotNetOpenAuth;
using Google.Apis.Calendar.v3;
using Google.Apis.Calendar.v3.Data;
using Google.Apis.Util;
using Microsoft.Office.Interop.Outlook;

namespace OutlookGoogleSync {
    /// <summary>
    /// Description of GoogleCalendar.
    /// </summary>
    public class GoogleCalendar {

        private static GoogleCalendar instance;

        public static GoogleCalendar Instance {
            get {
                if (instance == null) instance = new GoogleCalendar();
                return instance;
            }
        }

        CalendarService service;

        public GoogleCalendar() {
            var provider = new NativeApplicationClient(GoogleAuthenticationServer.Description);
            provider.ClientIdentifier = "662204240419.apps.googleusercontent.com";
            provider.ClientSecret = "4nJPnk5fE8yJM_HNUNQEEvjU";
            service = new CalendarService(new OAuth2Authenticator<NativeApplicationClient>(provider, GetAuthentication));
            service.Key = "AIzaSyDRGFSAyMGondZKR8fww1RtRARYtCbBC4k";
        }


        private static IAuthorizationState GetAuthentication(NativeApplicationClient arg) {
            // Get the auth URL:
            IAuthorizationState state = new AuthorizationState(new[] { CalendarService.Scopes.Calendar.GetStringValue() });
            state.Callback = new Uri(NativeApplicationClient.OutOfBandCallbackUrl);
            state.RefreshToken = Settings.Instance.RefreshToken;
            Uri authUri = arg.RequestUserAuthorization(state);

            IAuthorizationState result = null;

            if (state.RefreshToken == "") {
                // Request authorization from the user (by opening a browser window):
                Process.Start(authUri.ToString());

                EnterAuthorizationCode eac = new EnterAuthorizationCode();
                if (eac.ShowDialog() == DialogResult.OK) {
                    // Retrieve the access/refresh tokens by using the authorization code:
                    result = arg.ProcessUserAuthorization(eac.authcode, state);

                    //save the refresh token for future use
                    Settings.Instance.RefreshToken = result.RefreshToken;
                    XMLManager.export(Settings.Instance, MainForm.FILENAME);

                    return result;
                } else {
                    return null;
                }
            } else {
                arg.RefreshToken(state, null);
                result = state;
                return result;
            }

        }

        public List<MyCalendarListEntry> getCalendars() {
            CalendarList request = null;
            request = service.CalendarList.List().Fetch();
            
            if (request != null) {

                List<MyCalendarListEntry> result = new List<MyCalendarListEntry>();
                foreach (CalendarListEntry cle in request.Items) {
                    result.Add(new MyCalendarListEntry(cle));
                }
                return result;
            }
            return null;
        }



        public List<Event> getCalendarEntriesInRange() {
            List<Event> result = new List<Event>();
            Events request = null;
            String pageToken = null;

            do {
                EventsResource.ListRequest lr = service.Events.List(Settings.Instance.UseGoogleCalendar.Id);

                lr.TimeMin = GoogleTimeFrom(DateTime.Now.AddDays(-Settings.Instance.DaysInThePast));
                lr.TimeMax = GoogleTimeFrom(DateTime.Now.AddDays(+Settings.Instance.DaysInTheFuture + 1));
                //lr.OrderBy = EventsResource.OrderBy.StartTime;
                lr.PageToken = pageToken;

                request = lr.Fetch();
                pageToken = request.NextPageToken;

                if (request != null) {
                    if (request.Items != null) result.AddRange(request.Items);
                }
            } while (pageToken != null);

            if (Settings.Instance.CreateTextFiles) {
                TextWriter tw = new StreamWriter("export_found_in_google.txt");
                foreach (Event ev in result) {
                    tw.WriteLine(signature(ev));
                }
                tw.Close();
            }

            return result;
        }

        public void addEntry(Event e) {
            var result = service.Events.Insert(e, Settings.Instance.UseGoogleCalendar.Id).Fetch();
        }

        public void updateCalendarEntry(Event e) {
            var request = service.Events.Update(e, Settings.Instance.UseGoogleCalendar.Id, e.Id).Fetch();
        }

        public void deleteCalendarEntry(Event e) {
            string request = service.Events.Delete(Settings.Instance.UseGoogleCalendar.Id, e.Id).Fetch();
        }

        #region STATIC FUNCTIONS
        //returns the Google Time Format String of a given .Net DateTime value
        //Google Time Format = "2012-08-20T00:00:00+02:00"
        public static string GoogleTimeFrom(DateTime dt) {
            string timezone = TimeZoneInfo.Local.GetUtcOffset(dt).ToString();
            if (timezone[0] != '-') timezone = '+' + timezone;
            timezone = timezone.Substring(0, 6);

            string result = dt.GetDateTimeFormats('s')[0] + timezone;
            return result;
        }
        
        public static string signature(Event ev) {
            if (ev.Start.DateTime == null) ev.Start.DateTime = GoogleCalendar.GoogleTimeFrom(DateTime.Parse(ev.Start.Date));
            if (ev.End.DateTime == null) ev.End.DateTime = GoogleCalendar.GoogleTimeFrom(DateTime.Parse(ev.End.Date));
            return (ev.Start.DateTime + ";" + ev.End.DateTime + ";" + ev.Summary + ";" + ev.Location).Trim();
        }

        public static EventAttendee AddAttendee(Recipient recipient, AppointmentItem ai) {
            EventAttendee ea = new EventAttendee();
            ea.DisplayName = recipient.Name;
            Microsoft.Office.Interop.Outlook.PropertyAccessor pa = recipient.PropertyAccessor;
            ea.Email = pa.GetProperty(OutlookCalendar.PR_SMTP_ADDRESS).ToString();
            ea.Optional = (ai.OptionalAttendees != null && ai.OptionalAttendees.Contains(recipient.Name));
            ea.Organizer = (ai.Organizer == recipient.Name);
            ea.Self = (OutlookCalendar.Instance.CurrentUserName == recipient.Name);
            switch (recipient.MeetingResponseStatus) {
                case OlResponseStatus.olResponseNone: ea.ResponseStatus = "needsAction"; break;
                case OlResponseStatus.olResponseAccepted: ea.ResponseStatus = "accepted"; break;
                case OlResponseStatus.olResponseDeclined: ea.ResponseStatus = "declined"; break;
                case OlResponseStatus.olResponseTentative: ea.ResponseStatus = "tentative"; break;
            }
            return ea;
        }

        //<summary>New logic for comparing Outlook and Google events works as follows:
        //      1.  Scan through both lists looking for duplicates
        //      2.  Remove found duplicates from both lists
        //      3.  Items remaining in Outlook list are new and need to be created
        //      4.  Items remaining in Google list need to be deleted
        //</summary>
        public static void IdentifyEventDifferences(
            List<AppointmentItem> outlook,
            List<Event> google,
            Dictionary<AppointmentItem, Event> compare) {
            // Count backwards so that we can remove found items without affecting the order of remaining items
            for (int o = outlook.Count - 1; o >= 0; o--) {
                for (int g = google.Count - 1; g >= 0; g--) {
                    if (google[g].ExtendedProperties != null &&
                        google[g].ExtendedProperties.Private.ContainsKey("outlook_EntryID") &&
                        outlook[o].EntryID == google[g].ExtendedProperties.Private["outlook_EntryID"]) {
                        compare.Add(outlook[o], google[g]);
                        outlook.Remove(outlook[o]);
                        google.Remove(google[g]);
                        break;
                    }
                }
            }
            if (Settings.Instance.CreateTextFiles) {
                //Google Deletions
                TextWriter tw = new StreamWriter("export_to_be_deleted.txt");
                foreach (Event ev in google) {
                    tw.WriteLine(signature(ev));
                }
                tw.Close();

                //Google Creations
                tw = new StreamWriter("export_to_be_created.txt");
                foreach (AppointmentItem ai in outlook) {
                    tw.WriteLine(OutlookCalendar.signature(ai));
                }
                tw.Close();
            }
        }
        #endregion
    }
}
