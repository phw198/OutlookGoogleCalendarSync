using System;
using System.Collections.Generic;
using System.Windows.Forms;
using System.Diagnostics;
using DotNetOpenAuth.OAuth2;
using Google.Apis.Authentication;
using Google.Apis.Authentication.OAuth2;
using Google.Apis.Authentication.OAuth2.DotNetOpenAuth;
using Google.Apis.Calendar.v3;
using Google.Apis.Calendar.v3.Data;
using Google.Apis.Util;


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
            try {
                request = service.CalendarList.List().Fetch();
            } catch (Exception ex) {
                MainForm.Instance.HandleException(ex);
            }

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

            try {
                EventsResource.ListRequest lr = service.Events.List(Settings.Instance.UseGoogleCalendar.Id);

                lr.TimeMin = GoogleTimeFrom(DateTime.Now.AddDays(-Settings.Instance.DaysInThePast));
                lr.TimeMax = GoogleTimeFrom(DateTime.Now.AddDays(+Settings.Instance.DaysInTheFuture + 1));

                request = lr.Fetch();
            } catch (Exception ex) {
                MainForm.Instance.HandleException(ex);
            }

            if (request != null) {
                if (request.Items != null) result.AddRange(request.Items);
            }
            return result;
        }

        public void deleteCalendarEntry(Event e) {
            string request;

            try {
                request = service.Events.Delete(Settings.Instance.UseGoogleCalendar.Id, e.Id).Fetch();
            } catch (Exception ex) {
                MainForm.Instance.HandleException(ex);
            }
        }

        public void addEntry(Event e) {
            try {
                var result = service.Events.Insert(e, Settings.Instance.UseGoogleCalendar.Id).Fetch();
            } catch (Exception ex) {
                MainForm.Instance.HandleException(ex);
            }
        }


        //returns the Google Time Format String of a given .Net DateTime value
        //Google Time Format = "2012-08-20T00:00:00+02:00"
        public string GoogleTimeFrom(DateTime dt) {
            string timezone = TimeZoneInfo.Local.GetUtcOffset(dt).ToString();
            if (timezone[0] != '-') timezone = '+' + timezone;
            timezone = timezone.Substring(0, 6);

            string result = dt.GetDateTimeFormats('s')[0] + timezone;
            return result;
        }


    }
}
