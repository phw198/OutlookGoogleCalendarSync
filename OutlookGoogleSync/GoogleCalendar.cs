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

        private CalendarService service;
        private const String oEntryID = "outlook_EntryID";

        public GoogleCalendar() {
            var provider = new NativeApplicationClient(GoogleAuthenticationServer.Description);
            provider.ClientIdentifier = "653617509806-2nq341ol8ejgqhh2ku4j45m7q2bgdimv.apps.googleusercontent.com";
            provider.ClientSecret = "tAi-gZLWtasS58i8CcCwVwsq";
            service = new CalendarService(new OAuth2Authenticator<NativeApplicationClient>(provider, getAuthentication));
        }

        private static IAuthorizationState getAuthentication(NativeApplicationClient arg) {
            // Get the auth URL:
            IAuthorizationState state = new AuthorizationState(new[] { CalendarService.Scopes.Calendar.GetStringValue() });
            state.Callback = new Uri(NativeApplicationClient.OutOfBandCallbackUrl);
            state.RefreshToken = Settings.Instance.RefreshToken;
            Uri authUri = arg.RequestUserAuthorization(state);

            IAuthorizationState result = null;

            if (state.RefreshToken == "") {
                // Request authorization from the user (by opening a browser window):
                Process.Start(authUri.ToString());

                frmGoogleAuthorizationCode eac = new frmGoogleAuthorizationCode();
                if (eac.ShowDialog() == DialogResult.OK) {
                    // Retrieve the access/refresh tokens by using the authorization code:
                    result = arg.ProcessUserAuthorization(eac.authcode, state);

                    //save the refresh token for future use
                    Settings.Instance.RefreshToken = result.RefreshToken;
                    XMLManager.export(Settings.Instance, MainForm.FILENAME);

                    return result;
                } else {
                    String noAuth = "Sorry, but this application will not work if you don't give it access to your Google Calendar :(";
                    MessageBox.Show(noAuth, "Authorisation not given", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    throw new System.Exception(noAuth);
                }
            } else {
                arg.RefreshToken(state, null);
                result = state;
                return result;
            }

        }

        public List<MyCalendarListEntry> GetCalendars() {
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

        public List<Event> GetCalendarEntriesInRange() {
            List<Event> result = new List<Event>();
            Events request = null;
            String pageToken = null;

            do {
                EventsResource.ListRequest lr = service.Events.List(Settings.Instance.UseGoogleCalendar.Id);

                lr.TimeMin = GoogleTimeFrom(DateTime.Now.AddDays(-Settings.Instance.DaysInThePast));
                lr.TimeMax = GoogleTimeFrom(DateTime.Now.AddDays(+Settings.Instance.DaysInTheFuture));
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

        private void addCalendarEntry(Event e) {
            var result = service.Events.Insert(e, Settings.Instance.UseGoogleCalendar.Id).Fetch();
        }

        private void updateCalendarEntry(Event e) {
            var request = service.Events.Update(e, Settings.Instance.UseGoogleCalendar.Id, e.Id).Fetch();
        }

        private void deleteCalendarEntry(Event e) {
            string request = service.Events.Delete(Settings.Instance.UseGoogleCalendar.Id, e.Id).Fetch();
        }

        public void CreateCalendarEntries(List<AppointmentItem> appointments) {
            foreach (AppointmentItem ai in appointments) {
                Event ev = new Event();

                //Add the Outlook appointment ID into Google event.
                //This will make comparison more efficient and set the scene for 2-way sync.
                ev.ExtendedProperties = new Event.ExtendedPropertiesData();
                ev.ExtendedProperties.Private = new Event.ExtendedPropertiesData.PrivateData();
                ev.ExtendedProperties.Private.Add(oEntryID, ai.EntryID);

                ev.Start = new EventDateTime();
                ev.End = new EventDateTime();

                if (ai.AllDayEvent) {
                    ev.Start.Date = ai.Start.ToString("yyyy-MM-dd");
                    ev.End.Date = ai.End.ToString("yyyy-MM-dd");
                } else {
                    ev.Start.DateTime = GoogleCalendar.GoogleTimeFrom(ai.Start);
                    ev.End.DateTime = GoogleCalendar.GoogleTimeFrom(ai.End);
                }
                ev.Summary = ai.Subject;
                if (Settings.Instance.AddDescription) ev.Description = ai.Body;
                ev.Location = ai.Location;
                ev.Visibility = (ai.Sensitivity == OlSensitivity.olNormal) ? "default" : "private";
                ev.Transparency = (ai.BusyStatus == OlBusyStatus.olFree) ? "transparent" : "opaque";

                //This always reverts to the Google Calendar! So this code doesn't work :(
                //Will leave it in anyway - doesn't do any harm...
                //Organiser may not be an attendee, so do this first.
                ev.Organizer = new Event.OrganizerData();
                ev.Organizer.Self = (ai.Organizer == OutlookCalendar.Instance.CurrentUserName);
                ev.Organizer.DisplayName = ai.Organizer;
                ev.Organizer.Email = ai.GetOrganizer().Address;

                if (ai.Recipients.Count > 1) { //Don't add attendees if there's only 1 (me)
                    ev.Attendees = new List<EventAttendee>();
                    foreach (Microsoft.Office.Interop.Outlook.Recipient recipient in ai.Recipients) {
                        EventAttendee ea = GoogleCalendar.AddAttendee(recipient, ai);
                        if (Settings.Instance.AddAttendees) ev.Attendees.Add(ea);
                    }
                }

                //Reminder alert
                if (Settings.Instance.AddReminders && ai.ReminderSet) {
                    ev.Reminders = new Event.RemindersData();
                    ev.Reminders.UseDefault = false;
                    EventReminder reminder = new EventReminder();
                    reminder.Method = "popup";
                    reminder.Minutes = ai.ReminderMinutesBeforeStart;
                    ev.Reminders.Overrides = new List<EventReminder>();
                    ev.Reminders.Overrides.Add(reminder);
                }

                MainForm.Instance.Logboxout(getEventSummary(ev), verbose: true);
                GoogleCalendar.Instance.addCalendarEntry(ev);
            }
        }

        public void UpdateCalendarEntries(Dictionary<AppointmentItem, Event> entriesToBeCompared, ref int entriesUpdated) {
            entriesUpdated = 0;
            foreach (KeyValuePair<AppointmentItem, Event> compare in entriesToBeCompared) {
                AppointmentItem ai = compare.Key;
                Event ev = compare.Value;
                if (DateTime.Parse(ev.Updated) > DateTime.Parse(GoogleCalendar.GoogleTimeFrom(ai.LastModificationTime))) continue;

                int itemModified = 0;
                System.Text.StringBuilder sb = new System.Text.StringBuilder();
                sb.AppendLine(ai.Subject);
                if (ai.AllDayEvent) {
                    if (MainForm.CompareAttribute("Start time", ev.Start.Date, ai.Start.ToString("yyyy-MM-dd"), sb, ref itemModified)) {
                        ev.Start.Date = ai.Start.ToString("yyyy-MM-dd");
                        ev.Start.DateTime = null;
                    }
                    if (MainForm.CompareAttribute("End time", ev.End.Date, ai.End.ToString("yyyy-MM-dd"), sb, ref itemModified)) {
                        ev.End.Date = ai.End.ToString("yyyy-MM-dd");
                        ev.End.DateTime = null;
                    }
                } else {
                    if (MainForm.CompareAttribute("Start time", ev.Start.DateTime, GoogleCalendar.GoogleTimeFrom(ai.Start), sb, ref itemModified)) {
                        ev.Start.DateTime = GoogleCalendar.GoogleTimeFrom(ai.Start);
                        ev.Start.Date = null;
                    }
                    if (MainForm.CompareAttribute("End time", ev.End.DateTime, GoogleCalendar.GoogleTimeFrom(ai.End), sb, ref itemModified)) {
                        ev.End.DateTime = GoogleCalendar.GoogleTimeFrom(ai.End);
                        ev.End.Date = null;
                    }
                }
                if (MainForm.CompareAttribute("Subject", ev.Summary, ai.Subject, sb, ref itemModified)) {
                    ev.Summary = ai.Subject;
                }
                if (Settings.Instance.AddDescription) {
                    if (MainForm.CompareAttribute("Description", ev.Description, ai.Body, sb, ref itemModified)) ev.Description = ai.Body;
                }
                if (MainForm.CompareAttribute("Location", ev.Location, ai.Location, sb, ref itemModified)) ev.Location = ai.Location;

                String oPrivacy = (ai.Sensitivity == OlSensitivity.olNormal) ? "default" : "private";
                if (MainForm.CompareAttribute("Private", (ev.Visibility==null?"default":ev.Visibility), oPrivacy, sb, ref itemModified)) {
                    ev.Visibility = oPrivacy;
                }
                String oFreeBusy = (ai.BusyStatus == OlBusyStatus.olFree) ? "transparent" : "opaque";
                if (MainForm.CompareAttribute("Free/Busy", (ev.Transparency==null?"opaque":ev.Transparency), oFreeBusy, sb, ref itemModified)) {
                    ev.Transparency = oFreeBusy;
                }

                /*Organiser always reverts to the Google Calendar! So this doesn't work :(
                if (MainForm.CompareAttribute("Organiser", ev.Organizer.DisplayName, ai.Organizer, sb, ref itemModified)) {
                    ev.Organizer.DisplayName = ai.Organizer;
                    ev.Organizer.Self = (ai.Organizer == OutlookCalendar.Instance.CurrentUserName);
                    ev.Organizer.Email = ai.GetOrganizer().Address;
                }*/

                if (Settings.Instance.AddAttendees) {
                    //Build a list of Google attendees. Any remaining at the end of the diff must be deleted.
                    List<EventAttendee> removeAttendee = new List<EventAttendee>();
                    if (ev.Attendees != null) {
                        foreach (EventAttendee ea in ev.Attendees) {
                            removeAttendee.Add(ea);
                        }
                    }
                    if (ai.Recipients.Count > 1) {
                        for (int o = ai.Recipients.Count; o > 0; o--) {
                            bool foundAttendee = false;
                            Recipient recipient = ai.Recipients[o];
                            Microsoft.Office.Interop.Outlook.PropertyAccessor pa = recipient.PropertyAccessor;

                            if (ev.Attendees == null) break;
                            for (int g = ev.Attendees.Count - 1; g >= 0; g--) {
                                String recipientSMTP = pa.GetProperty(OutlookCalendar.PR_SMTP_ADDRESS).ToString().ToLower();
                                if (recipientSMTP == ev.Attendees[g].Email) {
                                    foundAttendee = true;
                                    removeAttendee.Remove(ev.Attendees[g]);

                                    //Optional attendee
                                    bool oOptional = (ai.OptionalAttendees != null && ai.OptionalAttendees.Contains(recipient.Name));
                                    bool gOptional = (ev.Attendees[g].Optional == null) ? false : (bool)ev.Attendees[g].Optional;
                                    if (MainForm.CompareAttribute("Attendee " + ev.Attendees[g].DisplayName + " - Optional",
                                        gOptional, oOptional, sb, ref itemModified)) {
                                        ev.Attendees[g].Optional = oOptional;
                                    }
                                    //Response
                                    switch (recipient.MeetingResponseStatus) {
                                        case OlResponseStatus.olResponseNone:
                                            if (MainForm.CompareAttribute("Attendee " + ev.Attendees[g].DisplayName + " - Response Status",
                                                ev.Attendees[g].ResponseStatus, "needsAction", sb, ref itemModified)) {
                                                ev.Attendees[g].ResponseStatus = "needsAction";
                                            }
                                            break;
                                        case OlResponseStatus.olResponseAccepted:
                                            if (MainForm.CompareAttribute("Attendee " + ev.Attendees[g].DisplayName + " - Response Status",
                                                ev.Attendees[g].ResponseStatus, "accepted", sb, ref itemModified)) {
                                                ev.Attendees[g].ResponseStatus = "accepted";
                                            }
                                            break;
                                        case OlResponseStatus.olResponseDeclined:
                                            if (MainForm.CompareAttribute("Attendee " + ev.Attendees[g].DisplayName + " - Response Status",
                                                ev.Attendees[g].ResponseStatus, "declined", sb, ref itemModified)) {
                                                ev.Attendees[g].ResponseStatus = "declined";
                                            }
                                            break;
                                        case OlResponseStatus.olResponseTentative:
                                            if (MainForm.CompareAttribute("Attendee " + ev.Attendees[g].DisplayName + " - Response Status",
                                                ev.Attendees[g].ResponseStatus, "tentative", sb, ref itemModified)) {
                                                ev.Attendees[g].ResponseStatus = "tentative";
                                            }
                                            break;
                                    }
                                }
                            }
                            if (!foundAttendee) {
                                sb.AppendLine("Attendee added: " + recipient.Name);
                                ev.Attendees.Add(GoogleCalendar.AddAttendee(ai.Recipients[o], ai));
                                itemModified++;
                            }
                        }
                    } //more than just 1 (me) recipients

                    foreach (EventAttendee ea in removeAttendee) {
                        sb.AppendLine("Attendee removed: " + ea.DisplayName);
                        ev.Attendees.Remove(ea);
                        itemModified++;
                    }
                    //Reminders
                    if (Settings.Instance.AddReminders) {
                        if (ev.Reminders.Overrides != null) {
                            //Find the popup reminder in Google
                            for (int r = ev.Reminders.Overrides.Count - 1; r >= 0; r--) {
                                EventReminder reminder = ev.Reminders.Overrides[r];
                                if (reminder.Method == "popup") {
                                    if (ai.ReminderSet) {
                                        if (MainForm.CompareAttribute("Reminder", reminder.Minutes.ToString(), ai.ReminderMinutesBeforeStart.ToString(), sb, ref itemModified)) {
                                            reminder.Minutes = ai.ReminderMinutesBeforeStart;
                                        }
                                    } else {
                                        sb.AppendLine("Reminder: " + reminder.Minutes + " => removed");
                                        ev.Reminders.Overrides.Remove(reminder);
                                        ev.Reminders.UseDefault = true;
                                        itemModified++;
                                    } //if Outlook reminders set
                                } //if google reminder found
                            } //foreach reminder

                        } else { //no google reminders set
                            if (ai.ReminderSet) {
                                sb.AppendLine("Reminder: nothing => " + ai.ReminderMinutesBeforeStart);
                                ev.Reminders.UseDefault = false;
                                EventReminder newReminder = new EventReminder();
                                newReminder.Method = "popup";
                                newReminder.Minutes = ai.ReminderMinutesBeforeStart;
                                ev.Reminders.Overrides = new List<EventReminder>();
                                ev.Reminders.Overrides.Add(newReminder);
                                itemModified++;
                            }
                        }
                    }
                }
                if (itemModified > 0) {
                    MainForm.Instance.Logboxout(sb.ToString(), false, verbose:true);
                    MainForm.Instance.Logboxout(itemModified + " attributes updated.", verbose:true);
                    System.Windows.Forms.Application.DoEvents();
                
                    GoogleCalendar.Instance.updateCalendarEntry(ev);
                    entriesUpdated++;
                }
            }
        }

        public void DeleteCalendarEntries(List<Event> events) {
            foreach (Event ev in events) {
                String eventSummary = getEventSummary(ev);
                Boolean delete = true;

                if (Settings.Instance.ConfirmOnDelete) {
                    if (MessageBox.Show("Delete " + eventSummary + "?", "Deletion Confirmation",
                        MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No) {
                        delete = false;
                        MainForm.Instance.Logboxout("Not deleted: " + eventSummary);
                    }
                } else {
                    MainForm.Instance.Logboxout(eventSummary, verbose:true);
                }
                if (delete) {
                    GoogleCalendar.Instance.deleteCalendarEntry(ev);
                    if (Settings.Instance.ConfirmOnDelete) MainForm.Instance.Logboxout("Deleted: " + eventSummary);
                }
            }
                
        }

        public void ReclaimOrphanCalendarEntries(ref List<Event> gEvents, ref List<AppointmentItem> oAppointments) {
            //This is needed for people migrating from other tools, which do not have our OutlookID extendedProperty
            int unclaimed = 0;
            List<Event> unclaimedEvents = new List<Event>();

            foreach (Event ev in gEvents) {
                //Find entries with no Outlook ID
                if (ev.ExtendedProperties == null ||
                    ev.ExtendedProperties.Private == null ||
                    !ev.ExtendedProperties.Private.ContainsKey(oEntryID))
                {
                    unclaimedEvents.Add(ev);
                    foreach (AppointmentItem ai in oAppointments) {
                        //Use simple matching on start,end,subject,location to pair events
                        String a = signature(ev);
                        String b = OutlookCalendar.signature(ai);
                        if (signature(ev) == OutlookCalendar.signature(ai)) {
                            if (ev.ExtendedProperties == null) ev.ExtendedProperties = new Event.ExtendedPropertiesData();
                            if (ev.ExtendedProperties.Private == null) ev.ExtendedProperties.Private = new Event.ExtendedPropertiesData.PrivateData();
                            ev.ExtendedProperties.Private.Add(oEntryID, ai.EntryID);
                            updateCalendarEntry(ev);
                            unclaimedEvents.Remove(ev);
                            MainForm.Instance.Logboxout("Reclaimed: " + getEventSummary(ev), verbose: true);
                            break;
                        }
                    }
                }
            }
            if ((Settings.Instance.SyncDirection == SyncDirection.OutlookToGoogle ||
                    Settings.Instance.SyncDirection == SyncDirection.Bidirectional ) &&
                unclaimedEvents.Count > 0 &&
                !Settings.Instance.MergeItems && !Settings.Instance.DisableDelete && !Settings.Instance.ConfirmOnDelete) {
                    
                if (MessageBox.Show(unclaimed + " Google calendar events can't be matched to Outlook.\r\n" +
                    "Remember, it's recommended to have a dedicated Google calendar to sync with, "+
                    "or you may wish to merge with unmatched events. Continue with deletions?",
                    "Delete unmatched Google events?", MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation) == DialogResult.No) {
                        
                    foreach (Event e in unclaimedEvents) {
                        gEvents.Remove(e);
                    }
                } 
            }
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
            ev.Start.DateTime = (ev.Start.DateTime == null) ? 
                GoogleTimeFrom(DateTime.Parse(ev.Start.Date)) :
                GoogleTimeFrom(DateTime.Parse(ev.Start.DateTime));
            ev.End.DateTime = (ev.End.DateTime == null) ?
                GoogleTimeFrom(DateTime.Parse(ev.End.Date)) :
                GoogleTimeFrom(DateTime.Parse(ev.End.DateTime));

            return (ev.Start.DateTime + ";" + ev.End.DateTime + ";" + ev.Summary + ";" + ev.Location).Trim();
        }

        public static string getEventSummary(Event ev) {
            String eventSummary = "";
            if (ev.Start.DateTime != null)
                eventSummary += DateTime.Parse(ev.Start.DateTime.ToString()).ToString("dd/MM/yyyy hh:mm");
            else
                eventSummary += DateTime.Parse(ev.Start.Date.ToString()).ToString("dd/MM/yyyy");
            eventSummary += " => ";
            eventSummary += '"' + ev.Summary + '"';
            return eventSummary;
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
            ref List<AppointmentItem> outlook,
            ref List<Event> google,
            Dictionary<AppointmentItem, Event> compare) {
            // Count backwards so that we can remove found items without affecting the order of remaining items
            for (int o = outlook.Count - 1; o >= 0; o--) {
                for (int g = google.Count - 1; g >= 0; g--) {
                    if (google[g].ExtendedProperties != null &&
                        google[g].ExtendedProperties.Private != null &&
                        google[g].ExtendedProperties.Private.ContainsKey(oEntryID)) {
                        
                        if (outlook[o].EntryID == google[g].ExtendedProperties.Private[oEntryID]) {
                            compare.Add(outlook[o], google[g]);
                            outlook.Remove(outlook[o]);
                            google.Remove(google[g]);
                            break;
                        }
                    } else if (Settings.Instance.MergeItems && !Settings.Instance.DisableDelete) {
                        //Remove the non-Outlook item so it doesn't get deleted
                        google.Remove(google[g]);
                    }
                }
            }
            if (Settings.Instance.DisableDelete) {
                google = new List<Event>();
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
