using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using DotNetOpenAuth.OAuth2;
using Google.Apis.Authentication.OAuth2;
using Google.Apis.Authentication.OAuth2.DotNetOpenAuth;
using Google.Apis.Calendar.v3;
using Google.Apis.Calendar.v3.Data;
using Google.Apis.Util;
using log4net;
using Microsoft.Office.Interop.Outlook;

namespace OutlookGoogleCalendarSync {
    /// <summary>
    /// Description of GoogleCalendar.
    /// </summary>
    public class GoogleCalendar {
        private static readonly ILog log = LogManager.GetLogger(typeof(GoogleCalendar));

        private static GoogleCalendar instance;
        public static GoogleCalendar Instance {
            get {
                if (instance == null) {
                    instance = new GoogleCalendar();
                    instance.initCalendarService();
                    instance.UserSubscriptionCheck();
                }
                return instance;
            }
        }

        private CalendarService service;
        public const String oEntryID = "outlook_EntryID";
        
        public static Boolean APIlimitReached_attendee = false;
        private const int backoffLimit = 5;
        private enum apiException {
            justContinue,
            backoffThenRetry,
            freeAPIexhausted,
            throwException
        }
        private static Random random = new Random();
        public long MinDefaultReminder = long.MinValue;

        public GoogleCalendar() { }

        public void initCalendarService() {
            var provider = new NativeApplicationClient(GoogleAuthenticationServer.Description);
            if (Settings.Instance.Subscribed != null && Settings.Instance.Subscribed != DateTime.Parse("01-Jan-2000")) {
                provider.ClientIdentifier = "550071650559-44lnvhdu5liq5kftj5t8k0aasgei5g7t.apps.googleusercontent.com";
                provider.ClientSecret = "MGUFapefXClJa2ysS4WNGS4k";
            } else {
                provider.ClientIdentifier = "653617509806-2nq341ol8ejgqhh2ku4j45m7q2bgdimv.apps.googleusercontent.com";
                provider.ClientSecret = "tAi-gZLWtasS58i8CcCwVwsq";
            }
            service = new CalendarService(new OAuth2Authenticator<NativeApplicationClient>(provider, getAuthentication));
        }

        public void Reset() {
            Settings.Instance.RefreshToken = "";
            initCalendarService();
        }

        public String ChangeAPIkeys() {
            String msg = "Your authorisation token is no longer valid and has been removed.\r\n" +
                        "The process to reauthorise access to your Google account will now begin...";
            MessageBox.Show(msg, "Authorisation token invalid", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            Reset();
            return msg;
        }

        public List<MyGoogleCalendarListEntry> GetCalendars() {
            CalendarList request = null;
            int backoff = 0;
            while (backoff < backoffLimit) {
                try {
                    request = service.CalendarList.List().Fetch();
                    break;
                } catch (Google.GoogleApiException ex) {
                    switch (handleAPIlimits(ex, null)) {
                        case apiException.throwException: throw;
                        case apiException.freeAPIexhausted: throw;
                        case apiException.backoffThenRetry: {
                            backoff++;
                            if (backoff == backoffLimit) {
                                log.Error("API limit backoff was not successful. Retrieve calendar list failed.");
                                throw;
                            } else {
                                log.Warn("API rate limit reached. Backing off " + backoff + "sec before retry.");
                                System.Threading.Thread.Sleep(backoff * 1000);
                            }
                            break;
                        }
                    }
                }
            }

            if (request != null) {
                List<MyGoogleCalendarListEntry> result = new List<MyGoogleCalendarListEntry>();
                foreach (CalendarListEntry cle in request.Items) {
                    result.Add(new MyGoogleCalendarListEntry(cle));
                }
                return result;
            } else {
                log.Error("Handshaking with the Google calendar service failed.");
            }
            return null;
        }
        
        public List<Event> GetCalendarEntriesInRecurrence(String recurringEventId) {
            List<Event> result = new List<Event>();
            Events request = null;
            String pageToken = null;
            Int16 pageNum = 1;

            try {
                log.Debug("Retrieving all recurring event instances from Google.");
                do {
                    EventsResource.InstancesRequest ir = service.Events.Instances(Settings.Instance.UseGoogleCalendar.Id, recurringEventId);
                    ir.ShowDeleted = true;
                    ir.PageToken = pageToken;
                    int backoff = 0;
                    while (backoff < backoffLimit) {
                        try {
                            request = ir.Fetch();
                            log.Debug("Page " + pageNum + " received.");
                            break;
                        } catch (Google.GoogleApiException ex) {
                            switch (handleAPIlimits(ex, null)) {
                                case apiException.throwException: throw;
                                case apiException.freeAPIexhausted: throw;
                                case apiException.backoffThenRetry: {
                                        backoff++;
                                        if (backoff == backoffLimit) {
                                            log.Error("API limit backoff was not successful. Paginated retrieve failed.");
                                            throw;
                                        } else {
                                            log.Warn("API rate limit reached. Backing off " + backoff + "sec before retry.");
                                            System.Threading.Thread.Sleep(backoff * 1000);
                                        }
                                        break;
                                    }
                            }
                        }
                    }

                    if (request != null) {
                        pageToken = request.NextPageToken;
                        pageNum++;
                        if (request.Items != null) result.AddRange(request.Items);
                    }
                } while (pageToken != null);
                return result;

            } catch (System.Exception ex) {
                MainForm.Instance.Logboxout("ERROR: Failed to retrieve recurring events");
                log.Error(ex.Message);
                return null;
            }
        }

        public Event GetCalendarEntry(String eventId) {
            Event request = null;

            try {
                log.Debug("Retrieving specific Event with ID " + eventId);
                EventsResource.GetRequest gr = service.Events.Get(Settings.Instance.UseGoogleCalendar.Id, eventId);
                int backoff = 0;
                while (backoff < backoffLimit) {
                    try {
                        request = gr.Fetch();
                        break;
                    } catch (Google.GoogleApiException ex) {
                        switch (handleAPIlimits(ex, null)) {
                            case apiException.throwException: throw;
                            case apiException.freeAPIexhausted: throw;
                            case apiException.backoffThenRetry: {
                                backoff++;
                                if (backoff == backoffLimit) {
                                    log.Error("API limit backoff was not successful. Retrieve failed.");
                                    throw;
                                } else {
                                    log.Warn("API rate limit reached. Backing off " + backoff + "sec before retry.");
                                    System.Threading.Thread.Sleep(backoff * 1000);
                                }
                                break;
                            }
                        }
                    }
                }

                if (request != null)
                    return request;
                else
                    throw new System.Exception("Returned null");
            } catch (System.Exception ex) {
                MainForm.Instance.Logboxout("ERROR: Failed to retrieve Google event");
                log.Error(ex.Message);
                return null;
            }
        }

        public List<Event> GetCalendarEntriesInRange() {
            return GetCalendarEntriesInRange(Settings.Instance.SyncStart, Settings.Instance.SyncEnd);
        }

        public List<Event> GetCalendarEntriesInRange(DateTime from, DateTime to) {
            List<Event> result = new List<Event>();
            Events request = null;
            String pageToken = null;
            Int16 pageNum = 1;
            
            log.Debug("Retrieving all events from Google: " + from.ToShortDateString() + " -> " + to.ToShortDateString());
            do {
                EventsResource.ListRequest lr = service.Events.List(Settings.Instance.UseGoogleCalendar.Id);
                
                lr.TimeMin = GoogleTimeFrom(from);
                lr.TimeMax = GoogleTimeFrom(to);
                lr.PageToken = pageToken;
                lr.SingleEvents = false;
                
                int backoff = 0;
                while (backoff < backoffLimit) {
                    try {
                        request = lr.Fetch();
                        log.Debug("Page " + pageNum + " received.");
                        break;
                    } catch (Google.GoogleApiException ex) {
                        switch (handleAPIlimits(ex, null)) {
                            case apiException.throwException: throw;
                            case apiException.freeAPIexhausted: throw;
                            case apiException.backoffThenRetry: {
                                backoff++;
                                if (backoff == backoffLimit) {
                                    log.Error("API limit backoff was not successful. Retrieve failed.");
                                    throw;
                                } else {
                                    log.Warn("API rate limit reached. Backing off " + backoff + "sec before retry.");
                                    System.Threading.Thread.Sleep(backoff * 1000);
                                }
                                break;
                            }
                        }
                    }
                }

                if (request != null) {
                    pageToken = request.NextPageToken;
                    pageNum++;
                    if (request.Items != null) result.AddRange(request.Items);
                }
            } while (pageToken != null);

            if (Settings.Instance.CreateCSVFiles) {
                ExportToCSV("Outputting all Events to CSV", "google_events.csv", result);
            }

            return result;
        }

        #region Create
        public void CreateCalendarEntries(List<AppointmentItem> appointments) {
            foreach (AppointmentItem ai in appointments) {
                Event newEvent = new Event();
                try {
                    newEvent = createCalendarEntry(ai);
                } catch (System.Exception ex) {
                    if (!Settings.Instance.VerboseOutput) MainForm.Instance.Logboxout(OutlookCalendar.GetEventSummary(ai));
                    MainForm.Instance.Logboxout("WARNING: Event creation failed.\r\n" + ex.Message);
                    log.Error(ex.StackTrace);
                    if (MessageBox.Show("Google event creation failed. Continue with synchronisation?", "Sync item failed", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                        continue;
                    else {
                        throw new UserCancelledSyncException("User chose not to continue sync.");
                    }
                }

                Event createdEvent = new Event();
                try {
                    createdEvent = createCalendarEntry_save(newEvent, ai);
                } catch (System.Exception ex) {
                    MainForm.Instance.Logboxout("WARNING: New event failed to save.\r\n" + ex.Message);
                    log.Error(ex.StackTrace);
                    if (MessageBox.Show("New Google event failed to save. Continue with synchronisation?", "Sync item failed", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                        continue;
                    else 
                        throw new UserCancelledSyncException("User chose not to continue sync.");
                }

                if (ai.IsRecurring && Recurrence.HasExceptions(ai) && createdEvent != null) {
                    MainForm.Instance.Logboxout("This is a recurring item with some exceptions:-");
                    Recurrence.CreateGoogleExceptions(ai, createdEvent.Id);
                    MainForm.Instance.Logboxout("Recurring exceptions completed.");
                }
            }
        }

        private Event createCalendarEntry(AppointmentItem ai) {
            string itemSummary = OutlookCalendar.GetEventSummary(ai);
            log.Debug("Processing >> " + itemSummary);
            MainForm.Instance.Logboxout(itemSummary, verbose: true);

            Event ev = new Event(); 
            //Add the Outlook appointment ID into Google event
            AddOutlookID(ref ev, ai);
                
            ev.Start = new EventDateTime();
            ev.End = new EventDateTime();

            ev.Recurrence = Recurrence.Instance.BuildGooglePattern(ai, ev);
            if (ev.Recurrence != null) {
                ev = OutlookCalendar.Instance.IOutlook.IANAtimezone_set(ev, ai);
            }
                            
            if (ai.AllDayEvent) {
                ev.Start.Date = ai.Start.ToString("yyyy-MM-dd");
                ev.End.Date = ai.End.ToString("yyyy-MM-dd");
            } else {
                ev.Start.DateTime = GoogleCalendar.GoogleTimeFrom(ai.Start);
                ev.End.DateTime = GoogleCalendar.GoogleTimeFrom(ai.End);
            }
            ev.Summary = Obfuscate.ApplyRegex(ai.Subject, SyncDirection.OutlookToGoogle);
            if (Settings.Instance.AddDescription) ev.Description = ai.Body;
            ev.Location = ai.Location;
            ev.Visibility = (ai.Sensitivity == OlSensitivity.olNormal) ? "default" : "private";
            ev.Transparency = (ai.BusyStatus == OlBusyStatus.olFree) ? "transparent" : "opaque";

            ev.Attendees = new List<EventAttendee>();
            if (Settings.Instance.AddAttendees && ai.Recipients.Count > 1 && !APIlimitReached_attendee) { //Don't add attendees if there's only 1 (me)
                if (ai.Recipients.Count >= 200) {
                    MainForm.Instance.Logboxout("ALERT: Attendees will not be synced for this meeting as it has " +
                        "more than 200, which Google does not allow.");
                } else {
                    foreach (Microsoft.Office.Interop.Outlook.Recipient recipient in ai.Recipients) {
                        EventAttendee ea = GoogleCalendar.CreateAttendee(recipient);
                        ev.Attendees.Add(ea);
                    }
                }
            }

            //Reminder alert
            if (Settings.Instance.AddReminders) {
                ev.Reminders = new Event.RemindersData();
                if (OutlookCalendar.Instance.IsOKtoSyncReminder(ai)) {
                    if (ai.ReminderSet) {
                        ev.Reminders.UseDefault = false;
                        EventReminder reminder = new EventReminder();
                        reminder.Method = "popup";
                        reminder.Minutes = ai.ReminderMinutesBeforeStart;
                        ev.Reminders.Overrides = new List<EventReminder>();
                        ev.Reminders.Overrides.Add(reminder);
                    } else {
                        ev.Reminders.UseDefault = Settings.Instance.UseGoogleDefaultReminder;
                    }
                } else {
                    ev.Reminders.UseDefault = false;
                }
            }
            return ev;
        }

        private Event createCalendarEntry_save(Event ev, AppointmentItem ai) {
            if (Settings.Instance.SyncDirection == SyncDirection.Bidirectional) {
                log.Debug("Saving timestamp when OGCS updated event.");
                setOGCSlastModified(ref ev);
            }
            if (Settings.Instance.APIlimit_inEffect) {
                addOGCSproperty(ref ev, Program.APIlimitHit, "True");
            }

            Event createdEvent = new Event();
            int backoff = 0;
            while (backoff < backoffLimit) {
                try {
                    createdEvent = service.Events.Insert(ev, Settings.Instance.UseGoogleCalendar.Id).Fetch();
                    if (Settings.Instance.AddAttendees && Settings.Instance.APIlimit_inEffect) {
                        log.Info("API limit for attendee sync lifted :-)");
                        Settings.Instance.APIlimit_inEffect = false;
                    }
                    break;
                } catch (Google.GoogleApiException ex) {
                    switch (handleAPIlimits(ex, ev)) {
                        case apiException.throwException: throw; 
                        case apiException.freeAPIexhausted: throw;
                        case apiException.justContinue: break;
                        case apiException.backoffThenRetry: {
                            backoff++;
                            if (backoff == backoffLimit) {
                                log.Error("API limit backoff was not successful. Save failed.");
                                throw;
                            } else {
                                log.Warn("API rate limit reached. Backing off " + backoff + "sec before retry.");
                                System.Threading.Thread.Sleep(backoff * 1000);
                            }
                            break;
                        }
                    }
                }
            }

            Boolean gKeyExists = false;
            try {
                UserProperty up = ai.UserProperties.Find(OutlookCalendar.gEventID);
                gKeyExists = (up != null);
            } catch { }
            if (Settings.Instance.SyncDirection == SyncDirection.Bidirectional || gKeyExists) {
                log.Debug("Storing the Google event ID in Outlook appointment.");
                OutlookCalendar.AddOGCSproperty(ref ai, OutlookCalendar.gEventID, createdEvent.Id);
                ai.Save();
            }
            //DOS ourself by triggering API limit
            //for (int i = 1; i <= 30; i++) {
            //    MainForm.Instance.Logboxout("Add #" + i, verbose:true);
            //    Event result = service.Events.Insert(e, Settings.Instance.UseGoogleCalendar.Id).Fetch();
            //    System.Threading.Thread.Sleep(300);
            //    GoogleCalendar.Instance.deleteCalendarEntry(result);
            //    System.Threading.Thread.Sleep(300);
            //}
            return createdEvent;
        }
        #endregion

        #region Update
        public void UpdateCalendarEntries(Dictionary<AppointmentItem, Event> entriesToBeCompared, ref int entriesUpdated) {
            entriesUpdated = 0;
            for (int i = 0; i < entriesToBeCompared.Count; i++) {
                KeyValuePair<AppointmentItem, Event> compare = entriesToBeCompared.ElementAt(i);
                int itemModified = 0;
                Event ev = new Event();
                try {
                    ev = UpdateCalendarEntry(compare.Key, compare.Value, ref itemModified);
                } catch (System.Exception ex) {
                    if (!Settings.Instance.VerboseOutput) MainForm.Instance.Logboxout(OutlookCalendar.GetEventSummary(compare.Key));
                    MainForm.Instance.Logboxout("WARNING: Event update failed.\r\n" + ex.Message);
                    log.Error(ex.StackTrace);
                    if (MessageBox.Show("Google event update failed. Continue with synchronisation?", "Sync item failed", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                        continue;
                    else
                        throw new UserCancelledSyncException("User chose not to continue sync.");
                }

                if (itemModified > 0) {
                    try {
                        UpdateCalendarEntry_save(ref ev);
                        entriesUpdated++;
                    } catch (System.Exception ex) {
                        MainForm.Instance.Logboxout("WARNING: Updated event failed to save.\r\n" + ex.Message);
                        log.Error(ex.StackTrace);
                        if (MessageBox.Show("Updated Google event failed to save. Continue with synchronisation?", "Sync item failed", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                            continue;
                        else
                            throw new UserCancelledSyncException("User chose not to continue sync.");
                    }

                }

                //Have to do this *before* any dummy update, else all the exceptions inherit the updated timestamp of the parent recurring event
                Recurrence.UpdateGoogleExceptions(compare.Key, ev ?? compare.Value);  

                if (itemModified == 0 && ev != null) {
                    log.Debug("Doing a dummy update in order to update the last modified date of "+ 
                        (ev.RecurringEventId == null && ev.Recurrence != null ? "recurring event" : "single instance"));
                    setOGCSlastModified(ref ev);
                    try {
                        UpdateCalendarEntry_save(ref ev);
                        entriesToBeCompared[compare.Key] = ev;
                    } catch (System.Exception ex) {
                        MainForm.Instance.Logboxout("WARNING: Updated event failed to save.\r\n" + ex.Message);
                        log.Error(ex.StackTrace);
                        if (MessageBox.Show("Updated Google event failed to save. Continue with synchronisation?", "Sync item failed", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                            continue;
                        else
                            throw new UserCancelledSyncException("User chose not to continue sync.");
                    }
                }

            }
        }

        public Event UpdateCalendarEntry(AppointmentItem ai, Event ev, ref int itemModified, Boolean forceCompare = false) {
            if (!Settings.Instance.APIlimit_inEffect && GetOGCSproperty(ev, Program.APIlimitHit)) {
                log.Fine("Back processing Event affected by attendee API limit.");
            } else {
                if (!(MainForm.Instance.ManualForceCompare || forceCompare)) { //Needed if the exception has just been created, but now needs updating
                    if (Settings.Instance.SyncDirection != SyncDirection.Bidirectional) {
                        if (DateTime.Parse(ev.Updated) > ai.LastModificationTime)
                            return null;
                    } else {
                        if (OutlookCalendar.GetOGCSlastModified(ai).AddSeconds(5) >= ai.LastModificationTime)
                            //Outlook last modified by OGCS
                            return null;
                        if (DateTime.Parse(ev.Updated) > DateTime.Parse(GoogleCalendar.GoogleTimeFrom(ai.LastModificationTime)))
                            return null;
                    }
                }
            }
                
            String aiSummary = OutlookCalendar.GetEventSummary(ai);
            log.Debug("Processing >> " + aiSummary);

            System.Text.StringBuilder sb = new System.Text.StringBuilder();
            sb.AppendLine(aiSummary);
                
            //Handle an event's all-day attribute being toggled
            String evStart = ev.Start.Date ?? ev.Start.DateTime;
            String evEnd = ev.End.Date ?? ev.End.DateTime;
            if (ai.AllDayEvent) {
                ev.Start.DateTime = null;
                ev.End.DateTime = null;
                if (MainForm.CompareAttribute("Start time", SyncDirection.OutlookToGoogle, evStart, ai.Start.ToString("yyyy-MM-dd"), sb, ref itemModified)) {
                    ev.Start.Date = ai.Start.ToString("yyyy-MM-dd");
                }
                if (MainForm.CompareAttribute("End time", SyncDirection.OutlookToGoogle, evEnd, ai.End.ToString("yyyy-MM-dd"), sb, ref itemModified)) {
                    ev.End.Date = ai.End.ToString("yyyy-MM-dd");
                }
            } else {
                //Handle: Google = all-day; Outlook = not all day, but midnight values (so effectively all day!)
                if (ev.Start.Date != null &&
                    GoogleCalendar.GoogleTimeFrom(DateTime.Parse(evStart)) == GoogleCalendar.GoogleTimeFrom(ai.Start) &&
                    GoogleCalendar.GoogleTimeFrom(DateTime.Parse(evEnd)) == GoogleCalendar.GoogleTimeFrom(ai.End)) 
                {
                    sb.AppendLine("All-Day: true => false");
                    ev.Start.DateTime = GoogleCalendar.GoogleTimeFrom(ai.Start);
                    ev.End.DateTime = GoogleCalendar.GoogleTimeFrom(ai.End);
                    itemModified++;
                }
                ev.Start.Date = null;
                ev.End.Date = null;
                if (MainForm.CompareAttribute("Start time", SyncDirection.OutlookToGoogle,
                    GoogleCalendar.GoogleTimeFrom(DateTime.Parse(evStart)), GoogleCalendar.GoogleTimeFrom(ai.Start), sb, ref itemModified)) {
                    ev.Start.DateTime = GoogleCalendar.GoogleTimeFrom(ai.Start);
                }
                if (MainForm.CompareAttribute("End time", SyncDirection.OutlookToGoogle, 
                    GoogleCalendar.GoogleTimeFrom(DateTime.Parse(evEnd)), GoogleCalendar.GoogleTimeFrom(ai.End), sb, ref itemModified)) {
                    ev.End.DateTime = GoogleCalendar.GoogleTimeFrom(ai.End);
                }
            } 

            List<String> oRrules = Recurrence.Instance.BuildGooglePattern(ai, ev);
            if (ev.Recurrence != null) {
                for (int r = 0; r < ev.Recurrence.Count; r++) {
                    String rrule = ev.Recurrence[r];
                    if (rrule.StartsWith("RRULE:")) {
                        if (oRrules != null) {
                            String[] oRrule_bits = oRrules.First().TrimStart("RRULE:".ToCharArray()).Split(';');
                            foreach (String oRrule_bit in oRrule_bits) {
                                if (!rrule.Contains(oRrule_bit)) {
                                    if (MainForm.CompareAttribute("Recurrence", SyncDirection.OutlookToGoogle, rrule, oRrules.First(), sb, ref itemModified)) {
                                        ev.Recurrence[r] = oRrules.First();
                                    }
                                }
                            }
                        } else {
                            log.Debug("Converting to non-recurring event.");
                            MainForm.CompareAttribute("Recurrence", SyncDirection.OutlookToGoogle, rrule, null, sb, ref itemModified);
                            ev.Recurrence[r] = null;
                        }
                        break;
                    }
                }
            } else {
                if (oRrules != null && ev.RecurringEventId == null) {
                    log.Debug("Converting to recurring event.");
                    MainForm.CompareAttribute("Recurrence", SyncDirection.OutlookToGoogle, null, oRrules.First(), sb, ref itemModified);
                    ev.Recurrence = oRrules;
                }
            }
            if (ev.Recurrence != null && ev.RecurringEventId == null) {
                ev = OutlookCalendar.Instance.IOutlook.IANAtimezone_set(ev, ai);
            }
            
            String subjectObfuscated = Obfuscate.ApplyRegex(ai.Subject, SyncDirection.OutlookToGoogle);
            if (MainForm.CompareAttribute("Subject", SyncDirection.OutlookToGoogle, ev.Summary, subjectObfuscated, sb, ref itemModified)) {
                ev.Summary = subjectObfuscated;
            }
            if (!Settings.Instance.AddDescription) ai.Body = "";
            String outlookBody = ai.Body;
            //Check for Google description truncated @ 8Kb
            if (!string.IsNullOrEmpty(ai.Body) && !string.IsNullOrEmpty(ev.Description)
                && ev.Description.Length == 8 * 1024  
                && ai.Body.Length > 8 * 1024) {
                outlookBody = ai.Body.Substring(0, 8 * 1024);
            }
            if (MainForm.CompareAttribute("Description", SyncDirection.OutlookToGoogle, ev.Description, outlookBody, sb, ref itemModified)) 
                ev.Description = outlookBody;
            
            if (MainForm.CompareAttribute("Location", SyncDirection.OutlookToGoogle, ev.Location, ai.Location, sb, ref itemModified)) 
                ev.Location = ai.Location;

            String oPrivacy = (ai.Sensitivity == OlSensitivity.olNormal) ? "default" : "private";
            String gPrivacy = ev.Visibility ?? "default";
            if (MainForm.CompareAttribute("Private", SyncDirection.OutlookToGoogle, gPrivacy, oPrivacy, sb, ref itemModified)) {
                ev.Visibility = oPrivacy;
            }
            String oFreeBusy = (ai.BusyStatus == OlBusyStatus.olFree) ? "transparent" : "opaque";
            String gFreeBusy = ev.Transparency ?? "opaque";
            if (MainForm.CompareAttribute("Free/Busy", SyncDirection.OutlookToGoogle, gFreeBusy, oFreeBusy, sb, ref itemModified)) {
                ev.Transparency = oFreeBusy;
            }

            if (Settings.Instance.AddAttendees && ai.Recipients.Count > 1 && !APIlimitReached_attendee) {
                if (ai.Recipients.Count >= 200) {
                    MainForm.Instance.Logboxout("ALERT: Attendees will not be synced for this meeting as it has " +
                        "more than 200, which Google does not allow.");
                    ev.Attendees = new List<EventAttendee>();
                } else {
                    try {
                        CompareRecipientsToAttendees(ai, ev, sb, ref itemModified);
                    } catch (System.Exception ex) {
                        if (OutlookCalendar.Instance.IOutlook.ExchangeConnectionMode().ToString().Contains("Disconnected")) {
                            MainForm.Instance.Logboxout("Outlook is currently disconnected from Exchange, so it's not possible to sync attendees.");
                            MainForm.Instance.Logboxout("Please reconnect or do not sync attendees.");
                            throw new System.Exception("Outlook has disconnected from Exchange.");
                        } else {
                            MainForm.Instance.Logboxout("WARNING: Unable to sync attendees.\r\n" + ex.Message);
                        }
                    }
                }
            }
                        
            //Reminders
            if (Settings.Instance.AddReminders) {
                Boolean OKtoSyncReminder = OutlookCalendar.Instance.IsOKtoSyncReminder(ai);
                if (ev.Reminders.Overrides != null) {
                    //Find the popup reminder in Google
                    for (int r = ev.Reminders.Overrides.Count - 1; r >= 0; r--) {
                        EventReminder reminder = ev.Reminders.Overrides[r];
                        if (reminder.Method == "popup") {
                            if (OKtoSyncReminder) {
                                if (ai.ReminderSet) {
                                    if (MainForm.CompareAttribute("Reminder", SyncDirection.OutlookToGoogle, reminder.Minutes.ToString(), ai.ReminderMinutesBeforeStart.ToString(), sb, ref itemModified)) {
                                        reminder.Minutes = ai.ReminderMinutesBeforeStart;
                                    }
                                } else {
                                    sb.AppendLine("Reminder: " + reminder.Minutes + " => removed");
                                    ev.Reminders.Overrides.Remove(reminder);
                                    if (ev.Reminders.Overrides == null || ev.Reminders.Overrides.Count == 0) {
                                        ev.Reminders.UseDefault = Settings.Instance.UseGoogleDefaultReminder;
                                    }
                                    itemModified++;
                                } //if Outlook reminders set
                            } else {
                                sb.AppendLine("Reminder: " + reminder.Minutes + " => removed");
                                ev.Reminders.Overrides.Remove(reminder);
                                ev.Reminders.UseDefault = false;
                                itemModified++;
                            }
                        } //if google reminder found
                    } //foreach reminder

                } else { //no google reminders set
                    if (ai.ReminderSet && OKtoSyncReminder) {
                        sb.AppendLine("Reminder: nothing => " + ai.ReminderMinutesBeforeStart);
                        ev.Reminders.UseDefault = false;
                        EventReminder newReminder = new EventReminder();
                        newReminder.Method = "popup";
                        newReminder.Minutes = ai.ReminderMinutesBeforeStart;
                        ev.Reminders.Overrides = new List<EventReminder>();
                        ev.Reminders.Overrides.Add(newReminder);
                        itemModified++;
                    } else {
                        if (MainForm.CompareAttribute("Reminder Default", SyncDirection.OutlookToGoogle, ev.Reminders.UseDefault.ToString(), OKtoSyncReminder ? Settings.Instance.UseGoogleDefaultReminder.ToString() : "False", sb, ref itemModified)) {
                            ev.Reminders.UseDefault = OKtoSyncReminder ? Settings.Instance.UseGoogleDefaultReminder : false;
                        }
                    }
                }
            }
            if (itemModified > 0) {
                MainForm.Instance.Logboxout(sb.ToString(), false, verbose: true);
                MainForm.Instance.Logboxout(itemModified + " attributes updated.", verbose: true);
                System.Windows.Forms.Application.DoEvents();
            }
            return ev;
        }

        public void UpdateCalendarEntry_save(ref Event ev) {
            if (Settings.Instance.SyncDirection == SyncDirection.Bidirectional) {
                log.Debug("Saving timestamp when OGCS updated event.");
                setOGCSlastModified(ref ev);
            }
            if (Settings.Instance.APIlimit_inEffect)
                addOGCSproperty(ref ev, Program.APIlimitHit, "True");
            else
                removeOGCSproperty(ref ev, Program.APIlimitHit);

            int backoff = 0;
            while (backoff < backoffLimit) {
                try {
                    ev = service.Events.Update(ev, Settings.Instance.UseGoogleCalendar.Id, ev.Id).Fetch();
                    if (Settings.Instance.AddAttendees && Settings.Instance.APIlimit_inEffect) {
                        log.Info("API limit for attendee sync lifted :-)");
                        Settings.Instance.APIlimit_inEffect = false;
                    }
                    break;
                } catch (Google.GoogleApiException ex) {
                    switch (handleAPIlimits(ex, ev)) {
                        case apiException.throwException: throw;
                        case apiException.freeAPIexhausted: throw;
                        case apiException.backoffThenRetry: {
                            backoff++;
                            if (backoff == backoffLimit) {
                                log.Error("API limit backoff was not successful. Save failed.");
                                throw;
                            } else {
                                log.Warn("API rate limit reached. Backing off " + backoff + "sec before retry.");
                                System.Threading.Thread.Sleep(backoff * 1000);
                            }
                            break;
                        }
                    }
                }
            }
        }
        #endregion

        #region Delete
        public void DeleteCalendarEntries(List<Event> events) {
            for (int g=events.Count-1; g>=0; g--) {
                Event ev = events[g];
                Boolean doDelete = false;
                try {
                    doDelete = deleteCalendarEntry(ev);
                } catch (System.Exception ex) {
                    if (!Settings.Instance.VerboseOutput) MainForm.Instance.Logboxout(GoogleCalendar.GetEventSummary(ev));
                    MainForm.Instance.Logboxout("WARNING: Event deletion failed.\r\n" + ex.Message);
                    log.Error(ex.StackTrace);
                    if (MessageBox.Show("Google event deletion failed. Continue with synchronisation?", "Sync item failed", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                        continue;
                    else {
                        throw new UserCancelledSyncException("User chose not to continue sync.");
                    }
                }

                try {
                    if (doDelete) deleteCalendarEntry_save(ev);
                    else events.Remove(ev);
                } catch (System.Exception ex) {
                    MainForm.Instance.Logboxout("WARNING: Deleted event failed to remove.\r\n" + ex.Message);
                    log.Error(ex.StackTrace);
                    if (MessageBox.Show("Deleted Google event failed to remove. Continue with synchronisation?", "Sync item failed", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                        continue;
                    else {
                        throw new UserCancelledSyncException("User chose not to continue sync.");
                    }
                }
            }
        }
        
        private Boolean deleteCalendarEntry(Event ev) {
            String eventSummary = GetEventSummary(ev);
            Boolean doDelete = true;

            if (Settings.Instance.ConfirmOnDelete) {
                if (MessageBox.Show("Delete " + eventSummary + "?", "Deletion Confirmation",
                    MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No) {
                    doDelete = false;
                    MainForm.Instance.Logboxout("Not deleted: " + eventSummary);
                } else {
                    MainForm.Instance.Logboxout("Deleted: " + eventSummary);
                }
            } else {
                MainForm.Instance.Logboxout(eventSummary, verbose:true);
            }
            return doDelete;
        }

        private void deleteCalendarEntry_save(Event ev) {
            int backoff = 0;
            while (backoff < backoffLimit) {
                try {
                    string request = service.Events.Delete(Settings.Instance.UseGoogleCalendar.Id, ev.Id).Fetch();
                    break;
                } catch (Google.GoogleApiException ex) {
                    switch (handleAPIlimits(ex, ev)) {
                        case apiException.throwException: throw;
                        case apiException.freeAPIexhausted: throw;
                        case apiException.backoffThenRetry: {
                            backoff++;
                            if (backoff == backoffLimit) {
                                log.Error("API limit backoff was not successful. Save failed.");
                                throw;
                            } else {
                                log.Warn("API rate limit reached. Backing off " + backoff + "sec before retry.");
                                System.Threading.Thread.Sleep(backoff * 1000);
                            }
                            break;
                        }
                    }
                }
            }
        }
        #endregion

        public void ReclaimOrphanCalendarEntries(ref List<Event> gEvents, ref List<AppointmentItem> oAppointments, Boolean neverDelete = false) {
            log.Debug("Looking for orphaned events to reclaim...");

            //This is needed for people migrating from other tools, which do not have our OutlookID extendedProperty
            List<Event> unclaimedEvents = new List<Event>();

            for (int g = gEvents.Count-1; g >= 0; g--) {
                Event ev = gEvents[g];
                    
                //Find entries with no Outlook ID
                if (!GetOGCSproperty(ev, oEntryID)) {
                    unclaimedEvents.Add(ev);
                    
                    //Use simple matching on start,end,subject,location to pair events
                    String sigEv = signature(ev);
                    if (String.IsNullOrEmpty(sigEv)) {
                        gEvents.Remove(ev);
                        break;
                    }
                    foreach (AppointmentItem ai in oAppointments) {
                        String sigAi = OutlookCalendar.signature(ai);

                        if (Settings.Instance.Obfuscation.Enabled) {
                            if (Settings.Instance.Obfuscation.Direction == SyncDirection.OutlookToGoogle)
                                sigAi = Obfuscate.ApplyRegex(sigAi, SyncDirection.OutlookToGoogle);
                            else
                                sigEv = Obfuscate.ApplyRegex(sigEv, SyncDirection.GoogleToOutlook);
                        }
                        if (sigEv == sigAi) {                            
                            AddOutlookID(ref ev, ai);
                            UpdateCalendarEntry_save(ref ev);
                            unclaimedEvents.Remove(ev);
                            MainForm.Instance.Logboxout("Reclaimed: " + GetEventSummary(ev), verbose: true);
                            gEvents[g] = ev;
                            break;
                        }
                    }
                }
            }
            log.Debug(unclaimedEvents.Count + " unclaimed.");
            if (!neverDelete && unclaimedEvents.Count > 0 &&
                (Settings.Instance.SyncDirection == SyncDirection.OutlookToGoogle ||
                 Settings.Instance.SyncDirection == SyncDirection.Bidirectional )) 
            {
                log.Info(unclaimedEvents.Count + " unclaimed orphan events found.");
                if (Settings.Instance.MergeItems || Settings.Instance.DisableDelete || Settings.Instance.ConfirmOnDelete) {
                    log.Info("These will be kept due to configuration settings.");
                } else if (Settings.Instance.SyncDirection == SyncDirection.Bidirectional) {
                    log.Debug("These 'orphaned' items must not be deleted - they need syncing up.");
                } else {
                    if (MessageBox.Show(unclaimedEvents.Count + " Google calendar events can't be matched to Outlook.\r\n" +
                        "Remember, it's recommended to have a dedicated Google calendar to sync with, " +
                        "or you may wish to merge with unmatched events. Continue with deletions?",
                        "Delete unmatched Google events?", MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation) == DialogResult.No) {

                        log.Info("User has requested to keep them.");
                        foreach (Event e in unclaimedEvents) {
                            gEvents.Remove(e);
                        }
                    } else {
                        log.Info("User has opted to delete them.");
                    }
                }
            }
        }

        //<summary>New logic for comparing Outlook and Google events works as follows:
        //      1.  Scan through both lists looking for duplicates
        //      2.  Remove found duplicates from both lists
        //      3.  Items remaining in Outlook list are new and need to be created
        //      4.  Items remaining in Google list need to be deleted
        //</summary>
        public void IdentifyEventDifferences(
            ref List<AppointmentItem> outlook,  //need creating
            ref List<Event> google,             //need deleting
            Dictionary<AppointmentItem, Event> compare) {
            log.Debug("Comparing Outlook items to Google events...");

            // Count backwards so that we can remove found items without affecting the order of remaining items
            String compare_gEntryID;
            int migrated = 0;
            for (int g = google.Count - 1; g >= 0; g--) {
                if (GetOGCSproperty(google[g], oEntryID, out compare_gEntryID)) {
                    for (int o = outlook.Count - 1; o >= 0; o--) {
                        try {
                            log.Fine("Checking "+ OutlookCalendar.GetEventSummary(outlook[o]));
                            String compare_oGlobalID = OutlookCalendar.Instance.IOutlook.GetGlobalApptID(outlook[o]);
                            if (compare_oGlobalID == outlook[o].EntryID) {
                                log.Debug("Can't migrate Event to GlobalAppointmentID.");
                            } else if (compare_gEntryID.StartsWith(outlook[o].EntryID)) {
                                if (migrated == 0) log.Info("Migrating Events from EntryID to GlobalAppointmentID...");
                                //Should have used AppointmentItem Global ID, not Object ID (which changes when accepting invites!)
                                //To prevent existing users from having everything deleted and recreated, need to migrate them.
                                //This change was effective from v2.0.2.3
                                //*** It's not the most efficient, so should be scrapped after a few Beta releases!
                                Event ev = google[g];
                                AddOutlookID(ref ev, outlook[o]);
                                UpdateCalendarEntry_save(ref ev);
                                google[g] = ev;
                                compare_gEntryID = compare_oGlobalID;
                                migrated++;
                                //There could be a lot of these, so let's not batter Google too hard - it doesn't like >4/sec
                                System.Threading.Thread.Sleep(250);
                            }

                            //For format of Global ID: https://msdn.microsoft.com/en-us/library/ee157690%28v=exchg.80%29.aspx
                            //For items copied from someone elses calendar, it appears the Global ID is generated for each access?! (Creation Time changes)
                            //I guess the copied item doesn't really have its "own" ID. Anyway, we could consider just comparing
                            //the "data" section of the byte array, which "ensures uniqueness" and doesn't include ID creation time
                            //For now, let's just compare the whole ID
                            if (compare_gEntryID == compare_oGlobalID) {
                                compare.Add(outlook[o], google[g]);
                                outlook.Remove(outlook[o]);
                                google.Remove(google[g]);
                                break;
                            }
                        } catch (System.Exception ex) {
                            if (!log.IsFineEnabled()) {
                                try {
                                    log.Info(OutlookCalendar.GetEventSummary(outlook[o]));
                                } catch { }
                            } 
                            if (ex.Message == "An error occurred in the underlying security system. An internal error occurred.") {
                                log.Warn("Item corrupted / inaccessible due to security certificate.");
                                outlook.Remove(outlook[o]);
                            } else {
                                log.Error(ex.Message);
                            }
                        }
                    }
                } else if (Settings.Instance.MergeItems) {
                    //Remove the non-Outlook item so it doesn't get deleted
                    google.Remove(google[g]);
                }
            }
            if (migrated > 0) log.Info(migrated + " migrated.");

            if (Settings.Instance.DisableDelete) {
                google = new List<Event>();
            }
            if (Settings.Instance.SyncDirection == SyncDirection.Bidirectional) {
                //Don't recreate any items that have been deleted in Google
                for (int o = outlook.Count - 1; o >= 0; o--) {
                    if (outlook[o].UserProperties[OutlookCalendar.gEventID] != null)
                        outlook.Remove(outlook[o]);
                }
                //Don't delete any items that aren't yet in Outlook or just created in Outlook during this sync
                for (int g = google.Count - 1; g >= 0; g--) {
                    if (!GetOGCSproperty(google[g], oEntryID) ||
                        DateTime.Parse(google[g].Updated) > Settings.Instance.LastSyncDate)
                        google.Remove(google[g]);
                }
            }
            if (Settings.Instance.CreateCSVFiles) {
                ExportToCSV("Events for deletion in Google", "google_delete.csv", google);
                OutlookCalendar.ExportToCSV("Appointments for creation in Google", "google_create.csv", outlook);
            }
        }
        
        public Boolean CompareRecipientsToAttendees(AppointmentItem ai, Event ev, StringBuilder sb, ref int itemModified) {
            log.Fine("Comparing Recipients");
            //Build a list of Google attendees. Any remaining at the end of the diff must be deleted.
            List<EventAttendee> removeAttendee = new List<EventAttendee>();
            foreach (EventAttendee ea in ev.Attendees ?? Enumerable.Empty<EventAttendee>()) {
                removeAttendee.Add(ea);
            }
            if (ai.Recipients.Count > 1) {
                for (int o = ai.Recipients.Count; o > 0; o--) {
                    bool foundAttendee = false;
                    Recipient recipient = ai.Recipients[o];
                    log.Fine("Comparing Outlook recipient: " + recipient.Name);
                    String recipientSMTP = OutlookCalendar.Instance.IOutlook.GetRecipientEmail(recipient);
                    foreach (EventAttendee attendee in ev.Attendees ?? Enumerable.Empty<EventAttendee>()) {
                        if (attendee.Email != null && (recipientSMTP.ToLower() == attendee.Email.ToLower())) {
                            foundAttendee = true;
                            removeAttendee.Remove(attendee);

                            //Optional attendee
                            bool oOptional = (recipient.Type == (int)OlMeetingRecipientType.olOptional);
                            bool gOptional = (attendee.Optional == null) ? false : (bool)attendee.Optional;
                            String attendeeIdentifier = (attendee.DisplayName == null) ? attendee.Email : attendee.DisplayName;
                            if (MainForm.CompareAttribute("Attendee " + attendeeIdentifier + " - Optional Check",
                                SyncDirection.OutlookToGoogle, gOptional, oOptional, sb, ref itemModified)) {
                                attendee.Optional = oOptional;
                            }

                            //Response
                            switch (recipient.MeetingResponseStatus) {
                                case OlResponseStatus.olResponseNone:
                                    if (MainForm.CompareAttribute("Attendee " + attendeeIdentifier + " - Response Status",
                                        SyncDirection.OutlookToGoogle,
                                        attendee.ResponseStatus, "needsAction", sb, ref itemModified)) {
                                        attendee.ResponseStatus = "needsAction";
                                    }
                                    break;
                                case OlResponseStatus.olResponseAccepted:
                                    if (MainForm.CompareAttribute("Attendee " + attendeeIdentifier + " - Response Status",
                                        SyncDirection.OutlookToGoogle,
                                        attendee.ResponseStatus, "accepted", sb, ref itemModified)) {
                                        attendee.ResponseStatus = "accepted";
                                    }
                                    break;
                                case OlResponseStatus.olResponseDeclined:
                                    if (MainForm.CompareAttribute("Attendee " + attendeeIdentifier + " - Response Status",
                                        SyncDirection.OutlookToGoogle,
                                        attendee.ResponseStatus, "declined", sb, ref itemModified)) {
                                        attendee.ResponseStatus = "declined";
                                    }
                                    break;
                                case OlResponseStatus.olResponseTentative:
                                    if (MainForm.CompareAttribute("Attendee " + attendeeIdentifier + " - Response Status",
                                        SyncDirection.OutlookToGoogle,
                                        attendee.ResponseStatus, "tentative", sb, ref itemModified)) {
                                        attendee.ResponseStatus = "tentative";
                                    }
                                    break;
                            }
                        }
                    } //each attendee

                    if (!foundAttendee) {
                        log.Fine("Attendee added: " + recipient.Name);
                        sb.AppendLine("Attendee added: " + recipient.Name);
                        if (ev.Attendees == null) ev.Attendees = new List<EventAttendee>();
                        ev.Attendees.Add(GoogleCalendar.CreateAttendee(recipient));
                        itemModified++;
                    }
                }
            } //more than just 1 (me) recipients

            foreach (EventAttendee ea in removeAttendee) {
                log.Fine("Attendee removed: " + (ea.DisplayName ?? ea.Email), ea.Email);
                sb.AppendLine("Attendee removed: " + (ea.DisplayName ?? ea.Email));
                ev.Attendees.Remove(ea);
                itemModified++;
            }
            return (itemModified > 0);
        }

        public void GetSetting(string setting) {
            try {
                service.Settings.Get(setting).FetchAsync();
            } catch { }
        }
        public void GetCalendarSettings() {
            if (!Settings.Instance.UseGoogleDefaultReminder) return;
            try {
                CalendarListResource.GetRequest request = service.CalendarList.Get(Settings.Instance.UseGoogleCalendar.Id);
                CalendarListEntry cal = request.Fetch();
                if (cal.DefaultReminders.Count == 0)
                    this.MinDefaultReminder = long.MinValue;
                else
                    this.MinDefaultReminder = cal.DefaultReminders.Where(x => x.Method.Equals("popup")).OrderBy(x => x.Minutes.Value).First().Minutes.Value;
            } catch (System.Exception ex) {
                log.Error("Failed to get calendar settings.");
                log.Error(ex.Message);
            }
        }

        public Boolean UserSubscriptionCheck() {
            List<Event> result = new List<Event>();
            Events request = null;
            String pageToken = null;
            Int16 pageNum = 1;
            
            log.Debug("Retrieving all subscribers from past year.");
            try {
                do {
                    EventsResource.ListRequest lr = service.Events.List("pqeo689qhvpl1g09bcnma1uaoo@group.calendar.google.com");

                    lr.PageToken = pageToken;
                    lr.SingleEvents = true;
                    lr.OrderBy = EventsResource.OrderBy.StartTime;
                    request = lr.Fetch();
                    log.Debug("Page " + pageNum + " received.");

                    if (request != null) {
                        pageToken = request.NextPageToken;
                        pageNum++;
                        if (request.Items != null) result.AddRange(request.Items);
                    }
                } while (pageToken != null);

                if (String.IsNullOrEmpty(Settings.Instance.GaccountEmail)) { //This gets retrieved via the above lr.Fetch()
                    log.Warn("User's Google account username is not present - cannot check if they have subscribed.");
                    return false;
                }
            } catch (System.Exception ex) {
                log.Error("Failed to retrieve subscribers - cannot check if they have subscribed.");
                log.Error(ex.Message);
                return false;
            }

            log.Debug("Searching for subscription for: "+ Settings.Instance.GaccountEmail_masked());
            List<Event> subscriptions = result.Where(x => x.Summary.Equals(Settings.Instance.GaccountEmail)).ToList();
            if (subscriptions.Count == 0) {
                log.Fine("This user has never subscribed.");
                Settings.Instance.Subscribed = DateTime.Parse("01-Jan-2000");
                return false;
            } else {
                Boolean subscribed;
                Event subscription = subscriptions.Last();
                DateTime subscriptionStart = DateTime.Parse(subscription.Start.Date ?? subscription.Start.DateTime).Date;
                Double subscriptionRemaining = (subscriptionStart.AddYears(1) - DateTime.Now.Date).TotalDays;
                if (subscriptionRemaining > 0) {
                    if (subscriptionRemaining > 360)
                        MainForm.Instance.syncNote(MainForm.SyncNotes.RecentSubscription, null);
                    if (subscriptionRemaining < 28)
                        MainForm.Instance.syncNote(MainForm.SyncNotes.SubscriptionPendingExpire, subscriptionStart.AddYears(1));
                    subscribed = true;
                } else {
                    Settings.Instance.Subscribed = DateTime.Parse("01-Jan-2000");
                    if (subscriptionRemaining > -14)
                        MainForm.Instance.syncNote(MainForm.SyncNotes.SubscriptionExpired, subscriptionStart.AddYears(1));
                    subscribed = false;
                }

                DateTime prevSubscriptionStart = Settings.Instance.Subscribed;
                Settings.Instance.Subscribed = subscriptionStart;
                if (prevSubscriptionStart != subscriptionStart) {
                    if (prevSubscriptionStart == DateTime.Parse("01-Jan-2000") || subscriptionStart == DateTime.Parse("01-Jan-2000")) {
                        GoogleCalendar.Instance.ChangeAPIkeys();
                    }
                }
                return subscribed;
            }
        }

        private IAuthorizationState getAuthentication(NativeApplicationClient arg) {
            log.Debug("Authenticating with Google calendar service...");
            // Get the auth URLs:
            IAuthorizationState state = new AuthorizationState(new[] { CalendarService.Scopes.Calendar.GetStringValue(), "https://www.googleapis.com/auth/userinfo.email" });
            state.Callback = new Uri(NativeApplicationClient.OutOfBandCallbackUrl);
            state.RefreshToken = Settings.Instance.RefreshToken;
            Uri authUri = arg.RequestUserAuthorization(state);

            IAuthorizationState result = null;

            if (state.RefreshToken == "") {
                log.Info("No refresh token available - need user authorisation.");

                // Request authorization from the user (by opening a browser window):
                Process.Start(authUri.ToString());

                frmGoogleAuthorizationCode eac = new frmGoogleAuthorizationCode();
                if (eac.ShowDialog() == DialogResult.OK) {
                    if (string.IsNullOrEmpty(eac.authcode))
                        log.Debug("User continued but did not provide a code! This isn't going to work...");
                    else
                        log.Debug("User has provided authentication code.");

                    // Retrieve the access/refresh tokens by using the authorization code:
                    result = arg.ProcessUserAuthorization(eac.authcode, state);

                    //save the refresh token for future use
                    Settings.Instance.RefreshToken = result.RefreshToken;
                    Settings.Instance.Save();
                    log.Info("Refresh and Access token successfully retrieved.");

                    getGaccountEmail(result.AccessToken, true);
                } else {
                    log.Info("User declined to provide authorisation code. Sync will not be able to work.");
                    String noAuth = "Sorry, but this application will not work if you don't give it access to your Google Calendar :(";
                    MessageBox.Show(noAuth, "Authorisation not given", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    throw new System.ApplicationException(noAuth);
                }
            } else {
                arg.RefreshToken(state, null);
                if (string.IsNullOrEmpty(state.AccessToken))
                    log.Error("Failed to retrieve Access token.");
                else
                    log.Debug("Access token refreshed - expires " + ((DateTime)state.AccessTokenExpirationUtc).ToLocalTime().ToString());
                result = state;
                getGaccountEmail(result.AccessToken, false);
            }
            return result;
        }

        #region STATIC FUNCTIONS        
        private void getGaccountEmail(String accessToken, Boolean newRefreshToken) {
            try {
                System.Net.WebClient wc = new System.Net.WebClient();
                wc.Headers.Add("user-agent", "Mozilla/5.0 (Windows NT 6.1; WOW64; rv:37.0) Gecko/20100101 Firefox/37.0");
                String xmlString = wc.DownloadString("https://www.googleapis.com/userinfo/email?alt=xml&access_token=" + accessToken);
                System.Xml.Linq.XDocument xmlDoc = System.Xml.Linq.XDocument.Parse(xmlString);

                if (Settings.Instance.GaccountEmail != xmlDoc.Element("email").Value) {
                    Settings.Instance.GaccountEmail = xmlDoc.Element("email").Value;
                    log.Debug("Updating Google account username: " + Settings.Instance.GaccountEmail_masked());
                    if (!String.IsNullOrEmpty(Settings.Instance.GaccountEmail) && !newRefreshToken) {
                        log.Debug("Looks like you've been tampering with the Google account username value :-O");
                    }
                }
            } catch (System.Net.WebException ex) {
                if (ex.Message.Contains("The remote server returned an error: (403) Forbidden.") || ex.Message == "Insufficient Permission") {
                    String msg = ChangeAPIkeys();
                    throw new System.ApplicationException(msg);
                }
            } catch (System.Exception ex) {
                log.Error("Failed to retrieve Google account username.");
                log.Error(ex.Message);
                log.Debug("Using previously retrieved username: " + Settings.Instance.GaccountEmail_masked());
            }
        }

        //returns the Google Time Format String of a given .Net DateTime value
        //Google Time Format = "2012-08-20T00:00:00+02:00"
        public static string GoogleTimeFrom(DateTime dt) {
            return dt.ToString("yyyy-MM-ddTHH:mm:sszzz", new System.Globalization.CultureInfo("en-US"));
        }
        
        public static string signature(Event ev) {
            String signature = "";
            try {
                if (ev.RecurringEventId != null && ev.Status == "cancelled" && ev.OriginalStartTime != null) {
                    signature += (ev.Summary ?? "[cancelled]");
                    signature += ";" + GoogleTimeFrom(DateTime.Parse(ev.OriginalStartTime.Date ?? ev.OriginalStartTime.DateTime));
                } else {
                    signature += ev.Summary;
                    signature += ";" + GoogleTimeFrom(DateTime.Parse(ev.Start.Date ?? ev.Start.DateTime)) + ";";
                    if (!(ev.EndTimeUnspecified != null && (Boolean)ev.EndTimeUnspecified)) {
                        signature += GoogleTimeFrom(DateTime.Parse(ev.End.Date ?? ev.End.DateTime));
                    }
                }
            } catch {
                log.Warn("Failed to create signature: " + signature);
                log.Warn("This Event cannot be synced.");
                try { log.Warn("  ev.Summary: " + ev.Summary); } catch { }
                try { log.Warn("  ev.Start: " + (ev.Start == null ? "null!" : ev.Start.Date ?? ev.Start.DateTime)); } catch { }
                try { log.Warn("  ev.End: " + (ev.End == null ? "null!" : ev.End.Date ?? ev.End.DateTime)); } catch { }
                try { log.Warn("  ev.Status: " + ev.Status ?? "null!"); } catch { }
                return "";
            }
            return signature.Trim();
        }

        public static void ExportToCSV(String action, String filename, List<Event> events) {
            log.Debug(action);

            TextWriter tw;
            try {
                tw = new StreamWriter(Path.Combine(Program.UserFilePath, filename));
            } catch (System.Exception ex) {
                MainForm.Instance.Logboxout("Failed to create CSV file '"+ filename +"'.");
                log.Error("Error opening file '"+ filename +"' for writing.");
                log.Error(ex.Message);
                return;
            } 
            try {
                String CSVheader = "Start Time,Finish Time,Subject,Location,Description,Privacy,FreeBusy,";
                CSVheader += "Required Attendees,Optional Attendees,Reminder Set,Reminder Minutes,Google ID,Outlook ID";

                tw.WriteLine(CSVheader);

                foreach (Event ev in events) {
                    try {
                        tw.WriteLine(exportToCSV(ev));
                    } catch {
                        MainForm.Instance.Logboxout("Failed to output following Google event to CSV:-");
                        MainForm.Instance.Logboxout(GetEventSummary(ev));
                    }
                }
            } catch {
                MainForm.Instance.Logboxout("Failed to output Google events to CSV.");
            } finally {
                if (tw != null) tw.Close();
            }
            log.Debug("CSV export done.");
        }
        private static String exportToCSV(Event ev) {
            System.Text.StringBuilder csv = new System.Text.StringBuilder();

            csv.Append(ev.Start.Date ?? ev.Start.DateTime + ",");
            csv.Append(ev.End.Date ?? ev.End.DateTime + ",");
            csv.Append("\"" + ev.Summary + "\",");
            
            if (ev.Location == null) csv.Append(",");
            else csv.Append("\"" + ev.Location + "\",");
            
            if (ev.Description == null) csv.Append(",");
            else {
                String csvDescription = ev.Description.Replace("\"", "");
                csvDescription = csvDescription.Replace("\r\n", " ");
                csv.Append("\"" + csvDescription.Substring(0, System.Math.Min(csvDescription.Length, 100)) + "\",");
            }
            csv.Append("\"" + ev.Visibility + "\",");
            csv.Append("\"" + ev.Transparency + "\",");
            System.Text.StringBuilder required = new System.Text.StringBuilder();
            System.Text.StringBuilder optional = new System.Text.StringBuilder();
            if (ev.Attendees != null) {
                foreach (EventAttendee ea in ev.Attendees) {
                    if (ea.Optional != null && (bool)ea.Optional) { optional.Append(ea.DisplayName + ";"); }
                    else { required.Append(ea.DisplayName + ";"); }
                }
                csv.Append("\"" + required + "\",");
                csv.Append("\"" + optional + "\",");
            } else
                csv.Append(",,");

            bool foundReminder = false;
            if (ev.Reminders != null && ev.Reminders.Overrides != null) {
                foreach (EventReminder er in ev.Reminders.Overrides) {
                    if (er.Method == "popup") {
                        csv.Append("true," + er.Minutes +",");
                        foundReminder = true;
                    }
                    break;
                }
            }
            if (!foundReminder) csv.Append(",,");

            csv.Append(ev.Id + ",");
            String gEntryID;
            if (GetOGCSproperty(ev, oEntryID, out gEntryID)) {
                csv.Append(gEntryID);
            }

            return csv.ToString();
        }

        public static string GetEventSummary(Event ev) {
            String eventSummary = "";
            if (ev.Start.DateTime != null) {
                DateTime gDate = DateTime.Parse(ev.Start.DateTime);
                eventSummary += gDate.ToShortDateString() + " " + gDate.ToShortTimeString();
            } else
                eventSummary += DateTime.Parse(ev.Start.Date).ToShortDateString();
            if ((ev.Recurrence != null && ev.RecurringEventId == null) || ev.RecurringEventId != null)
                eventSummary += " (R)";
            eventSummary += " => \"" + ev.Summary + "\"";
            return eventSummary;
        }

        public static EventAttendee CreateAttendee(Recipient recipient) {
            EventAttendee ea = new EventAttendee();
            log.Fine("Creating attendee " + recipient.Name);
            ea.DisplayName = recipient.Name;
            ea.Email = OutlookCalendar.Instance.IOutlook.GetRecipientEmail(recipient);
            ea.Optional = (recipient.Type == (int)OlMeetingRecipientType.olOptional);
            //Readonly: ea.Organizer = (ai.Organizer == recipient.Name);
            switch (recipient.MeetingResponseStatus) {
                case OlResponseStatus.olResponseNone: ea.ResponseStatus = "needsAction"; break;
                case OlResponseStatus.olResponseAccepted: ea.ResponseStatus = "accepted"; break;
                case OlResponseStatus.olResponseDeclined: ea.ResponseStatus = "declined"; break;
                case OlResponseStatus.olResponseTentative: ea.ResponseStatus = "tentative"; break;
            }
            return ea;
        }

        private static apiException handleAPIlimits(System.Exception ex, Event ev) {
            //https://developers.google.com/analytics/devguides/reporting/core/v3/coreErrors

            if (Settings.Instance.AddAttendees && ex.Message.Contains("Calendar usage limits exceeded. [403]") && ev != null) {
                //"Google.Apis.Requests.RequestError\r\nCalendar usage limits exceeded. [403]\r\nErrors [\r\n\tMessage[Calendar usage limits exceeded.] Location[ - ] Reason[quotaExceeded] Domain[usageLimits]\r\n]\r\n"
                //This happens because too many attendees have been added in a short period of time.
                //See https://support.google.com/a/answer/2905486?hl=en-uk&hlrm=en

                MainForm.Instance.Logboxout("ALERT: You have added enough meeting attendees to have reached the Google API limit.");
                MainForm.Instance.Logboxout("Don't worry, this only lasts for an hour or two, but until then attendees will not be synced.");
                
                APIlimitReached_attendee = true;
                Settings.Instance.APIlimit_inEffect = true;
                Settings.Instance.APIlimit_lastHit = DateTime.Now;

                ev.Attendees = new List<EventAttendee>();
                return apiException.justContinue;

            } else if (ex.Message.Contains("Rate Limit Exceeded")) {
                return apiException.backoffThenRetry;

            } else if (ex.Message.Contains("Daily Limit Exceeded [403]")) {
                log.Warn(ex.Message);
                log.Warn("Google's free Calendar quota has been exhausted! New quota comes into effect 08:00 GMT");
                MainForm.Instance.syncNote(MainForm.SyncNotes.QuotaExhaustedInfo, null);
                return apiException.freeAPIexhausted;

            } else if (ex.Message.Equals("The remote server returned an error: (401) Unauthorized.")) {
                log.Warn(ex.Message);
                log.Debug("This error seems to be a new transient issue, so treating it with exponential backoff...");
                return apiException.backoffThenRetry;

            } else {
                return apiException.throwException;
            }
        }

        #region OGCS event properties
        public static void AddOutlookID(ref Event ev, AppointmentItem ai) {
            //Add the Outlook appointment ID into Google event.
            //This will make comparison more efficient and set the scene for 2-way sync.

            addOGCSproperty(ref ev, oEntryID, OutlookCalendar.Instance.IOutlook.GetGlobalApptID(ai));
        }

        private static void addOGCSproperty(ref Event ev, String key, String value) {
            if (ev.ExtendedProperties == null) ev.ExtendedProperties = new Event.ExtendedPropertiesData();
            if (ev.ExtendedProperties.Private == null) ev.ExtendedProperties.Private = new Event.ExtendedPropertiesData.PrivateData();
            if (ev.ExtendedProperties.Private.ContainsKey(key))
                ev.ExtendedProperties.Private[key] = value;
            else
                ev.ExtendedProperties.Private.Add(key, value);
        }
        private static void addOGCSproperty(ref Event ev, String key, DateTime value) {
            addOGCSproperty(ref ev, key, value.ToString("yyyyMMddHHmmss", System.Globalization.CultureInfo.InvariantCulture));
        }

        public static Boolean GetOGCSproperty(Event ev, String key) {
            String throwAway;
            return GoogleCalendar.GetOGCSproperty(ev, key, out throwAway);
        }
        public static Boolean GetOGCSproperty(Event ev, String key, out String value) {
            if (ev.ExtendedProperties != null &&
                ev.ExtendedProperties.Private != null &&
                ev.ExtendedProperties.Private.ContainsKey(key)) {
                value = ev.ExtendedProperties.Private[key];
                return true;
            } else {
                value = null;
                return false;
            }
        }

        private static void removeOGCSproperty(ref Event ev, String key) {
            if (GetOGCSproperty(ev, key))
                ev.ExtendedProperties.Private.Remove(Program.APIlimitHit);
        }

        public static DateTime GetOGCSlastModified(Event ev) {
            String lastModded = null;
            if (GetOGCSproperty(ev, Program.OGCSmodified, out lastModded)) {
                try {
                    return DateTime.ParseExact(lastModded, "yyyyMMddHHmmss", System.Globalization.CultureInfo.InvariantCulture);
                } catch (System.FormatException) {
                    //Bugfix <= v2.2, 
                    log.Fine("Date wasn't stored as invariant culture.");
                    DateTime retDate;
                    if (DateTime.TryParse(lastModded, out retDate)) {
                        log.Fine("Fall back to current culture successful.");
                        return retDate;
                    } else {
                        log.Debug("Fall back to current culture for date failed. Last resort: setting to a month ago.");
                        return DateTime.Now.AddMonths(-1);
                    }
                }
            } else {
                return new DateTime();
            }
        }
        private static void setOGCSlastModified(ref Event ev) {
            addOGCSproperty(ref ev, Program.OGCSmodified, DateTime.Now);
        }
        #endregion
        #endregion
    }
}
