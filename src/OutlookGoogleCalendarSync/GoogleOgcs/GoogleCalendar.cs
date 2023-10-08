using Google.Apis.Calendar.v3;
using Google.Apis.Calendar.v3.Data;
using log4net;
using Microsoft.Office.Interop.Outlook;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace OutlookGoogleCalendarSync.GoogleOgcs {
    /// <summary>
    /// Description of GoogleOgcs.Calendar.
    /// </summary>
    public class Calendar {
        private static readonly ILog log = LogManager.GetLogger(typeof(Calendar));

        private static Calendar instance;
        public static Boolean IsInstanceNull { get { return instance == null; } }
        public static Calendar Instance {
            get {
                if (instance == null) {
                    instance = new GoogleOgcs.Calendar {
                        Authenticator = new GoogleOgcs.Authenticator()
                    };
                    instance.Authenticator.GetAuthenticated();
                    if (instance.Authenticator.Authenticated)
                        instance.Authenticator.OgcsUserStatus();
                    else {
                        instance = null;
                        if (Forms.Main.Instance.Console.DocumentText.Contains("Authorisation to allow OGCS to manage your Google calendar was cancelled."))
                            throw new OperationCanceledException();
                        else
                            throw new ApplicationException("Google handshake failed.");
                    }
                }
                return instance;
            }
        }
        public Calendar() { }
        private Boolean openedIssue1593 = false;
        public GoogleOgcs.Authenticator Authenticator;

        private GoogleOgcs.EventColour colourPalette;

        public static Boolean IsColourPaletteNull { get { return instance?.colourPalette == null; } }
        public GoogleOgcs.EventColour ColourPalette {
            get {
                if (colourPalette == null)
                    colourPalette = new EventColour();
                if (Authenticator.Authenticated && !colourPalette.IsCached())
                    colourPalette.Get();
                return colourPalette;
            }
        }

        private CalendarService service;
        public CalendarService Service {
            get {
                if (service == null) {
                    log.Debug("Google service not yet instantiated.");
                    Authenticator = new GoogleOgcs.Authenticator();
                    Authenticator.GetAuthenticated();
                    if (Authenticator.Authenticated)
                        Authenticator.OgcsUserStatus();
                    else {
                        service = null;
                        throw new ApplicationException("Google handshake failed.");
                    }
                }
                return service;
            }
            set { service = value; }
        }
        public static Boolean APIlimitReached_attendee = false;
        public const int BackoffLimit = 5;
        public enum ApiException {
            justContinue,
            backoffThenRetry,
            freeAPIexhausted,
            throwException
        }
        private static Random random = new Random();
        public int MinDefaultReminder = int.MinValue;
        public Int16 UTCoffset { get; internal set; }
        public String SubscriptionInvite {
            get {
                String invite = "Google's free calendar quota ran out! You'll need to wait for fresh quota";
                if (string.IsNullOrEmpty(Settings.Instance.GaccountEmail))
                    invite += ".";
                else {
                    String url = "https://www.paypal.com/cgi-bin/webscr?cmd=_s-xclick&hosted_button_id=E595EQ7SNDBHA&item_name=" + "OGCS Premium for " + Settings.Instance.GaccountEmail;
                    invite += " or <a href='" + url + "' target='_blank'>get guaranteed quota</a> for just £1/month.";
                }
                return invite;
            }
        }

        public EphemeralProperties EphemeralProperties = new EphemeralProperties();

        private List<GoogleCalendarListEntry> calendarList = new List<GoogleCalendarListEntry>();
        public List<GoogleCalendarListEntry> CalendarList {
            get { return calendarList; }
            protected set { calendarList = value; }
        }

        public void GetCalendars() {
            CalendarList request = null;
            String pageToken = null;
            List<GoogleCalendarListEntry> result = new List<GoogleCalendarListEntry>();
            int backoff = 0;

            do {
                while (backoff < BackoffLimit) {
                    try {
                        CalendarListResource.ListRequest lr = Service.CalendarList.List();
                        lr.PageToken = pageToken;
                        lr.ShowHidden = true;
                        request = lr.Execute();
                        break;
                    } catch (Google.GoogleApiException ex) {
                        switch (HandleAPIlimits(ref ex, null)) {
                            case ApiException.throwException: throw;
                            case ApiException.freeAPIexhausted:
                                OGCSexception.LogAsFail(ref ex);
                                OGCSexception.Analyse(ex);
                                System.ApplicationException aex = new System.ApplicationException(SubscriptionInvite, ex);
                                OGCSexception.LogAsFail(ref aex);
                                throw aex;
                            case ApiException.backoffThenRetry:
                                backoff++;
                                if (backoff == BackoffLimit) {
                                    log.Error("API limit backoff was not successful. Retrieve calendar list failed.");
                                    throw;
                                } else {
                                    int backoffDelay = (int)Math.Pow(2, backoff);
                                    log.Warn("API rate limit reached. Backing off " + backoffDelay + "sec before retry.");
                                    System.Threading.Thread.Sleep(backoffDelay * 1000);
                                }
                                break;
                        }
                    }
                }

                if (request != null) {
                    pageToken = request.NextPageToken;
                    foreach (CalendarListEntry cle in request.Items) {
                        result.Add(new GoogleCalendarListEntry(cle));
                    }
                } else {
                    log.Error("Handshaking with the Google calendar service failed.");
                }
            } while (pageToken != null);

            this.CalendarList = result;
        }

        public List<Event> GetCalendarEntriesInRecurrence(String recurringEventId) {
            List<Event> result = new List<Event>();
            Events request = null;
            String pageToken = null;
            Int16 pageNum = 1;

            try {
                log.Debug("Retrieving all recurring event instances from Google for " + recurringEventId);
                do {
                    EventsResource.InstancesRequest ir = Service.Events.Instances(Sync.Engine.Calendar.Instance.Profile.UseGoogleCalendar.Id, recurringEventId);
                    ir.ShowDeleted = true;
                    ir.PageToken = pageToken;
                    int backoff = 0;
                    while (backoff < BackoffLimit) {
                        try {
                            request = ir.Execute();
                            log.Debug("Page " + pageNum + " received.");
                            break;
                        } catch (Google.GoogleApiException ex) {
                            switch (HandleAPIlimits(ref ex, null)) {
                                case ApiException.throwException: throw;
                                case ApiException.freeAPIexhausted:
                                    OGCSexception.LogAsFail(ref ex);
                                    OGCSexception.Analyse(ex);
                                    System.ApplicationException aex = new System.ApplicationException(SubscriptionInvite, ex);
                                    OGCSexception.LogAsFail(ref aex);
                                    throw aex;
                                case ApiException.backoffThenRetry:
                                    backoff++;
                                    if (backoff == BackoffLimit) {
                                        log.Error("API limit backoff was not successful. Paginated retrieve failed.");
                                        throw;
                                    } else {
                                        int backoffDelay = (int)Math.Pow(2, backoff);
                                        log.Warn("API rate limit reached. Backing off " + backoffDelay + "sec before retry.");
                                        System.Threading.Thread.Sleep(backoffDelay * 1000);
                                    }
                                    break;
                            }
                        }
                    }

                    if (request != null) {
                        pageToken = request.NextPageToken;
                        pageNum++;
                        if (request.Items != null) result.AddRange(request.Items);
                    }
                } while (pageToken != null);
                log.Fine(request.Items.Count + " recurring event instances found.");
                return result;

            } catch (System.Exception ex) {
                Forms.Main.Instance.Console.UpdateWithError("Failed to retrieve recurring events.", OGCSexception.LogAsFail(ex));
                OGCSexception.Analyse("recurringEventId: " + recurringEventId, ex);
                return null;
            }
        }

        public Event GetCalendarEntry(String eventId) {
            Event request = null;

            try {
                log.Debug("Retrieving specific Event with ID " + eventId);
                EventsResource.GetRequest gr = Service.Events.Get(Sync.Engine.Calendar.Instance.Profile.UseGoogleCalendar.Id, eventId);
                int backoff = 0;
                while (backoff < BackoffLimit) {
                    try {
                        request = gr.Execute();
                        break;
                    } catch (Google.GoogleApiException ex) {
                        if (ex.Error.Code == 404) { //Not found
                            log.Fail("Could not find Google Event with specified ID " + eventId);
                            return null;
                        }
                        switch (HandleAPIlimits(ref ex, null)) {
                            case ApiException.throwException: throw;
                            case ApiException.freeAPIexhausted:
                                OGCSexception.LogAsFail(ref ex);
                                OGCSexception.Analyse(ex);
                                System.ApplicationException aex = new System.ApplicationException(SubscriptionInvite, ex);
                                OGCSexception.LogAsFail(ref aex);
                                throw aex;
                            case ApiException.backoffThenRetry:
                                backoff++;
                                if (backoff == BackoffLimit) {
                                    log.Error("API limit backoff was not successful. Retrieve failed.");
                                    throw;
                                } else {
                                    int backoffDelay = (int)Math.Pow(2, backoff);
                                    log.Warn("API rate limit reached. Backing off " + backoffDelay + "sec before retry.");
                                    System.Threading.Thread.Sleep(backoffDelay * 1000);
                                }
                                break;
                        }
                    }
                }

                if (request != null)
                    return request;
                else
                    throw new System.Exception("Returned null");
            } catch (System.Exception ex) {
                if (ex is ApplicationException) throw;
                Forms.Main.Instance.Console.Update("Failed to retrieve Google event", Console.Markup.error);
                return null;
            }
        }

        public List<Event> GetCalendarEntriesInRange() {
            SettingsStore.Calendar profile = Settings.Profile.InPlay();
            return GetCalendarEntriesInRange(profile.SyncStart, profile.SyncEnd);
        }

        public List<Event> GetCalendarEntriesInRange(DateTime from, DateTime to) {
            List<Event> result = new List<Event>();
            Events request = null;
            String pageToken = null;
            Int16 pageNum = 1;

            SettingsStore.Calendar profile = Settings.Profile.InPlay();

            log.Debug("Retrieving all events from Google: " + from.ToShortDateString() + " -> " + to.ToShortDateString());
            do {
                EventsResource.ListRequest lr = Service.Events.List(profile.UseGoogleCalendar.Id);

                lr.TimeMin = from;
                lr.TimeMax = to;
                lr.PageToken = pageToken;
                lr.ShowDeleted = false;
                lr.SingleEvents = false;

                int backoff = 0;
                while (backoff < BackoffLimit) {
                    try {
                        request = lr.Execute();
                        log.Debug("Page " + pageNum + " received.");
                        break;
                    } catch (Google.GoogleApiException ex) {
                        switch (HandleAPIlimits(ref ex, null)) {
                            case ApiException.throwException: throw;
                            case ApiException.freeAPIexhausted:
                                OGCSexception.LogAsFail(ref ex);
                                OGCSexception.Analyse(ex);
                                System.ApplicationException aex = new System.ApplicationException(SubscriptionInvite, ex);
                                OGCSexception.LogAsFail(ref aex);
                                throw aex;
                            case ApiException.backoffThenRetry:
                                backoff++;
                                if (backoff == BackoffLimit) {
                                    log.Error("API limit backoff was not successful. Retrieve failed.");
                                    aex = new System.ApplicationException(SubscriptionInvite, ex);
                                    OGCSexception.LogAsFail(ref aex);
                                    throw aex;
                                } else {
                                    int backoffDelay = (int)Math.Pow(2, backoff);
                                    log.Warn("API rate limit reached. Backing off " + backoffDelay + "sec before retry.");
                                    System.Threading.Thread.Sleep(backoffDelay * 1000);
                                }
                                break;
                        }
                    }
                }

                if (request != null) {
                    pageToken = request.NextPageToken;
                    pageNum++;
                    if (request.Items != null) result.AddRange(request.Items);
                }
            } while (pageToken != null);

            //Remove cancelled non-recurring Events - don't know how these exist, but some users have them!
            List<Event> cancelled = result.Where(ev =>
                ev.Status == "cancelled" && string.IsNullOrEmpty(ev.RecurringEventId) &&
                ev.Start == null && ev.End == null && string.IsNullOrEmpty(ev.Summary)).ToList();
            if (cancelled.Count > 0) {
                log.Debug(cancelled.Count + " Google Events are cancelled and will be excluded.");
                result = result.Except(cancelled).ToList();
            }

            List<Event> endsOnSyncStart = result.Where(ev => (ev.End != null && ev.End.SafeDateTime() == from)).ToList();
            if (endsOnSyncStart.Count > 0) {
                log.Debug(endsOnSyncStart.Count + " Google Events end at midnight of the sync start date window.");
                result = result.Except(endsOnSyncStart).ToList();
            }

            if (profile.SyncDirection.Id != Sync.Direction.OutlookToGoogle.Id) { //Sync direction means G->O will delete previously synced all-days
                if (profile.ExcludeFree) {
                    List<Event> availability = result.Where(ev => profile.ExcludeFree && ev.Transparency == "transparent").ToList();
                    if (availability.Count > 0) {
                        log.Debug(availability.Count + " Google Free items excluded.");
                        result = result.Except(availability).ToList();
                    }
                }
                if (profile.ExcludeAllDays) {
                    List<Event> allDays = result.Where(ev => ev.AllDayEvent(true) && (profile.ExcludeFreeAllDays ? ev.Transparency == "transparent" : true)).ToList();
                    if (allDays.Count > 0) {
                        log.Debug(allDays.Count + " Google all-day items excluded.");
                        result = result.Except(allDays).ToList();
                    }
                }
                if (profile.ExcludePrivate) {
                    List<Event> privacy = result.Where(ev => profile.ExcludePrivate && ev.Visibility == "private").ToList();
                    if (privacy.Count > 0) {
                        log.Debug(privacy.Count + " Google Private items excluded.");
                        result = result.Except(privacy).ToList();
                    }
                }
            }

            if (profile.ExcludeDeclinedInvites) {
                List<Event> declined = result.Where(ev => string.IsNullOrEmpty(ev.RecurringEventId) && ev.Attendees != null && ev.Attendees.Count(a => a.Self == true && a.ResponseStatus == "declined") == 1).ToList();
                if (declined.Count > 0) {
                    log.Debug(declined.Count + " Google Event invites have been declined and will be excluded.");
                    result = result.Except(declined).ToList();
                }
            }

            if ((IsDefaultCalendar() ?? true) && profile.ExcludeGoals) {
                List<Event> goals = result.Where(ev =>
                    !string.IsNullOrEmpty(ev.Description) && ev.Description.Contains("This event was added from Goals in Google Calendar.") &&
                    ev.Organizer != null && ev.Organizer.Email == "unknownorganizer@calendar.google.com" && ev.Organizer.DisplayName == "Google Calendar").ToList();
                if (goals.Count > 0) {
                    log.Debug(goals.Count + " Google Events are Goals and will be excluded.");
                    result = result.Except(goals).ToList();
                }
            }

            return result;
        }

        #region Create
        public void CreateCalendarEntries(List<AppointmentItem> appointments) {
            foreach (AppointmentItem ai in appointments) {
                if (Sync.Engine.Instance.CancellationPending) return;

                Event newEvent = new Event();
                try {
                    newEvent = createCalendarEntry(ai);
                } catch (System.Exception ex) {
                    if (ex is ApplicationException) {
                        Forms.Main.Instance.Console.Update(OutlookOgcs.Calendar.GetEventSummary(ai, true) + "Event creation skipped.", Console.Markup.warning);
                        continue;
                    } else {
                        Forms.Main.Instance.Console.UpdateWithError(OutlookOgcs.Calendar.GetEventSummary(ai, true) + "Event creation failed.", ex);
                        OGCSexception.Analyse(ex, true);
                        if (OgcsMessageBox.Show("Google event creation failed. Continue with synchronisation?", "Sync item failed", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                            continue;
                        else
                            throw new UserCancelledSyncException("User chose not to continue sync.");
                    }
                }

                Event createdEvent = new Event();
                try {
                    createdEvent = createCalendarEntry_save(newEvent, ai);
                } catch (System.Exception ex) {
                    if (ex.Message.Contains("You need to have writer access to this calendar")) {
                        Forms.Main.Instance.Console.Update("The Google calendar being synced with must not be read-only.<br/>Cannot continue sync.", Console.Markup.fail, newLine: false);
                        return;
                    }
                    Forms.Main.Instance.Console.UpdateWithError(OutlookOgcs.Calendar.GetEventSummary(ai, true) + "New event failed to save.", ex);
                    OGCSexception.Analyse(ex, true);
                    if (OgcsMessageBox.Show("New Google event failed to save. Continue with synchronisation?", "Sync item failed", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                        continue;
                    else
                        throw new UserCancelledSyncException("User chose not to continue sync.");
                }
                if (ai.IsRecurring && Recurrence.HasExceptions(ai) && createdEvent != null) {
                    Forms.Main.Instance.Console.Update("This is a recurring item with some exceptions:-", verbose: true);
                    Recurrence.CreateGoogleExceptions(ai, createdEvent.Id);
                    Forms.Main.Instance.Console.Update("Recurring exceptions completed.", verbose: true);
                }
            }
        }

        private Event createCalendarEntry(AppointmentItem ai) {
            SettingsStore.Calendar profile = Sync.Engine.Calendar.Instance.Profile;

            string itemSummary = OutlookOgcs.Calendar.GetEventSummary(ai);
            log.Debug("Processing >> " + itemSummary);
            Forms.Main.Instance.Console.Update(itemSummary, Console.Markup.calendar, verbose: true);

            Event ev = new Event();

            ev.Recurrence = Recurrence.Instance.BuildGooglePattern(ai, ev);
            ev.Start = new EventDateTime();
            ev.End = new EventDateTime();

            if (ai.AllDayEvent) {
                ev.Start.Date = ai.Start.ToString("yyyy-MM-dd");
                ev.End.Date = ai.End.ToString("yyyy-MM-dd");
            } else {
                ev.Start.DateTime = ai.Start;
                ev.End.DateTime = ai.End;
            }
            ev = OutlookOgcs.Calendar.Instance.IOutlook.IANAtimezone_set(ev, ai);

            ev.Summary = Obfuscate.ApplyRegex(ai.Subject, null, Sync.Direction.OutlookToGoogle);
            if (profile.AddDescription) {
                try {
                    ev.Description = ai.Body;
                } catch (System.Exception ex) {
                    if (OGCSexception.GetErrorCode(ex) == "0x80004004") {
                        Forms.Main.Instance.Console.Update("You do not have the rights to programmatically access Outlook appointment descriptions.<br/>" +
                            "It may be best to stop syncing the Description attribute.", Console.Markup.warning);
                    } else throw;
                }
            }
            if (profile.AddLocation)
                ev.Location = ai.Location;
            ev.Visibility = getPrivacy(ai.Sensitivity, null);
            ev.Transparency = getAvailability(ai.BusyStatus, null);
            ev.ColorId = getColour(ai.Categories, null)?.Id ?? EventColour.Palette.NullPalette.Id;

            ev.Attendees = new List<Google.Apis.Calendar.v3.Data.EventAttendee>();
            if (profile.AddAttendees && ai.Recipients.Count > 1 && !APIlimitReached_attendee) { //Don't add attendees if there's only 1 (me)
                if (ai.Recipients.Count > profile.MaxAttendees) {
                    log.Warn("This Outlook appointment has " + ai.Recipients.Count + " attendees, more than the user configured maximum.");
                    if (ai.Recipients.Count >= 200) {
                        Forms.Main.Instance.Console.Update("Attendees will not be synced for this meeting as it has " +
                            "more than 200, which Google does not allow.", Console.Markup.warning);
                    }
                } else {
                    foreach (Microsoft.Office.Interop.Outlook.Recipient recipient in ai.Recipients) {
                        Google.Apis.Calendar.v3.Data.EventAttendee ea = GoogleOgcs.Calendar.CreateAttendee(recipient, ai.Organizer == recipient.Name);
                        ev.Attendees.Add(ea);
                    }
                }
            }

            //Reminder alert
            ev.Reminders = new Event.RemindersData();
            if (profile.AddReminders) {
                if (OutlookOgcs.Calendar.Instance.IsOKtoSyncReminder(ai)) {
                    if (ai.ReminderSet) {
                        ev.Reminders.UseDefault = false;
                        EventReminder reminder = new EventReminder {
                            Method = "popup",
                            Minutes = ai.ReminderMinutesBeforeStart
                        };
                        ev.Reminders.Overrides = new List<EventReminder> { reminder };
                    } else {
                        ev.Reminders.UseDefault = profile.UseGoogleDefaultReminder;
                    }
                } else {
                    ev.Reminders.UseDefault = false;
                }
            } else
                ev.Reminders.UseDefault = profile.UseGoogleDefaultReminder;

            //Add the Outlook appointment ID into Google event
            CustomProperty.AddOutlookIDs(ref ev, ai);

            return ev;
        }

        private Event createCalendarEntry_save(Event ev, AppointmentItem ai) {
            SettingsStore.Calendar profile = Sync.Engine.Calendar.Instance.Profile;

            if (profile.SyncDirection.Id == Sync.Direction.Bidirectional.Id) {
                log.Debug("Saving timestamp when OGCS created event.");
                CustomProperty.SetOGCSlastModified(ref ev);
            }
            if (Settings.Instance.APIlimit_inEffect) {
                CustomProperty.Add(ref ev, CustomProperty.MetadataId.apiLimitHit, "True");
            }

            Event createdEvent = new Event();
            int backoff = 0;
            while (backoff < BackoffLimit) {
                try {
                    EventsResource.InsertRequest request = Service.Events.Insert(ev, profile.UseGoogleCalendar.Id);
                    request.SendUpdates = EventsResource.InsertRequest.SendUpdatesEnum.None;
                    createdEvent = request.Execute();
                    if (profile.AddAttendees && Settings.Instance.APIlimit_inEffect) {
                        log.Info("API limit for attendee sync lifted :-)");
                        Settings.Instance.APIlimit_inEffect = false;
                    }
                    break;
                } catch (Google.GoogleApiException ex) {
                    switch (HandleAPIlimits(ref ex, ev)) {
                        case ApiException.throwException: throw;
                        case ApiException.freeAPIexhausted:
                            OGCSexception.LogAsFail(ref ex);
                            OGCSexception.Analyse(ex);
                            System.ApplicationException aex = new System.ApplicationException(SubscriptionInvite, ex);
                            OGCSexception.LogAsFail(ref aex);
                            throw aex;
                        case ApiException.justContinue: break;
                        case ApiException.backoffThenRetry:
                            backoff++;
                            if (backoff == BackoffLimit) {
                                log.Error("API limit backoff was not successful. Save failed.");
                                throw;
                            } else {
                                int backoffDelay = (int)Math.Pow(2, backoff);
                                log.Warn("API rate limit reached. Backing off " + backoffDelay + "sec before retry.");
                                System.Threading.Thread.Sleep(backoffDelay * 1000);
                            }
                            break;
                    }
                }
            }

            if (profile.SyncDirection.Id == Sync.Direction.Bidirectional.Id || OutlookOgcs.CustomProperty.ExistsAny(ai)) {
                log.Debug("Storing the Google event IDs in Outlook appointment.");
                OutlookOgcs.CustomProperty.AddGoogleIDs(ref ai, createdEvent);
                OutlookOgcs.CustomProperty.SetOGCSlastModified(ref ai);
                ai.Save();
            }
            //DOS ourself by triggering API limit
            //for (int i = 1; i <= 100; i++) {
            //    Forms.Main.Instance.Console.Update("Add #" + i, verbose: true);
            //    Event result = service.Events.Insert(ev, Settings.Instance.UseGoogleCalendar.Id).Execute();
            //    System.Threading.Thread.Sleep(300);
            //    GoogleOgcs.Calendar.Instance.deleteCalendarEntry_save(result);
            //    System.Threading.Thread.Sleep(300);
            //}
            return createdEvent;
        }
        #endregion

        #region Update
        public void UpdateCalendarEntries(Dictionary<AppointmentItem, Event> entriesToBeCompared, ref int entriesUpdated) {
            entriesUpdated = 0;
            for (int i = 0; i < entriesToBeCompared.Count; i++) {
                if (Sync.Engine.Instance.CancellationPending) return;

                KeyValuePair<AppointmentItem, Event> compare = entriesToBeCompared.ElementAt(i);
                int itemModified = 0;
                Boolean eventExceptionCacheDirty = false;
                Event ev = new Event();
                try {
                    ev = UpdateCalendarEntry(compare.Key, compare.Value, ref itemModified);
                } catch (System.Exception ex) {
                    if (ex is ApplicationException) {
                        Forms.Main.Instance.Console.Update(OutlookOgcs.Calendar.GetEventSummary(compare.Key, true) + "Event update skipped.", Console.Markup.warning);
                        continue;
                    } else {
                        Forms.Main.Instance.Console.UpdateWithError(OutlookOgcs.Calendar.GetEventSummary(compare.Key, true) + "Event update failed.", ex);
                        if (ex is System.Runtime.InteropServices.COMException) throw;
                        OGCSexception.Analyse(ex, true);
                        if (OgcsMessageBox.Show("Google event update failed. Continue with synchronisation?", "Sync item failed", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                            continue;
                        else
                            throw new UserCancelledSyncException("User chose not to continue sync.");
                    }
                }

                if (itemModified > 0) {
                    try {
                        UpdateCalendarEntry_save(ref ev);
                        entriesUpdated++;
                        eventExceptionCacheDirty = true;
                    } catch (System.Exception ex) {
                        Forms.Main.Instance.Console.UpdateWithError(OutlookOgcs.Calendar.GetEventSummary(compare.Key, true) + "Updated event failed to save.", ex);
                        OGCSexception.Analyse(ex, true);
                        if (OgcsMessageBox.Show("Updated Google event failed to save. Continue with synchronisation?", "Sync item failed", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                            continue;
                        else
                            throw new UserCancelledSyncException("User chose not to continue sync.");
                    }

                }

                //Have to do this *before* any dummy update, else all the exceptions inherit the updated timestamp of the parent recurring event
                Recurrence.UpdateGoogleExceptions(compare.Key, ev ?? compare.Value, eventExceptionCacheDirty);

                if (itemModified == 0) {
                    if (ev == null) {
                        if (compare.Value.Updated < compare.Key.LastModificationTime || CustomProperty.Exists(compare.Value, CustomProperty.MetadataId.forceSave))
                            ev = compare.Value;
                        else
                            continue;
                    }
                    log.Debug("Doing a dummy update in order to update the last modified date of " +
                        (ev.RecurringEventId == null && ev.Recurrence != null ? "recurring master event" : "single instance"));
                    CustomProperty.SetOGCSlastModified(ref ev);
                    try {
                        UpdateCalendarEntry_save(ref ev);
                        entriesToBeCompared[compare.Key] = ev;
                    } catch (System.Exception ex) {
                        Forms.Main.Instance.Console.UpdateWithError(OutlookOgcs.Calendar.GetEventSummary(compare.Key, true) + "Updated event failed to save.", ex);
                        OGCSexception.Analyse(ex, true);
                        if (OgcsMessageBox.Show("Updated Google event failed to save. Continue with synchronisation?", "Sync item failed", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                            continue;
                        else
                            throw new UserCancelledSyncException("User chose not to continue sync.");
                    }
                }
            }
        }

        public Event UpdateCalendarEntry(AppointmentItem ai, Event ev, ref int itemModified, Boolean forceCompare = false) {
            SettingsStore.Calendar profile = Sync.Engine.Calendar.Instance.Profile;

            if (!Settings.Instance.APIlimit_inEffect && CustomProperty.Exists(ev, CustomProperty.MetadataId.apiLimitHit)) {
                log.Fine("Back processing Event affected by attendee API limit.");
            } else {
                if (!(Sync.Engine.Instance.ManualForceCompare || forceCompare)) { //Needed if the exception has just been created, but now needs updating
                    if (profile.SyncDirection.Id != Sync.Direction.Bidirectional.Id) {
                        if (ev.Updated > ai.LastModificationTime)
                            return null;
                    } else {
                        if (OutlookOgcs.CustomProperty.GetOGCSlastModified(ai).AddSeconds(5) >= ai.LastModificationTime) {
                            log.Fine("Outlook last modified by OGCS.");
                            return null;
                        }
                        if (ev.Updated > ai.LastModificationTime)
                            return null;
                    }
                }
            }

            String aiSummary = OutlookOgcs.Calendar.GetEventSummary(ai);
            log.Debug("Processing >> " + aiSummary);

            if (!(ev.Creator.Self ?? (ev.Creator.Email == Settings.Instance.GaccountEmail)) && ev.Recurrence != null) {
                log.Debug("Not being the recurring Event owner, comparison for update is futile - changes won't take effect/fail.");
                log.Fine("Owner: " + ev.Creator.Email);
                return ev;
            }

            System.Text.StringBuilder sb = new System.Text.StringBuilder();
            sb.AppendLine(aiSummary);

            //Handle an event's all-day attribute being toggled
            DateTime evStart = ev.Start.SafeDateTime();
            DateTime evEnd = ev.End.SafeDateTime();
            if (ai.AllDayEvent && ai.Start.TimeOfDay == new TimeSpan(0, 0, 0)) {
                ev.Start.DateTime = null;
                ev.End.DateTime = null;
                if (Sync.Engine.CompareAttribute("Start time", Sync.Direction.OutlookToGoogle, evStart, ai.Start.Date, sb, ref itemModified)) {
                    ev.Start.Date = ai.Start.ToString("yyyy-MM-dd");
                }
                if (Sync.Engine.CompareAttribute("End time", Sync.Direction.OutlookToGoogle, evEnd, ai.End.Date, sb, ref itemModified)) {
                    ev.End.Date = ai.End.ToString("yyyy-MM-dd");
                }
                //If there was no change in the start/end time, make sure we still have dates populated
                if (ev.Start.Date == null) ev.Start.Date = ai.Start.ToString("yyyy-MM-dd");
                if (ev.End.Date == null) ev.End.Date = ai.End.ToString("yyyy-MM-dd");

            } else {
                //Handle: Google = all-day; Outlook = not all day, but midnight values (so effectively all day!)
                if (ev.AllDayEvent() && evStart == ai.Start && evEnd == ai.End) {
                    sb.AppendLine("All-Day: true => false");
                    ev.Start.DateTime = ai.Start;
                    ev.End.DateTime = ai.End;
                    itemModified++;
                }
                ev.Start.Date = null;
                ev.End.Date = null;
                if (Sync.Engine.CompareAttribute("Start time", Sync.Direction.OutlookToGoogle, evStart, ai.Start, sb, ref itemModified)) {
                    ev.Start.DateTime = ai.Start;
                }
                if (Sync.Engine.CompareAttribute("End time", Sync.Direction.OutlookToGoogle, evEnd, ai.End, sb, ref itemModified)) {
                    ev.End.DateTime = ai.End;
                }
                //If there was no change in the start/end time, make sure we still have dates populated
                if (ev.Start.DateTime == null) ev.Start.DateTime = ai.Start;
                if (ev.End.DateTime == null) ev.End.DateTime = ai.End;
            }

            List<String> oRrules = Recurrence.Instance.BuildGooglePattern(ai, ev);
            if (ev.Recurrence != null) {
                for (int r = 0; r < ev.Recurrence.Count; r++) {
                    String rrule = ev.Recurrence[r];
                    if (rrule.StartsWith("RRULE:")) {
                        if (oRrules != null) {
                            String[] gRrule_bits = rrule.Split(';');
                            String[] oRrule_bits = oRrules.First().TrimStart("RRULE:".ToCharArray()).Split(';');
                            if (gRrule_bits.Count() != oRrule_bits.Count()) {
                                if (Sync.Engine.CompareAttribute("Recurrence", Sync.Direction.OutlookToGoogle, rrule, oRrules.First(), sb, ref itemModified)) {
                                    ev.Recurrence[r] = oRrules.First();
                                    break;
                                }
                            }
                            foreach (String oRrule_bit in oRrule_bits) {
                                if (!rrule.Contains(oRrule_bit)) {
                                    if (Sync.Engine.CompareAttribute("Recurrence", Sync.Direction.OutlookToGoogle, rrule, oRrules.First(), sb, ref itemModified)) {
                                        ev.Recurrence[r] = oRrules.First();
                                        break;
                                    }
                                }
                            }
                        } else {
                            log.Debug("Converting to non-recurring event.");
                            Sync.Engine.CompareAttribute("Recurrence", Sync.Direction.OutlookToGoogle, rrule, null, sb, ref itemModified);
                            ev.Recurrence[r] = null;
                        }
                        break;
                    }
                }
            } else {
                if (oRrules != null && ev.RecurringEventId == null) {
                    if (!(ev.Creator.Self ?? (ev.Creator.Email == Settings.Instance.GaccountEmail))) {
                        log.Warn("Cannot convert Event organised by another to a recurring series.");
                    } else {
                        log.Debug("Converting to recurring event.");
                        Sync.Engine.CompareAttribute("Recurrence", Sync.Direction.OutlookToGoogle, null, oRrules.First(), sb, ref itemModified);
                        ev.Recurrence = oRrules;
                    }
                }
            }

            //TimeZone
            if (ev.Start.DateTime != null) {
                String currentStartTZ = ev.Start.TimeZone;
                String currentEndTZ = ev.End.TimeZone;
                ev = OutlookOgcs.Calendar.Instance.IOutlook.IANAtimezone_set(ev, ai);
                if (ev.Recurrence != null && ev.Start.TimeZone != ev.End.TimeZone) {
                    log.Warn("Outlook recurring series has a different start and end timezone, which Google does not allow. Setting both to the start timezone.");
                    ev.End.TimeZone = ev.Start.TimeZone;
                }
                Sync.Engine.CompareAttribute("Start Timezone", Sync.Direction.OutlookToGoogle, currentStartTZ, ev.Start.TimeZone, sb, ref itemModified);
                Sync.Engine.CompareAttribute("End Timezone", Sync.Direction.OutlookToGoogle, currentEndTZ, ev.End.TimeZone, sb, ref itemModified);
            }

            String subjectObfuscated = Obfuscate.ApplyRegex(ai.Subject, ev.Summary, Sync.Direction.OutlookToGoogle);
            if (Sync.Engine.CompareAttribute("Subject", Sync.Direction.OutlookToGoogle, ev.Summary, subjectObfuscated, sb, ref itemModified)) {
                ev.Summary = subjectObfuscated;
            }
            if (profile.AddDescription) {
                String outlookBody = ai.Body;
                if (profile.SyncDirection == Sync.Direction.Bidirectional && profile.AddDescription_OnlyToGoogle &&
                    string.IsNullOrEmpty(outlookBody) && !string.IsNullOrEmpty(ev.Description))
                {
                    log.Warn("Avoided loss of Google description, as none exists in Outlook.");
                } else {
                    //Check for Google description truncated @ 8Kb
                    if (!string.IsNullOrEmpty(outlookBody) && !string.IsNullOrEmpty(ev.Description)
                        && ev.Description.Length == 8 * 1024
                        && outlookBody.Length > 8 * 1024) 
                    {
                        outlookBody = outlookBody.Substring(0, 8 * 1024);
                    }
                    if (Sync.Engine.CompareAttribute("Description", Sync.Direction.OutlookToGoogle, ev.Description, outlookBody, sb, ref itemModified))
                        ev.Description = outlookBody;
                }
            }

            if (profile.AddLocation && Sync.Engine.CompareAttribute("Location", Sync.Direction.OutlookToGoogle, ev.Location, ai.Location, sb, ref itemModified))
                ev.Location = ai.Location;

            String gPrivacy = ev.Visibility ?? "default";
            String oPrivacy = getPrivacy(ai.Sensitivity, gPrivacy);
            if (Sync.Engine.CompareAttribute("Privacy", Sync.Direction.OutlookToGoogle, gPrivacy, oPrivacy, sb, ref itemModified)) {
                ev.Visibility = oPrivacy;
            }

            String gFreeBusy = ev.Transparency ?? "opaque";
            String oFreeBusy = getAvailability(ai.BusyStatus, gFreeBusy);
            if (Sync.Engine.CompareAttribute("Free/Busy", Sync.Direction.OutlookToGoogle, gFreeBusy, oFreeBusy, sb, ref itemModified)) {
                ev.Transparency = oFreeBusy;
            }

            if (profile.AddColours || profile.SetEntriesColour) {
                EventColour.Palette gColour = this.ColourPalette.GetColour(ev.ColorId);
                EventColour.Palette oColour = getColour(ai.Categories, gColour);
                if (!string.IsNullOrEmpty(ai.Categories) && oColour == null)
                    log.Warn("Not comparing colour as there is a problem with the mapping.");
                else {
                    oColour ??= EventColour.Palette.NullPalette;
                    if (Sync.Engine.CompareAttribute("Colour", Sync.Direction.OutlookToGoogle, gColour.Name, oColour.Name, sb, ref itemModified)) {
                        ev.ColorId = oColour.Id;
                    }
                }
            }

            if (profile.AddAttendees && ai.Recipients.Count > 1 && !APIlimitReached_attendee) {
                if (ai.Recipients.Count > profile.MaxAttendees) {
                    log.Warn("This Outlook appointment has " + ai.Recipients.Count + " attendees, more than the user configured maximum.");
                    if (ai.Recipients.Count >= 200) {
                        Forms.Main.Instance.Console.Update(OutlookOgcs.Calendar.GetEventSummary(ai) + "<br/>Attendees will not be synced for this meeting as it has " +
                            "more than 200, which Google does not allow.", Console.Markup.warning);
                    }
                } else if (profile.SyncDirection.Id == Sync.Direction.Bidirectional.Id &&
                        ev.Attendees != null && ev.Attendees.Count > profile.MaxAttendees && ai.Recipients.Count <= profile.MaxAttendees) {
                    log.Warn("This Google event has " + ev.Attendees.Count + " attendees, more than the user configured maximum. They can't safely be compared.");
                } else {
                    try {
                        CompareRecipientsToAttendees(ai, ev, sb, ref itemModified);
                    } catch (System.Exception ex) {
                        if (OutlookOgcs.Calendar.Instance.IOutlook.ExchangeConnectionMode().ToString().Contains("Disconnected")) {
                            Forms.Main.Instance.Console.Update("Outlook is currently disconnected from Exchange, so it's not possible to sync attendees.<br/>" +
                                "Please reconnect or do not sync attendees.", Console.Markup.error);
                            throw new System.Exception("Outlook has disconnected from Exchange.");
                        } else {
                            Forms.Main.Instance.Console.UpdateWithError("Unable to sync attendees.", ex);
                        }
                    }
                }
            }

            #region Reminders
            if (profile.AddReminders) {
                Boolean OKtoSyncReminder = OutlookOgcs.Calendar.Instance.IsOKtoSyncReminder(ai);
                if (ev.Reminders.Overrides != null && ev.Reminders.Overrides.Any(r => r.Method == "popup")) {
                    //Find the popup reminder(s) in Google
                    for (int r = ev.Reminders.Overrides.Count - 1; r >= 0; r--) {
                        EventReminder reminder = ev.Reminders.Overrides[r];
                        if (reminder.Method == "popup") {
                            if (OKtoSyncReminder) {
                                if (ai.ReminderSet) {
                                    if (Sync.Engine.CompareAttribute("Reminder", Sync.Direction.OutlookToGoogle, reminder.Minutes.ToString(), ai.ReminderMinutesBeforeStart.ToString(), sb, ref itemModified)) {
                                        reminder.Minutes = ai.ReminderMinutesBeforeStart;
                                    }
                                } else {
                                    sb.AppendLine("Reminder: " + reminder.Minutes + " => removed");
                                    ev.Reminders.Overrides.Remove(reminder);
                                    itemModified++;
                                } //if Outlook reminders set
                            } else {
                                sb.AppendLine("Reminder: " + reminder.Minutes + " => removed");
                                ev.Reminders.Overrides.Remove(reminder);
                                ev.Reminders.UseDefault = false;
                                itemModified++;
                            }
                        } //if Google reminder found
                    } //foreach reminder

                } else { //no Google popup reminders set
                    if (ai.ReminderSet && OKtoSyncReminder) {
                        sb.AppendLine("Reminder: nothing => " + ai.ReminderMinutesBeforeStart);
                        ev.Reminders.UseDefault = false;
                        EventReminder newReminder = new EventReminder {
                            Method = "popup",
                            Minutes = ai.ReminderMinutesBeforeStart
                        };
                        ev.Reminders.Overrides = new List<EventReminder> { newReminder };
                        itemModified++;

                    } else if (ev.Reminders.Overrides == null) { //No Google email reminders either
                        Boolean newVal = OKtoSyncReminder ? profile.UseGoogleDefaultReminder : false;

                        //Google bug?! For all-day events, default notifications are added as overrides and UseDefault=false
                        //Which means it keeps adding the default back in!! Let's stop that:
                        if (newVal && ev.AllDayEvent()) {
                            log.Warn("Evading Google bug - not allowing default calendar notification to be (re?)set for all-day event.");
                            newVal = false;
                        }

                        if (Sync.Engine.CompareAttribute("Reminder Default", Sync.Direction.OutlookToGoogle, ev.Reminders.UseDefault.ToString(), newVal.ToString(), sb, ref itemModified)) {
                            ev.Reminders.UseDefault = newVal;
                        }
                    }
                }
            } else {
                if (ev.Reminders.Overrides == null) {
                    if (Sync.Engine.CompareAttribute("Reminder Default", Sync.Direction.OutlookToGoogle, ev.Reminders.UseDefault.ToString(), profile.UseGoogleDefaultReminder.ToString(), sb, ref itemModified))
                        ev.Reminders.UseDefault = profile.UseGoogleDefaultReminder;
                }
            }
            #endregion

            if (itemModified > 0) {
                Forms.Main.Instance.Console.FormatEventChanges(sb);
                Forms.Main.Instance.Console.Update(itemModified + " attributes updated.", Console.Markup.appointmentEnd, verbose: true, newLine: false);
                System.Windows.Forms.Application.DoEvents();
            }
            return ev;
        }

        public void UpdateCalendarEntry_save(ref Event ev) {
            SettingsStore.Calendar profile = Sync.Engine.Calendar.Instance.Profile;

            if (profile.SyncDirection.Id == Sync.Direction.Bidirectional.Id) {
                log.Debug("Saving timestamp when OGCS updated event.");
                CustomProperty.SetOGCSlastModified(ref ev);
            }
            if (Settings.Instance.APIlimit_inEffect)
                CustomProperty.Add(ref ev, CustomProperty.MetadataId.apiLimitHit, "True");
            else
                CustomProperty.Remove(ref ev, CustomProperty.MetadataId.apiLimitHit);

            CustomProperty.Remove(ref ev, CustomProperty.MetadataId.forceSave);

            int backoff = 0;
            while (backoff < BackoffLimit) {
                try {
                    EventsResource.UpdateRequest request = Service.Events.Update(ev, profile.UseGoogleCalendar.Id, ev.Id);
                    request.SendUpdates = EventsResource.UpdateRequest.SendUpdatesEnum.None;
                    ev = request.Execute();
                    if (profile.AddAttendees && Settings.Instance.APIlimit_inEffect) {
                        log.Info("API limit for attendee sync lifted :-)");
                        Settings.Instance.APIlimit_inEffect = false;
                    }
                    break;
                } catch (Google.GoogleApiException ex) {
                    ApiException handled = HandleAPIlimits(ref ex, ev);
                    switch (handled) {
                        case ApiException.throwException: throw;
                        case ApiException.freeAPIexhausted:
                            OGCSexception.LogAsFail(ref ex);
                            OGCSexception.Analyse(ex);
                            System.ApplicationException aex = new System.ApplicationException(SubscriptionInvite, ex);
                            OGCSexception.LogAsFail(ref aex);
                            throw aex;
                        case ApiException.backoffThenRetry:
                            backoff++;
                            if (backoff == BackoffLimit) {
                                log.Error("API limit backoff was not successful. Save failed.");
                                throw;
                            } else {
                                int backoffDelay = (int)Math.Pow(2, backoff);
                                log.Warn("API rate limit reached. Backing off " + backoffDelay + "sec before retry.");
                                System.Threading.Thread.Sleep(backoffDelay * 1000);
                            }
                            break;
                        case ApiException.justContinue:
                            backoff = BackoffLimit;
                            break;
                    }
                    if (handled != ApiException.justContinue && ex.Error?.Code == 412 && !this.openedIssue1593) { //Precondition failed
                        OgcsMessageBox.Show("A 'PreCondition Failed [412]' error was encountered.\r\nPlease see issue #1593 on GitHub for further information.",
                        "PreCondition Failed: Issue #1593", MessageBoxButtons.OK, MessageBoxIcon.Hand);
                        Helper.OpenBrowser("https://github.com/phw198/OutlookGoogleCalendarSync/issues/1593");
                        this.openedIssue1593 = true;
                    }
                }
            }
        }
        #endregion

        //void ShowError(String message, Window windowToBlock) {
        //    if (this.Dispatcher.CheckAccess())
        //        OgcsMessageBox.Show(windowToBlock, message, "Error", MessageBoxButton.OK, MessageBoxImage.Error);
        //    else {
        //        this.Dispatcher.Invoke(
        //            new Action(() => {
        //                OgcsMessageBox.Show(windowToBlock, message, "Error", MessageBoxButton.OK, MessageBoxImage.Error);
        //            }));
        //    }
        //}

        #region Delete
        public void DeleteCalendarEntries(List<Event> events) {
            for (int g = events.Count - 1; g >= 0; g--) {
                if (Sync.Engine.Instance.CancellationPending) return;

                Event ev = events[g];
                Boolean doDelete = false;
                try {
                    doDelete = deleteCalendarEntry(ev);
                } catch (System.Exception ex) {
                    Forms.Main.Instance.Console.UpdateWithError(GoogleOgcs.Calendar.GetEventSummary(ev, true) + "Event deletion failed.", ex);
                    OGCSexception.Analyse(ex, true);
                    if (OgcsMessageBox.Show("Google event deletion failed. Continue with synchronisation?", "Sync item failed", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                        continue;
                    else {
                        throw new UserCancelledSyncException("User chose not to continue sync.");
                    }
                }

                try {
                    if (doDelete) deleteCalendarEntry_save(ev);
                    else events.Remove(ev);
                } catch (System.Exception ex) {
                    if (ex is Google.GoogleApiException) {
                        Google.GoogleApiException gex = ex as Google.GoogleApiException;
                        if (gex.Error != null && gex.Error.Code == 410) { //Resource has been deleted
                            log.Fail("This event is already deleted! Ignoring failed request to delete.");
                            continue;
                        }
                    }
                    Forms.Main.Instance.Console.UpdateWithError(GoogleOgcs.Calendar.GetEventSummary(ev, true) + "Deleted event failed to remove.", ex);
                    OGCSexception.Analyse(ex, true);
                    if (OgcsMessageBox.Show("Deleted Google event failed to remove. Continue with synchronisation?", "Sync item failed", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
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

            if (Sync.Engine.Calendar.Instance.Profile.ConfirmOnDelete) {
                if (OgcsMessageBox.Show("Delete " + eventSummary + "?", "Confirm Deletion From Google",
                    MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2) == DialogResult.No) {
                    doDelete = false;
                    Forms.Main.Instance.Console.Update("Not deleted: " + eventSummary, Console.Markup.calendar);
                } else {
                    Forms.Main.Instance.Console.Update("Deleted: " + eventSummary, Console.Markup.calendar);
                }
            } else {
                Forms.Main.Instance.Console.Update(eventSummary, Console.Markup.calendar, verbose: true);
            }
            return doDelete;
        }

        private void deleteCalendarEntry_save(Event ev) {
            int backoff = 0;
            while (backoff < BackoffLimit) {
                try {
                    EventsResource.DeleteRequest request = Service.Events.Delete(Sync.Engine.Calendar.Instance.Profile.UseGoogleCalendar.Id, ev.Id);
                    request.SendUpdates = EventsResource.DeleteRequest.SendUpdatesEnum.None;
                    string result = request.Execute();
                    break;
                } catch (Google.GoogleApiException ex) {
                    switch (HandleAPIlimits(ref ex, ev)) {
                        case ApiException.throwException: throw;
                        case ApiException.freeAPIexhausted:
                            OGCSexception.LogAsFail(ref ex);
                            OGCSexception.Analyse(ex);
                            System.ApplicationException aex = new System.ApplicationException(SubscriptionInvite, ex);
                            OGCSexception.LogAsFail(ref aex);
                            throw aex;
                        case ApiException.backoffThenRetry:
                            backoff++;
                            if (backoff == BackoffLimit) {
                                log.Error("API limit backoff was not successful. Save failed.");
                                throw;
                            } else {
                                int backoffDelay = (int)Math.Pow(2, backoff);
                                log.Warn("API rate limit reached. Backing off " + backoffDelay + "sec before retry.");
                                System.Threading.Thread.Sleep(backoffDelay * 1000);
                            }
                            break;
                    }
                }
            }
        }
        #endregion

        public void ReclaimOrphanCalendarEntries(ref List<Event> gEvents, ref List<AppointmentItem> oAppointments, Boolean neverDelete = false) {
            SettingsStore.Calendar profile = Sync.Engine.Calendar.Instance.Profile;

            if (profile.SyncDirection.Id == Sync.Direction.GoogleToOutlook.Id) return;

            if (!neverDelete) Forms.Main.Instance.Console.Update("Checking for orphaned Google items", verbose: true);
            try {
                log.Debug("Scanning " + gEvents.Count + " Google events for orphans to reclaim...");
                String consoleTitle = "Reclaiming Google calendar entries";

                //This is needed for people migrating from other tools, which do not have our OutlookID extendedProperty
                List<Event> unclaimedEvents = new List<Event>();

                for (int g = gEvents.Count - 1; g >= 0; g--) {
                    if (Sync.Engine.Instance.CancellationPending) return;
                    Event ev = gEvents[g];
                    CustomProperty.LogProperties(ev, Program.MyFineLevel);

                    //Find entries with no Outlook ID
                    if (!CustomProperty.Exists(ev, CustomProperty.MetadataId.oEntryId)) {

                        //Use simple matching on start,end,subject,location to pair events
                        String sigEv = signature(ev);
                        if (String.IsNullOrEmpty(sigEv)) {
                            gEvents.Remove(ev);
                            continue;
                        }

                        unclaimedEvents.Add(ev);
                        for (int o = oAppointments.Count - 1; o >= 0; o--) {
                            AppointmentItem ai = oAppointments[o];
                            if (SignaturesMatch(sigEv, OutlookOgcs.Calendar.signature(ai))) {
                                try {
                                    Event originalEv = ev;
                                    CustomProperty.AddOutlookIDs(ref ev, ai);
                                    UpdateCalendarEntry_save(ref ev);
                                    unclaimedEvents.Remove(originalEv);
                                    if (consoleTitle != "") Forms.Main.Instance.Console.Update("<span class='em em-reclaim'></span>" + consoleTitle, Console.Markup.h2, newLine: false, verbose: true);
                                    consoleTitle = "";
                                    Forms.Main.Instance.Console.Update("Reclaimed: " + GetEventSummary(ev), verbose: true);
                                    gEvents[g] = ev;
                                    if (profile.SyncDirection.Id == Sync.Direction.Bidirectional.Id || OutlookOgcs.CustomProperty.ExistsAny(ai)) {
                                        log.Debug("Updating the Google event IDs in Outlook appointment.");
                                        OutlookOgcs.CustomProperty.AddGoogleIDs(ref ai, ev);
                                        ai.Save();
                                    }
                                } catch (System.Exception ex) {
                                    log.Error("Failed to reclaim Event: " + GetEventSummary(ev));
                                    log.Debug(ex.Message);
                                    log.Debug("Event status: " + ev.Status);
                                }
                                break;
                            }
                        }
                    }
                    if (Sync.Engine.Instance.CancellationPending) return;
                }
                log.Debug(unclaimedEvents.Count + " unclaimed.");
                if (!neverDelete && unclaimedEvents.Count > 0 &&
                    (profile.SyncDirection.Id == Sync.Direction.OutlookToGoogle.Id ||
                     profile.SyncDirection.Id == Sync.Direction.Bidirectional.Id)) 
                {
                    log.Info(unclaimedEvents.Count + " unclaimed orphan events found.");
                    if (profile.MergeItems || profile.DisableDelete || profile.ConfirmOnDelete) {
                        log.Info("These will be kept due to configuration settings.");
                    } else if (profile.SyncDirection.Id == Sync.Direction.Bidirectional.Id) {
                        log.Debug("These 'orphaned' items must not be deleted - they need syncing up.");
                    } else {
                        if (OgcsMessageBox.Show(unclaimedEvents.Count + " Google calendar events can't be matched to Outlook.\r\n" +
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
            } catch (System.Exception) {
                Forms.Main.Instance.Console.Update("Unable to reclaim orphan calendar entries in Google calendar.", Console.Markup.error);
                throw;
            }
        }

        public void CleanDuplicateEntries(ref List<Event> google) {
            //If a recurring series is altered for "this and following events", Google duplicates the original series.
            //This includes the private ExtendedProperties, containing the Outlook IDs - not good, these need to be detected and removed

            log.Debug("Checking for Events that have been duplicated.");

            try {
                List<Event> duplicateCheck = google.Where(w => CustomProperty.Exists(w, CustomProperty.MetadataId.oEntryId)).ToList();
                duplicateCheck = duplicateCheck.
                    GroupBy(e => new { e.Created, oEntryId = CustomProperty.Get(e, CustomProperty.MetadataId.oEntryId) }).
                    Where(g => g.Count() > 1).
                    SelectMany(x => x).ToList();
                if (duplicateCheck.Count() == 0) return;

                log.Warn(duplicateCheck.Count() + " Events found with same creation date and Outlook EntryID.");
                duplicateCheck.Sort((x, y) => (x.CreatedRaw + ":" + (x.Sequence ?? 0)).CompareTo(y.CreatedRaw + ":" + (y.Sequence ?? 0)));
                //Skip the first one, the original 
                DateTime? lastSeenDuplicateSet = null;
                for (int g = 0; g < duplicateCheck.Count(); g++) {
                    Event ev = duplicateCheck[g];
                    if (lastSeenDuplicateSet == null || lastSeenDuplicateSet != ev.Created) {
                        lastSeenDuplicateSet = ev.Created;
                        continue;
                    }
                    log.Info("Cleaning duplicate metadata from: " + GetEventSummary(ev));
                    google.Remove(ev);
                    CustomProperty.RemoveAll(ref ev);
                    this.UpdateCalendarEntry_save(ref ev);
                    google.Add(ev);
                    lastSeenDuplicateSet = ev.Created;
                }
            } catch (System.Exception ex) {
                OGCSexception.Analyse(ex);
            }
        }

        public void IdentifyEventDifferences_Simple(
            SettingsStore.Calendar profile,
            ref List<AppointmentItem> outlook,  //need creating
            ref List<Event> google,             //need deleting
            ref Dictionary<AppointmentItem, Event> compare)
        {
            Forms.Main.Instance.Console.Update("Matching calendar items using simple method...");

            //Order by start date (same as Outlook) for quickest matching
            google.Sort((x, y) => (x.Start.DateTimeRaw ?? x.Start.Date).CompareTo((y.Start.DateTimeRaw ?? y.Start.Date)));

            // Count backwards so that we can remove found items without affecting the order of remaining items
            for (int g = google.Count - 1; g >= 0; g--) {
                if (Sync.Engine.Instance.CancellationPending) return;
                log.Fine("Checking " + GoogleOgcs.Calendar.GetEventSummary(google[g]));

                //Use simple matching on start,end,subject,location to pair events
                String sigEv = signature(google[g]);
                if (String.IsNullOrEmpty(sigEv)) {
                    google.Remove(google[g]);
                    continue;
                }

                Boolean foundMatch = false;
                for (int o = outlook.Count - 1; o >= 0; o--) {
                    try {
                        if (log.IsUltraFineEnabled()) log.UltraFine("Checking " + OutlookOgcs.Calendar.GetEventSummary(outlook[o]));

                        if (SignaturesMatch(sigEv, OutlookOgcs.Calendar.signature(outlook[o]))) {
                            foundMatch = true;
                            compare.Add(outlook[o], google[g]);
                            outlook.Remove(outlook[o]);
                            google.Remove(google[g]);
                            break;
                        }
                    } catch (System.Exception ex) {
                        if (!log.IsUltraFineEnabled()) {
                            try {
                                log.Info(OutlookOgcs.Calendar.GetEventSummary(outlook[o]));
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
                if (!foundMatch && profile.MergeItems)
                    google.Remove(google[g]);
            }
        }

        public void IdentifyEventDifferences_IDs(
            SettingsStore.Calendar profile,
            ref List<AppointmentItem> outlook,  //need creating
            ref List<Event> google,             //need deleting
            ref Dictionary<AppointmentItem, Event> compare) 
        {
            Forms.Main.Instance.Console.Update("Matching calendar items...");

            //Order by start date (same as Outlook) for quickest matching
            google.Sort((x, y) => (x.Start.DateTimeRaw ?? x.Start.Date).CompareTo((y.Start.DateTimeRaw ?? y.Start.Date)));

            // Count backwards so that we can remove found items without affecting the order of remaining items
            int metadataEnhanced = 0;
            for (int g = google.Count - 1; g >= 0; g--) {
                if (Sync.Engine.Instance.CancellationPending) return;
                log.Fine("Checking " + GoogleOgcs.Calendar.GetEventSummary(google[g]));

                if (CustomProperty.Exists(google[g], CustomProperty.MetadataId.oEntryId)) {
                    String compare_gEntryID = CustomProperty.Get(google[g], CustomProperty.MetadataId.oEntryId);
                    Boolean outlookIDmissing = CustomProperty.OutlookIdMissing(google[g]);
                    Boolean foundMatch = false;

                    for (int o = outlook.Count - 1; o >= 0; o--) {
                        try {
                            if (log.IsUltraFineEnabled()) log.UltraFine("Checking " + OutlookOgcs.Calendar.GetEventSummary(outlook[o]));

                            String compare_oID;
                            if (outlookIDmissing && compare_gEntryID.StartsWith(OutlookOgcs.Calendar.GlobalIdPattern)) {
                                //compare_gEntryID actually holds GlobalID up to v2.3.2.3 - yes, confusing I know, but we're sorting this now
                                compare_oID = OutlookOgcs.Calendar.Instance.IOutlook.GetGlobalApptID(outlook[o]);
                            } else {
                                compare_oID = outlook[o].EntryID;
                            }
                            if (compare_gEntryID == compare_oID && outlookIDmissing) {
                                log.Info("Enhancing event's metadata...");
                                Event ev = google[g];
                                CustomProperty.AddOutlookIDs(ref ev, outlook[o]);
                                //Don't want to save right now, else may make modified timestamp newer than a change in Outlook
                                //which would no longer sync.
                                CustomProperty.Add(ref ev, CustomProperty.MetadataId.forceSave, "True");
                                google[g] = ev;
                                metadataEnhanced++;
                            }

                            Event evCheck = google[g];
                            if (ItemIDsMatch(ref evCheck, outlook[o])) {
                                foundMatch = true;
                                google[g] = evCheck;
                                compare.Add(outlook[o], google[g]);
                                outlook.Remove(outlook[o]);
                                google.Remove(google[g]);
                                break;
                            }
                        } catch (System.Exception ex) {
                            if (!log.IsUltraFineEnabled()) {
                                try {
                                    log.Info(OutlookOgcs.Calendar.GetEventSummary(outlook[o]));
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
                    if (!foundMatch && profile.MergeItems &&
                        GoogleOgcs.CustomProperty.Get(google[g], CustomProperty.MetadataId.oCalendarId) != profile.UseOutlookCalendar.Id)
                        google.Remove(google[g]);

                } else if (profile.MergeItems) {
                    //Remove the non-Outlook item so it doesn't get deleted
                    google.Remove(google[g]);
                }
            }
            if (metadataEnhanced > 0) log.Info(metadataEnhanced + " item's metadata enhanced.");
        }

        //<summary>Logic for comparing Outlook and Google events works as follows:
        //      1.  Scan through both lists looking for matches
        //      2.  Remove matches from both lists and add to the compare dictionary
        //      3.  Items remaining in Outlook list are new and need to be created
        //      4.  Items remaining in Google list need to be deleted
        //</summary>
        public void IdentifyEventDifferences(
            ref List<AppointmentItem> outlook,  //need creating
            ref List<Event> google,             //need deleting
            ref Dictionary<AppointmentItem, Event> compare)
        {
            log.Debug("Comparing Outlook items to Google events...");
            SettingsStore.Calendar profile = Sync.Engine.Calendar.Instance.Profile;

            if (profile.SimpleMatch)
                IdentifyEventDifferences_Simple(profile, ref outlook, ref google, ref compare);
            else
                IdentifyEventDifferences_IDs(profile, ref outlook, ref google, ref compare);

            if (Sync.Engine.Instance.CancellationPending) return;

            if (outlook.Count > 0 && profile.OnlyRespondedInvites) {
                //Check if Outlook items to be created in Google have invitations not yet responded to
                int responseFiltered = 0;
                for (int o = outlook.Count - 1; o >= 0; o--) {
                    if (outlook[o].ResponseStatus == OlResponseStatus.olResponseNotResponded) {
                        outlook.Remove(outlook[o]);
                        responseFiltered++;
                    }
                }
                if (responseFiltered > 0) log.Info(responseFiltered + " Outlook items will not be created due to only syncing invites that have been responded to.");
            }

            if (google.Count > 0 && OutlookOgcs.Calendar.Instance.ExcludedByCategory.Count > 0 && profile.SyncDirection.Id == Sync.Direction.Bidirectional.Id && !profile.DeleteWhenCategoryExcluded) {
                //Check if Google items to be deleted were filtered out from Outlook
                for (int g = google.Count - 1; g >= 0; g--) {
                    if (CustomProperty.Exists(google[g], CustomProperty.MetadataId.oEntryId) &&
                        OutlookOgcs.Calendar.Instance.ExcludedByCategory.Contains(CustomProperty.Get(google[g], CustomProperty.MetadataId.oEntryId))) {
                        google.Remove(google[g]);
                    }
                }
            }

            if (profile.DisableDelete) {
                if (google.Count > 0) {
                    Forms.Main.Instance.Console.Update(google.Count + " Google items would have been deleted, but you have deletions disabled.", Console.Markup.warning);
                    for (int g = 0; g < google.Count; g++)
                        Forms.Main.Instance.Console.Update(GetEventSummary(google[g]), verbose: true);
                }
                google = new List<Event>();
            }
            if (profile.SyncDirection.Id == Sync.Direction.Bidirectional.Id) {
                //Don't recreate any items that have been deleted in Google
                for (int o = outlook.Count - 1; o >= 0; o--) {
                    if (OutlookOgcs.CustomProperty.Exists(outlook[o], OutlookOgcs.CustomProperty.MetadataId.gEventID))
                        outlook.Remove(outlook[o]);
                }
                //Don't delete any items that aren't yet in Outlook or just created in Outlook during this sync
                for (int g = google.Count - 1; g >= 0; g--) {
                    if (!CustomProperty.Exists(google[g], CustomProperty.MetadataId.oEntryId) ||
                        google[g].Updated > Sync.Engine.Instance.SyncStarted)
                        google.Remove(google[g]);
                }
            }
            if (Settings.Instance.CreateCSVFiles) {
                ExportToCSV("Events for deletion in Google", "google_delete.csv", google);
                OutlookOgcs.Calendar.ExportToCSV("Appointments for creation in Google", "google_create.csv", outlook);
            }
        }

        public static Boolean ItemIDsMatch(ref Event ev, AppointmentItem ai) {
            //AppointmentItem Entry IDs change when accepting invites; causes item to be recreated if only match on that
            //So first match on the Global Appointment ID

            //For format of Entry ID : https://msdn.microsoft.com/en-us/library/ee201952(v=exchg.80).aspx
            //For format of Global ID: https://msdn.microsoft.com/en-us/library/ee157690%28v=exchg.80%29.aspx
            log.Fine("Comparing Outlook GlobalID");

            SettingsStore.Calendar profile = Sync.Engine.Calendar.Instance.Profile;

            if (CustomProperty.Exists(ev, CustomProperty.MetadataId.oGlobalApptId)) {
                String gCompareID = CustomProperty.Get(ev, CustomProperty.MetadataId.oGlobalApptId);
                String oGlobalID = OutlookOgcs.Calendar.Instance.IOutlook.GetGlobalApptID(ai);

                //For items copied from someone elses calendar, it appears the Global ID is generated for each access?! (Creation Time changes)
                //I guess the copied item doesn't really have its "own" ID. So, we'll just compare
                //the "data" section of the byte array, which "ensures uniqueness" and doesn't include ID creation time

                if ((OutlookOgcs.Factory.OutlookVersionName == OutlookOgcs.Factory.OutlookVersionNames.Outlook2003 && oGlobalID == gCompareID) //Actually simple compare of EntryId for O2003
                    ||
                    (OutlookOgcs.Factory.OutlookVersionName != OutlookOgcs.Factory.OutlookVersionNames.Outlook2003 &&
                        (
                            (oGlobalID.StartsWith(OutlookOgcs.Calendar.GlobalIdPattern) &&
                            gCompareID.StartsWith(OutlookOgcs.Calendar.GlobalIdPattern) &&
                            gCompareID.Substring(72) == oGlobalID.Substring(72))             //We've got bonafide Global IDs match
                            ||
                            (!oGlobalID.StartsWith(OutlookOgcs.Calendar.GlobalIdPattern) &&
                            !gCompareID.StartsWith(OutlookOgcs.Calendar.GlobalIdPattern) &&
                            gCompareID.Remove(gCompareID.Length - 16) == oGlobalID.Remove(oGlobalID.Length - 16)) //Or it's really a Entry ID (failsafe match)
                        )
                    ))
                {
                    log.Fine("Comparing Outlook CalendarID");
                    gCompareID = CustomProperty.Get(ev, CustomProperty.MetadataId.oCalendarId);
                    if (gCompareID == profile.UseOutlookCalendar.Id) {

                        //But...if an appointment is copied within ones own calendar, the DATA part is the same (only the creation time changes)!
                        //So now compare the Entry ID too.
                        log.Fine("Comparing Outlook EntryID");
                        gCompareID = CustomProperty.Get(ev, CustomProperty.MetadataId.oEntryId);
                        if (gCompareID == ai.EntryID) {
                            return true;
                        } else if (!string.IsNullOrEmpty(gCompareID) &&
                            gCompareID.Remove(gCompareID.Length - 16) == ai.EntryID.Remove(ai.EntryID.Length - 16))
                        {
                            //Worse still, both a locally copied item AND a rescheduled appointment by someone else 
                            //will have the MessageGlobalCounter bytes incremented (last 8-bytes)
                            //The former is identified by ExplorerWatcher adding a special flag
                            if (OutlookOgcs.CustomProperty.Get(ai, OutlookOgcs.CustomProperty.MetadataId.locallyCopied) == true.ToString()) {
                                log.Fine("This appointment was copied by the user. Incorrect match avoided.");
                                return false;
                            } else {
                                if (profile.OutlookGalBlocked || ai.Organizer != OutlookOgcs.Calendar.Instance.IOutlook.CurrentUserName()) {
                                    if (profile.OutlookGalBlocked)
                                        log.Warn("It looks like the organiser changed time of appointment, but due to GAL policy we can't check who they are.");
                                    else
                                        log.Fine("Organiser changed time of appointment.");
                                    CustomProperty.AddOutlookIDs(ref ev, ai); //update EntryID
                                    CustomProperty.Add(ref ev, CustomProperty.MetadataId.forceSave, "True");
                                    return true;
                                } else {
                                    log.Warn("Organiser changed time of appointment...but the organiser is you! (Shouldn't have ended up here)");
                                    return false;
                                }
                            }

                        } else {
                            log.Fine("EntryID has changed - invite accepted?");
                            if (SignaturesMatch(signature(ev), OutlookOgcs.Calendar.signature(ai))) {
                                CustomProperty.AddOutlookIDs(ref ev, ai); //update EntryID
                                CustomProperty.Add(ref ev, CustomProperty.MetadataId.forceSave, "True");
                                return true;
                            }
                        }
                    }

                } else if (profile.SyncDirection.Id == Sync.Direction.Bidirectional.Id &&
                    oGlobalID.StartsWith(OutlookOgcs.Calendar.GlobalIdPattern) &&
                    gCompareID.StartsWith(OutlookOgcs.Calendar.GlobalIdPattern) &&
                    gCompareID.Substring(72) != oGlobalID.Substring(72) &&
                    OutlookOgcs.CustomProperty.Get(ai, OutlookOgcs.CustomProperty.MetadataId.gEventID) == ev.Id &&
                    SignaturesMatch(signature(ev), OutlookOgcs.Calendar.signature(ai))) 
                {
                    //Apple iCloud completely recreates the GlobalID and zeros out the timestamp element! Issue #447.
                    log.Warn("Appointment GlobalID has completely changed, but Google Event ID matches so relying on that!");
                    log.Debug("Google's Event Id: " + ev.Id);
                    log.Debug("Google's Outlook Global Id: " + gCompareID);
                    log.Debug("Outlook's new Global Id: " + oGlobalID);
                    CustomProperty.AddOutlookIDs(ref ev, ai); //update GlobalID
                    CustomProperty.Add(ref ev, CustomProperty.MetadataId.forceSave, "True");
                    return true;
                }
            } else {
                if (profile.MergeItems)
                    log.Fine("Could not find global Appointment ID against Google Event.");
                else
                    log.Warn("Could not find global Appointment ID against Google Event.");
            }
            return false;
        }

        public Boolean CompareRecipientsToAttendees(AppointmentItem ai, Event ev, StringBuilder sb, ref int itemModified) {
            log.Fine("Comparing Recipients");
            //Build a list of Google attendees. Any remaining at the end of the diff must be deleted.
            List<Google.Apis.Calendar.v3.Data.EventAttendee> removeAttendee = new List<Google.Apis.Calendar.v3.Data.EventAttendee>();
            foreach (Google.Apis.Calendar.v3.Data.EventAttendee ea in ev.Attendees ?? Enumerable.Empty<Google.Apis.Calendar.v3.Data.EventAttendee>()) {
                removeAttendee.Add(ea);
            }
            if (ai.Recipients.Count > 1) {
                for (int o = ai.Recipients.Count; o > 0; o--) {
                    bool foundAttendee = false;
                    Recipient recipient = ai.Recipients[o];
                    log.Fine("Comparing Outlook recipient: " + recipient.Name);
                    String recipientSMTP = OutlookOgcs.Calendar.Instance.IOutlook.GetRecipientEmail(recipient);
                    foreach (Google.Apis.Calendar.v3.Data.EventAttendee attendee in ev.Attendees ?? Enumerable.Empty<Google.Apis.Calendar.v3.Data.EventAttendee>()) {
                        GoogleOgcs.EventAttendee ogcsAttendee = new GoogleOgcs.EventAttendee(attendee);
                        if (ogcsAttendee.Email != null && (recipientSMTP.ToLower() == ogcsAttendee.Email.ToLower())) {
                            foundAttendee = true;
                            removeAttendee.Remove(attendee);

                            if (Sync.Engine.Calendar.Instance.Profile.CloakEmail != ogcsAttendee.IsCloaked()) {
                                Sync.Engine.CompareAttribute("Attendee updated", Sync.Direction.OutlookToGoogle, attendee.Email, EventAttendee.CloakEmail(attendee.Email), sb, ref itemModified);
                                attendee.Email = EventAttendee.CloakEmail(attendee.Email);
                            }

                            //Optional attendee
                            bool oOptional = (recipient.Type == (int)OlMeetingRecipientType.olOptional);
                            bool gOptional = (attendee.Optional == null) ? false : (bool)attendee.Optional;
                            String attendeeIdentifier = attendee.DisplayName ?? ogcsAttendee.Email;
                            if (Sync.Engine.CompareAttribute("Attendee " + attendeeIdentifier + " - Optional Check",
                                Sync.Direction.OutlookToGoogle, gOptional, oOptional, sb, ref itemModified)) {
                                attendee.Optional = oOptional;
                            }

                            //Response
                            if (attendeeIdentifier == Settings.Instance.GaccountEmail) {
                                log.Fine("The Outlook attendee is the Google organiser, therefore not touching response status.");
                                break;
                            } else if (ai.Organizer == attendeeIdentifier) {
                                if (Sync.Engine.CompareAttribute("Organiser " + attendeeIdentifier + " - Response Status",
                                    Sync.Direction.OutlookToGoogle,
                                    attendee.ResponseStatus, "accepted", sb, ref itemModified)) {
                                    log.Fine("Forcing the Outlook organiser to have accepted the 'invite' in Google");
                                    attendee.ResponseStatus = "accepted";
                                }
                                break;
                            }

                            switch (recipient.MeetingResponseStatus) {
                                case OlResponseStatus.olResponseNone:
                                    if (Sync.Engine.CompareAttribute("Attendee " + attendeeIdentifier + " - Response Status",
                                        Sync.Direction.OutlookToGoogle,
                                        attendee.ResponseStatus, "needsAction", sb, ref itemModified)) {
                                        attendee.ResponseStatus = "needsAction";
                                    }
                                    break;
                                case OlResponseStatus.olResponseAccepted:
                                    if (Sync.Engine.CompareAttribute("Attendee " + attendeeIdentifier + " - Response Status",
                                        Sync.Direction.OutlookToGoogle,
                                        attendee.ResponseStatus, "accepted", sb, ref itemModified)) {
                                        attendee.ResponseStatus = "accepted";
                                    }
                                    break;
                                case OlResponseStatus.olResponseDeclined:
                                    if (Sync.Engine.CompareAttribute("Attendee " + attendeeIdentifier + " - Response Status",
                                        Sync.Direction.OutlookToGoogle,
                                        attendee.ResponseStatus, "declined", sb, ref itemModified)) {
                                        attendee.ResponseStatus = "declined";
                                    }
                                    break;
                                case OlResponseStatus.olResponseTentative:
                                    if (Sync.Engine.CompareAttribute("Attendee " + attendeeIdentifier + " - Response Status",
                                        Sync.Direction.OutlookToGoogle,
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
                        if (ev.Attendees == null) ev.Attendees = new List<Google.Apis.Calendar.v3.Data.EventAttendee>();
                        ev.Attendees.Add(GoogleOgcs.Calendar.CreateAttendee(recipient, ai.Organizer == recipient.Name));
                        itemModified++;
                    }
                }
            } //more than just 1 (me) recipients

            foreach (Google.Apis.Calendar.v3.Data.EventAttendee gea in removeAttendee) {
                GoogleOgcs.EventAttendee ea = new GoogleOgcs.EventAttendee(gea);
                log.Fine("Attendee removed: " + (ea.DisplayName ?? ea.Email), ea.Email);
                sb.AppendLine("Attendee removed: " + (ea.DisplayName ?? ea.Email));
                ev.Attendees.Remove(gea);
                itemModified++;
            }
            return (itemModified > 0);
        }

        /// <summary>
        /// Get the global Calendar settings
        /// </summary>
        public void GetSettings() {
            int backoff = 0;
            String stage = "retrieve Google calendar's global timezone";
            while (backoff < BackoffLimit) {
                try {
                    log.Fine("Get the timezone offset - convert from IANA string to UTC offset integer.");
                    Setting setting = Service.Settings.Get("timezone").Execute();
                    this.UTCoffset = TimezoneDB.GetUtcOffset(setting.Value);
                    log.Info("Google account timezone: " + setting.Value);
                    stage = "retrieve settings for synced Google calendar";
                    getCalendarSettings();
                    break;
                } catch (Google.GoogleApiException ex) {
                    switch (HandleAPIlimits(ref ex, null)) {
                        case ApiException.throwException: throw;
                        case ApiException.freeAPIexhausted:
                            OGCSexception.LogAsFail(ref ex);
                            OGCSexception.Analyse("Not able to " + stage, ex);
                            System.ApplicationException aex = new System.ApplicationException(SubscriptionInvite, ex);
                            OGCSexception.LogAsFail(ref aex);
                            throw aex;
                        case ApiException.backoffThenRetry:
                            backoff++;
                            if (backoff == BackoffLimit) {
                                log.Error("API limit backoff was not successful. Save failed.");
                                throw;
                            } else {
                                int backoffDelay = (int)Math.Pow(2, backoff);
                                log.Warn("API rate limit reached. Backing off " + backoffDelay + "sec before retry.");
                                System.Threading.Thread.Sleep(backoffDelay * 1000);
                            }
                            break;
                    }
                    OGCSexception.Analyse("Not able to " + stage, ex);
                    throw new System.ApplicationException("Unable to " + stage + ".", ex);

                } catch (System.Exception ex) {
                    OGCSexception.Analyse("Not able to " + stage, ex);
                    throw;
                }
            }
        }
        private void getCalendarSettings() {
            SettingsStore.Calendar profile = Settings.Profile.InPlay();
            CalendarListResource.GetRequest request = Service.CalendarList.Get(profile.UseGoogleCalendar.Id);
            CalendarListEntry cal = request.Execute();
            log.Info("Google calendar timezone: " + cal.TimeZone);

            if (!profile.AddReminders) return;

            if (cal.DefaultReminders.Count == 0)
                this.MinDefaultReminder = int.MinValue;
            else
                this.MinDefaultReminder = cal.DefaultReminders.Where(x => x.Method.Equals("popup")).OrderBy(x => x.Minutes.Value).FirstOrDefault()?.Minutes.Value ?? int.MinValue;
        }

        /// <summary>
        /// Determine Event's privacy setting
        /// </summary>
        /// <param name="oSensitivity">Outlook's current setting</param>
        /// <param name="gVisibility">Google's current setting</param>
        /// <param name="direction">Direction of sync</param>
        private String getPrivacy(OlSensitivity oSensitivity, String gVisibility) {
            SettingsStore.Calendar profile = Sync.Engine.Calendar.Instance.Profile;

            if (!profile.SetEntriesPrivate)
                return (oSensitivity == OlSensitivity.olNormal) ? "default" : "private";

            if (profile.SyncDirection.Id != Sync.Direction.Bidirectional.Id) {
                return (profile.PrivacyLevel == OlSensitivity.olPrivate.ToString()) ? "private" : "public";
            } else {
                if (profile.TargetCalendar.Id == Sync.Direction.GoogleToOutlook.Id) { //Privacy enforcement is in other direction
                    if (gVisibility == null)
                        return (oSensitivity == OlSensitivity.olNormal) ? "default" : "private";
                    else if (gVisibility == "private" && oSensitivity != OlSensitivity.olPrivate) {
                        log.Fine("Source of truth for privacy is already set private and target is NOT - so syncing this back.");
                        return "default";
                    } else
                        return gVisibility;
                } else {
                    if (!profile.CreatedItemsOnly || (profile.CreatedItemsOnly && gVisibility == null))
                        return (profile.PrivacyLevel == OlSensitivity.olPrivate.ToString()) ? "private" : "public";
                    else
                        return (oSensitivity == OlSensitivity.olNormal) ? "default" : "private";
                }
            }
        }

        /// <summary>
        /// Determine Event's availability setting
        /// </summary>
        /// <param name="oSsensitivity">Outlook's current setting</param>
        /// <param name="gTransparency">Google's current setting</param>
        private String getAvailability(OlBusyStatus oBusyStatus, String gTransparency) {
            SettingsStore.Calendar profile = Sync.Engine.Calendar.Instance.Profile;

            if (!profile.SetEntriesAvailable)
                return (oBusyStatus == OlBusyStatus.olFree) ? "transparent" : "opaque";

            String overrideTransparency = "transparent";
            OlBusyStatus fbStatus = OlBusyStatus.olFree;
            try {
                Enum.TryParse(profile.AvailabilityStatus, out fbStatus);
                if (fbStatus != OlBusyStatus.olFree)
                    overrideTransparency = "opaque";
            } catch (System.Exception ex) {
                OGCSexception.Analyse("Could not convert string '" + profile.AvailabilityStatus + "' to OlBusyStatus type. Defaulting override to available.", ex);
            }

            if (profile.SyncDirection.Id != Sync.Direction.Bidirectional.Id) {
                return overrideTransparency;
            } else {
                if (profile.TargetCalendar.Id == Sync.Direction.GoogleToOutlook.Id) { //Availability enforcement is in other direction
                    if (gTransparency == null)
                        return (oBusyStatus == OlBusyStatus.olFree) ? "transparent" : "opaque";
                    else
                        return gTransparency;
                } else {
                    if (!profile.CreatedItemsOnly || (profile.CreatedItemsOnly && gTransparency == null))
                        return overrideTransparency;
                    else
                        return (oBusyStatus == OlBusyStatus.olFree) ? "transparent" : "opaque";
                }
            }
        }

        /// <summary>
        /// Get the Google palette colour from a list of Outlook categories
        /// </summary>
        /// <param name="aiCategories">The appointment item "categories" field</param>
        /// <param name="gColour">The Google palette, if already assigned to Event</param>
        /// <returns>A match or a "null" Palette signifying no match</returns>
        private EventColour.Palette getColour(String aiCategories, EventColour.Palette gColour) {
            SettingsStore.Calendar profile = Sync.Engine.Calendar.Instance.Profile;

            if (!profile.AddColours && !profile.SetEntriesColour) return EventColour.Palette.NullPalette;

            OlCategoryColor? categoryColour = null;

            if (profile.SetEntriesColour) {
                if (profile.TargetCalendar.Id == Sync.Direction.GoogleToOutlook.Id) { //Colour forced to sync in other direction
                    if (gColour == null) //Creating item
                        return this.ColourPalette.ActivePalette[Convert.ToInt16(profile.SetEntriesColourGoogleId)];
                    else return gColour;

                } else {
                    if (!profile.CreatedItemsOnly || (profile.CreatedItemsOnly && gColour == null)) {
                        return this.ColourPalette.ActivePalette[Convert.ToInt16(profile.SetEntriesColourGoogleId)];
                    } else return gColour;
                }

            } else {
                getOutlookCategoryColour(aiCategories, ref categoryColour);
            }
            if (categoryColour == null)
                return null;
            else if (categoryColour == OlCategoryColor.olCategoryColorNone)
                return EventColour.Palette.NullPalette;
            else
                return GetColour((OlCategoryColor)categoryColour);
        }

        public EventColour.Palette GetColour(OlCategoryColor categoryColour) {
            EventColour.Palette gColour = null;

            SettingsStore.Calendar profile = Settings.Profile.InPlay();
            if (profile.ColourMaps.Count > 0) {
                KeyValuePair<String, String> kvp = profile.ColourMaps.FirstOrDefault(cm => OutlookOgcs.Calendar.Categories.OutlookColour(cm.Key) == categoryColour);
                if (kvp.Key != null) {
                    gColour = ColourPalette.ActivePalette.FirstOrDefault(ap => ap.Id == kvp.Value);
                    if (gColour != null) {
                        log.Debug("Colour mapping used: " + kvp.Key + " => " + kvp.Value + ":" + gColour.Name);
                        return gColour;
                    }
                }
            }
            //Algorithmic closest colour matching
            System.Drawing.Color color = OutlookOgcs.Categories.Map.RgbColour((OlCategoryColor)categoryColour);
            EventColour.Palette closest = ColourPalette.GetClosestColour(color);
            return (closest.Id == "0") ? EventColour.Palette.NullPalette : closest;
        }

        /// <summary>
        /// Get the first Outlook category colour from any defined against an Appointment's category(ies)
        /// </summary>
        /// <param name="aiCategories">The appointment categories assigned</param>
        /// <param name="categoryColour">The category colour identified</param>
        private void getOutlookCategoryColour(String aiCategories, ref OlCategoryColor? categoryColour) {
            if (!string.IsNullOrEmpty(aiCategories)) {
                log.Fine("Categories: " + aiCategories);
                try {
                    String category = aiCategories.Split(new[] { OutlookOgcs.Calendar.Categories.Delimiter }, StringSplitOptions.None).FirstOrDefault();
                    categoryColour = OutlookOgcs.Calendar.Categories.OutlookColour(category);
                } catch (System.Exception ex) {
                    log.Error("Failed determining colour for Event from AppointmentItem categories: " + aiCategories);
                    OGCSexception.Analyse(ex);
                }
            }
        }

        #region STATIC FUNCTIONS
        public static string signature(Event ev) {
            String signature = "";
            try {
                if (ev.RecurringEventId != null && ev.Status == "cancelled" && ev.OriginalStartTime != null) {
                    signature += (ev.Summary ?? "[cancelled]");
                    signature += ";" + ev.OriginalStartTime.SafeDateTime().ToPreciseString();
                } else {
                    signature += ev.Summary;
                    signature += ";" + ev.Start.SafeDateTime().ToPreciseString() + ";";
                    if (!(ev.EndTimeUnspecified != null && (Boolean)ev.EndTimeUnspecified)) {
                        signature += ev.End.SafeDateTime().ToPreciseString();
                    }
                }
            } catch {
                log.Warn("Failed to create signature: " + signature);
                log.Warn("This Event cannot be synced.");
                try { log.Warn("  ev.Summary: " + ev.Summary); } catch { }
                try { log.Warn("  ev.Start: " + (ev.Start == null ? "null!" : string.IsNullOrEmpty(ev.Start.Date) ? ev.Start.DateTime.ToString() : ev.Start.Date)); } catch { }
                try { log.Warn("  ev.End: " + (ev.End == null ? "null!" : string.IsNullOrEmpty(ev.End.Date) ? ev.End.DateTime.ToString() : ev.End.Date)); } catch { }
                try { log.Warn("  ev.Status: " + ev.Status ?? "null!"); } catch { }
                try { log.Warn("  ev.RecurringEventId: " + ev.RecurringEventId ?? "null"); } catch { }
                return "";
            }
            return signature.Trim();
        }

        public static Boolean SignaturesMatch(String sigEv, String sigAi) {
            //Use simple matching on start,end,subject to pair events
            SettingsStore.Calendar profile = Sync.Engine.Calendar.Instance.Profile;

            if (profile.Obfuscation.Enabled) {
                if (profile.Obfuscation.Direction.Id == Sync.Direction.OutlookToGoogle.Id)
                    sigAi = Obfuscate.ApplyRegex(sigAi, null, Sync.Direction.OutlookToGoogle);
                else
                    sigEv = Obfuscate.ApplyRegex(sigEv, null, Sync.Direction.GoogleToOutlook);
            }
            return (sigEv == sigAi);
        }

        public static void ExportToCSV(String action, String filename, List<Event> events) {
            if (!Settings.Instance.CreateCSVFiles) return;

            log.Debug("CSV export: " + action);

            String fullFilename = Path.Combine(Program.UserFilePath, filename);
            try {
                if (File.Exists(fullFilename)) {
                    String backupFilename = Path.Combine(Program.UserFilePath, Path.GetFileNameWithoutExtension(filename) + "-prev") + Path.GetExtension(filename);
                    if (File.Exists(backupFilename)) File.Delete(backupFilename);
                    File.Move(fullFilename, backupFilename);
                    log.Debug("Previous export renamed to " + backupFilename);
                }
            } catch (System.Exception ex) {
                OGCSexception.Analyse("Failed to backup previous CSV file.", ex);
            }

            Stream stream = null;
            TextWriter tw = null;
            try {
                try {
                    stream = new FileStream(Path.Combine(Program.UserFilePath, filename), FileMode.Create, FileAccess.Write);
                    tw = new StreamWriter(stream, Encoding.UTF8);
                } catch (System.Exception ex) {
                    Forms.Main.Instance.Console.Update("Failed to create CSV file '" + filename + "'.", Console.Markup.error);
                    OGCSexception.Analyse("Error opening file '" + filename + "' for writing.", ex);
                    return;
                }
                try {
                    String CSVheader = "Start Time,Finish Time,Subject,Location,Description,Privacy,FreeBusy,";
                    CSVheader += "Required Attendees,Optional Attendees,Reminder Set,Reminder Minutes,";
                    CSVheader += "Google EventID,Google CalendarID,";
                    CSVheader += "Outlook EntryID,Outlook GlobalID,Outlook CalendarID,";
                    CSVheader += "OGCS Modified,Force Save,API Limited";

                    tw.WriteLine(CSVheader);

                    foreach (Event ev in events) {
                        try {
                            tw.WriteLine(exportToCSV(ev));
                        } catch (System.Exception ex) {
                            Forms.Main.Instance.Console.Update("Failed to output following Google event to CSV:-<br/>" + GetEventSummary(ev), Console.Markup.warning);
                            OGCSexception.Analyse(ex, true);
                        }
                    }
                } catch (System.Exception ex) {
                    Forms.Main.Instance.Console.Update("Failed to output Google events to CSV.", Console.Markup.error);
                    OGCSexception.Analyse(ex);
                }
            } finally {
                if (tw != null) tw.Close();
                if (stream != null) stream.Close();
            }
            log.Fine("CSV export done.");
        }
        private static String exportToCSV(Event ev) {
            System.Text.StringBuilder csv = new System.Text.StringBuilder();

            csv.Append((ev.Start == null ? "null" : (string.IsNullOrEmpty(ev.Start.Date) ? ev.Start.DateTime.ToString() : ev.Start.Date)) + ",");
            csv.Append((ev.End == null ? "null" : (string.IsNullOrEmpty(ev.End.Date) ? ev.End.DateTime.ToString() : ev.End.Date)) + ",");
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
                foreach (Google.Apis.Calendar.v3.Data.EventAttendee ea in ev.Attendees) {
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
                        csv.Append("true," + er.Minutes + ",");
                        foundReminder = true;
                    }
                    break;
                }
            }
            if (!foundReminder) csv.Append(",,");

            csv.Append(ev.Id + "," + Sync.Engine.Calendar.Instance.Profile.UseGoogleCalendar.Id + ",");
            csv.Append((CustomProperty.Get(ev, CustomProperty.MetadataId.oEntryId) ?? "") + ",");
            csv.Append((CustomProperty.Get(ev, CustomProperty.MetadataId.oGlobalApptId) ?? "") + ",");
            csv.Append((CustomProperty.Get(ev, CustomProperty.MetadataId.oCalendarId) ?? "") + ",");
            csv.Append((CustomProperty.Get(ev, CustomProperty.MetadataId.ogcsModified) ?? "") + ",");
            csv.Append((CustomProperty.Get(ev, CustomProperty.MetadataId.forceSave) ?? "") + ",");
            csv.Append(CustomProperty.Get(ev, CustomProperty.MetadataId.apiLimitHit) ?? "");

            return csv.ToString();
        }

        /// <summary>
        /// Get the summary of an event.
        /// </summary>
        /// <param name="ai">The event</param>
        /// <param name="onlyIfNotVerbose">Only return if user doesn't have Verbose output on. Useful for indicating offending item during errors.</param>
        /// <returns></returns>
        public static string GetEventSummary(Event ev, Boolean onlyIfNotVerbose = false) {
            String eventSummary = "";
            if (!onlyIfNotVerbose || onlyIfNotVerbose && !Settings.Instance.VerboseOutput) {
                try {
                    if (ev.Start.DateTime != null) {
                        DateTime gDate = (DateTime)ev.Start.DateTime;
                        eventSummary += gDate.ToShortDateString() + " " + gDate.ToShortTimeString();
                    } else
                        eventSummary += DateTime.Parse(ev.Start.Date).ToShortDateString();
                    if ((ev.Recurrence != null && ev.RecurringEventId == null) || ev.RecurringEventId != null)
                        eventSummary += " (R)";
                    eventSummary += " => \"" + ev.Summary + "\"";
                    if (onlyIfNotVerbose) eventSummary += "<br/>";
                } catch {
                    log.Warn("Failed to create Event summary: " + eventSummary);
                    log.Warn("This Event cannot be synced.");
                    try { log.Warn("  ev.Summary: " + ev.Summary); } catch { }
                    try { log.Warn("  ev.Start: " + (ev.Start == null ? "null!" : string.IsNullOrEmpty(ev.Start.Date) ? ev.Start.DateTime.ToString() : ev.Start.Date)); } catch { }
                    try { log.Warn("  ev.End: " + (ev.End == null ? "null!" : string.IsNullOrEmpty(ev.End.Date) ? ev.End.DateTime.ToString() : ev.End.Date)); } catch { }
                    try { log.Warn("  ev.Status: " + ev.Status ?? "null!"); } catch { }
                    try { log.Warn("  ev.RecurringEventId: " + ev.RecurringEventId ?? "null"); } catch { }
                }
            }
            return eventSummary;
        }

        public static Google.Apis.Calendar.v3.Data.EventAttendee CreateAttendee(Recipient recipient, Boolean isOrganiser) {
            GoogleOgcs.EventAttendee ea = new GoogleOgcs.EventAttendee();
            log.Fine("Creating attendee " + recipient.Name);
            ea.DisplayName = recipient.Name;
            ea.Email = OutlookOgcs.Calendar.Instance.IOutlook.GetRecipientEmail(recipient);
            ea.Optional = (recipient.Type == (int)OlMeetingRecipientType.olOptional);
            if (isOrganiser) {
                //ea.Organizer = true; This is read-only. The best we can do is force them to have accepted the "invite"
                ea.ResponseStatus = "accepted";
                return ea;
            }
            switch (recipient.MeetingResponseStatus) {
                case OlResponseStatus.olResponseNone: ea.ResponseStatus = "needsAction"; break;
                case OlResponseStatus.olResponseAccepted: ea.ResponseStatus = "accepted"; break;
                case OlResponseStatus.olResponseDeclined: ea.ResponseStatus = "declined"; break;
                case OlResponseStatus.olResponseTentative: ea.ResponseStatus = "tentative"; break;
            }
            return ea;
        }

        public static ApiException HandleAPIlimits(ref Google.GoogleApiException ex, Event ev) {
            //https://developers.google.com/analytics/devguides/reporting/core/v3/coreErrors

            log.Fail(ex.Message);

            try {
                Telemetry.GA4Event.Event apiGa4Ev = new(Telemetry.GA4Event.Event.Name.error);
                apiGa4Ev.AddParameter("api_google_error", ex.Message);
                apiGa4Ev.AddParameter("code", ex.Error?.Code);
                apiGa4Ev.AddParameter("domain", ex.Error?.Errors?.First().Domain);
                apiGa4Ev.AddParameter("reason", ex.Error?.Errors?.First().Reason);
                apiGa4Ev.AddParameter("message", ex.Error?.Errors?.First().Message);
                apiGa4Ev.Send();
            } catch (System.Exception gaEx) {
                OGCSexception.Analyse(gaEx);
            }

            SettingsStore.Calendar profile = Sync.Engine.Calendar.Instance.Profile;
            if (profile.AddAttendees && ex.Message.Contains("Calendar usage limits exceeded. [403]") && ev != null) {
                //"Google.Apis.Requests.RequestError\r\nCalendar usage limits exceeded. [403]\r\nErrors [\r\n\tMessage[Calendar usage limits exceeded.] Location[ - ] Reason[quotaExceeded] Domain[usageLimits]\r\n]\r\n"
                //This happens because too many attendees have been added in a short period of time.
                //See https://support.google.com/a/answer/2905486?hl=en-uk&hlrm=en

                Forms.Main.Instance.Console.Update("You have added enough meeting attendees to have reached the Google API limit.<br/>" +
                    "Don't worry, this only lasts for an hour or two, but until then attendees will not be synced.", Console.Markup.warning);

                APIlimitReached_attendee = true;
                Settings.Instance.APIlimit_inEffect = true;
                Settings.Instance.APIlimit_lastHit = DateTime.Now;

                ev.Attendees = new List<Google.Apis.Calendar.v3.Data.EventAttendee>();
                return ApiException.justContinue;

            }

            if (ex.Error?.Code == 400 && ex.Error.Message.Contains("Invalid time zone definition") &&
                (ev.Start.TimeZone == "Europe/Kyiv" || ev.End.TimeZone == "Europe/Kyiv")) {

                //Mar-2023: Google has updated its definition to Kyiv, so this is not longer required...but helpful to leave here in case another time zone changes in the future
                log.Warn("Reverting to old IANA timezone definition: Europe/Kiev");
                TimezoneDB.Instance.RevertKyiv = true;
                ev.Start.TimeZone = ev.Start.TimeZone.Replace("Europe/Kyiv", "Europe/Kiev");
                ev.End.TimeZone = ev.End.TimeZone.Replace("Europe/Kyiv", "Europe/Kiev");
                return ApiException.backoffThenRetry;

            } else if (ex.Error?.Code == 401 && ex.Error.Message.Contains("Unauthorized")) {
                log.Debug("This error seems to be a new transient issue, so treating it with exponential backoff...");
                return ApiException.backoffThenRetry;

            } else if (ex.Error?.Code == 403 && ex.Error.Errors?.First().Domain == "usageLimits") {
                if (ex.Error.Errors.First().Reason == "rateLimitExceeded") {
                    if (ex.Message.Contains("limit 'Queries per minute'")) {
                        log.Fail(OGCSexception.FriendlyMessage(ex));
                        OGCSexception.LogAsFail(ref ex);
                        return ApiException.backoffThenRetry;

                    } else if (ex.Message.Contains("limit 'Queries per day'") || ex.Message.Contains("Daily Limit Exceeded")) {
                        log.Warn("Google's free Calendar quota has been exhausted! New quota comes into effect 08:00 GMT.");
                        Forms.Main.Instance.SyncNote(Forms.Main.SyncNotes.DailyQuotaExhaustedInfo, null);

                        //Delay next scheduled sync until after the new quota
                        if (profile.SyncInterval != 0) {
                            DateTime utcNow = DateTime.UtcNow;
                            DateTime quotaReset = utcNow.Date.AddHours(8).AddMinutes(utcNow.Minute);
                            if ((quotaReset - utcNow).Ticks < 0) quotaReset = quotaReset.AddDays(1);
                            int delayMins = (int)(quotaReset - utcNow).TotalMinutes;
                            profile.OgcsTimer.SetNextSync(delayMins, fromNow: true, calculateInterval: false);
                            Forms.Main.Instance.Console.Update("The next sync has been delayed by " + delayMins + " minutes, when new quota is available.", Console.Markup.warning);
                        }
                        return ApiException.freeAPIexhausted;

                    } else if (ex.Message.Contains("Rate Limit Exceeded")) {
                        if (Settings.Instance.Subscribed > DateTime.Now.AddYears(-1))
                            return ApiException.backoffThenRetry;
                        
                        log.Warn("Google's free Calendar quota is being exceeded!");
                        Forms.Main.Instance.SyncNote(Forms.Main.SyncNotes.QuotaExceededInfo, null);

                        //Delay next scheduled sync for an hour
                        if (profile.SyncInterval != 0) {
                            DateTime utcNow = DateTime.UtcNow;
                            DateTime nextSync = utcNow.AddMinutes(60 + new Random().Next(1,10));
                            int delayMins = (int)(nextSync - utcNow).TotalMinutes;
                            profile.OgcsTimer.SetNextSync(delayMins, fromNow: true, calculateInterval: false);
                            Forms.Main.Instance.Console.Update("The next sync has been delayed by " + delayMins + " minutes to let free quota rebuild.", Console.Markup.warning);
                        }
                        return ApiException.freeAPIexhausted;
                    }

                } else if (ex.Error.Errors.First().Reason == "dailyLimitExceededUnreg") {
                    if (ex.Message.Contains("Daily Limit for Unauthenticated Use Exceeded. Continued use requires signup.")) {
                        Forms.Main.Instance.Console.Update("You are not properly authenticated to Google.<br/>" +
                            "On the Settings > Google tab, please disconnect and re-authenticate your account.", Console.Markup.error);
                        ex.Data.Add("OGCS", "Unauthenticated access to Google account attempted. Authentication required.");
                        return ApiException.throwException;
                    }
                }

            } else if (ex.Error?.Code == 412 && ex.Error.Message.Contains("Precondition Failed")) {
                log.Warn("The Event has changed since it was last retrieved - attempting to force an overwrite.");
                EventsResource.UpdateRequest request;
                try {
                    request = GoogleOgcs.Calendar.Instance.Service.Events.Update(ev, profile.UseGoogleCalendar.Id, ev.Id);
                    request.ETagAction = Google.Apis.ETagAction.Ignore;
                    request.SendUpdates = EventsResource.UpdateRequest.SendUpdatesEnum.None;
                    ev = request.Execute();
                    log.Debug("Successfully forced save by ignoring eTag values.");
                } catch (System.Exception ex2) {
                    try {
                        OGCSexception.Analyse("Failed forcing save with ETagAction.Ignore", OGCSexception.LogAsFail(ex2));
                        log.Fine("Current eTag: " + ev.ETag);
                        log.Fine("Current Updated: " + ev.UpdatedRaw);
                        log.Fine("Current Sequence: " + ev.Sequence);
                        log.Debug("Refetching event from Google.");
                        Event remoteEv = GoogleOgcs.Calendar.Instance.GetCalendarEntry(ev.Id);
                        log.Fine("Remote eTag: " + remoteEv.ETag);
                        log.Fine("Remote Updated: " + remoteEv.UpdatedRaw);
                        log.Fine("Remote Sequence: " + remoteEv.Sequence);
                        log.Warn("Attempting trample of remote version...");
                        ev.ETag = remoteEv.ETag;
                        ev.Sequence = remoteEv.Sequence;
                        request = GoogleOgcs.Calendar.Instance.Service.Events.Update(ev, profile.UseGoogleCalendar.Id, ev.Id);
                        request.SendUpdates = EventsResource.UpdateRequest.SendUpdatesEnum.None;
                        ev = request.Execute();
                        log.Debug("Successful!");
                    } catch {
                        return ApiException.throwException;
                    }
                }
                return ApiException.justContinue;

            } else if (ex.Error?.Code == 500) {
                log.Fail(OGCSexception.FriendlyMessage(ex));
                OGCSexception.LogAsFail(ref ex);
                return ApiException.backoffThenRetry;
            }

            log.Warn("Unhandled API exception.");
            return ApiException.throwException;
        }

        public static Boolean? IsDefaultCalendar() {
            try {
                SettingsStore.Calendar profile = Sync.Engine.Calendar.Instance.Profile;
                if (!Settings.InstanceInitialiased() || (profile?.UseGoogleCalendar?.Id == null || string.IsNullOrEmpty(Settings.Instance.GaccountEmail)))
                return null;

                return profile.UseGoogleCalendar.Id == Settings.Instance.GaccountEmail;
            } catch (System.Exception ex) {
                OGCSexception.Analyse(ex);
                return null;
        }
        }
        #endregion

        /// <summary>
        /// This is solely for purposefully causing an error to assist when developing
        /// </summary>
        private void throwApiException() {
            Google.GoogleApiException ex = new Google.GoogleApiException("Service", "Rate Limit Exceeded");
            Google.Apis.Requests.SingleError err = new Google.Apis.Requests.SingleError { Domain = "usageLimits", Reason = "rateLimitExceeded" };
            ex.Error = new Google.Apis.Requests.RequestError { Errors = new List<Google.Apis.Requests.SingleError>(), Code = 403 };
            ex.Error.Errors.Add(err);
            throw ex;
        }
    }
}
