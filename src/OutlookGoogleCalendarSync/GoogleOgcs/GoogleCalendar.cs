﻿using Google.Apis.Calendar.v3;
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
                    instance = new GoogleOgcs.Calendar();
                    instance.Authenticator = new GoogleOgcs.Authenticator();
                    instance.Authenticator.GetAuthenticated();
                    if (instance.Authenticator.Authenticated)
                        instance.Authenticator.OgcsUserStatus();
                    else {
                        instance = null;
                        throw new ApplicationException("Google handshake failed.");
                    }
                }
                return instance;
            }
        }
        public Calendar() { }
        public GoogleOgcs.Authenticator Authenticator;
        private GoogleOgcs.EventColour colourPalette;
        public GoogleOgcs.EventColour ColourPalette {
            get {
                if (colourPalette == null) {
                    colourPalette = new EventColour();
                    colourPalette.Get();
                }
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
        private const int backoffLimit = 5;
        private enum apiException {
            justContinue,
            backoffThenRetry,
            freeAPIexhausted,
            throwException
        }
        private static Random random = new Random();
        public long MinDefaultReminder = long.MinValue;

        public EphemeralProperties EphemeralProperties = new EphemeralProperties();

        public List<GoogleCalendarListEntry> GetCalendars() {
            CalendarList request = null;
            int backoff = 0;
            while (backoff < backoffLimit) {
                try {
                    request = Service.CalendarList.List().Execute();
                    break;
                } catch (Google.GoogleApiException ex) {
                    switch (handleAPIlimits(ex, null)) {
                        case apiException.throwException: throw;
                        case apiException.freeAPIexhausted:
                            System.ApplicationException aex = new System.ApplicationException("Google's free daily Calendar quota has been exhausted! New quota comes into effect 08:00 GMT.", ex);
                            OGCSexception.LogAsFail(ref aex);
                            throw aex;
                        case apiException.backoffThenRetry:
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

            if (request != null) {
                List<GoogleCalendarListEntry> result = new List<GoogleCalendarListEntry>();
                foreach (CalendarListEntry cle in request.Items) {
                    result.Add(new GoogleCalendarListEntry(cle));
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
                    EventsResource.InstancesRequest ir = Service.Events.Instances(Settings.Instance.UseGoogleCalendar.Id, recurringEventId);
                    ir.ShowDeleted = true;
                    ir.PageToken = pageToken;
                    int backoff = 0;
                    while (backoff < backoffLimit) {
                        try {
                            request = ir.Execute();
                            log.Debug("Page " + pageNum + " received.");
                            break;
                        } catch (Google.GoogleApiException ex) {
                            switch (handleAPIlimits(ex, null)) {
                                case apiException.throwException: throw;
                                case apiException.freeAPIexhausted:
                                    System.ApplicationException aex = new System.ApplicationException("Google's free daily Calendar quota has been exhausted! New quota comes into effect 08:00 GMT.", ex);
                                    OGCSexception.LogAsFail(ref aex);
                                    throw aex;
                                case apiException.backoffThenRetry:
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

                    if (request != null) {
                        pageToken = request.NextPageToken;
                        pageNum++;
                        if (request.Items != null) result.AddRange(request.Items);
                    }
                } while (pageToken != null);
                return result;

            } catch (System.Exception ex) {
                Forms.Main.Instance.Console.Update("Failed to retrieve recurring events", Console.Markup.error);
                log.Error(ex.Message);
                return null;
            }
        }

        public Event GetCalendarEntry(String eventId) {
            Event request = null;

            try {
                log.Debug("Retrieving specific Event with ID " + eventId);
                EventsResource.GetRequest gr = Service.Events.Get(Settings.Instance.UseGoogleCalendar.Id, eventId);
                int backoff = 0;
                while (backoff < backoffLimit) {
                    try {
                        request = gr.Execute();
                        break;
                    } catch (Google.GoogleApiException ex) {
                        if (ex.Error.Code == 404) { //Not found
                            log.Fail("Could not find Google Event with specified ID " + eventId);
                            return null;
                        }
                        switch (handleAPIlimits(ex, null)) {
                            case apiException.throwException: throw;
                            case apiException.freeAPIexhausted:
                                System.ApplicationException aex = new System.ApplicationException("Google's free daily Calendar quota has been exhausted! New quota comes into effect 08:00 GMT.", ex);
                                OGCSexception.LogAsFail(ref aex);
                                throw aex;
                            case apiException.backoffThenRetry:
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

                if (request != null)
                    return request;
                else
                    throw new System.Exception("Returned null");
            } catch (System.Exception ex) {
                Forms.Main.Instance.Console.Update("Failed to retrieve Google event", Console.Markup.error);
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
                EventsResource.ListRequest lr = Service.Events.List(Settings.Instance.UseGoogleCalendar.Id);

                lr.TimeMin = from;
                lr.TimeMax = to;
                lr.PageToken = pageToken;
                lr.ShowDeleted = false;
                lr.SingleEvents = false;

                int backoff = 0;
                while (backoff < backoffLimit) {
                    try {
                        request = lr.Execute();
                        log.Debug("Page " + pageNum + " received.");
                        break;
                    } catch (Google.GoogleApiException ex) {
                        switch (handleAPIlimits(ex, null)) {
                            case apiException.throwException: throw;
                            case apiException.freeAPIexhausted:
                                System.ApplicationException aex = new System.ApplicationException("Google's free daily Calendar quota has been exhausted! New quota comes into effect 08:00 GMT.", ex);
                                OGCSexception.LogAsFail(ref aex);
                                throw aex;
                            case apiException.backoffThenRetry:
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
                        if (MessageBox.Show("Google event creation failed. Continue with synchronisation?", "Sync item failed", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                            continue;
                        else
                            throw new UserCancelledSyncException("User chose not to continue sync.");
                    }
                }

                Event createdEvent = new Event();
                try {
                    createdEvent = createCalendarEntry_save(newEvent, ai);
                } catch (System.Exception ex) {
                    Forms.Main.Instance.Console.UpdateWithError(OutlookOgcs.Calendar.GetEventSummary(ai, true) + "New event failed to save.", ex);
                    OGCSexception.Analyse(ex, true);
                    if (MessageBox.Show("New Google event failed to save. Continue with synchronisation?", "Sync item failed", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
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

            ev.Summary = Obfuscate.ApplyRegex(ai.Subject, Sync.Direction.OutlookToGoogle);
            if (Settings.Instance.AddDescription) {
                try {
                    ev.Description = ai.Body;
                } catch (System.Exception ex) {
                    if (OGCSexception.GetErrorCode(ex) == "0x80004004") {
                        Forms.Main.Instance.Console.Update("You do not have the rights to programmatically access Outlook appointment descriptions.<br/>" +
                            "It may be best to stop syncing the Description attribute.", Console.Markup.warning);
                    } else throw ex;
                }
            }
            if (Settings.Instance.AddLocation)
                ev.Location = ai.Location;
            ev.Visibility = getPrivacy(ai.Sensitivity, null);
            ev.Transparency = getAvailability(ai.BusyStatus, null);
            ev.ColorId = getColour(ai.Categories, null).Id;

            ev.Attendees = new List<Google.Apis.Calendar.v3.Data.EventAttendee>();
            if (Settings.Instance.AddAttendees && ai.Recipients.Count > 1 && !APIlimitReached_attendee) { //Don't add attendees if there's only 1 (me)
                if (ai.Recipients.Count >= 200) {
                    Forms.Main.Instance.Console.Update("Attendees will not be synced for this meeting as it has " +
                        "more than 200, which Google does not allow.", Console.Markup.warning);
                } else {
                    foreach (Microsoft.Office.Interop.Outlook.Recipient recipient in ai.Recipients) {
                        Google.Apis.Calendar.v3.Data.EventAttendee ea = GoogleOgcs.Calendar.CreateAttendee(recipient);
                        ev.Attendees.Add(ea);
                    }
                }
            }

            //Reminder alert
            ev.Reminders = new Event.RemindersData();
            if (Settings.Instance.AddReminders) {
                if (OutlookOgcs.Calendar.Instance.IsOKtoSyncReminder(ai)) {
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
            } else
                ev.Reminders.UseDefault = Settings.Instance.UseGoogleDefaultReminder;

            //Add the Outlook appointment ID into Google event
            CustomProperty.AddOutlookIDs(ref ev, ai);

            return ev;
        }

        private Event createCalendarEntry_save(Event ev, AppointmentItem ai) {
            if (Settings.Instance.SyncDirection == Sync.Direction.Bidirectional) {
                log.Debug("Saving timestamp when OGCS updated event.");
                CustomProperty.SetOGCSlastModified(ref ev);
            }
            if (Settings.Instance.APIlimit_inEffect) {
                CustomProperty.Add(ref ev, CustomProperty.MetadataId.apiLimitHit, "True");
            }

            Event createdEvent = new Event();
            int backoff = 0;
            while (backoff < backoffLimit) {
                try {
                    createdEvent = Service.Events.Insert(ev, Settings.Instance.UseGoogleCalendar.Id).Execute();
                    if (Settings.Instance.AddAttendees && Settings.Instance.APIlimit_inEffect) {
                        log.Info("API limit for attendee sync lifted :-)");
                        Settings.Instance.APIlimit_inEffect = false;
                    }
                    break;
                } catch (Google.GoogleApiException ex) {
                    switch (handleAPIlimits(ex, ev)) {
                        case apiException.throwException: throw;
                        case apiException.freeAPIexhausted:
                            System.ApplicationException aex = new System.ApplicationException("Google's free daily Calendar quota has been exhausted! New quota comes into effect 08:00 GMT.", ex);
                            OGCSexception.LogAsFail(ref aex);
                            throw aex;
                        case apiException.justContinue: break;
                        case apiException.backoffThenRetry:
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

            if (Settings.Instance.SyncDirection == Sync.Direction.Bidirectional || OutlookOgcs.CustomProperty.ExistsAny(ai)) {
                log.Debug("Storing the Google event IDs in Outlook appointment.");
                OutlookOgcs.CustomProperty.AddGoogleIDs(ref ai, createdEvent);
                ai.Save();
            }
            //DOS ourself by triggering API limit
            //for (int i = 1; i <= 30; i++) {
            //    Forms.Main.Instance.Console.Update("Add #" + i, verbose:true);
            //    Event result = service.Events.Insert(e, Settings.Instance.UseGoogleCalendar.Id).Fetch();
            //    System.Threading.Thread.Sleep(300);
            //    GoogleOgcs.Calendar.Instance.deleteCalendarEntry(result);
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
                        OGCSexception.Analyse(ex, true);
                        if (MessageBox.Show("Google event update failed. Continue with synchronisation?", "Sync item failed", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
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
                        if (MessageBox.Show("Updated Google event failed to save. Continue with synchronisation?", "Sync item failed", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                            continue;
                        else
                            throw new UserCancelledSyncException("User chose not to continue sync.");
                    }

                }

                //Have to do this *before* any dummy update, else all the exceptions inherit the updated timestamp of the parent recurring event
                Recurrence.UpdateGoogleExceptions(compare.Key, ev ?? compare.Value, eventExceptionCacheDirty);

                if (itemModified == 0) {
                    if (ev == null && CustomProperty.Exists(compare.Value, CustomProperty.MetadataId.forceSave))
                        ev = compare.Value;

                    if (ev == null || compare.Value.Updated > compare.Key.LastModificationTime) continue;

                    log.Debug("Doing a dummy update in order to update the last modified date of " +
                        (ev.RecurringEventId == null && ev.Recurrence != null ? "recurring master event" : "single instance"));
                    CustomProperty.SetOGCSlastModified(ref ev);
                    try {
                        UpdateCalendarEntry_save(ref ev);
                        entriesToBeCompared[compare.Key] = ev;
                    } catch (System.Exception ex) {
                        Forms.Main.Instance.Console.UpdateWithError("Updated event failed to save.", ex);
                        OGCSexception.Analyse(ex, true);
                        if (MessageBox.Show("Updated Google event failed to save. Continue with synchronisation?", "Sync item failed", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                            continue;
                        else
                            throw new UserCancelledSyncException("User chose not to continue sync.");
                    }
                }
            }
        }

        public Event UpdateCalendarEntry(AppointmentItem ai, Event ev, ref int itemModified, Boolean forceCompare = false) {
            if (!Settings.Instance.APIlimit_inEffect && CustomProperty.Exists(ev, CustomProperty.MetadataId.apiLimitHit)) {
                log.Fine("Back processing Event affected by attendee API limit.");
            } else {
                if (!(Sync.Engine.Instance.ManualForceCompare || forceCompare)) { //Needed if the exception has just been created, but now needs updating
                    if (Settings.Instance.SyncDirection != Sync.Direction.Bidirectional) {
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

            System.Text.StringBuilder sb = new System.Text.StringBuilder();
            sb.AppendLine(aiSummary);

            //Handle an event's all-day attribute being toggled
            DateTime evStart = ev.Start.DateTime ?? DateTime.Parse(ev.Start.Date);
            DateTime evEnd = ev.End.DateTime ?? DateTime.Parse(ev.End.Date);
            if (ai.AllDayEvent && ai.Start.TimeOfDay == new TimeSpan(0,0,0)) {
                ev.Start.DateTime = null;
                ev.End.DateTime = null;
                if (Sync.Engine.CompareAttribute("Start time", Sync.Direction.OutlookToGoogle, evStart, ai.Start.Date, sb, ref itemModified)) {
                    ev.Start.Date = ai.Start.ToString("yyyy-MM-dd");
                }
                if (Sync.Engine.CompareAttribute("End time", Sync.Direction.OutlookToGoogle, evEnd, ai.End.Date, sb, ref itemModified)) {
                    ev.End.Date = ai.End.ToString("yyyy-MM-dd");
                }
            } else {
                //Handle: Google = all-day; Outlook = not all day, but midnight values (so effectively all day!)
                if (ev.Start.DateTime == null && evStart == ai.Start && evEnd == ai.End) {
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
                                    if (Sync.Engine.CompareAttribute("Recurrence", Sync.Direction.OutlookToGoogle, rrule, oRrules.First(), sb, ref itemModified)) {
                                        ev.Recurrence[r] = oRrules.First();
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
                    log.Debug("Converting to recurring event.");
                    Sync.Engine.CompareAttribute("Recurrence", Sync.Direction.OutlookToGoogle, null, oRrules.First(), sb, ref itemModified);
                    ev.Recurrence = oRrules;
                }
            }

            //TimeZone
            if (ev.Start.DateTime != null) {
                String currentStartTZ = ev.Start.TimeZone;
                String currentEndTZ = ev.End.TimeZone;
                ev = OutlookOgcs.Calendar.Instance.IOutlook.IANAtimezone_set(ev, ai);
                Sync.Engine.CompareAttribute("Start Timezone", Sync.Direction.OutlookToGoogle, currentStartTZ, ev.Start.TimeZone, sb, ref itemModified);
                Sync.Engine.CompareAttribute("End Timezone", Sync.Direction.OutlookToGoogle, currentEndTZ, ev.End.TimeZone, sb, ref itemModified);
            }

            String subjectObfuscated = Obfuscate.ApplyRegex(ai.Subject, Sync.Direction.OutlookToGoogle);
            if (Sync.Engine.CompareAttribute("Subject", Sync.Direction.OutlookToGoogle, ev.Summary, subjectObfuscated, sb, ref itemModified)) {
                ev.Summary = subjectObfuscated;
            }
            if (Settings.Instance.AddDescription) {
                String outlookBody = ai.Body;
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

            if (Settings.Instance.AddLocation && Sync.Engine.CompareAttribute("Location", Sync.Direction.OutlookToGoogle, ev.Location, ai.Location, sb, ref itemModified))
                ev.Location = ai.Location;

            String gPrivacy = (ev.Visibility == null || ev.Visibility == "public") ? "default" : ev.Visibility;
            String oPrivacy = getPrivacy(ai.Sensitivity, gPrivacy);
            if (Sync.Engine.CompareAttribute("Privacy", Sync.Direction.OutlookToGoogle, gPrivacy, oPrivacy, sb, ref itemModified)) {
                ev.Visibility = oPrivacy;
            }

            String gFreeBusy = ev.Transparency ?? "opaque";
            String oFreeBusy = getAvailability(ai.BusyStatus, gFreeBusy);
            if (Sync.Engine.CompareAttribute("Free/Busy", Sync.Direction.OutlookToGoogle, gFreeBusy, oFreeBusy, sb, ref itemModified)) {
                ev.Transparency = oFreeBusy;
            }

            Palette gColour = this.ColourPalette.GetColour(ev.ColorId);
            Palette oColour = getColour(ai.Categories, gColour);
            if (Sync.Engine.CompareAttribute("Colour", Sync.Direction.OutlookToGoogle, gColour.HexValue, oColour.HexValue, sb, ref itemModified)) {
                ev.ColorId = oColour.Id;
            }

            if (Settings.Instance.AddAttendees && ai.Recipients.Count > 1 && !APIlimitReached_attendee) {
                if (ai.Recipients.Count >= 200) {
                    Forms.Main.Instance.Console.Update(OutlookOgcs.Calendar.GetEventSummary(ai) + "<br/>Attendees will not be synced for this meeting as it has " +
                        "more than 200, which Google does not allow.", Console.Markup.warning);
                    ev.Attendees = new List<Google.Apis.Calendar.v3.Data.EventAttendee>();
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

            //Reminders
            if (Settings.Instance.AddReminders) {
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

                } else { //no Google reminders set
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
                        Boolean newVal = OKtoSyncReminder ? Settings.Instance.UseGoogleDefaultReminder : false;

                        //Google bug?! For all-day events, default notifications are added as overrides and UseDefault=false
                        //Which means it keeps adding the default back in!! Let's stop that:
                        if (newVal && ev.Start.Date != null) {
                            log.Warn("Evading Google bug - not allowing default calendar notification to be (re?)set for all-day event.");
                            newVal = false;
                        }

                        if (Sync.Engine.CompareAttribute("Reminder Default", Sync.Direction.OutlookToGoogle, ev.Reminders.UseDefault.ToString(), newVal.ToString(), sb, ref itemModified)) {
                            ev.Reminders.UseDefault = newVal;
                        }
                    }
                }
            } else {
                if (Sync.Engine.CompareAttribute("Reminder Default", Sync.Direction.OutlookToGoogle, ev.Reminders.UseDefault.ToString(), Settings.Instance.UseGoogleDefaultReminder.ToString(), sb, ref itemModified))
                    ev.Reminders.UseDefault = Settings.Instance.UseGoogleDefaultReminder;
            }

            if (itemModified > 0) {
                Forms.Main.Instance.Console.FormatEventChanges(sb);
                Forms.Main.Instance.Console.Update(itemModified + " attributes updated.", Console.Markup.appointmentEnd, verbose: true, newLine: false);
                System.Windows.Forms.Application.DoEvents();
            }
            return ev;
        }

        public void UpdateCalendarEntry_save(ref Event ev) {
            if (Settings.Instance.SyncDirection == Sync.Direction.Bidirectional) {
                log.Debug("Saving timestamp when OGCS updated event.");
                CustomProperty.SetOGCSlastModified(ref ev);
            }
            if (Settings.Instance.APIlimit_inEffect)
                CustomProperty.Add(ref ev, CustomProperty.MetadataId.apiLimitHit, "True");
            else
                CustomProperty.Remove(ref ev, CustomProperty.MetadataId.apiLimitHit);

            CustomProperty.Remove(ref ev, CustomProperty.MetadataId.forceSave);

            int backoff = 0;
            while (backoff < backoffLimit) {
                try {
                    ev = Service.Events.Update(ev, Settings.Instance.UseGoogleCalendar.Id, ev.Id).Execute();
                    if (Settings.Instance.AddAttendees && Settings.Instance.APIlimit_inEffect) {
                        log.Info("API limit for attendee sync lifted :-)");
                        Settings.Instance.APIlimit_inEffect = false;
                    }
                    break;
                } catch (Google.GoogleApiException ex) {
                    switch (handleAPIlimits(ex, ev)) {
                        case apiException.throwException: throw;
                        case apiException.freeAPIexhausted:
                            System.ApplicationException aex = new System.ApplicationException("Google's free daily Calendar quota has been exhausted! New quota comes into effect 08:00 GMT.", ex);
                            OGCSexception.LogAsFail(ref aex);
                            throw aex;
                        case apiException.backoffThenRetry:
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
        #endregion

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
                    Forms.Main.Instance.Console.UpdateWithError(GoogleOgcs.Calendar.GetEventSummary(ev, true) + "Deleted event failed to remove.", ex);
                    OGCSexception.Analyse(ex, true);
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
            while (backoff < backoffLimit) {
                try {
                    string request = Service.Events.Delete(Settings.Instance.UseGoogleCalendar.Id, ev.Id).Execute();
                    break;
                } catch (Google.GoogleApiException ex) {
                    switch (handleAPIlimits(ex, ev)) {
                        case apiException.throwException: throw;
                        case apiException.freeAPIexhausted:
                            System.ApplicationException aex = new System.ApplicationException("Google's free daily Calendar quota has been exhausted! New quota comes into effect 08:00 GMT.", ex);
                            OGCSexception.LogAsFail(ref aex);
                            throw aex;
                        case apiException.backoffThenRetry:
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
        #endregion

        public void ReclaimOrphanCalendarEntries(ref List<Event> gEvents, ref List<AppointmentItem> oAppointments, Boolean neverDelete = false) {
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
                    foreach (AppointmentItem ai in oAppointments) {
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
                            } catch (System.Exception ex) {
                                log.Error("Failed to reclaim Event: " + GetEventSummary(ev));
                                log.Debug(ex.Message);
                                log.Debug("Event status: " + ev.Status);
                            }
                            break;
                        }
                    }
                }
            }
            log.Debug(unclaimedEvents.Count + " unclaimed.");
            if (!neverDelete && unclaimedEvents.Count > 0 &&
                (Settings.Instance.SyncDirection == Sync.Direction.OutlookToGoogle ||
                 Settings.Instance.SyncDirection == Sync.Direction.Bidirectional)) 
            {
                log.Info(unclaimedEvents.Count + " unclaimed orphan events found.");
                if (Settings.Instance.MergeItems || Settings.Instance.DisableDelete || Settings.Instance.ConfirmOnDelete) {
                    log.Info("These will be kept due to configuration settings.");
                } else if (Settings.Instance.SyncDirection == Sync.Direction.Bidirectional) {
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
                            log.UltraFine("Checking " + OutlookOgcs.Calendar.GetEventSummary(outlook[o]));

                            String compare_oID;
                            if (outlookIDmissing && compare_gEntryID.StartsWith("040000008200E00074C5B7101A82E008")) {
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
                            if (!log.IsFineEnabled()) {
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
                    if (!foundMatch && Settings.Instance.MergeItems &&
                        GoogleOgcs.CustomProperty.Get(google[g], CustomProperty.MetadataId.oCalendarId) != OutlookOgcs.Calendar.Instance.UseOutlookCalendar.EntryID)
                        google.Remove(google[g]);

                } else if (Settings.Instance.MergeItems) {
                    //Remove the non-Outlook item so it doesn't get deleted
                    google.Remove(google[g]);
                }
            }
            if (metadataEnhanced > 0) log.Info(metadataEnhanced + " item's metadata enhanced.");

            if (Settings.Instance.DisableDelete) {
                if (google.Count > 0)
                    Forms.Main.Instance.Console.Update(google.Count + " Google items would have been deleted, but you have deletions disabled.", Console.Markup.warning);
                google = new List<Event>();
            }
            if (Settings.Instance.SyncDirection == Sync.Direction.Bidirectional) {
                //Don't recreate any items that have been deleted in Google
                for (int o = outlook.Count - 1; o >= 0; o--) {
                    if (OutlookOgcs.CustomProperty.Exists(outlook[o], OutlookOgcs.CustomProperty.MetadataId.gEventID))
                        outlook.Remove(outlook[o]);
                }
                //Don't delete any items that aren't yet in Outlook or just created in Outlook during this sync
                for (int g = google.Count - 1; g >= 0; g--) {
                    if (!CustomProperty.Exists(google[g], CustomProperty.MetadataId.oEntryId) ||
                        google[g].Updated > Settings.Instance.LastSyncDate)
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

            if (CustomProperty.Exists(ev, CustomProperty.MetadataId.oGlobalApptId)) {
                String gCompareID = CustomProperty.Get(ev, CustomProperty.MetadataId.oGlobalApptId);
                String oGlobalID = OutlookOgcs.Calendar.Instance.IOutlook.GetGlobalApptID(ai);

                //For items copied from someone elses calendar, it appears the Global ID is generated for each access?! (Creation Time changes)
                //I guess the copied item doesn't really have its "own" ID. So, we'll just compare
                //the "data" section of the byte array, which "ensures uniqueness" and doesn't include ID creation time

                if ((oGlobalID.StartsWith(OutlookOgcs.Calendar.GlobalIdPattern) &&
                    gCompareID.StartsWith(OutlookOgcs.Calendar.GlobalIdPattern) &&
                    gCompareID.Substring(72) == oGlobalID.Substring(72))             //We've got bonafide Global IDs match
                    ||
                    (!oGlobalID.StartsWith(OutlookOgcs.Calendar.GlobalIdPattern) &&
                    !gCompareID.StartsWith(OutlookOgcs.Calendar.GlobalIdPattern) &&
                    gCompareID.Remove(gCompareID.Length - 16) == oGlobalID.Remove(oGlobalID.Length - 16))) //Or it's really a Entry ID (failsafe match)
                {
                    log.Fine("Comparing Outlook CalendarID");
                    gCompareID = CustomProperty.Get(ev, CustomProperty.MetadataId.oCalendarId);
                    if (gCompareID == OutlookOgcs.Calendar.Instance.UseOutlookCalendar.EntryID) {

                        //But...if an appointment is copied within ones own calendar, the DATA part is the same (only the creation time changes)!
                        //So now compare the Entry ID too.
                        log.Fine("Comparing Outlook EntryID");
                        gCompareID = CustomProperty.Get(ev, CustomProperty.MetadataId.oEntryId);
                        if (gCompareID == ai.EntryID) {
                            return true;
                        } else if (!string.IsNullOrEmpty(gCompareID) &&
                            gCompareID.Remove(gCompareID.Length-16) == ai.EntryID.Remove(ai.EntryID.Length-16)) 
                        {
                            //Worse still, both a locally copied item AND a rescheduled appointment by someone else 
                            //will have the MessageGlobalCounter bytes incremented (last 8-bytes)
                            //The former is identified by ExplorerWatcher adding a special flag
                            if (OutlookOgcs.CustomProperty.Get(ai, OutlookOgcs.CustomProperty.MetadataId.locallyCopied) == true.ToString()) {
                                log.Fine("This appointment was copied by the user. Incorrect match avoided.");
                                return false;
                            } else {
                                if (OutlookOgcs.Calendar.GetEventOrganizer(ai) != OutlookOgcs.Calendar.Instance.IOutlook.CurrentUserName()) {
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

                } else if (Settings.Instance.SyncDirection == Sync.Direction.Bidirectional &&
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
                if (Settings.Instance.MergeItems)
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

                            //Optional attendee
                            bool oOptional = (recipient.Type == (int)OlMeetingRecipientType.olOptional);
                            bool gOptional = (attendee.Optional == null) ? false : (bool)attendee.Optional;
                            String attendeeIdentifier = (attendee.DisplayName == null) ? ogcsAttendee.Email : attendee.DisplayName;
                            if (Sync.Engine.CompareAttribute("Attendee " + attendeeIdentifier + " - Optional Check",
                                Sync.Direction.OutlookToGoogle, gOptional, oOptional, sb, ref itemModified)) {
                                attendee.Optional = oOptional;
                            }

                            //Response
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
                        ev.Attendees.Add(GoogleOgcs.Calendar.CreateAttendee(recipient));
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

        public void GetSetting(string setting) {
            try {
                Service.Settings.Get(setting).ExecuteAsync();
            } catch { }
        }
        public void GetCalendarSettings() {
            if (!Settings.Instance.AddReminders || !Settings.Instance.UseGoogleDefaultReminder) return;
            try {
                CalendarListResource.GetRequest request = Service.CalendarList.Get(Settings.Instance.UseGoogleCalendar.Id);
                CalendarListEntry cal = request.Execute();
                if (cal.DefaultReminders.Count == 0)
                    this.MinDefaultReminder = long.MinValue;
                else
                    this.MinDefaultReminder = cal.DefaultReminders.Where(x => x.Method.Equals("popup")).OrderBy(x => x.Minutes.Value).First().Minutes.Value;
            } catch (System.Exception ex) {
                log.Error("Failed to get calendar settings.");
                log.Error(ex.Message);
            }
        }

        /// <summary>
        /// Determine Event's privacy setting
        /// </summary>
        /// <param name="oSensitivity">Outlook's current setting</param>
        /// <param name="gVisibility">Google's current setting</param>
        /// <param name="direction">Direction of sync</param>
        private String getPrivacy(OlSensitivity oSensitivity, String gVisibility) {
            if (!Settings.Instance.SetEntriesPrivate)
                return (oSensitivity == OlSensitivity.olNormal) ? "default" : "private";

            if (Settings.Instance.SyncDirection != Sync.Direction.Bidirectional) {
                return "private";
            } else {
                if (Settings.Instance.TargetCalendar == Sync.Direction.GoogleToOutlook) { //Privacy enforcement is in other direction
                    if (gVisibility == null)
                        return (oSensitivity == OlSensitivity.olNormal) ? "default" : "private";
                    else if (gVisibility == "private" && oSensitivity != OlSensitivity.olPrivate) {
                        log.Fine("Source of truth for privacy is already set private and target is NOT - so syncing this back.");
                        return "default";
                    } else
                        return gVisibility;
                } else {
                    if (!Settings.Instance.CreatedItemsOnly || (Settings.Instance.CreatedItemsOnly && gVisibility == null))
                        return "private";
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
            if (!Settings.Instance.SetEntriesAvailable)
                return (oBusyStatus == OlBusyStatus.olFree) ? "transparent" : "opaque";

            if (Settings.Instance.SyncDirection != Sync.Direction.Bidirectional) {
                return "transparent";
            } else {
                if (Settings.Instance.TargetCalendar == Sync.Direction.GoogleToOutlook) { //Availability enforcement is in other direction
                    if (gTransparency == null)
                        return (oBusyStatus == OlBusyStatus.olFree) ? "transparent" : "opaque";
                    else if (gTransparency == "transparent" && oBusyStatus != OlBusyStatus.olFree) {
                        log.Fine("Source of truth for Availability is already set available and target is NOT - so syncing this back.");
                        return "opaque";
                    } else
                        return gTransparency;
                } else {
                    if (!Settings.Instance.CreatedItemsOnly || (Settings.Instance.CreatedItemsOnly && gTransparency == null))
                        return "transparent";
                    else
                        return (oBusyStatus == OlBusyStatus.olFree) ? "transparent" : "opaque";
                }
            }
        }

        /// <summary>
        /// Get the Google palette colour from a list of Outlook categories
        /// </summary>
        /// <param name="aiCategories">The appointment item "categories" field</param>
        /// <returns>A match or a "null" Palette signifying no match</returns>
        private Palette getColour(String aiCategories, Palette gColour) {
            if (!Settings.Instance.AddColours && !Settings.Instance.SetEntriesColour) return Palette.NullPalette;

            OlCategoryColor? categoryColour = null;

            if (Settings.Instance.SetEntriesColour) {
                if (Settings.Instance.TargetCalendar == Sync.Direction.GoogleToOutlook) { //Colour forced to sync in other direction
                    if (gColour == null) //Creating item
                        getOutlookCategoryColour(aiCategories, ref categoryColour);
                    else return gColour;

                } else {
                    if (!Settings.Instance.CreatedItemsOnly || (Settings.Instance.CreatedItemsOnly && gColour == null)) {
                        categoryColour = OutlookOgcs.CategoryMap.Colours.Where(c => c.Key.ToString() == Settings.Instance.SetEntriesColourValue).FirstOrDefault().Key;
                        if (categoryColour == null) log.Warn("Could not convert category name '" + Settings.Instance.SetEntriesColourValue + "' into Outlook category type.");
                    } else return gColour;
                }

            } else {
                getOutlookCategoryColour(aiCategories, ref categoryColour);
            }
            if (categoryColour == null) return Palette.NullPalette;
            else {
                System.Drawing.Color color = OutlookOgcs.CategoryMap.RgbColour((OlCategoryColor)categoryColour);
                Palette closest = ColourPalette.GetClosestColour(color);
                return (closest.Id == "Custom") ? Palette.NullPalette : closest;
            }
        }
        private void getOutlookCategoryColour(String aiCategories, ref OlCategoryColor? categoryColour) {
            if (!string.IsNullOrEmpty(aiCategories)) {
                log.Fine("Categories: " + aiCategories);
                try {
                    String category = aiCategories.Split(new[] { OutlookOgcs.Calendar.Categories.Delimiter }, StringSplitOptions.None).FirstOrDefault();
                    categoryColour = OutlookOgcs.Calendar.Categories.OutlookColour(category);
                    if (categoryColour == null) log.Warn("Could not convert category name '" + category + "' into Outlook category type.");
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
                    signature += ";" + (ev.OriginalStartTime.DateTime ?? DateTime.Parse(ev.OriginalStartTime.Date)).ToPreciseString();
                } else {
                    signature += ev.Summary;
                    signature += ";" + (ev.Start.DateTime ?? DateTime.Parse(ev.Start.Date)).ToPreciseString() + ";";
                    if (!(ev.EndTimeUnspecified != null && (Boolean)ev.EndTimeUnspecified)) {
                        signature += (ev.End.DateTime ?? DateTime.Parse(ev.End.Date)).ToPreciseString();
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
            if (Settings.Instance.Obfuscation.Enabled) {
                if (Settings.Instance.Obfuscation.Direction.Id == Sync.Direction.OutlookToGoogle.Id)
                    sigAi = Obfuscate.ApplyRegex(sigAi, Sync.Direction.OutlookToGoogle);
                else
                    sigEv = Obfuscate.ApplyRegex(sigEv, Sync.Direction.GoogleToOutlook);
            }
            return (sigEv == sigAi);
        }

        public static void ExportToCSV(String action, String filename, List<Event> events) {
            if (!Settings.Instance.CreateCSVFiles) return;

            log.Debug("CSV export: " + action);

            TextWriter tw;
            try {
                tw = new StreamWriter(Path.Combine(Program.UserFilePath, filename));
            } catch (System.Exception ex) {
                Forms.Main.Instance.Console.Update("Failed to create CSV file '" + filename + "'.", Console.Markup.error);
                log.Error("Error opening file '" + filename + "' for writing.");
                log.Error(ex.Message);
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
                        OGCSexception.Analyse(ex);
                    }
                }
            } catch (System.Exception ex) {
                Forms.Main.Instance.Console.Update("Failed to output Google events to CSV.", Console.Markup.error);
                OGCSexception.Analyse(ex);
            } finally {
                if (tw != null) tw.Close();
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

            csv.Append(ev.Id + "," + Settings.Instance.UseGoogleCalendar.Id + ",");
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

        public static Google.Apis.Calendar.v3.Data.EventAttendee CreateAttendee(Recipient recipient) {
            GoogleOgcs.EventAttendee ea = new GoogleOgcs.EventAttendee();
            log.Fine("Creating attendee " + recipient.Name);
            ea.DisplayName = recipient.Name;
            ea.Email = OutlookOgcs.Calendar.Instance.IOutlook.GetRecipientEmail(recipient);
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

        private static apiException handleAPIlimits(Google.GoogleApiException ex, Event ev) {
            //https://developers.google.com/analytics/devguides/reporting/core/v3/coreErrors

            if (Settings.Instance.AddAttendees && ex.Message.Contains("Calendar usage limits exceeded. [403]") && ev != null) {
                //"Google.Apis.Requests.RequestError\r\nCalendar usage limits exceeded. [403]\r\nErrors [\r\n\tMessage[Calendar usage limits exceeded.] Location[ - ] Reason[quotaExceeded] Domain[usageLimits]\r\n]\r\n"
                //This happens because too many attendees have been added in a short period of time.
                //See https://support.google.com/a/answer/2905486?hl=en-uk&hlrm=en

                Forms.Main.Instance.Console.Update("You have added enough meeting attendees to have reached the Google API limit.<br/>" +
                    "Don't worry, this only lasts for an hour or two, but until then attendees will not be synced.", Console.Markup.warning);

                APIlimitReached_attendee = true;
                Settings.Instance.APIlimit_inEffect = true;
                Settings.Instance.APIlimit_lastHit = DateTime.Now;

                ev.Attendees = new List<Google.Apis.Calendar.v3.Data.EventAttendee>();
                return apiException.justContinue;

            } else if (ex.Message.Contains("Rate Limit Exceeded")) {
                return apiException.backoffThenRetry;

            } else if (ex.Message.Contains("Daily Limit Exceeded")) {
                log.Warn(ex.Message);
                log.Warn("Google's free Calendar quota has been exhausted! New quota comes into effect 08:00 GMT.");
                Forms.Main.Instance.SyncNote(Forms.Main.SyncNotes.QuotaExhaustedInfo, null);

                //Delay next scheduled sync until after the new quota
                DateTime utcNow = DateTime.UtcNow;
                DateTime quotaReset = utcNow.Date.AddHours(8).AddMinutes(utcNow.Minute);
                if ((quotaReset - utcNow).Ticks < 0) quotaReset = quotaReset.AddDays(1);
                Sync.Engine.Instance.OgcsTimer.SetNextSync((int)(quotaReset - DateTime.Now).TotalMinutes, fromNow: true);

                return apiException.freeAPIexhausted;

            } else if (ex.Message.Contains("Daily Limit for Unauthenticated Use Exceeded. Continued use requires signup. [403]")) {
                log.Warn(ex.Message);
                Forms.Main.Instance.Console.Update("You are not properly authenticated to Google.<br/>" +
                    "On the Settings > Google tab, please disconnect and re-authenticate your account.", Console.Markup.error);
                ex.Data.Add("OGCS", "Unauthenticated access to Google account attempted. Authentication required.");
                return apiException.throwException;

            } else if (ex.Error.Code == 401 && ex.Error.Message.Contains("Unauthorized")) {
                log.Warn(ex.Message);
                log.Debug("This error seems to be a new transient issue, so treating it with exponential backoff...");
                return apiException.backoffThenRetry;

            } else {
                log.Warn(ex.Message);
                return apiException.throwException;
            }
        }
        #endregion
    }
}
