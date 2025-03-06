using Google.Apis.Calendar.v3;
using Google.Apis.Calendar.v3.Data;
using log4net;
//using Microsoft.Office.Interop.Outlook;
using OutlookGoogleCalendarSync.Extensions;
using OutlookGoogleCalendarSync.GraphExtension;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using GcalData = Google.Apis.Calendar.v3.Data;
using Ogcs = OutlookGoogleCalendarSync;

namespace OutlookGoogleCalendarSync.Google.Graph {
    /// <summary>
    /// Description of Ogcs.Google.Calendar.
    /// </summary>
    public class Calendar {
        private static readonly ILog log = LogManager.GetLogger(typeof(Calendar));

        /*private static Calendar instance;
        public static Boolean IsInstanceNull { get { return instance == null; } }
        public static Calendar Instance {
            get {
                if (instance == null) {
                    instance = new Ogcs.Google.Calendar {
                        Authenticator = new Ogcs.Google.Authenticator()
                    };
                    instance.Authenticator.GetAuthenticated();
                    if (instance.Authenticator.Authenticated) {
                        instance.Authenticator.OgcsUserStatus();
                        _ = instance.ColourPalette;
                    } else {
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
        public Ogcs.Google.Authenticator Authenticator;

        /// <summary>Google Events excluded through user config <Event.Id, Appt.EntryId></summary>
        public Dictionary<String, String> ExcludedByColour { get; private set; }

        private Ogcs.Google.EventColour colourPalette;

        public static Boolean IsColourPaletteNull { get { return instance?.colourPalette == null; } }
        public Ogcs.Google.EventColour ColourPalette {
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
                    Authenticator = new Ogcs.Google.Authenticator();
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

        private static String[] permittedEventTypes = new String[] { "default", "focusTime", "outOfOffice" }; //Excluding workingLocation
        
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
                    } catch (global::Google.GoogleApiException ex) {
                        switch (HandleAPIlimits(ref ex, null)) {
                            case ApiException.throwException: throw;
                            case ApiException.freeAPIexhausted:
                                Ogcs.Exception.LogAsFail(ref ex);
                                Ogcs.Exception.Analyse(ex);
                                System.ApplicationException aex = new System.ApplicationException(SubscriptionInvite, ex);
                                Ogcs.Exception.LogAsFail(ref aex);
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

            this.calendarList = result;
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
                        } catch (global::Google.GoogleApiException ex) {
                            switch (HandleAPIlimits(ref ex, null)) {
                                case ApiException.throwException: throw;
                                case ApiException.freeAPIexhausted:
                                    Ogcs.Exception.LogAsFail(ref ex);
                                    Ogcs.Exception.Analyse(ex);
                                    System.ApplicationException aex = new System.ApplicationException(SubscriptionInvite, ex);
                                    Ogcs.Exception.LogAsFail(ref aex);
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
                Forms.Main.Instance.Console.UpdateWithError("Failed to retrieve recurring events.", Ogcs.Exception.LogAsFail(ex));
                ex.Analyse("recurringEventId: " + recurringEventId);
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
                        if (!permittedEventTypes.Contains(request.EventType)) {
                            log.Warn($"Non-consumer version of EventType '{request.EventType}' found - excluding.");
                            return null;
                        }
                        break;
                    } catch (global::Google.GoogleApiException ex) {
                        if (ex.Error.Code == 404) { //Not found
                            log.Fail("Could not find Google Event with specified ID " + eventId);
                            return null;
                        }
                        switch (HandleAPIlimits(ref ex, null)) {
                            case ApiException.throwException: throw;
                            case ApiException.freeAPIexhausted:
                                Ogcs.Exception.LogAsFail(ref ex);
                                Ogcs.Exception.Analyse(ex);
                                System.ApplicationException aex = new System.ApplicationException(SubscriptionInvite, ex);
                                Ogcs.Exception.LogAsFail(ref aex);
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

        public List<Event> GetCalendarEntriesInRange(System.DateTime from, System.DateTime to) {
            List<Event> result = new List<Event>();
            ExcludedByColour = new Dictionary<String, String>();
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
                lr.EventTypes = permittedEventTypes;

                int backoff = 0;
                while (backoff < BackoffLimit) {
                    try {
                        request = lr.Execute();
                        log.Debug("Page " + pageNum + " received.");
                        break;
                    } catch (global::Google.GoogleApiException ex) {
                        switch (HandleAPIlimits(ref ex, null)) {
                            case ApiException.throwException: throw;
                            case ApiException.freeAPIexhausted:
                                Ogcs.Exception.LogAsFail(ref ex);
                                Ogcs.Exception.Analyse(ex);
                                System.ApplicationException aex = new System.ApplicationException(SubscriptionInvite, ex);
                                Ogcs.Exception.LogAsFail(ref aex);
                                throw aex;
                            case ApiException.backoffThenRetry:
                                backoff++;
                                if (backoff == BackoffLimit) {
                                    log.Error("API limit backoff was not successful. Retrieve failed.");
                                    aex = new System.ApplicationException(SubscriptionInvite, ex);
                                    Ogcs.Exception.LogAsFail(ref aex);
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

            List<Event> availability = new();
            List<Event> allDays = new();
            List<Event> privacy = new();
            List<Event> declined = new();
            List<Event> subject = new();
            List<Event> goals = new();
            List<Event> colour = new();

            //Colours
            if (profile.ColoursRestrictBy == SettingsStore.Calendar.RestrictBy.Include) {
                colour = result.Where(ev => string.IsNullOrEmpty(ev.RecurringEventId) &&
                    (profile.Colours.Count() == 0 || (String.IsNullOrEmpty(ev.ColorId) && !profile.Colours.Contains("<Default calendar colour>")) ||
                        !String.IsNullOrEmpty(ev.ColorId) && !profile.Colours.Contains(EventColour.Palette.GetColourName(ev.ColorId))
                    )
                ).ToList();
            } else if (profile.ColoursRestrictBy == SettingsStore.Calendar.RestrictBy.Exclude) {
                colour = result.Where(ev => string.IsNullOrEmpty(ev.RecurringEventId) &&
                    (profile.Colours.Count() > 0 && (String.IsNullOrEmpty(ev.ColorId) && profile.Colours.Contains("<Default calendar colour>")) ||
                        !String.IsNullOrEmpty(ev.ColorId) && profile.Colours.Contains(EventColour.Palette.GetColourName(ev.ColorId))
                    )
                ).ToList();
            }
            if (colour.Count > 0) {
                log.Debug(colour.Count + " Google items contain a colour that is filtered out.");
            }
            foreach (Event ev in colour) {
                ExcludedByColour.Add(ev.Id, CustomProperty.Get(ev, CustomProperty.MetadataId.oEntryId));
            }
            result = result.Except(colour).ToList();

            //Availability, Privacy
            if (profile.SyncDirection.Id != Sync.Direction.OutlookToGoogle.Id) { //Sync direction means G->O will delete previously synced all-days
                if (profile.ExcludeFree) {
                    availability = result.Where(ev => String.IsNullOrEmpty(ev.RecurringEventId) && ev.Transparency == "transparent").ToList();
                    if (availability.Count > 0) {
                        log.Debug(availability.Count + " Google Free items excluded.");
                        result = result.Except(availability).ToList();
                    }
                }
                if (profile.ExcludeAllDays) {
                    allDays = result.Where(ev => String.IsNullOrEmpty(ev.RecurringEventId) && ev.AllDayEvent(true) && (profile.ExcludeFreeAllDays ? ev.Transparency == "transparent" : true)).ToList();
                    if (allDays.Count > 0) {
                        log.Debug(allDays.Count + " Google all-day items excluded.");
                        result = result.Except(allDays).ToList();
                    }
                }
                if (profile.ExcludePrivate) {
                    privacy = result.Where(ev => String.IsNullOrEmpty(ev.RecurringEventId) && ev.Visibility == "private").ToList();
                    if (privacy.Count > 0) {
                        log.Debug(privacy.Count + " Google Private items excluded.");
                        result = result.Except(privacy).ToList();
                    }
                }
                if (profile.ExcludeSubject && !String.IsNullOrEmpty(profile.ExcludeSubjectText)) {
                    Regex rgx = new Regex(profile.ExcludeSubjectText, RegexOptions.IgnoreCase);
                    subject = result.Where(ev => String.IsNullOrEmpty(ev.RecurringEventId) && rgx.IsMatch(ev.Summary ?? "")).ToList();
                    if (subject.Count > 0) {
                        log.Debug(subject.Count + " Google items excluded with Subject containing '" + profile.ExcludeSubjectText + "'");
                        result = result.Except(subject).ToList();
                    }
                }
            }

            //Invitation
            if (profile.ExcludeDeclinedInvites) {
                declined = result.Where(ev => string.IsNullOrEmpty(ev.RecurringEventId) && ev.Attendees != null && ev.Attendees.Count(a => a.Self == true && a.ResponseStatus == "declined") == 1).ToList();
                if (declined.Count > 0) {
                    log.Debug(declined.Count + " Google Event invites have been declined and will be excluded.");
                    result = result.Except(declined).ToList();
                }
            }

            //Goals
            if ((IsDefaultCalendar() ?? true) && profile.ExcludeGoals) {
                goals = result.Where(ev =>
                    !string.IsNullOrEmpty(ev.Description) && ev.Description.Contains("This event was added from Goals in Google Calendar.") &&
                    ev.Organizer != null && ev.Organizer.Email == "unknownorganizer@calendar.google.com" && ev.Organizer.DisplayName == "Google Calendar").ToList();
                if (goals.Count > 0) {
                    log.Debug(goals.Count + " Google Events are Goals and will be excluded.");
                    result = result.Except(goals).ToList();
                }
            }

            if (profile.SyncDirection.Id == Sync.Direction.Bidirectional.Id) {
                List<Event> allExcluded = colour.Concat(availability).Concat(allDays).Concat(privacy).Concat(subject).Concat(declined).Concat(goals).ToList();
                for (int g = 0; g < allExcluded.Count(); g++) {
                    Event ev = allExcluded[g];
                    if (CustomProperty.ExistAnyOutlookIDs(ev)) {
                        log.Debug("Previously synced Google item is now excluded. Removing Outlook metadata.");
                        CustomProperty.RemoveOutlookIDs(ref ev);
                        UpdateCalendarEntry_save(ref ev);
                    }
                }
            }

            return result;
        }

        /// <summary>
        /// Update just the attributes provided.
        /// </summary>
        /// <param name="ev">The Event with the subset of attribute values.</param>
        /// <param name="patchId">The Event ID to patch. If not passed, the attribute must be included in the ev parameter object.</param>
        /// <returns>The complete target Event after patch applied.</returns>
        private Event patchEvent(Event ev, String patchId = null) {
            int backoff = 0;
            while (backoff < BackoffLimit) {
                try {
                    EventsResource.PatchRequest pr = Service.Events.Patch(ev, Sync.Engine.Calendar.Instance.Profile.UseGoogleCalendar.Id, patchId ?? ev.Id);
                    if (String.IsNullOrEmpty(ev.ConferenceData?.ConferenceSolution?.Name)) {
                        log.Debug("Updating conference data.");
                        pr.ConferenceDataVersion = 1;
                    }
                    return pr.Execute();

                } catch (global::Google.GoogleApiException ex) {
                    switch (HandleAPIlimits(ref ex, ev)) {
                        case ApiException.throwException: throw;
                        case ApiException.freeAPIexhausted:
                            Ogcs.Exception.LogAsFail(ref ex);
                            Ogcs.Exception.Analyse(ex);
                            System.ApplicationException aex = new System.ApplicationException(SubscriptionInvite, ex);
                            Ogcs.Exception.LogAsFail(ref aex);
                            throw aex;
                        case ApiException.justContinue: break;
                        case ApiException.backoffThenRetry:
                            backoff++;
                            if (backoff == BackoffLimit) {
                                log.Error("API limit backoff was not successful. Patch failed.");
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
            return null;
        }*/
        #region Create
        public static void CreateCalendarEntries(List<Microsoft.Graph.Event> appointments) {
            foreach (Microsoft.Graph.Event ai in appointments) {
                if (Sync.Engine.Instance.CancellationPending) return;

                GcalData.Event newEvent = new();
                try {
                    newEvent = createCalendarEntry(ai);
                } catch (System.Exception ex) {
                    if (ex is ApplicationException) {
                        String summary = Outlook.Graph.Calendar.GetEventSummary("Event creation skipped.<br/>" + ex.Message, ai, out String anonSummary);
                        Forms.Main.Instance.Console.Update(summary, anonSummary, Console.Markup.warning);
                        if (ex.InnerException is global::Google.GoogleApiException) break;
                        continue;
                    } else {
                        String summary = Outlook.Graph.Calendar.GetEventSummary("Event creation failed.", ai, out String anonSummary);
                        Forms.Main.Instance.Console.UpdateWithError(summary, ex, logEntry: anonSummary);
                        Ogcs.Exception.Analyse(ex, true);
                        if (Ogcs.Extensions.MessageBox.Show("Google event creation failed. Continue with synchronisation?", "Sync item failed", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
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
                    Forms.Main.Instance.Console.UpdateWithError(Outlook.Graph.Calendar.GetEventSummary("New event failed to save.", ai, out String anonSummary, true), ex, logEntry: anonSummary);
                    Ogcs.Exception.Analyse(ex, true);
                    if (Ogcs.Extensions.MessageBox.Show("New Google event failed to save. Continue with synchronisation?", "Sync item failed", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                        continue;
                    else
                        throw new UserCancelledSyncException("User chose not to continue sync.");
                }

                //Recurrence.CreateGoogleExceptions(ai, ref createdEvent);
            }
        }

        private static GcalData.Event createCalendarEntry(Microsoft.Graph.Event ai) {
            SettingsStore.Calendar profile = Sync.Engine.Calendar.Instance.Profile;

            string itemSummary = Outlook.Graph.Calendar.GetEventSummary(ai, out String anonSummary);
            log.Debug("Processing >> " + (anonSummary ?? itemSummary));
            Forms.Main.Instance.Console.Update(itemSummary, anonSummary, Console.Markup.calendar, verbose: true);

            GcalData.Event ev = new();

            ev.Recurrence = Recurrence.BuildGooglePattern(ai, ev);
            ev.Start = new EventDateTime();
            ev.End = new EventDateTime();

            if (ai.IsAllDay ?? false) {
                ev.Start.Date = ai.Start.SafeDateTime().ToString("yyyy-MM-dd");
                ev.End.Date = ai.End.SafeDateTime().ToString("yyyy-MM-dd");
            } else {
                ev.Start.DateTimeRaw = ai.Start.SafeDateTime().ToPreciseString();
                String startTimeZone = string.IsNullOrEmpty(ai.OriginalStartTimeZone) ? "UTC" : ai.OriginalStartTimeZone;
                ev.Start.TimeZone = TimezoneDB.IANAtimezone(startTimeZone, startTimeZone);

                ev.End.DateTimeRaw = ai.End.SafeDateTime().ToPreciseString();
                String endTimeZone = string.IsNullOrEmpty(ai.OriginalEndTimeZone) ? "UTC" : ai.OriginalEndTimeZone;
                ev.End.TimeZone = startTimeZone == endTimeZone ? ev.Start.TimeZone : TimezoneDB.IANAtimezone(endTimeZone, endTimeZone);
            }
            
            ev.Summary = Obfuscate.ApplyRegex(Obfuscate.Property.Subject, ai.Subject, null, Sync.Direction.OutlookToGoogle);
            if (profile.AddDescription)
                ev.Description = Obfuscate.ApplyRegex(Obfuscate.Property.Description, ai.Body.BodyInnerHtml(), null, Sync.Direction.OutlookToGoogle);
            if (profile.AddLocation)
                ev.Location = Obfuscate.ApplyRegex(Obfuscate.Property.Location, ai.Location.DisplayName, null, Sync.Direction.OutlookToGoogle);
            ev.Visibility = getPrivacy(ai.Sensitivity, null);
            ev.Transparency = getAvailability(ai.ShowAs, null);
            //ev.ColorId = getColour(ai.Categories, null)?.Id ?? EventColour.Palette.NullPalette.Id;

            ev.Attendees = new List<GcalData.EventAttendee>();
            if (profile.AddAttendees && !Ogcs.Google.Calendar.APIlimitReached_attendee) {
                if (ai.Attendees.Count() > profile.MaxAttendees) {
                    log.Warn("This Outlook appointment has " + ai.Attendees.Count() + " attendees, more than the user configured maximum.");
                    if (ai.Attendees.Count() >= 200) {
                        Forms.Main.Instance.Console.Update("Attendees will not be synced for this meeting as it has " +
                            "more than 200, which Google does not allow.", Console.Markup.warning);
                    }
                } else {
                    foreach (Microsoft.Graph.Attendee recipient in ai.Attendees) {
                        if (Settings.Instance.GaccountEmail.ToLower() == recipient.EmailAddress.Address) continue;

                        GcalData.EventAttendee ea = CreateAttendee(recipient, ai.Organizer.EmailAddress == recipient.EmailAddress);
                        ev.Attendees.Add(ea);
                    }
                }
            }

            //Reminder alert
            ev.Reminders = new Event.RemindersData();
            if (profile.AddReminders) {
                if (Outlook.Graph.Calendar.Instance.IsOKtoSyncReminder(ai)) {
                    if (ai.IsReminderOn ?? false) {
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

        private static Event createCalendarEntry_save(Event ev, Microsoft.Graph.Event ai) {
            SettingsStore.Calendar profile = Sync.Engine.Calendar.Instance.Profile;

            if (profile.SyncDirection.Id == Sync.Direction.Bidirectional.Id) {
                log.Debug("Saving timestamp when OGCS created event.");
                Google.CustomProperty.SetOGCSlastModified(ref ev);
            }
            if (Settings.Instance.APIlimit_inEffect) {
                Google.CustomProperty.Add(ref ev, Google.CustomProperty.MetadataId.apiLimitHit, "True");
            }

            Event createdEvent = new Event();
            int backoff = 0;
            while (backoff < Google.Calendar.BackoffLimit) {
                try {
                    EventsResource.InsertRequest request = Google.Calendar.Instance.Service.Events.Insert(ev, profile.UseGoogleCalendar.Id);
                    request.SendUpdates = EventsResource.InsertRequest.SendUpdatesEnum.None;
                    createdEvent = request.Execute();
                    if (profile.AddAttendees && Settings.Instance.APIlimit_inEffect) {
                        log.Info("API limit for attendee sync lifted :-)");
                        Settings.Instance.APIlimit_inEffect = false;
                    }
                    break;
                } catch (global::Google.GoogleApiException ex) {
                    switch (Google.Calendar.HandleAPIlimits(ref ex, ev)) {
                        case Google.Calendar.ApiException.throwException: throw;
                        case Google.Calendar.ApiException.freeAPIexhausted:
                            Ogcs.Exception.LogAsFail(ref ex);
                            Ogcs.Exception.Analyse(ex);
                            System.ApplicationException aex = new System.ApplicationException(Google.Calendar.Instance.SubscriptionInvite, ex);
                            Ogcs.Exception.LogAsFail(ref aex);
                            throw aex;
                        case Google.Calendar.ApiException.justContinue: break;
                        case Google.Calendar.ApiException.backoffThenRetry:
                            backoff++;
                            if (backoff == Google.Calendar.BackoffLimit) {
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

            if (!String.IsNullOrEmpty(createdEvent.Id) && (profile.SyncDirection.Id == Sync.Direction.Bidirectional.Id || Outlook.Graph.CustomProperty.ExistAnyGoogleIDs(ai))) {
                log.Debug("Storing the Google event IDs in Outlook appointment.");
                Outlook.Graph.CustomProperty.AddGoogleIDs(ref ai, createdEvent);
                Outlook.Graph.CustomProperty.SetOGCSlastModified(ref ai);
                Microsoft.Graph.Event aiPatch = new() { Id = ai.Id, Extensions = ai.Extensions };
                Outlook.Graph.Calendar.Instance.UpdateCalendarEntry_save(ref aiPatch);
                ai = aiPatch;
            }

            /*if (profile.AddGMeet && Outlook.GMeet.BodyHasGmeetUrl(ai)) {
                log.Info("Adding GMeet conference details.");
                String outlookGMeet = Outlook.GMeet.RgxGmeetUrl().Match(ai.Body).Value;
                Ogcs.Google.GMeet.GoogleMeet(createdEvent, outlookGMeet);
                createdEvent = patchEvent(createdEvent) ?? createdEvent;
                log.Fine("Conference data added.");
            }*/

            #region DOS ourself by triggering API limit
            //for (int i = 1; i <= 100; i++) {
            //    Forms.Main.Instance.Console.Update("Add #" + i, verbose: true);
            //    Event result = service.Events.Insert(ev, Settings.Instance.UseGoogleCalendar.Id).Execute();
            //    System.Threading.Thread.Sleep(300);
            //    Ogcs.Google.Calendar.Instance.deleteCalendarEntry_save(result);
            //    System.Threading.Thread.Sleep(300);
            //}
            #endregion

            return createdEvent;
        }
        #endregion

        #region Update
        public static void UpdateCalendarEntries(Dictionary<Microsoft.Graph.Event, GcalData.Event> entriesToBeCompared, ref int entriesUpdated) {
            for (int i = 0; i < entriesToBeCompared.Count; i++) {
                if (Sync.Engine.Instance.CancellationPending) return;

                KeyValuePair<Microsoft.Graph.Event, Event> compare = entriesToBeCompared.ElementAt(i);
                int itemModified = 0;
                Boolean eventExceptionCacheDirty = false;
                Event ev = new Event();
                try {
                    ev = UpdateCalendarEntry(compare.Key, compare.Value, ref itemModified);
                } catch (System.Exception ex) {
                    if (ex is ApplicationException) {
                        String summary = Outlook.Graph.Calendar.GetEventSummary("<br/>Event update skipped.<br/>" + ex.Message, compare.Key, out String anonSummary);
                        Forms.Main.Instance.Console.Update(summary, anonSummary, Console.Markup.warning);
                        if (ex.InnerException is global::Google.GoogleApiException) break;
                        continue;
                    } else {
                        String summary = Outlook.Graph.Calendar.GetEventSummary("<br/>Event update failed.", compare.Key, out String anonSummary);
                        Forms.Main.Instance.Console.UpdateWithError(summary, ex, logEntry: anonSummary);
                        if (ex is System.Runtime.InteropServices.COMException) throw;
                        Ogcs.Exception.Analyse(ex, true);
                        if (Ogcs.Extensions.MessageBox.Show("Google event update failed. Continue with synchronisation?", "Sync item failed", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                            continue;
                        else
                            throw new UserCancelledSyncException("User chose not to continue sync.");
                    }
                }

                if (itemModified > 0) {
                    try {
                        Ogcs.Google.Calendar.Instance.UpdateCalendarEntry_save(ref ev);
                        entriesUpdated++;
                        eventExceptionCacheDirty = true;
                    } catch (System.Exception ex) {
                        Forms.Main.Instance.Console.UpdateWithError(Outlook.Graph.Calendar.GetEventSummary("Updated event failed to save.", compare.Key, out String anonSummary, true), ex, logEntry: anonSummary);
                        log.Debug(Newtonsoft.Json.JsonConvert.SerializeObject(ev));
                        Ogcs.Exception.Analyse(ex, true);
                        if (Ogcs.Extensions.MessageBox.Show("Updated Google event failed to save. Continue with synchronisation?", "Sync item failed", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                            continue;
                        else
                            throw new UserCancelledSyncException("User chose not to continue sync.");
                    }

                }

                //Have to do this *before* any dummy update, else all the exceptions inherit the updated timestamp of the parent recurring event
                entriesUpdated += Recurrence.UpdateGoogleExceptions(compare.Key, ev ?? compare.Value, eventExceptionCacheDirty);

                if (itemModified == 0) {
                    if (ev == null) {
                        if (compare.Value.UpdatedDateTimeOffset < compare.Key.LastModifiedDateTime || Google.CustomProperty.Exists(compare.Value, Google.CustomProperty.MetadataId.forceSave))
                            ev = compare.Value;
                        else
                            continue;
                    }
                    log.Debug("Doing a dummy update in order to update the last modified date of " +
                        (ev.RecurringEventId == null && ev.Recurrence != null ? "recurring master event" : "single instance"));
                    Google.CustomProperty.SetOGCSlastModified(ref ev);
                    try {
                        Ogcs.Google.Calendar.Instance.UpdateCalendarEntry_save(ref ev);
                        entriesToBeCompared[compare.Key] = ev;
                    } catch (System.Exception ex) {
                        Forms.Main.Instance.Console.UpdateWithError(Outlook.Graph.Calendar.GetEventSummary("Updated event failed to save.", compare.Key, out String anonSummary, true), ex, logEntry: anonSummary);
                        log.Debug(Newtonsoft.Json.JsonConvert.SerializeObject(ev));
                        Ogcs.Exception.Analyse(ex, true);
                        if (Ogcs.Extensions.MessageBox.Show("Updated Google event failed to save. Continue with synchronisation?", "Sync item failed", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                            continue;
                        else
                            throw new UserCancelledSyncException("User chose not to continue sync.");
                    }
                }
            }
        }

        public static Event UpdateCalendarEntry(Microsoft.Graph.Event ai, GcalData.Event ev, ref int itemModified, Boolean forceCompare = false) {
            SettingsStore.Calendar profile = Sync.Engine.Calendar.Instance.Profile;

            if (!Settings.Instance.APIlimit_inEffect && Google.CustomProperty.Exists(ev, Google.CustomProperty.MetadataId.apiLimitHit)) {
                log.Fine("Back processing Event affected by attendee API limit.");
            } else {
                if (!(Sync.Engine.Instance.ManualForceCompare || forceCompare)) { //Needed if the exception has just been created, but now needs updating
                    if (profile.SyncDirection.Id != Sync.Direction.Bidirectional.Id) {
                        if (ev.UpdatedDateTimeOffset > ai.LastModifiedDateTime)
                            return null;
                    } else {
                        if (Outlook.Graph.CustomProperty.GetOGCSlastModified(ai).AddSeconds(5) >= ai.LastModifiedDateTime?.ToLocalTime()) {
                            log.Fine("Outlook last modified by OGCS.");
                            return null;
                        }
                        if (ev.UpdatedDateTimeOffset > ai.LastModifiedDateTime)
                            return null;
                    }
                }
            }

            String aiSummary = Outlook.Graph.Calendar.GetEventSummary(ai, out String anonSummary);
            log.Debug("Processing >> " + (anonSummary ?? aiSummary));

            if (!(ev.Creator.Self ?? (ev.Creator.Email == Settings.Instance.GaccountEmail)) && ev.Recurrence != null) {
                log.Debug("Not being the recurring Event owner, comparison for update is futile - changes won't take effect/fail.");
                log.Fine("Owner: " + ev.Creator.Email);
                return ev;
            }

            System.Text.StringBuilder sb = new System.Text.StringBuilder();
            sb.AppendLine(aiSummary);

            #region Date/Time & Time Zone
            //Handle an event's all-day attribute being toggled
            Boolean evAllDay = ev.AllDayEvent();
            OgcsDateTime evStart = new(ev.Start.SafeDateTime(), evAllDay);
            OgcsDateTime evEnd = new(ev.End.SafeDateTime(), evAllDay);
            if ((bool)ai.IsAllDay) {
                ev.Start.Date = ai.Start.SafeDateTime().ToString("yyyy-MM-dd");
                ev.End.Date = ai.End.SafeDateTime().ToString("yyyy-MM-dd");
                ev.Start.DateTimeDateTimeOffset = null;
                ev.End.DateTimeDateTimeOffset = null;
                Sync.Engine.CompareAttribute("All-Day", Sync.Direction.OutlookToGoogle, evAllDay, true, sb, ref itemModified);
                Sync.Engine.CompareAttribute("Start time", Sync.Direction.OutlookToGoogle, evStart, new OgcsDateTime(ai.Start.SafeDateTime(), true), sb, ref itemModified);
                Sync.Engine.CompareAttribute("End time", Sync.Direction.OutlookToGoogle, evEnd, new OgcsDateTime(ai.End.SafeDateTime(), true), sb, ref itemModified);
            } else {
                ev.Start.Date = null;
                ev.End.Date = null;
                ev.Start.DateTimeDateTimeOffset = ai.Start.SafeDateTime();
                ev.End.DateTimeDateTimeOffset = ai.End.SafeDateTime();
                Sync.Engine.CompareAttribute("All-Day", Sync.Direction.OutlookToGoogle, evAllDay, false, sb, ref itemModified);
                Sync.Engine.CompareAttribute("Start time", Sync.Direction.OutlookToGoogle, evStart, new OgcsDateTime(ai.Start.SafeDateTime(), false), sb, ref itemModified);
                Sync.Engine.CompareAttribute("End time", Sync.Direction.OutlookToGoogle, evEnd, new OgcsDateTime(ai.End.SafeDateTime(), false), sb, ref itemModified) ;
            }

            List<String> oRrules = Recurrence.BuildGooglePattern(ai, ev);
            Google.Recurrence.CompareGooglePattern(oRrules, ev, sb, ref itemModified);

            //TimeZone
            if (string.IsNullOrEmpty(ev.Start.Date)) {
                String startTimeZone = string.IsNullOrEmpty(ai.OriginalStartTimeZone) ? "UTC" : ai.OriginalStartTimeZone;
                startTimeZone = TimezoneDB.IANAtimezone(startTimeZone, startTimeZone);
                if (Sync.Engine.CompareAttribute("Start Timezone", Sync.Direction.OutlookToGoogle, ev.Start.TimeZone, startTimeZone, sb, ref itemModified))
                    ev.Start.TimeZone = startTimeZone;

                if (ev.Recurrence != null && ev.Start.TimeZone != ev.End.TimeZone) {
                    log.Warn("Outlook recurring series has a different start and end timezone, which Google does not allow. Setting both to the start timezone.");
                    ev.End.TimeZone = ev.Start.TimeZone;
                }
                String endTimeZone = string.IsNullOrEmpty(ai.OriginalEndTimeZone) ? "UTC" : ai.OriginalEndTimeZone;
                endTimeZone = TimezoneDB.IANAtimezone(endTimeZone, endTimeZone);
                if (Sync.Engine.CompareAttribute("End Timezone", Sync.Direction.OutlookToGoogle, ev.End.TimeZone, endTimeZone, sb, ref itemModified))
                    ev.End.TimeZone = endTimeZone;
            }
            #endregion

            String subjectObfuscated = Obfuscate.ApplyRegex(Obfuscate.Property.Subject, ai.Subject, ev.Summary, Sync.Direction.OutlookToGoogle);
            if (Sync.Engine.CompareAttribute("Subject", Sync.Direction.OutlookToGoogle, ev.Summary, subjectObfuscated, sb, ref itemModified)) {
                ev.Summary = subjectObfuscated;
            }
            if (profile.AddDescription) {
                String outlookBody = ai.Body.BodyInnerHtml();
                if (profile.SyncDirection.Id == Sync.Direction.Bidirectional.Id && profile.AddDescription_OnlyToGoogle &&
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
                    String bodyObfuscated = Obfuscate.ApplyRegex(Obfuscate.Property.Description, outlookBody, ev.Description, Sync.Direction.OutlookToGoogle);

                    //Remove HTML markup from Console output
                    String evTagsStripped = Regex.Replace(ev.Description ?? "", "<.*?>", String.Empty);
                    String aiTagsStripped = Regex.Replace(bodyObfuscated, "<.*?>", String.Empty);
                    StringBuilder currentSB = new(sb.Capacity);
                    currentSB.Append(sb);
                    
                    if (Sync.Engine.CompareAttribute("Description", Sync.Direction.OutlookToGoogle, ev.Description, bodyObfuscated, sb, ref itemModified)) {
                        ev.Description = bodyObfuscated;
                        String googleAttr_stub = ((evTagsStripped.Length > 50) ? evTagsStripped.Substring(0, 47) + "..." : evTagsStripped).RemoveLineBreaks();
                        String outlookAttr_stub = ((aiTagsStripped.Length > 50) ? aiTagsStripped.Substring(0, 47) + "..." : aiTagsStripped).RemoveLineBreaks();
                        sb = currentSB.AppendLine("Description" + ": " + googleAttr_stub + " => " + outlookAttr_stub);
                    }

                    /*if (profile.AddGMeet) {
                        String outlookGMeet = Outlook.GMeet.RgxGmeetUrl().Match(ai.Body ?? "")?.Value;
                        if (Sync.Engine.CompareAttribute("Google Meet", Sync.Direction.OutlookToGoogle, ev.HangoutLink, outlookGMeet, sb, ref itemModified)) {
                            try {
                                Ogcs.Google.GMeet.GoogleMeet(ev, outlookGMeet);
                                ev = patchEvent(ev) ?? ev;
                                log.Fine("Conference data change successfully saved.");
                            } catch (System.Exception ex) {
                                ex.Analyse("Could not update conference data in existing Event.");
                            }
                        }
                    }*/
                }
            }

            if (profile.AddLocation) {
                String locationObfuscated = Obfuscate.ApplyRegex(Obfuscate.Property.Location, ai.Location.DisplayName, ev.Location, Sync.Direction.OutlookToGoogle);
                if (Sync.Engine.CompareAttribute("Location", Sync.Direction.OutlookToGoogle, ev.Location, locationObfuscated, sb, ref itemModified))
                    ev.Location = locationObfuscated;
            }

            String gPrivacy = ev.Visibility ?? "default";
            String oPrivacy = getPrivacy(ai.Sensitivity, gPrivacy);
            if (Sync.Engine.CompareAttribute("Privacy", Sync.Direction.OutlookToGoogle, gPrivacy, oPrivacy, sb, ref itemModified)) {
                ev.Visibility = oPrivacy;
            }

            String gFreeBusy = ev.Transparency ?? "opaque";
            String oFreeBusy = getAvailability(ai.ShowAs, gFreeBusy);
            if (Sync.Engine.CompareAttribute("Free/Busy", Sync.Direction.OutlookToGoogle, gFreeBusy, oFreeBusy, sb, ref itemModified)) {
                ev.Transparency = oFreeBusy;
            }

            /*if (profile.AddColours || profile.SetEntriesColour) {
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
            }*/

            if (profile.AddAttendees && !Google.Calendar.APIlimitReached_attendee) {
                if (ai.Attendees.Count() > profile.MaxAttendees) {
                    log.Warn("This Outlook appointment has " + ai.Attendees.Count() + " attendees, more than the user configured maximum.");
                    if (ai.Attendees.Count() >= 200) {
                        Forms.Main.Instance.Console.Update(aiSummary + "<br/>Attendees will not be synced for this meeting as it has " +
                            "more than 200, which Google does not allow.", anonSummary, Console.Markup.warning);
                    }
                } else if (profile.SyncDirection.Id == Sync.Direction.Bidirectional.Id &&
                        ev.Attendees != null && ev.Attendees.Count > profile.MaxAttendees && ai.Attendees.Count() <= profile.MaxAttendees) {
                    log.Warn("This Google event has " + ev.Attendees.Count + " attendees, more than the user configured maximum. They can't safely be compared.");
                } else {
                    try {
                        CompareRecipientsToAttendees(ai, ev, sb, ref itemModified);
                    } catch (System.Exception ex) {
                        Forms.Main.Instance.Console.UpdateWithError("Unable to sync attendees.", ex);
                    }
                }
            }

            #region Reminders
            if (profile.AddReminders) {
                Boolean OKtoSyncReminder = Outlook.Graph.Calendar.Instance.IsOKtoSyncReminder(ai);
                if (ev.Reminders.Overrides != null && ev.Reminders.Overrides.Any(r => r.Method == "popup")) {
                    //Find the popup reminder(s) in Google
                    for (int r = ev.Reminders.Overrides.Count - 1; r >= 0; r--) {
                        EventReminder reminder = ev.Reminders.Overrides[r];
                        if (reminder.Method == "popup") {
                            if (OKtoSyncReminder) {
                                if ((bool)ai.IsReminderOn) {
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
                    if ((bool)ai.IsReminderOn && OKtoSyncReminder) {
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
                Forms.Main.Instance.Console.FormatEventChanges(sb, sb.ToString().Replace(aiSummary, anonSummary));
                Forms.Main.Instance.Console.Update(itemModified + " attributes updated.", Console.Markup.appointmentEnd, verbose: true, newLine: false);
                System.Windows.Forms.Application.DoEvents();
            }
            return ev;
        }
        #endregion

        public static void ReclaimOrphanCalendarEntries(ref List<Event> gEvents, ref List<Microsoft.Graph.Event> oAppointments, Boolean neverDelete = false) {
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
                    Google.CustomProperty.LogProperties(ev, Program.MyFineLevel);

                    //Find entries with no Outlook ID
                    if (!Google.CustomProperty.Exists(ev, Google.CustomProperty.MetadataId.oEntryId)) {

                        //Use simple matching on start,end,subject,location to pair events
                        String sigEv = Google.Calendar.Signature(ev);
                        if (String.IsNullOrEmpty(sigEv)) {
                            gEvents.Remove(ev);
                            continue;
                        }

                        unclaimedEvents.Add(ev);
                        for (int o = oAppointments.Count - 1; o >= 0; o--) {
                            Microsoft.Graph.Event ai = oAppointments[o];
                            if (Google.Calendar.SignaturesMatch(sigEv, Outlook.Graph.Calendar.Signature(ai))) {
                                try {
                                    Event originalEv = ev;
                                    CustomProperty.AddOutlookIDs(ref ev, ai);
                                    Google.Calendar.Instance.UpdateCalendarEntry_save(ref ev);
                                    unclaimedEvents.Remove(originalEv);
                                    if (consoleTitle != "") Forms.Main.Instance.Console.Update("<span class='em em-reclaim'></span>" + consoleTitle, Console.Markup.h2, newLine: false, verbose: true);
                                    consoleTitle = "";
                                    Forms.Main.Instance.Console.Update(Google.Calendar.GetEventSummary("Reclaimed: ", ev, out String anonSummary, appendContext: false), anonSummary, verbose: true);
                                    gEvents[g] = ev;
                                    if (profile.SyncDirection.Id == Sync.Direction.Bidirectional.Id || Outlook.Graph.CustomProperty.ExistAnyGoogleIDs(ai)) {
                                        log.Debug("Updating the Google event IDs in Outlook appointment.");
                                        Outlook.Graph.CustomProperty.AddGoogleIDs(ref ai, ev);
                                        Microsoft.Graph.Event aiPatch = new() { Id = ai.Id, Extensions = ai.Extensions };
                                        Outlook.Graph.Calendar.Instance.UpdateCalendarEntry_save(ref aiPatch);
                                        ai = aiPatch;
                                    }
                                } catch (System.Exception ex) {
                                    log.Error("Failed to reclaim Event: " + Google.Calendar.GetEventSummary(ev));
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
                     profile.SyncDirection.Id == Sync.Direction.Bidirectional.Id)) //
                {
                    log.Info(unclaimedEvents.Count + " unclaimed orphan events found.");
                    if (profile.MergeItems || profile.DisableDelete || profile.ConfirmOnDelete) {
                        log.Info("These will be kept due to configuration settings.");
                    } else if (profile.SyncDirection.Id == Sync.Direction.Bidirectional.Id) {
                        log.Debug("These 'orphaned' items must not be deleted - they need syncing up.");
                    } else {
                        if (Ogcs.Extensions.MessageBox.Show(unclaimedEvents.Count + " Google calendar events can't be matched to Outlook.\r\n" +
                            "Remember, it's recommended to have a dedicated Google calendar to sync with, " +
                            "or you may wish to merge with unmatched events. Continue with deletions?",
                            "Delete unmatched Google events?", MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation) == DialogResult.No) //
                        {
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

        /*
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
                        System.DateTime? lastSeenDuplicateSet = null;
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
                        Ogcs.Exception.Analyse(ex);
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
                        log.Fine("Checking " + Ogcs.Google.Calendar.GetEventSummary(google[g]));

                        //Use simple matching on start,end,subject,location to pair events
                        String sigEv = signature(google[g]);
                        if (String.IsNullOrEmpty(sigEv)) {
                            google.Remove(google[g]);
                            continue;
                        }

                        Boolean foundMatch = false;
                        for (int o = outlook.Count - 1; o >= 0; o--) {
                            try {
                                if (log.IsUltraFineEnabled()) log.UltraFine("Checking " + Outlook.Calendar.GetEventSummary(outlook[o]));

                                if (SignaturesMatch(sigEv, Outlook.Calendar.signature(outlook[o]))) {
                                    foundMatch = true;
                                    compare.Add(outlook[o], google[g]);
                                    outlook.Remove(outlook[o]);
                                    google.Remove(google[g]);
                                    break;
                                }
                            } catch (System.Exception ex) {
                                if (!log.IsUltraFineEnabled()) {
                                    try {
                                        log.Info(Outlook.Calendar.GetEventSummary(outlook[o]));
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
        */
        public static void IdentifyEventDifferences_IDs(
            SettingsStore.Calendar profile,
            ref List<Microsoft.Graph.Event> outlook,  //need creating
            ref List<GcalData.Event> google,          //need deleting
            ref Dictionary<Microsoft.Graph.Event, GcalData.Event> compare)
        {
            Forms.Main.Instance.Console.Update("Matching calendar items...");

            //Order by start date (same as Outlook) for quickest matching
            google.Sort((x, y) => (x.Start.DateTimeRaw ?? x.Start.Date).CompareTo((y.Start.DateTimeRaw ?? y.Start.Date)));

            // Count backwards so that we can remove found items without affecting the order of remaining items
            int metadataEnhanced = 0;
            for (int g = google.Count - 1; g >= 0; g--) {
                if (Sync.Engine.Instance.CancellationPending) return;
                log.Fine("Checking " + Ogcs.Google.Calendar.GetEventSummary(google[g]));

                if (Ogcs.Google.CustomProperty.Exists(google[g], Ogcs.Google.CustomProperty.MetadataId.oEntryId)) {
                    String compare_gEntryID = Ogcs.Google.CustomProperty.Get(google[g], Ogcs.Google.CustomProperty.MetadataId.oEntryId);
                    Boolean outlookIDmissing = Ogcs.Google.CustomProperty.OutlookIdMissing(google[g]);
                    Boolean foundMatch = false;

                    for (int o = outlook.Count - 1; o >= 0; o--) {
                        try {
                            if (log.IsUltraFineEnabled()) log.UltraFine("Checking " + Outlook.Graph.Calendar.GetEventSummary(outlook[o]));

                            String compare_oID;
                            compare_oID = outlook[o].Id;
                            if (compare_gEntryID == compare_oID && outlookIDmissing) {
                                log.Info("Enhancing event's metadata...");
                                Event ev = google[g];
                                CustomProperty.AddOutlookIDs(ref ev, outlook[o]);
                                //Don't want to save right now, else may make modified timestamp newer than a change in Outlook
                                //which would no longer sync.
                                Ogcs.Google.CustomProperty.Add(ref ev, Ogcs.Google.CustomProperty.MetadataId.forceSave, "True");
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
                                    log.Info(Outlook.Graph.Calendar.GetEventSummary(outlook[o]));
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
                        Ogcs.Google.CustomProperty.Get(google[g], Ogcs.Google.CustomProperty.MetadataId.oCalendarId) != profile.UseOutlookCalendar.Id)
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
        public static void IdentifyEventDifferences(
            ref List<Microsoft.Graph.Event> outlook,  //need creating
            ref List<GcalData.Event> google,          //need deleting
            ref Dictionary<Microsoft.Graph.Event, GcalData.Event> compare) //
        {
            log.Debug("Comparing Outlook items to Google events...");
            SettingsStore.Calendar profile = Sync.Engine.Calendar.Instance.Profile;

            /*if (profile.SimpleMatch)
                IdentifyEventDifferences_Simple(profile, ref outlook, ref google, ref compare);
            else*/
            IdentifyEventDifferences_IDs(profile, ref outlook, ref google, ref compare);

            if (Sync.Engine.Instance.CancellationPending) return;

            if (outlook.Count > 0 && profile.OnlyRespondedInvites) {
                //Check if Outlook items to be created in Google have invitations not yet responded to
                int responseFiltered = 0;
                for (int o = outlook.Count - 1; o >= 0; o--) {
                    if (outlook[o].ResponseStatus.Response == Microsoft.Graph.ResponseType.NotResponded) {
                        outlook.Remove(outlook[o]);
                        responseFiltered++;
                    }
                }
                if (responseFiltered > 0) log.Info(responseFiltered + " Outlook items will not be created due to only syncing invites that have been responded to.");
            }

            /*
                        if (google.Count > 0 && Outlook.Calendar.Instance.ExcludedByCategory?.Count > 0 && !profile.DeleteWhenCategoryExcluded) {
                            //Check if Google items to be deleted were filtered out from Outlook
                            for (int g = google.Count - 1; g >= 0; g--) {
                                if (Outlook.Calendar.Instance.ExcludedByCategory.ContainsValue(google[g].Id) ||
                                    Outlook.Calendar.Instance.ExcludedByCategory.ContainsKey(Ogcs.Google.CustomProperty.Get(google[g], Ogcs.Google.CustomProperty.MetadataId.oEntryId) ?? "")) {
                                    google.Remove(google[g]);
                                }
                            }
                        }
                        if (outlook.Count > 0 && Ogcs.Google.Calendar.Instance.ExcludedByColour?.Count > 0) {
                            //Check if Outlook items to be created were filtered out from Google
                            for (int o = outlook.Count - 1; o >= 0; o--) {
                                if (Ogcs.Google.Calendar.Instance.ExcludedByColour.ContainsValue(outlook[o].Id) ||
                                    Ogcs.Google.Calendar.Instance.ExcludedByColour.ContainsKey(Outlook.Graph.CustomProperty.Get(outlook[o], Outlook.Graph.CustomProperty.MetadataId.gEventID) ?? "")) {
                                    outlook.Remove(outlook[o]);
                                }
                            }
                        }
            */
            if (profile.SyncDirection.Id == Sync.Direction.Bidirectional.Id) {
                //Don't recreate any items that have been deleted in Google
                for (int o = outlook.Count - 1; o >= 0; o--) {
                    if (Outlook.Graph.CustomProperty.Exists(outlook[o], Outlook.Graph.CustomProperty.MetadataId.gEventID))
                        outlook.Remove(outlook[o]);
                }
                //Don't delete any items that aren't yet in Outlook or just created in Outlook during this sync
                for (int g = google.Count - 1; g >= 0; g--) {
                    if (!Ogcs.Google.CustomProperty.Exists(google[g], Ogcs.Google.CustomProperty.MetadataId.oEntryId) ||
                        google[g].UpdatedDateTimeOffset > Sync.Engine.Instance.SyncStarted)
                        google.Remove(google[g]);
                }
            }
            if (profile.DisableDelete) {
                if (google.Count > 0) {
                    Forms.Main.Instance.Console.Update(google.Count + " Google items would have been deleted, but you have deletions disabled.", Console.Markup.warning);
                    for (int g = 0; g < google.Count; g++)
                        Forms.Main.Instance.Console.Update(Ogcs.Google.Calendar.GetEventSummary(google[g], out String anonSummary), anonSummary, verbose: true);
                }
                google = new List<Event>();
            }
            if (Settings.Instance.CreateCSVFiles) {
                Ogcs.Google.Calendar.ExportToCSV("Events for deletion in Google", "google_delete.csv", google);
                Outlook.Graph.Calendar.ExportToCSV("Appointments for creation in Google", "google_create.csv", outlook);
            }
        }
        public static Boolean ItemIDsMatch(ref GcalData.Event ev, Microsoft.Graph.Event ai) {
            SettingsStore.Calendar profile = Sync.Engine.Calendar.Instance.Profile;

            //First match on the Global Appointment ID
            //For format of Global ID: https://msdn.microsoft.com/en-us/library/ee157690%28v=exchg.80%29.aspx
            log.Fine("Comparing Outlook GlobalID");

            if (Ogcs.Google.CustomProperty.Exists(ev, Ogcs.Google.CustomProperty.MetadataId.oGlobalApptId)) {
                String gCompareID = Ogcs.Google.CustomProperty.Get(ev, Ogcs.Google.CustomProperty.MetadataId.oGlobalApptId);
                String oGlobalID = ai.ICalUId;

                //For items copied from someone elses calendar, it appears the Global ID is generated for each access?! (Creation Time changes)
                //I guess the copied item doesn't really have its "own" ID. So, we'll just compare
                //the "data" section of the byte array, which "ensures uniqueness" and doesn't include ID creation time

                if (oGlobalID == gCompareID ||
                    ((oGlobalID.StartsWith(Outlook.Calendar.GlobalIdPattern) &&
                        gCompareID.StartsWith(Outlook.Calendar.GlobalIdPattern) &&
                        gCompareID.Substring(72) == oGlobalID.Substring(72))             //We've got bonafide Global IDs match
                        ||
                        (!oGlobalID.StartsWith(Outlook.Calendar.GlobalIdPattern) &&
                        !gCompareID.StartsWith(Outlook.Calendar.GlobalIdPattern) &&
                        gCompareID.Remove(gCompareID.Length - 16) == oGlobalID.Remove(oGlobalID.Length - 16)) //Or it's really a Entry ID (failsafe match)
                    ))
                {
                    log.Fine("Comparing Outlook CalendarID");
                    gCompareID = Ogcs.Google.CustomProperty.Get(ev, Ogcs.Google.CustomProperty.MetadataId.oCalendarId);
                    if (gCompareID == profile.UseOutlookCalendar.Id) {

                        log.Fine("Comparing Outlook EntryID");
                        gCompareID = Ogcs.Google.CustomProperty.Get(ev, Ogcs.Google.CustomProperty.MetadataId.oEntryId);
                        if (gCompareID == ai.Id) {
                            return true;
                        } /*else if (!string.IsNullOrEmpty(gCompareID) &&
                            gCompareID.Remove(gCompareID.Length - 16) == ai.EntryID.Remove(ai.EntryID.Length - 16))
                        {
                            //Worse still, both a locally copied item AND a rescheduled appointment by someone else 
                            //will have the MessageGlobalCounter bytes incremented (last 8-bytes)
                            //The former is identified by ExplorerWatcher adding a special flag
                            if (Outlook.CustomProperty.Get(ai, Outlook.CustomProperty.MetadataId.locallyCopied) == true.ToString()) {
                                log.Fine("This appointment was copied by the user. Incorrect match avoided.");
                                return false;
                            } else {
                                if (profile.OutlookGalBlocked || ai.Organizer != Outlook.Calendar.Instance.IOutlook.CurrentUserName()) {
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
                            if (SignaturesMatch(signature(ev), Outlook.Calendar.signature(ai))) {
                                CustomProperty.AddOutlookIDs(ref ev, ai); //update EntryID
                                CustomProperty.Add(ref ev, CustomProperty.MetadataId.forceSave, "True");
                                return true;
                            }*/
                    }
                }

                /*} else if (profile.SyncDirection.Id == Sync.Direction.Bidirectional.Id &&
                        oGlobalID.StartsWith(Outlook.Calendar.GlobalIdPattern) &&
                        gCompareID.StartsWith(Outlook.Calendar.GlobalIdPattern) &&
                        gCompareID.Substring(72) != oGlobalID.Substring(72) &&
                        Outlook.CustomProperty.Get(ai, Outlook.CustomProperty.MetadataId.gEventID) == ev.Id &&
                        SignaturesMatch(signature(ev), Outlook.Calendar.signature(ai))) 
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
                */
            } else {
                if (profile.MergeItems)
                    log.Fine("Could not find global Appointment ID against Google Event.");
                else
                    log.Warn("Could not find global Appointment ID against Google Event.");
            }
            return false;
        }

        public static Boolean CompareRecipientsToAttendees(Microsoft.Graph.Event ai, GcalData.Event ev, StringBuilder sb, ref int itemModified) {
            log.Fine("Comparing Recipients");
            List<Microsoft.Graph.Attendee> recipients = ai.Attendees.ToList();
            //Build a list of Google attendees. Any remaining at the end of the diff must be deleted.
            List<GcalData.EventAttendee> removeAttendee = new();
            foreach (GcalData.EventAttendee ea in ev.Attendees ?? Enumerable.Empty<GcalData.EventAttendee>()) {
                removeAttendee.Add(ea);
            }
            for (int o = recipients.Count() - 1; o >= 0; o--) {
                bool foundAttendee = false;
                Microsoft.Graph.Attendee recipient = recipients[o];
                log.Fine("Comparing Outlook recipient: " + (recipient.EmailAddress.Name ?? recipient.EmailAddress.Address));
                //String recipientSMTP = Outlook.Calendar.Instance.IOutlook.GetRecipientEmail(recipient);
                foreach (GcalData.EventAttendee attendee in ev.Attendees ?? Enumerable.Empty<GcalData.EventAttendee>()) {
                    EventAttendee ogcsAttendee = new(attendee);
                    if (ogcsAttendee.Email != null && (recipient.EmailAddress.Address == ogcsAttendee.Email)) {
                        foundAttendee = true;
                        removeAttendee.Remove(attendee);

                        if (Sync.Engine.Calendar.Instance.Profile.CloakEmail != ogcsAttendee.IsCloaked()) {
                            Sync.Engine.CompareAttribute("Attendee updated", Sync.Direction.OutlookToGoogle, attendee.Email, EventAttendee.CloakEmail(attendee.Email), sb, ref itemModified);
                            attendee.Email = EventAttendee.CloakEmail(attendee.Email);
                        }

                        //Optional attendee
                        bool oOptional = (recipient.Type == Microsoft.Graph.AttendeeType.Optional);
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
                        } else if (ai.Organizer.EmailAddress.Address == attendeeIdentifier) {
                            if (Sync.Engine.CompareAttribute("Organiser " + attendeeIdentifier + " - Response Status",
                                Sync.Direction.OutlookToGoogle,
                                attendee.ResponseStatus, "accepted", sb, ref itemModified)) {
                                log.Fine("Forcing the Outlook organiser to have accepted the 'invite' in Google");
                                attendee.ResponseStatus = "accepted";
                            }
                            break;
                        }

                        String oResponse = null;
                        switch (recipient.Status.Response) {
                            case Microsoft.Graph.ResponseType.None: oResponse = "needsAction"; break;
                            case Microsoft.Graph.ResponseType.NotResponded: oResponse = "needsAction"; break;
                            case Microsoft.Graph.ResponseType.Accepted: oResponse = "accepted"; break;
                            case Microsoft.Graph.ResponseType.Declined: oResponse = "declined"; break;
                            case Microsoft.Graph.ResponseType.TentativelyAccepted: oResponse = "tentative"; break;
                        }
                        if (Sync.Engine.CompareAttribute("Attendee " + attendeeIdentifier + " - Response Status",
                            Sync.Direction.OutlookToGoogle, attendee.ResponseStatus, oResponse, sb, ref itemModified)) //
                        {
                            attendee.ResponseStatus = oResponse;
                        }
                    }
                } //each attendee

                if (!foundAttendee) {
                    log.Fine("Attendee added: " + (recipient.EmailAddress.Name ?? recipient.EmailAddress.Address));
                    sb.AppendLine("Attendee added: " + (recipient.EmailAddress.Name ?? recipient.EmailAddress.Address));
                    ev.Attendees ??= new List<GcalData.EventAttendee>();
                    ev.Attendees.Add(CreateAttendee(recipient, ai.Organizer.EmailAddress.Address == recipient.EmailAddress.Address));
                    itemModified++;
                }
            } //each recipient

            foreach (GcalData.EventAttendee gea in removeAttendee) {
                Ogcs.Google.EventAttendee ea = new Ogcs.Google.EventAttendee(gea);
                log.Fine("Attendee removed: " + (ea.DisplayName ?? ea.Email), ea.Email);
                sb.AppendLine("Attendee removed: " + (ea.DisplayName ?? ea.Email));
                ev.Attendees.Remove(gea);
                itemModified++;
            }
            return (itemModified > 0);
        }

        /// <summary>
        /// Determine Event's privacy setting
        /// </summary>
        /// <param name="oSensitivity">Outlook's current setting</param>
        /// <param name="gVisibility">Google's current setting</param>
        /// <param name="direction">Direction of sync</param>
        private static String getPrivacy(Microsoft.Graph.Sensitivity? oSensitivity, String gVisibility) {
            SettingsStore.Calendar profile = Sync.Engine.Calendar.Instance.Profile;
            oSensitivity ??= Microsoft.Graph.Sensitivity.Normal;

            if (!profile.SetEntriesPrivate)
                return (oSensitivity == Microsoft.Graph.Sensitivity.Normal) ? "default" : "private";

            if (profile.SyncDirection.Id != Sync.Direction.Bidirectional.Id) {
                return (profile.PrivacyLevel == Microsoft.Graph.Sensitivity.Private.ToString()) ? "private" : "public";
            } else {
                if (profile.TargetCalendar.Id == Sync.Direction.GoogleToOutlook.Id) { //Privacy enforcement is in other direction
                    if (gVisibility == null)
                        return (oSensitivity == Microsoft.Graph.Sensitivity.Normal) ? "default" : "private";
                    else if (gVisibility == "private" && oSensitivity != Microsoft.Graph.Sensitivity.Private) {
                        log.Fine("Source of truth for privacy is already set private and target is NOT - so syncing this back.");
                        return "default";
                    } else
                        return gVisibility;
                } else {
                    if (!profile.CreatedItemsOnly || (profile.CreatedItemsOnly && gVisibility == null))
                        return (profile.PrivacyLevel == Microsoft.Graph.Sensitivity.Private.ToString()) ? "private" : "public";
                    else
                        return (oSensitivity == Microsoft.Graph.Sensitivity.Normal) ? "default" : "private";
                }
            }
        }

        /// <summary>
        /// Determine Event's availability setting
        /// </summary>
        /// <param name="oBusyStatus">Outlook's current setting</param>
        /// <param name="gTransparency">Google's current setting</param>
        private static String getAvailability(Microsoft.Graph.FreeBusyStatus? oBusyStatus, String gTransparency) {
            SettingsStore.Calendar profile = Sync.Engine.Calendar.Instance.Profile;
            oBusyStatus ??= Microsoft.Graph.FreeBusyStatus.Busy;

            if (!profile.SetEntriesAvailable)
                return (oBusyStatus == Microsoft.Graph.FreeBusyStatus.Free) ? "transparent" : "opaque";

            String overrideTransparency = "transparent";
            Microsoft.Graph.FreeBusyStatus fbStatus = Microsoft.Graph.FreeBusyStatus.Free;
            try {
                Enum.TryParse(profile.AvailabilityStatus, out fbStatus);
                if (fbStatus != Microsoft.Graph.FreeBusyStatus.Free)
                    overrideTransparency = "opaque";
            } catch (System.Exception ex) {
                ex.Analyse("Could not convert string '" + profile.AvailabilityStatus + "' to FreeBusyStatus type. Defaulting override to available.");
            }

            if (profile.SyncDirection.Id != Sync.Direction.Bidirectional.Id) {
                return overrideTransparency;
            } else {
                if (profile.TargetCalendar.Id == Sync.Direction.GoogleToOutlook.Id) { //Availability enforcement is in other direction
                    if (gTransparency == null)
                        return (oBusyStatus == Microsoft.Graph.FreeBusyStatus.Free) ? "transparent" : "opaque";
                    else
                        return gTransparency;
                } else {
                    if (!profile.CreatedItemsOnly || (profile.CreatedItemsOnly && gTransparency == null))
                        return overrideTransparency;
                    else
                        return (oBusyStatus == Microsoft.Graph.FreeBusyStatus.Free) ? "transparent" : "opaque";
                }
            }
        }

        /*
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
                KeyValuePair<String, String> kvp = profile.ColourMaps.FirstOrDefault(cm => Outlook.Calendar.Categories.OutlookColour(cm.Key) == categoryColour);
                if (kvp.Key != null) {
                    gColour = ColourPalette.ActivePalette.FirstOrDefault(ap => ap.Id == kvp.Value);
                    if (gColour != null) {
                        log.Debug("Colour mapping used: " + kvp.Key + " => " + kvp.Value + ":" + gColour.Name);
                        return gColour;
                    }
                }
            }
            //Algorithmic closest colour matching
            System.Drawing.Color color = Outlook.Categories.Map.RgbColour((OlCategoryColor)categoryColour);
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
                    String category = aiCategories.Split(new[] { Outlook.Calendar.Categories.Delimiter }, StringSplitOptions.None).FirstOrDefault();
                    categoryColour = Outlook.Calendar.Categories.OutlookColour(category);
                } catch (System.Exception ex) {
                    log.Error("Failed determining colour for Event from AppointmentItem categories: " + aiCategories);
                    Ogcs.Exception.Analyse(ex);
                }
            }
        }
        */

        public static GcalData.EventAttendee CreateAttendee(Microsoft.Graph.Attendee recipient, Boolean isOrganiser) {
            Ogcs.Google.EventAttendee ea = new Ogcs.Google.EventAttendee();
            log.Fine("Creating attendee " + (string.IsNullOrEmpty(recipient.EmailAddress.Name) ? recipient.EmailAddress.Address : recipient.EmailAddress.Name));
            ea.DisplayName = (recipient.EmailAddress.Name != recipient.EmailAddress.Address ? recipient.EmailAddress.Name : null);
            ea.Email = recipient.EmailAddress.Address;
            ea.Optional = recipient.Type == Microsoft.Graph.AttendeeType.Optional;
            if (isOrganiser) {
                //ea.Organizer = true; This is read-only. The best we can do is force them to have accepted the "invite"
                ea.ResponseStatus = "accepted";
                return ea;
            }
            switch (recipient.Status.Response) {
                case Microsoft.Graph.ResponseType.NotResponded: ea.ResponseStatus = "needsAction"; break;
                case Microsoft.Graph.ResponseType.None: ea.ResponseStatus = "needsAction"; break;
                case Microsoft.Graph.ResponseType.Accepted: ea.ResponseStatus = "accepted"; break;
                case Microsoft.Graph.ResponseType.Declined: ea.ResponseStatus = "declined"; break;
                case Microsoft.Graph.ResponseType.TentativelyAccepted: ea.ResponseStatus = "tentative"; break;
            }
            return ea;
        }
    }
}
