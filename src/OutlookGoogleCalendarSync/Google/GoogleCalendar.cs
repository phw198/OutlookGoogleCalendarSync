using Google.Apis.Calendar.v3;
using Google.Apis.Calendar.v3.Data;
using log4net;
using Microsoft.Office.Interop.Outlook;
using OutlookGoogleCalendarSync.Extensions;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using Ogcs = OutlookGoogleCalendarSync;

namespace OutlookGoogleCalendarSync.Google {
    /// <summary>
    /// Description of Ogcs.Google.Calendar.
    /// </summary>
    public class Calendar {
        private static readonly ILog log = LogManager.GetLogger(typeof(Calendar));

        private static Calendar instance;
        public static Boolean IsInstanceNull { get { return instance == null; } }
        public static Calendar Instance {
            get {
                if (instance == null) {
                    instance = new Ogcs.Google.Calendar {
                        Authenticator = new Ogcs.Google.Authenticator()
                    };
                    _ = instance.Service;
                }
                return instance;
            }
        }
        public Calendar() { }
        private Boolean openedIssue1593 = false;
        public Ogcs.Google.Authenticator Authenticator;

        /// <summary>Google Events excluded through user config <Event.Id, Appt.EntryId></summary>
        public List<String> ExcludedByConfig { get; set; }
        /// <summary>Google Events excluded by colour through user config <Event.Id, Appt.EntryId></summary>
        public Dictionary<String, String> ExcludedByColour { get; set; }

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
                    Authenticator.GetAuthenticated();
                    if (Authenticator?.Authenticated ?? false) {
                        Authenticator.OgcsUserStatus();
                        _ = ColourPalette;
                    } else {
                        if (Forms.Main.Instance.Console.DocumentText.Contains("Authorisation to allow OGCS to manage your Google calendar was cancelled."))
                            throw new OperationCanceledException();
                        else if (Authenticator != null && !Authenticator.SufficientPermissions) {
                            throw new ApplicationException("OGCS has not been granted permission to manage your calendars. " +
                                "When authorising access to your Google account, please ensure permission is granted to <b>all the items</b> requested.");
                        } else {
                            instance = null;
                            throw new ApplicationException("Google handshake failed.");
                        }
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

                if (request != null) {
                    SettingsStore.Calendar profile = Settings.Profile.InPlay();
                    List<Event> evList = new List<Event>() { request };
                    applyExclusions(ref evList, profile);
                    return evList.FirstOrDefault();
                } else
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

        public List<Event> GetCalendarEntriesInRange(System.DateTime from, System.DateTime to, Boolean suppressAdvisories = false) {
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

            List<Event> allExcluded = applyExclusions(ref result, profile);
            if (allExcluded.Count > 0) {
                if (!suppressAdvisories) {
                    String filterWarning = "Due to your OGCS Google settings, " + (result.Count == 0 ? "all" : allExcluded.Count) + " Google items have been filtered out" + (result.Count == 0 ? "!" : ".");
                    Forms.Main.Instance.Console.Update(filterWarning, Console.Markup.config, newLine: false, notifyBubble: (result.Count == 0));

                    filterWarning = "";
                    if (profile.SyncDirection.Id != Sync.Direction.OutlookToGoogle.Id && ExcludedByColour.Count > 0 && profile.DeleteWhenColourExcluded) {
                        filterWarning = "If they exist in Outlook, they may get deleted. To avoid deletion, uncheck \"Delete synced items if excluded\".";
                        if (!profile.DisableDelete) {
                            filterWarning += " Recover unintentional deletions from the Outlook 'Deleted Items' folder.";
                            if (profile.ConfirmOnDelete)
                                filterWarning += "<p style='margin-top: 8px;'>If prompted to confirm deletion and you opt <i>not</i> to delete them, this will reoccur every sync. " +
                                    "Consider assigning an excluded category to those items in Outlook.</p>" +
                                    "<p style='margin-top: 8px;'>See the wiki for tips if needing to <a href='https://github.com/phw198/OutlookGoogleCalendarSync/wiki/FAQs#duplicates-due-to-colourcategory-exclusion'>resolve duplicates</a>.</p>";
                        }
                    }
                    if (!String.IsNullOrEmpty(filterWarning))
                        Forms.Main.Instance.Console.Update(filterWarning, Console.Markup.warning, newLine: false);
                }
                if (profile.SyncDirection.Id == Sync.Direction.Bidirectional.Id) {
                    for (int g = 0; g < allExcluded.Count; g++) {
                        Event ev = allExcluded[g];
                        if (CustomProperty.ExistAnyOutlookIDs(ev)) {
                            log.Debug("Previously synced Google item is now excluded. Removing Outlook metadata.");
                            //We don't want them getting automatically deleted if brought back in scope; better to create possible duplicate
                            CustomProperty.RemoveOutlookIDs(ref ev);
                            UpdateCalendarEntry_save(ref ev);
                        }
                    }
                }
            }

            log.Fine("Filtered down to " + result.Count);
            return result;
        }

        private List<Event> applyExclusions(ref List<Event> result, SettingsStore.Calendar profile) {
            List<Event> colour = new();
            List<Event> availability = new();
            List<Event> allDays = new();
            List<Event> privacy = new();
            List<Event> subject = new();
            List<Event> declined = new();
            List<Event> goals = new();

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
                if (!ExcludedByColour.ContainsKey(ev.Id))
                    ExcludedByColour.Add(ev.Id, CustomProperty.Get(ev, CustomProperty.MetadataId.oEntryId));
            }
            result = result.Except(colour).ToList();

            //Availability, All-Days, Privacy, Subject
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

            List<Event> allExcluded = colour.Concat(availability).Concat(allDays).Concat(privacy).Concat(subject).Concat(declined).Concat(goals).ToList();
            foreach (Event ev in allExcluded) {
                if (!ExcludedByConfig.Contains(ev.Id))
                    ExcludedByConfig.Add(ev.Id);
            }
            return allExcluded;
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
                        String summary = Outlook.Calendar.GetEventSummary("Event creation skipped.<br/>" + ex.Message, ai, out String anonSummary);
                        Forms.Main.Instance.Console.Update(summary, anonSummary, Console.Markup.warning);
                        if (ex.InnerException is global::Google.GoogleApiException) break;
                        continue;
                    } else {
                        String summary = Outlook.Calendar.GetEventSummary("Event creation failed.", ai, out String anonSummary);
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
                    Forms.Main.Instance.Console.UpdateWithError(Outlook.Calendar.GetEventSummary("New event failed to save.", ai, out String anonSummary, true), ex, logEntry: anonSummary);
                    Ogcs.Exception.Analyse(ex, true);
                    if (Ogcs.Extensions.MessageBox.Show("New Google event failed to save. Continue with synchronisation?", "Sync item failed", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                        continue;
                    else
                        throw new UserCancelledSyncException("User chose not to continue sync.");
                }
                if (ai.IsRecurring && Outlook.Recurrence.HasExceptions(ai) && createdEvent != null) {
                    Forms.Main.Instance.Console.Update("This is a recurring item with some exceptions:-", verbose: true);
                    Recurrence.CreateGoogleExceptions(ai, createdEvent.Id);
                    Forms.Main.Instance.Console.Update("Recurring exceptions completed.", verbose: true);
                }
            }
        }

        private Event createCalendarEntry(AppointmentItem ai) {
            SettingsStore.Calendar profile = Sync.Engine.Calendar.Instance.Profile;

            string itemSummary = Outlook.Calendar.GetEventSummary(ai, out String anonSummary);
            log.Debug("Processing >> " + (anonSummary ?? itemSummary));
            Forms.Main.Instance.Console.Update(itemSummary, anonSummary, Console.Markup.calendar, verbose: true);

            Event ev = new Event();

            ev.Recurrence = Recurrence.BuildGooglePattern(ai, ev);
            ev.Start = new EventDateTime();
            ev.End = new EventDateTime();

            if (ai.AllDayEvent) {
                ev.Start.Date = ai.Start.ToString("yyyy-MM-dd");
                ev.End.Date = ai.End.ToString("yyyy-MM-dd");
            } else {
                ev.Start.DateTime = ai.Start;
                ev.End.DateTime = ai.End;
            }
            ev = Outlook.Calendar.Instance.IOutlook.IANAtimezone_set(ev, ai);

            ev.Summary = Obfuscate.ApplyRegex(Obfuscate.Property.Subject, ai.Subject, null, Sync.Direction.OutlookToGoogle);
            if (profile.AddDescription) {
                try {
                    ev.Description = Obfuscate.ApplyRegex(Obfuscate.Property.Description, ai.Body, null, Sync.Direction.OutlookToGoogle);
                } catch (System.Exception ex) {
                    if (ex.GetErrorCode() == "0x80004004") {
                        Forms.Main.Instance.Console.Update("You do not have the rights to programmatically access Outlook appointment descriptions.<br/>" +
                            "It may be best to stop syncing the Description attribute.", Console.Markup.warning);
                    } else throw;
                }
            }
            if (profile.AddLocation)
                ev.Location = Obfuscate.ApplyRegex(Obfuscate.Property.Location, ai.Location, null, Sync.Direction.OutlookToGoogle);
            ev.Visibility = getPrivacy(ai.Sensitivity, null);
            ev.Transparency = getAvailability(ai.BusyStatus, null);
            ev.ColorId = getColour(ai.Categories, null)?.Id ?? EventColour.Palette.NullPalette.Id;

            ev.Attendees = new List<global::Google.Apis.Calendar.v3.Data.EventAttendee>();
            if (profile.AddAttendees && ai.Recipients.Count > 1 && !APIlimitReached_attendee) { //Don't add attendees if there's only 1 (me)
                if (ai.Recipients.Count > profile.MaxAttendees) {
                    log.Warn("This Outlook appointment has " + ai.Recipients.Count + " attendees, more than the user configured maximum.");
                    if (ai.Recipients.Count >= 200) {
                        Forms.Main.Instance.Console.Update("Attendees will not be synced for this meeting as it has " +
                            "more than 200, which Google does not allow.", Console.Markup.warning);
                    }
                } else {
                    foreach (Microsoft.Office.Interop.Outlook.Recipient recipient in ai.Recipients) {
                        global::Google.Apis.Calendar.v3.Data.EventAttendee ea = Ogcs.Google.Calendar.CreateAttendee(recipient, ai.Organizer == recipient.Name);
                        ev.Attendees.Add(ea);
                    }
                }
            }

            //Reminder alert
            ev.Reminders = new Event.RemindersData();
            if (profile.AddReminders) {
                if (Outlook.Calendar.Instance.IsOKtoSyncReminder(ai)) {
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

            if (!String.IsNullOrEmpty(createdEvent.Id) && (profile.SyncDirection.Id == Sync.Direction.Bidirectional.Id || Outlook.CustomProperty.ExistAnyGoogleIDs(ai))) {
                log.Debug("Storing the Google event IDs in Outlook appointment.");
                Outlook.CustomProperty.AddGoogleIDs(ref ai, createdEvent);
                Outlook.CustomProperty.SetOGCSlastModified(ref ai);
                ai.Save();
            }

            if (profile.AddGMeet && Outlook.GMeet.BodyHasGmeetUrl(ai)) {
                log.Info("Adding GMeet conference details.");
                String outlookGMeet = Outlook.GMeet.RgxGmeetUrl().Match(ai.Body).Value;
                Ogcs.Google.GMeet.GoogleMeet(createdEvent, outlookGMeet);
                createdEvent = patchEvent(createdEvent) ?? createdEvent;
                log.Fine("Conference data added.");
            }

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
                        String summary = Outlook.Calendar.GetEventSummary("<br/>Event update skipped.<br/>" + ex.Message, compare.Key, out String anonSummary);
                        Forms.Main.Instance.Console.Update(summary, anonSummary, Console.Markup.warning);
                        if (ex.InnerException is global::Google.GoogleApiException) break;
                        continue;
                    } else {
                        String summary = Outlook.Calendar.GetEventSummary("<br/>Event update failed.", compare.Key, out String anonSummary);
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
                        UpdateCalendarEntry_save(ref ev);
                        entriesUpdated++;
                        eventExceptionCacheDirty = true;
                    } catch (System.Exception ex) {
                        Forms.Main.Instance.Console.UpdateWithError(Outlook.Calendar.GetEventSummary("Updated event failed to save.", compare.Key, out String anonSummary, true), ex, logEntry: anonSummary);
                        log.Debug(Newtonsoft.Json.JsonConvert.SerializeObject(ev));
                        Ogcs.Exception.Analyse(ex, true);
                        if (Ogcs.Extensions.MessageBox.Show("Updated Google event failed to save. Continue with synchronisation?", "Sync item failed", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
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
                        Forms.Main.Instance.Console.UpdateWithError(Outlook.Calendar.GetEventSummary("Updated event failed to save.", compare.Key, out String anonSummary, true), ex, logEntry: anonSummary);
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
                        if (Outlook.CustomProperty.GetOGCSlastModified(ai).AddSeconds(5) >= ai.LastModificationTime) {
                            log.Fine("Outlook last modified by OGCS.");
                            return null;
                        }
                        if (ev.Updated > ai.LastModificationTime)
                            return null;
                    }
                }
            }

            String aiSummary = Outlook.Calendar.GetEventSummary(ai, out String anonSummary);
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
            Extensions.OgcsDateTime evStart = new OgcsDateTime(ev.Start.SafeDateTime(), evAllDay);
            Extensions.OgcsDateTime evEnd = new OgcsDateTime(ev.End.SafeDateTime(), evAllDay);
            if (ai.AllDayEvent) {
                ev.Start.Date = ai.Start.ToString("yyyy-MM-dd");
                ev.End.Date = ai.End.ToString("yyyy-MM-dd");
                ev.Start.DateTime = null;
                ev.End.DateTime = null;
                Sync.Engine.CompareAttribute("All-Day", Sync.Direction.OutlookToGoogle, evAllDay, true, sb, ref itemModified);
                Sync.Engine.CompareAttribute("Start time", Sync.Direction.OutlookToGoogle, evStart, new Extensions.OgcsDateTime(ai.Start, true), sb, ref itemModified);
                Sync.Engine.CompareAttribute("End time", Sync.Direction.OutlookToGoogle, evEnd, new Extensions.OgcsDateTime(ai.End, true), sb, ref itemModified);
            } else {
                ev.Start.Date = null;
                ev.End.Date = null;
                ev.Start.DateTime = ai.Start;
                ev.End.DateTime = ai.End;
                Sync.Engine.CompareAttribute("All-Day", Sync.Direction.OutlookToGoogle, evAllDay, false, sb, ref itemModified);
                Sync.Engine.CompareAttribute("Start time", Sync.Direction.OutlookToGoogle, evStart, new Extensions.OgcsDateTime(ai.Start, false), sb, ref itemModified);
                Sync.Engine.CompareAttribute("End time", Sync.Direction.OutlookToGoogle, evEnd, new Extensions.OgcsDateTime(ai.End, false), sb, ref itemModified);
            }

            List<String> oRrules = Recurrence.BuildGooglePattern(ai, ev);
            Recurrence.CompareGooglePattern(oRrules, ev, sb, ref itemModified);

            //TimeZone
            if (ev.Start.DateTime != null) {
                String currentStartTZ = ev.Start.TimeZone;
                String currentEndTZ = ev.End.TimeZone;
                ev = Outlook.Calendar.Instance.IOutlook.IANAtimezone_set(ev, ai);
                if (ev.Recurrence != null && ev.Start.TimeZone != ev.End.TimeZone) {
                    log.Warn("Outlook recurring series has a different start and end timezone, which Google does not allow. Setting both to the start timezone.");
                    ev.End.TimeZone = ev.Start.TimeZone;
                }
                Sync.Engine.CompareAttribute("Start Timezone", Sync.Direction.OutlookToGoogle, currentStartTZ, ev.Start.TimeZone, sb, ref itemModified);
                Sync.Engine.CompareAttribute("End Timezone", Sync.Direction.OutlookToGoogle, currentEndTZ, ev.End.TimeZone, sb, ref itemModified);
            }
            #endregion

            String subjectObfuscated = Obfuscate.ApplyRegex(Obfuscate.Property.Subject, ai.Subject, ev.Summary, Sync.Direction.OutlookToGoogle);
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
                    String bodyObfuscated = Obfuscate.ApplyRegex(Obfuscate.Property.Description, outlookBody, ev.Description, Sync.Direction.OutlookToGoogle);
                    if (Sync.Engine.CompareAttribute("Description", Sync.Direction.OutlookToGoogle, ev.Description, bodyObfuscated, sb, ref itemModified))
                        ev.Description = bodyObfuscated;

                    if (profile.AddGMeet) {
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
                    }
                }
            }

            if (profile.AddLocation) {
                String locationObfuscated = Obfuscate.ApplyRegex(Obfuscate.Property.Location, ai.Location, ev.Location, Sync.Direction.OutlookToGoogle);
                if (Sync.Engine.CompareAttribute("Location", Sync.Direction.OutlookToGoogle, ev.Location, locationObfuscated, sb, ref itemModified))
                    ev.Location = locationObfuscated;
            }

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
                        Forms.Main.Instance.Console.Update(aiSummary + "<br/>Attendees will not be synced for this meeting as it has " +
                            "more than 200, which Google does not allow.", anonSummary, Console.Markup.warning);
                    }
                } else if (profile.SyncDirection.Id == Sync.Direction.Bidirectional.Id &&
                        ev.Attendees != null && ev.Attendees.Count > profile.MaxAttendees && ai.Recipients.Count <= profile.MaxAttendees) {
                    log.Warn("This Google event has " + ev.Attendees.Count + " attendees, more than the user configured maximum. They can't safely be compared.");
                } else {
                    try {
                        CompareRecipientsToAttendees(ai, ev, sb, ref itemModified);
                    } catch (System.Exception ex) {
                        if (Outlook.Calendar.Instance.IOutlook.ExchangeConnectionMode().ToString().Contains("Disconnected")) {
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
                Boolean OKtoSyncReminder = Outlook.Calendar.Instance.IsOKtoSyncReminder(ai);
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
                Forms.Main.Instance.Console.FormatEventChanges(sb, sb.ToString().Replace(aiSummary, anonSummary));
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
                } catch (global::Google.GoogleApiException ex) {
                    ApiException handled = HandleAPIlimits(ref ex, ev);
                    switch (handled) {
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
                        Ogcs.Extensions.MessageBox.Show("A 'PreCondition Failed [412]' error was encountered.\r\nPlease see issue #1593 on GitHub for further information.",
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
        //        Ogcs.Extensions.MessageBox.Show(windowToBlock, message, "Error", MessageBoxButton.OK, MessageBoxImage.Error);
        //    else {
        //        this.Dispatcher.Invoke(
        //            new Action(() => {
        //                Ogcs.Extensions.MessageBox.Show(windowToBlock, message, "Error", MessageBoxButton.OK, MessageBoxImage.Error);
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
                    Forms.Main.Instance.Console.UpdateWithError(Ogcs.Google.Calendar.GetEventSummary("Event deletion failed.", ev, out String anonSummary, true), ex, logEntry: anonSummary);
                    Ogcs.Exception.Analyse(ex, true);
                    if (Ogcs.Extensions.MessageBox.Show("Google event deletion failed. Continue with synchronisation?", "Sync item failed", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                        continue;
                    else {
                        throw new UserCancelledSyncException("User chose not to continue sync.");
                    }
                }

                try {
                    if (doDelete) deleteCalendarEntry_save(ev);
                    else events.Remove(ev);
                } catch (System.Exception ex) {
                    if (ex is global::Google.GoogleApiException) {
                        global::Google.GoogleApiException gex = ex as global::Google.GoogleApiException;
                        if (gex.Error != null && gex.Error.Code == 410) { //Resource has been deleted
                            log.Fail("This event is already deleted! Ignoring failed request to delete.");
                            continue;
                        }
                    }
                    if (ex is ApplicationException) {
                        String summary = GetEventSummary("<br/>Event deletion skipped.<br/>" + ex.Message, ev, out String anonSummary);
                        Forms.Main.Instance.Console.Update(summary, anonSummary, Console.Markup.warning);
                        if (ex.InnerException is global::Google.GoogleApiException) break;
                        continue;
                    } else {
                        String summary = GetEventSummary("<br/>Event deletion failed.", ev, out String anonSummary);
                        Forms.Main.Instance.Console.UpdateWithError(summary, ex, logEntry: anonSummary);
                        Ogcs.Exception.Analyse(ex, true);
                        if (Ogcs.Extensions.MessageBox.Show("Google event deletion failed. Continue with synchronisation?", "Sync item failed", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                            continue;
                        else
                            throw new UserCancelledSyncException("User chose not to continue sync.");
                    }
                }
            }
        }

        private Boolean deleteCalendarEntry(Event ev) {
            String eventSummary = GetEventSummary(ev, out String anonSummary);
            Boolean doDelete = true;

            if (Sync.Engine.Calendar.Instance.Profile.ConfirmOnDelete) {
                if (Ogcs.Extensions.MessageBox.Show("Delete " + eventSummary + "?", "Confirm Deletion From Google",
                    MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2) == DialogResult.No) {
                    doDelete = false;
                    if (Sync.Engine.Calendar.Instance.Profile.SyncDirection.Id == Sync.Direction.Bidirectional.Id && CustomProperty.ExistAnyOutlookIDs(ev)) {
                        if (Ogcs.Outlook.Calendar.Instance.ExcludedByCategory.ContainsKey(CustomProperty.Get(ev, CustomProperty.MetadataId.oEntryId))) {
                            log.Fine("Refrained from removing Outlook metadata from Event; avoids duplication back into Outlook.");
                        } else {
                            CustomProperty.RemoveOutlookIDs(ref ev);
                            UpdateCalendarEntry_save(ref ev);
                        }
                    }
                    Forms.Main.Instance.Console.Update("Not deleted: " + eventSummary, anonSummary?.Prepend("Not deleted: "), Console.Markup.calendar);
                } else {
                    Forms.Main.Instance.Console.Update("Deleted: " + eventSummary, anonSummary?.Prepend("Deleted: "), Console.Markup.calendar);
                }
            } else {
                Forms.Main.Instance.Console.Update(eventSummary, anonSummary, Console.Markup.calendar, verbose: true);
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
                } catch (global::Google.GoogleApiException ex) {
                    switch (HandleAPIlimits(ref ex, ev)) {
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
                            if (SignaturesMatch(sigEv, Outlook.Calendar.signature(ai))) {
                                try {
                                    Event originalEv = ev;
                                    CustomProperty.AddOutlookIDs(ref ev, ai);
                                    UpdateCalendarEntry_save(ref ev);
                                    unclaimedEvents.Remove(originalEv);
                                    if (consoleTitle != "") Forms.Main.Instance.Console.Update("<span class='em em-reclaim'></span>" + consoleTitle, Console.Markup.h2, newLine: false, verbose: true);
                                    consoleTitle = "";
                                    Forms.Main.Instance.Console.Update(GetEventSummary("Reclaimed: ", ev, out String anonSummary, appendContext: false), anonSummary, verbose: true);
                                    gEvents[g] = ev;
                                    if (profile.SyncDirection.Id == Sync.Direction.Bidirectional.Id || Outlook.CustomProperty.ExistAnyGoogleIDs(ai)) {
                                        log.Debug("Updating the Google event IDs in Outlook appointment.");
                                        Outlook.CustomProperty.AddGoogleIDs(ref ai, ev);
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
                        if (Ogcs.Extensions.MessageBox.Show(unclaimedEvents.Count + " Google calendar events can't be matched to Outlook.\r\n" +
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
                log.Fine("Checking " + Ogcs.Google.Calendar.GetEventSummary(google[g]));

                if (CustomProperty.Exists(google[g], CustomProperty.MetadataId.oEntryId)) {
                    String compare_gEntryID = CustomProperty.Get(google[g], CustomProperty.MetadataId.oEntryId);
                    Boolean outlookIDmissing = CustomProperty.OutlookIdMissing(google[g]);
                    Boolean foundMatch = false;

                    for (int o = outlook.Count - 1; o >= 0; o--) {
                        try {
                            if (log.IsUltraFineEnabled()) log.UltraFine("Checking " + Outlook.Calendar.GetEventSummary(outlook[o]));

                            String compare_oID;
                            if (outlookIDmissing && compare_gEntryID.StartsWith(Outlook.Calendar.GlobalIdPattern)) {
                                //compare_gEntryID actually holds GlobalID up to v2.3.2.3 - yes, confusing I know, but we're sorting this now
                                compare_oID = Outlook.Calendar.Instance.IOutlook.GetGlobalApptID(outlook[o]);
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
                    if (!foundMatch && profile.MergeItems &&
                        Ogcs.Google.CustomProperty.Get(google[g], CustomProperty.MetadataId.oCalendarId) != profile.UseOutlookCalendar.Id)
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

            if (google.Count > 0 && Outlook.Calendar.Instance.ExcludedByCategory?.Count > 0 && !profile.DeleteWhenCategoryExcluded) {
                //Check if Google items to be deleted were filtered out from Outlook
                for (int g = google.Count - 1; g >= 0; g--) {
                    if (Outlook.Calendar.Instance.ExcludedByCategory.ContainsValue(google[g].Id) ||
                        Outlook.Calendar.Instance.ExcludedByCategory.ContainsKey(CustomProperty.Get(google[g], CustomProperty.MetadataId.oEntryId) ?? "")) {
                        google.Remove(google[g]);
                    }
                }
            }
            if (outlook.Count > 0 && Ogcs.Google.Calendar.Instance.ExcludedByColour?.Count > 0) {
                //Check if Outlook items to be created were filtered out from Google
                for (int o = outlook.Count - 1; o >= 0; o--) {
                    if (ExcludedByColour.ContainsValue(outlook[o].EntryID) ||
                        ExcludedByColour.ContainsKey(Outlook.CustomProperty.Get(outlook[o], Outlook.CustomProperty.MetadataId.gEventID) ?? "")) {
                        outlook.Remove(outlook[o]);
                    }
                }
            }

            if (profile.SyncDirection.Id == Sync.Direction.Bidirectional.Id) {
                //Don't recreate any items that have been deleted in Google
                for (int o = outlook.Count - 1; o >= 0; o--) {
                    if (Outlook.CustomProperty.Exists(outlook[o], Outlook.CustomProperty.MetadataId.gEventID))
                        outlook.Remove(outlook[o]);
                }
                //Don't delete any items that aren't yet in Outlook or just created in Outlook during this sync
                for (int g = google.Count - 1; g >= 0; g--) {
                    if (!CustomProperty.Exists(google[g], CustomProperty.MetadataId.oEntryId) ||
                        google[g].Updated > Sync.Engine.Instance.SyncStarted)
                        google.Remove(google[g]);
                }
            }
            if (profile.DisableDelete) {
                if (google.Count > 0) {
                    Forms.Main.Instance.Console.Update(google.Count + " Google items would have been deleted, but you have deletions disabled.", Console.Markup.warning);
                    for (int g = 0; g < google.Count; g++)
                        Forms.Main.Instance.Console.Update(GetEventSummary(google[g], out String anonSummary), anonSummary, verbose: true);
                }
                google = new List<Event>();
            }
            if (Settings.Instance.CreateCSVFiles) {
                ExportToCSV("Events for deletion in Google", "google_delete.csv", google);
                Outlook.Calendar.ExportToCSV("Appointments for creation in Google", "google_create.csv", outlook);
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
                String oGlobalID = Outlook.Calendar.Instance.IOutlook.GetGlobalApptID(ai);

                //For items copied from someone elses calendar, it appears the Global ID is generated for each access?! (Creation Time changes)
                //I guess the copied item doesn't really have its "own" ID. So, we'll just compare
                //the "data" section of the byte array, which "ensures uniqueness" and doesn't include ID creation time

                if ((Outlook.Factory.OutlookVersionName == Outlook.Factory.OutlookVersionNames.Outlook2003 && oGlobalID == gCompareID) //Actually simple compare of EntryId for O2003
                    ||
                    (Outlook.Factory.OutlookVersionName != Outlook.Factory.OutlookVersionNames.Outlook2003 &&
                        (
                            (oGlobalID.StartsWith(Outlook.Calendar.GlobalIdPattern) &&
                            gCompareID.StartsWith(Outlook.Calendar.GlobalIdPattern) &&
                            gCompareID.Substring(72) == oGlobalID.Substring(72))             //We've got bonafide Global IDs match
                            ||
                            (!oGlobalID.StartsWith(Outlook.Calendar.GlobalIdPattern) &&
                            !gCompareID.StartsWith(Outlook.Calendar.GlobalIdPattern) &&
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
                            }
                        }
                    }

                } else if (profile.SyncDirection.Id == Sync.Direction.Bidirectional.Id &&
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
            List<global::Google.Apis.Calendar.v3.Data.EventAttendee> removeAttendee = new List<global::Google.Apis.Calendar.v3.Data.EventAttendee>();
            foreach (global::Google.Apis.Calendar.v3.Data.EventAttendee ea in ev.Attendees ?? Enumerable.Empty<global::Google.Apis.Calendar.v3.Data.EventAttendee>()) {
                removeAttendee.Add(ea);
            }
            if (ai.Recipients.Count > 1) {
                for (int o = ai.Recipients.Count; o > 0; o--) {
                    bool foundAttendee = false;
                    Recipient recipient = ai.Recipients[o];
                    log.Fine("Comparing Outlook recipient: " + recipient.Name);
                    String recipientSMTP = Outlook.Calendar.Instance.IOutlook.GetRecipientEmail(recipient);
                    foreach (global::Google.Apis.Calendar.v3.Data.EventAttendee attendee in ev.Attendees ?? Enumerable.Empty<global::Google.Apis.Calendar.v3.Data.EventAttendee>()) {
                        Ogcs.Google.EventAttendee ogcsAttendee = new Ogcs.Google.EventAttendee(attendee);
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
                        if (ev.Attendees == null) ev.Attendees = new List<global::Google.Apis.Calendar.v3.Data.EventAttendee>();
                        ev.Attendees.Add(Ogcs.Google.Calendar.CreateAttendee(recipient, ai.Organizer == recipient.Name));
                        itemModified++;
                    }
                }
            } //more than just 1 (me) recipients

            foreach (global::Google.Apis.Calendar.v3.Data.EventAttendee gea in removeAttendee) {
                Ogcs.Google.EventAttendee ea = new Ogcs.Google.EventAttendee(gea);
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
                } catch (global::Google.GoogleApiException ex) {
                    switch (HandleAPIlimits(ref ex, null)) {
                        case ApiException.throwException: throw;
                        case ApiException.freeAPIexhausted:
                            Ogcs.Exception.LogAsFail(ref ex);
                            ex.Analyse("Not able to " + stage);
                            System.ApplicationException aex = new System.ApplicationException(SubscriptionInvite, ex);
                            Ogcs.Exception.LogAsFail(ref aex);
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
                    ex.Analyse("Not able to " + stage);
                    throw new System.ApplicationException("Unable to " + stage + ".", ex);

                } catch (System.Exception ex) {
                    ex.Analyse("Not able to " + stage);
                    throw;
                }
            }
        }
        private void getCalendarSettings() {
            SettingsStore.Calendar profile = Settings.Profile.InPlay();
            CalendarListResource.GetRequest request = Service.CalendarList.Get(profile.UseGoogleCalendar.Id);
            CalendarListEntry cal;
            try {
                cal = request.Execute();
            } catch (global::Google.GoogleApiException ex) {
                if (ex.InnerException is Newtonsoft.Json.JsonReaderException && ex.InnerException.Message.Contains("Unexpected character encountered while parsing value: <") && Settings.Instance.Proxy.Type != "None") {
                    log.Warn("Call to CalendarList API endpoint failed. Retrying with trailing '/' in case of poorly configured proxy.");
                    //The URI ends with "@group.calendar.google.com", which seemingly can cause confusion - see issue #1745
                    try {
                        System.Net.Http.HttpRequestMessage hrm = request.CreateRequest();
                        hrm.RequestUri = new System.Uri(hrm.RequestUri + "/");
                        System.Net.Http.HttpResponseMessage response = Service.HttpClient.SendAsync(hrm).Result;
                        String responseBody = response.Content.ReadAsStringAsync().Result;
                        cal = Newtonsoft.Json.JsonConvert.DeserializeObject<CalendarListEntry>(responseBody);
                    } catch (System.Exception ex2) {
                        ex2.Analyse("Failed retrieving calendarList via HttpRequestMessage.");
                        throw;
                    }
                } else throw;
            }
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

            String overridePrivacy = "public";
            try {
                Enum.TryParse(profile.PrivacyLevel,  out OlSensitivity olOverridePrivacy);
                overridePrivacy = olOverridePrivacy == OlSensitivity.olPrivate ? "private" : "public";
            } catch (System.Exception ex) {
                ex.Analyse("Could not convert string '" + profile.PrivacyLevel + "' to OlSensitivity type. Defaulting override to normal.");
            }

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
                    return overridePrivacy;
                else {
                    if (profile.CreatedItemsOnly) return gVisibility;
                    else return overridePrivacy;
                }
            }
        }

        /// <summary>
        /// Determine Event's availability setting
        /// </summary>
        /// <param name="oBusyStatus">Outlook's current setting</param>
        /// <param name="gTransparency">Google's current setting</param>
        private String getAvailability(OlBusyStatus oBusyStatus, String gTransparency) {
            SettingsStore.Calendar profile = Sync.Engine.Calendar.Instance.Profile;

            if (!profile.SetEntriesAvailable)
                return (oBusyStatus == OlBusyStatus.olFree) ? "transparent" : "opaque";

            String overrideTransparency = "opaque";
            OlBusyStatus fbStatus = OlBusyStatus.olBusy;
            try {
                Enum.TryParse(profile.AvailabilityStatus, out fbStatus);
                if (fbStatus == OlBusyStatus.olFree)
                    overrideTransparency = "transparent";
            } catch (System.Exception ex) {
                ex.Analyse("Could not convert string '" + profile.AvailabilityStatus + "' to OlBusyStatus type. Defaulting override to busy.");
            }

                if (profile.TargetCalendar.Id == Sync.Direction.GoogleToOutlook.Id) { //Availability enforcement is in other direction
                    if (gTransparency == null)
                        return (oBusyStatus == OlBusyStatus.olFree) ? "transparent" : "opaque";
                    else
                        return gTransparency;
                } else {
                    if (!profile.CreatedItemsOnly || (profile.CreatedItemsOnly && gTransparency == null))
                        return overrideTransparency;
                else {
                    if (profile.CreatedItemsOnly) return gTransparency;
                    else return overrideTransparency;
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
            EventColour.Palette overrideColour = this.ColourPalette.ActivePalette[Convert.ToInt16(profile.SetEntriesColourGoogleId)];

            if (profile.SetEntriesColour) {
                if (profile.TargetCalendar.Id == Sync.Direction.GoogleToOutlook.Id) { //Colour forced to sync in other direction
                    if (gColour == null) //Creating item
                        return EventColour.Palette.NullPalette;
                    else return gColour;

                } else {
                    if (!profile.CreatedItemsOnly || (profile.CreatedItemsOnly && gColour == null))
                        return overrideColour;
                    else {
                        if (profile.CreatedItemsOnly) return gColour;
                        else return overrideColour;
                    }
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
                    sigAi = Obfuscate.ApplyRegex(Obfuscate.Property.Subject, sigAi, null, Sync.Direction.OutlookToGoogle);
                else
                    sigEv = Obfuscate.ApplyRegex(Obfuscate.Property.Subject, sigEv, null, Sync.Direction.GoogleToOutlook);
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
                ex.Analyse("Failed to backup previous CSV file.");
            }

            Stream stream = null;
            TextWriter tw = null;
            try {
                try {
                    stream = new FileStream(Path.Combine(Program.UserFilePath, filename), FileMode.Create, FileAccess.Write);
                    tw = new StreamWriter(stream, Encoding.UTF8);
                } catch (System.Exception ex) {
                    Forms.Main.Instance.Console.Update("Failed to create CSV file '" + filename + "'.", Console.Markup.error);
                    ex.Analyse("Error opening file '" + filename + "' for writing.");
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
                            Forms.Main.Instance.Console.Update(GetEventSummary("Failed to output following Google event to CSV:-<br/>", ev, out String anonSummary, appendContext: false), anonSummary, Console.Markup.warning);
                            Ogcs.Exception.Analyse(ex, true);
                        }
                    }
                } catch (System.Exception ex) {
                    Forms.Main.Instance.Console.Update("Failed to output Google events to CSV.", Console.Markup.error);
                    Ogcs.Exception.Analyse(ex);
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
                foreach (global::Google.Apis.Calendar.v3.Data.EventAttendee ea in ev.Attendees) {
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
        /// Get the anonymised summary of an Event item, else standard summary.
        /// </summary>
        /// <param name="ev">The Event item.</param>
        /// <returns>The summary, anonymised if settings dictate.</returns>
        public static String GetEventSummary(Event ev) {
            String eventSummary = GetEventSummary(ev, out String anonymisedSummary, false);
            return anonymisedSummary ?? eventSummary;
        }

        /// <summary>
        /// Pre/Append context to the summary of an Event item.
        /// </summary>
        /// <param name="context">Text to add before/after the summary and anonymised summary.</param>
        /// <param name="ev">The Event item.</param>
        /// <param name="eventSummaryAnonymised">The anonymised summary with context.</param>
        /// <param name="onlyIfNotVerbose">Only return if user doesn't have Verbose output on. Useful for indicating offending item during errors.</param>
        /// <param name="appendContext">If the context should be before or after.</param>
        /// <returns>The standard summary.</returns>
        public static string GetEventSummary(String context, Event ev, out String eventSummaryAnonymised, Boolean onlyIfNotVerbose = false, Boolean appendContext = true) {
            String eventSummary = GetEventSummary(ev, out String anonymisedSummary, onlyIfNotVerbose);
            if (appendContext) {
                eventSummary = eventSummary + context;
                eventSummaryAnonymised = anonymisedSummary?.Append(context);
            } else {
                eventSummary = context + eventSummary;
                eventSummaryAnonymised = anonymisedSummary?.Prepend(context);
            }
            return eventSummary;
        }


        /// <summary>
        /// Get the summary of an Event.
        /// </summary>
        /// <param name="ev">The event.</param>
        /// <param name="eventSummaryAnonymised">Anonymised version of the returned summary string value.</param>
        /// <param name="onlyIfNotVerbose">Only return if user doesn't have Verbose output on. Useful for indicating offending item during errors.</param>
        /// <returns></returns>
        public static String GetEventSummary(Event ev, out String eventSummaryAnonymised, Boolean onlyIfNotVerbose = false) {
            String eventSummary = "";
            eventSummaryAnonymised = null;
            if (!onlyIfNotVerbose || onlyIfNotVerbose && !Settings.Instance.VerboseOutput) {
                try {
                    if (ev.Start.DateTime != null) {
                        System.DateTime gDate = (System.DateTime)ev.Start.DateTime;
                        eventSummary += gDate.ToShortDateString() + " " + gDate.ToShortTimeString();
                    } else
                        eventSummary += System.DateTime.Parse(ev.Start.Date).ToShortDateString();
                    if ((ev.Recurrence != null && ev.RecurringEventId == null) || ev.RecurringEventId != null)
                        eventSummary += " (R)";

                    if (Settings.Instance.AnonymiseLogs)
                        eventSummaryAnonymised = eventSummary + " => \"" + Authenticator.GetMd5(ev.Summary, silent: true) + "\"" + (onlyIfNotVerbose ? "<br/>" : "");
                    eventSummary += " => \"" + ev.Summary + "\"" + (onlyIfNotVerbose ? "<br/>" : "");

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

        public static global::Google.Apis.Calendar.v3.Data.EventAttendee CreateAttendee(Recipient recipient, Boolean isOrganiser) {
            Ogcs.Google.EventAttendee ea = new Ogcs.Google.EventAttendee();
            log.Fine("Creating attendee " + recipient.Name);
            ea.DisplayName = recipient.Name;
            ea.Email = Outlook.Calendar.Instance.IOutlook.GetRecipientEmail(recipient);
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

        public static ApiException HandleAPIlimits(ref global::Google.GoogleApiException ex, Event ev) {
            //https://developers.google.com/analytics/devguides/reporting/core/v3/coreErrors

            log.Fail(ex.Message);

            try {
                new Telemetry.GA4Event.Event(Telemetry.GA4Event.Event.Name.ogcs_error)
                    .AddParameter("api_google_error", ex.Message)
                    .AddParameter("code", ex.Error?.Code)
                    .AddParameter("domain", ex.Error?.Errors?.First().Domain)
                    .AddParameter("reason", ex.Error?.Errors?.First().Reason)
                    .AddParameter("message", ex.Error?.Errors?.First().Message)
                    .Send();
            } catch (System.Exception gaEx) {
                Ogcs.Exception.Analyse(gaEx);
            }

            SettingsStore.Calendar profile = Sync.Engine.Calendar.Instance.Profile;

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
                if (ex.Message.Contains("Calendar usage limits exceeded") && profile.AddAttendees && ev != null) {
                    //"global::Google.Apis.Requests.RequestError\r\nCalendar usage limits exceeded. [403]\r\nErrors [\r\n\tMessage[Calendar usage limits exceeded.] Location[ - ] Reason[quotaExceeded] Domain[usageLimits]\r\n]\r\n"
                    //This happens because too many attendees have been added in a short period of time.
                    //See https://support.google.com/a/answer/2905486?hl=en-uk&hlrm=en

                    Forms.Main.Instance.Console.Update("You have added enough meeting attendees to have reached the Google API limit.<br/>" +
                        "Don't worry, this only lasts for an hour or two, but until then attendees will not be synced.", Console.Markup.warning);

                    APIlimitReached_attendee = true;
                    Settings.Instance.APIlimit_inEffect = true;
                    Settings.Instance.APIlimit_lastHit = System.DateTime.Now;

                    ev.Attendees = new List<global::Google.Apis.Calendar.v3.Data.EventAttendee>();
                    return ApiException.backoffThenRetry;

                } else if (ex.Error.Errors.First().Reason == "rateLimitExceeded") {
                    if (ex.Message.Contains("limit 'Queries per minute'")) {
                        log.Fail(ex.FriendlyMessage());
                        Ogcs.Exception.LogAsFail(ref ex);
                        return ApiException.backoffThenRetry;

                    } else if (ex.Message.Contains("limit 'Queries per day'") || ex.Message.Contains("Daily Limit Exceeded")) {
                        log.Warn("Google's free Calendar quota has been exhausted! New quota comes into effect 08:00 GMT.");
                        Forms.Main.Instance.SyncNote(Forms.Main.SyncNotes.DailyQuotaExhaustedInfo, null);

                        //Delay next scheduled sync until after the new quota
                        if (profile.SyncInterval != 0) {
                            System.DateTime utcNow = System.DateTime.UtcNow;
                            System.DateTime quotaReset = utcNow.Date.AddHours(8).AddMinutes(utcNow.Minute);
                            if ((quotaReset - utcNow).Ticks < 0) quotaReset = quotaReset.AddDays(1);
                            int delayMins = (int)(quotaReset - utcNow).TotalMinutes;
                            profile.OgcsTimer.SetNextSync(delayMins, fromNow: true, calculateInterval: false);
                            Forms.Main.Instance.Console.Update("The next sync has been delayed by " + delayMins + " minutes, when new quota is available.", Console.Markup.warning);
                        }
                        return ApiException.freeAPIexhausted;

                    } else if (ex.Message.Contains("Rate Limit Exceeded")) {
                        if (Settings.Instance.Subscribed > System.DateTime.Now.AddYears(-1))
                            return ApiException.backoffThenRetry;

                        log.Warn("Google's free Calendar quota is being exceeded!");
                        Forms.Main.Instance.SyncNote(Forms.Main.SyncNotes.QuotaExceededInfo, null);

                        //Delay next scheduled sync for an hour
                        if (profile.SyncInterval != 0) {
                            System.DateTime utcNow = System.DateTime.UtcNow;
                            System.DateTime nextSync = utcNow.AddMinutes(60 + new Random().Next(1, 10));
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
                    request = Ogcs.Google.Calendar.Instance.Service.Events.Update(ev, profile.UseGoogleCalendar.Id, ev.Id);
                    request.ETagAction = global::Google.Apis.ETagAction.Ignore;
                    request.SendUpdates = EventsResource.UpdateRequest.SendUpdatesEnum.None;
                    ev = request.Execute();
                    log.Debug("Successfully forced save by ignoring eTag values.");
                } catch (System.Exception ex2) {
                    try {
                        ex2.LogAsFail().Analyse("Failed forcing save with ETagAction.Ignore");
                        log.Fine("Current eTag: " + ev.ETag);
                        log.Fine("Current Updated: " + ev.UpdatedRaw);
                        log.Fine("Current Sequence: " + ev.Sequence);
                        log.Debug("Refetching event from Google.");
                        Event remoteEv = Ogcs.Google.Calendar.Instance.GetCalendarEntry(ev.Id);
                        log.Fine("Remote eTag: " + remoteEv.ETag);
                        log.Fine("Remote Updated: " + remoteEv.UpdatedRaw);
                        log.Fine("Remote Sequence: " + remoteEv.Sequence);
                        log.Warn("Attempting trample of remote version...");
                        ev.ETag = remoteEv.ETag;
                        ev.Sequence = remoteEv.Sequence;
                        request = Ogcs.Google.Calendar.Instance.Service.Events.Update(ev, profile.UseGoogleCalendar.Id, ev.Id);
                        request.SendUpdates = EventsResource.UpdateRequest.SendUpdatesEnum.None;
                        ev = request.Execute();
                        log.Debug("Successful!");
                    } catch {
                        return ApiException.throwException;
                    }
                }
                return ApiException.justContinue;

            } else if (ex.Error?.Code == 500) {
                log.Fail(ex.FriendlyMessage());
                Ogcs.Exception.LogAsFail(ref ex);
                return ApiException.backoffThenRetry;
            }

            log.Warn("Unhandled API exception.");
            return ApiException.throwException;
        }

        public static Boolean? IsDefaultCalendar() {
            try {
                SettingsStore.Calendar profile = Sync.Engine.Calendar.Instance.Profile;
                if (!Settings.InstanceInitialiased || (profile?.UseGoogleCalendar?.Id == null || string.IsNullOrEmpty(Settings.Instance.GaccountEmail)))
                    return null;

                return profile.UseGoogleCalendar.Id == Settings.Instance.GaccountEmail;
            } catch (System.Exception ex) {
                Ogcs.Exception.Analyse(ex);
                return null;
            }
        }

        /// <summary>
        /// Build colour list from any saved in Settings, instead of downloading from Google.
        /// </summary>
        /// <param name="clb">The checklistbox to populate with the colours.</param>
        public static void BuildOfflineColourPicker(System.Windows.Forms.CheckedListBox clb) {
            if (IsInstanceNull || !Instance.Authenticator.Authenticated) {
                clb.BeginUpdate();
                clb.Items.Clear();
                foreach (String colour in Forms.Main.Instance.ActiveCalendarProfile.Colours) {
                    clb.Items.Add(colour, true);
                }
                clb.EndUpdate();
            } else {
                Instance.ColourPalette.BuildPicker(clb);
            }
        }
        #endregion

        /// <summary>
        /// This is solely for purposefully causing an error to assist when developing
        /// </summary>
        public void ThrowApiException() {
            global::Google.GoogleApiException ex = new global::Google.GoogleApiException("Service", "Rate Limit Exceeded");
            global::Google.Apis.Requests.SingleError err = new global::Google.Apis.Requests.SingleError { Domain = "usageLimits", Reason = "rateLimitExceeded" };
            ex.Error = new global::Google.Apis.Requests.RequestError { Errors = new List<global::Google.Apis.Requests.SingleError>(), Code = 403 };
            ex.Error.Errors.Add(err);
            throw ex;
        }
    }
}
