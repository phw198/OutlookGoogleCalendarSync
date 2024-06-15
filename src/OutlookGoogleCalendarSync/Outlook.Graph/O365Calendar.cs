using log4net;
using Microsoft.Graph;
using OutlookGoogleCalendarSync.Extensions;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;
using GcalData = Google.Apis.Calendar.v3.Data;
using Ogcs = OutlookGoogleCalendarSync;

namespace OutlookGoogleCalendarSync.Outlook.Graph {
    public class Calendar {
        private static readonly ILog log = LogManager.GetLogger(typeof(Calendar));

        private static Calendar instance;
        public static Boolean IsInstanceNull { get { return instance == null; } }
        public static Calendar Instance {
            get {
                if (instance == null) {
                    instance = new Ogcs.Outlook.Graph.Calendar {
                        Authenticator = new Ogcs.Outlook.Graph.Authenticator()
                    };
                }
                return instance;
            }
        }
        public Calendar() { }

        public Ogcs.Outlook.Graph.Authenticator Authenticator;
        private GraphServiceClient graphClient;
        public GraphServiceClient GraphClient {
            get {
                if (graphClient == null || !(Authenticator?.Authenticated ?? false)) {
                    log.Debug("MS Graph service not yet instantiated.");
                    Authenticator = new Ogcs.Outlook.Graph.Authenticator();
                    Authenticator.GetAuthenticated(nonInteractiveAuth: false);
                } else if (Authenticator.AgedAccessToken) {
                    log.Debug("MS Graph access token expired - refreshing...");
                    Authenticator.GetAuthenticated(nonInteractiveAuth: false);
                }
                if (!Authenticator.Authenticated) {
                    graphClient = null;
                    throw new ApplicationException("Microsoft handshake failed.");
                }
                return graphClient;
            }
            set { graphClient = value; }
        }

        public Graph.EphemeralProperties EphemeralProperties = new Graph.EphemeralProperties();

        private Dictionary<String, OutlookCalendarListEntry> calendarFolders = new Dictionary<string, OutlookCalendarListEntry>();
        public Dictionary<String, OutlookCalendarListEntry> CalendarFolders {
            get { return calendarFolders; }
        }

        /// <summary>Retrieve calendar list from the cloud.</summary>
        public Dictionary<String, OutlookCalendarListEntry> GetCalendars() {
            calendarFolders = new();
            List<Microsoft.Graph.Calendar> cals = new();

            var graphThread = new System.Threading.Thread(() => {
                try {
                    Microsoft.Graph.IUserCalendarsCollectionPage calPage = GraphClient.Me.Calendars.Request().GetAsync().Result;
                    cals.AddRange(calPage.CurrentPage);
                    while (calPage.NextPageRequest != null) {
                        calPage = calPage.NextPageRequest.GetAsync().Result;
                        cals.AddRange(calPage.CurrentPage);
                    }
                } catch (System.Exception ex) {
                    log.Debug(ex.ToString());
                }
            });
            graphThread.Start();
            while (graphThread.IsAlive) {
                System.Windows.Forms.Application.DoEvents();
                System.Threading.Thread.Sleep(250);
            }

            foreach (Microsoft.Graph.Calendar cal in cals) {
                if (cal.AdditionalData.ContainsKey("isDefaultCalendar") && (Boolean)cal.AdditionalData["isDefaultCalendar"])
                    cal.Name = "Default " + cal.Name;
                log.Debug(cal.Name);
                calendarFolders.Add(cal.Name, new OutlookCalendarListEntry(cal));
            }

            return calendarFolders;
        }

        /// <summary>
        /// Get all calendar entries within the defined date-range for sync
        /// </summary>
        /// <param name="suppressAdvisories">Don't give user feedback, eg during background Push sync</param>
        /// <returns></returns>
        public List<Microsoft.Graph.Event> GetCalendarEntriesInRange(SettingsStore.Calendar profile, Boolean suppressAdvisories) {
            List<Microsoft.Graph.Event> filtered = new List<Microsoft.Graph.Event>();
            try {
                filtered = FilterCalendarEntries(profile, suppressAdvisories: suppressAdvisories);
            } catch (System.Exception) {
                if (!suppressAdvisories) Forms.Main.Instance.Console.Update("Unable to access the Outlook calendar.", Console.Markup.error);
                throw;
            }
            return filtered;
        }

        public List<Microsoft.Graph.Event> FilterCalendarEntries(SettingsStore.Calendar profile, Boolean filterBySettings = true,
            Boolean noDateFilter = false, String extraFilter = "", Boolean suppressAdvisories = false) {
            //Filtering info @ https://msdn.microsoft.com/en-us/library/cc513841%28v=office.12%29.aspx

            List<Microsoft.Graph.Event> result = new List<Microsoft.Graph.Event>();
            //Items OutlookItems = null;
            List<Microsoft.Graph.Event> OutlookItems = new();
            //ExcludedByCategory = new();

            profile ??= Settings.Profile.InPlay();

            try {
                //MAPIFolder thisUseOutlookCalendar = IOutlook.GetFolderByID(profile.UseOutlookCalendar.Id);
                //OutlookItems = thisUseOutlookCalendar.Items;

                // Code snippets are only available for the latest version. Current version is 5.x

                // To initialize your graphClient, see https://learn.microsoft.com/en-us/graph/sdks/create-client?from=snippets&tabs=csharp
                //.e.Events.GetAsync((requestConfiguration) =>
                //{
                //    requestConfiguration.QueryParameters.Select = new string[] { "subject", "body", "bodyPreview", "organizer", "attendees", "start", "end", "location" };
                //    requestConfiguration.Headers.Add("Prefer", "outlook.timezone=\"Pacific Standard Time\"");
                //});

                try {
                    Int16 pageNum = 1;
                    ICalendarEventsCollectionRequest req = GraphClient.Me.Calendars[profile.UseOutlookCalendar.Id].Events.Request();

                    System.DateTime min = System.DateTime.MinValue;
                    System.DateTime max = System.DateTime.MaxValue;
                    if (!noDateFilter) {
                        min = profile.SyncStart;
                        max = profile.SyncEnd;
                    }

                    string filter = "end/dateTime ge '" + min.ToString("yyyy-MM-dd") +
                        "' and start/dateTime lt '" + max.ToString("yyyy-MM-dd") + "'" + extraFilter;
                    log.Fine("Filter string: " + filter);
                    req.Filter(filter);

                    req.Top(250);
                    //req.OrderBy("start");

                    ICalendarEventsCollectionPage eventPage = req.GetAsync().Result;
                    //IUserEventsCollectionPage eventPage = GraphClient.Me.Events.Request().Top(250)..GetAsync().Result;
                    OutlookItems.AddRange(eventPage.CurrentPage);
                    while (eventPage.NextPageRequest != null) {
                        pageNum++;
                        eventPage = eventPage.NextPageRequest.GetAsync().Result;
                        log.Debug("Page " + pageNum + " received.");
                        OutlookItems.AddRange(eventPage.CurrentPage);
                    }
                } catch (System.Exception ex) {
                    log.Debug(ex.ToString());
                }
            } catch {
                log.Fail("Could not open '" + Settings.Profile.Name(profile) + "' profile calendar folder with ID " + profile.UseOutlookCalendar.Id);
                throw;
            }

            if (OutlookItems != null) {
                log.Fine(OutlookItems.Count + " calendar items exist.");
/*              
                Int32 allDayFiltered = 0;
                Int32 availabilityFiltered = 0;
                Int32 privacyFiltered = 0;
                Int32 subjectFiltered = 0;
                Int32 responseFiltered = 0;

                foreach (Object obj in IOutlook.FilterItems(OutlookItems, filter)) {
                    AppointmentItem ai;
                    try {
                        ai = obj as AppointmentItem;
                    } catch {
                        log.Warn("Encountered a non-appointment item in the calendar.");
                        if (obj is MeetingItem) log.Debug("It is a meeting item.");
                        else if (obj is MailItem) log.Debug("It is a mail item.");
                        else if (obj is ContactItem) log.Debug("It is a contact item.");
                        else if (obj is TaskItem) log.Debug("It is a task item.");
                        else log.Debug("WTF is this item?!");
                        continue;
                    }
                    try {
                        if (ai.End == min) continue; //Required for midnight to midnight events 
                    } catch (System.NullReferenceException) {
                        log.Debug("NullReferenceException accessing ai.End");
                        try {
                            System.DateTime start = ai.Start;
                        } catch (System.NullReferenceException) {
                            try { log.Debug("Subject: " + ai.Subject); } catch { }
                            log.Fail("Appointment item seems unusable - no Start or End date! Discarding.");
                            continue;
                        }
                        log.Debug("Unable to get End date for: " + GetEventSummary(ai));
                        continue;

                    } catch (System.Exception ex) {
                        Ogcs.Exception.Analyse(ex, true);
                        log.Debug("Unable to get End date for: " + GetEventSummary(ai));
                        continue;
                    }

                    if (!filterBySettings) result.Add(ai);
                    else {
                        Boolean filtered = false;

                        try {
                            //Categories
                            try {
                                if (profile.CategoriesRestrictBy == SettingsStore.Calendar.RestrictBy.Include) {
                                    filtered = (profile.Categories.Count() == 0 || (ai.Categories == null && !profile.Categories.Contains("<No category assigned>")) ||
                                        (ai.Categories != null && ai.Categories.Split(new[] { Categories.Delimiter }, StringSplitOptions.None).Intersect(profile.Categories).Count() == 0));

                                } else if (profile.CategoriesRestrictBy == SettingsStore.Calendar.RestrictBy.Exclude) {
                                    filtered = (profile.Categories.Count() > 0 && (ai.Categories == null && profile.Categories.Contains("<No category assigned>")) ||
                                        (ai.Categories != null && ai.Categories.Split(new[] { Categories.Delimiter }, StringSplitOptions.None).Intersect(profile.Categories).Count() > 0));
                                }
                            } catch (System.Runtime.InteropServices.COMException ex) {
                                if (ex.TargetSite.Name == "get_Categories") {
                                    log.Warn("Could not access Categories property for " + GetEventSummary(ai));
                                    filtered = ((profile.CategoriesRestrictBy == SettingsStore.Calendar.RestrictBy.Include && !profile.Categories.Contains("<No category assigned>")) ||
                                        (profile.CategoriesRestrictBy == SettingsStore.Calendar.RestrictBy.Exclude && profile.Categories.Contains("<No category assigned>")));
                                } else throw;
                            }
                            if (filtered) { ExcludedByCategory.Add(ai.EntryID, CustomProperty.Get(ai, CustomProperty.MetadataId.gEventID)); continue; }

                            //Availability, Privacy, Subject
                            if (profile.SyncDirection.Id != Sync.Direction.GoogleToOutlook.Id) { //Sync direction means O->G will delete previously synced excluded items
                                if (filtered = ((profile.ExcludeTentative && ai.BusyStatus == OlBusyStatus.olTentative) ||
                                    (profile.ExcludeFree && ai.BusyStatus == OlBusyStatus.olFree))) {
                                    availabilityFiltered++; continue;
                                }

                                if (profile.ExcludeAllDays && ai.AllDayEvent(true)) {
                                    if (profile.ExcludeFreeAllDays)
                                        filtered = (ai.BusyStatus == OlBusyStatus.olFree);
                                    else
                                        filtered = true;
                                    if (filtered) { allDayFiltered++; continue; }
                                }

                                if (filtered = profile.ExcludePrivate && ai.Sensitivity == OlSensitivity.olPrivate) {
                                    privacyFiltered++; continue;
                                }

                                if (profile.ExcludeSubject && !String.IsNullOrEmpty(profile.ExcludeSubjectText)) {
                                    Regex rgx = new Regex(profile.ExcludeSubjectText, RegexOptions.IgnoreCase);
                                    if (rgx.IsMatch(ai.Subject ?? "")) {
                                        log.Fine("Regex has matched subject string: " + profile.ExcludeSubjectText);
                                        subjectFiltered++; continue;
                                    }
                                }
                            }

                            //Invitation
                            if (profile.OnlyRespondedInvites) {
                                //These are actually filtered out later on when identifying differences
                                if (filtered = ai.ResponseStatus == OlResponseStatus.olResponseNotResponded)
                                    responseFiltered++;
                            }
                        } finally {
                            if (filtered && profile.SyncDirection.Id == Sync.Direction.Bidirectional.Id && CustomProperty.ExistAnyGoogleIDs(ai)) {
                                log.Debug("Previously synced Outlook item is now excluded. Removing Google metadata.");
                                CustomProperty.RemoveGoogleIDs(ref ai);
                                ai.Save();
                            }
                        }

                        result.Add(ai);
                    }
                }
                if (!suppressAdvisories) {
                    if (availabilityFiltered > 0) log.Info(availabilityFiltered + " Outlook items excluded due to availability.");
                    if (allDayFiltered > 0) log.Info(allDayFiltered + " Outlook all day items excluded.");
                    if (ExcludedByCategory.Count > 0) log.Info(ExcludedByCategory.Count + " Outlook items contain a category that is filtered out.");
                    if (subjectFiltered > 0) log.Info(subjectFiltered + " Outlook items with subject containing '" + profile.ExcludeSubjectText + "' filtered out.");
                    if (responseFiltered > 0) log.Info(responseFiltered + " Outlook items are invites not yet responded to.");

                    if ((availabilityFiltered + allDayFiltered + ExcludedByCategory.Count + subjectFiltered + responseFiltered) > 0) {
                        if (result.Count == 0)
                            Forms.Main.Instance.Console.Update("Due to your OGCS Outlook settings, all Outlook items have been filtered out!", Console.Markup.config, notifyBubble: true);
                        else if (profile.SyncDirection.Id == Sync.Direction.GoogleToOutlook.Id)
                            Forms.Main.Instance.Console.Update("Due to your OGCS Outlook settings, Outlook items have been filtered out. " +
                                "If they exist in Google, they may be synced and appear as \"duplicates\".", Console.Markup.config);
                    }
                }*/
            }
            log.Fine("Filtered down to " + result.Count);
            return OutlookItems; // result;
        }

        #region Create
        public void CreateCalendarEntries(List<GcalData.Event> events) {
            for (int g = 0; g < events.Count; g++) {
                if (Sync.Engine.Instance.CancellationPending) return;

                GcalData.Event ev = events[g];
                Microsoft.Graph.Event newAi = new();
                //try {
                    try {
                        createCalendarEntry(ev, ref newAi);
                    } catch (System.Exception ex) {
                        if (ex.GetType() == typeof(ApplicationException)) {
                            Forms.Main.Instance.Console.Update(Ogcs.Google.Calendar.GetEventSummary("Appointment creation skipped: " + ex.Message, ev, out String anonSummary, true), anonSummary, Console.Markup.warning);
                            continue;
                        } else {
                            Forms.Main.Instance.Console.UpdateWithError(Ogcs.Google.Calendar.GetEventSummary("Appointment creation failed.", ev, out String anonSummary, true), ex, logEntry: anonSummary);
                            Ogcs.Exception.Analyse(ex, true);
                            if (Ogcs.Extensions.MessageBox.Show("Outlook appointment creation failed. Continue with synchronisation?", "Sync item failed", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                                continue;
                            else
                                throw new UserCancelledSyncException("User chose not to continue sync.");
                        }
                    }

                Event createdEvent = new Event();
                try {
                        /*createdEvent =*/ createCalendarEntry_save(newAi, ref ev);
                        events[g] = ev;
                    } catch (System.Exception ex) {
                        Forms.Main.Instance.Console.UpdateWithError(Ogcs.Google.Calendar.GetEventSummary("New appointment failed to save.", ev, out String anonSummary, true), ex, logEntry: anonSummary);
                        Ogcs.Exception.Analyse(ex, true);
                        if (Ogcs.Extensions.MessageBox.Show("New Outlook appointment failed to save. Continue with synchronisation?", "Sync item failed", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                            continue;
                        else
                            throw new UserCancelledSyncException("User chose not to continue sync.");
                    }

                    /*if (ev.Recurrence != null && ev.RecurringEventId == null && Recurrence.Instance.HasExceptions(ev)) {
                        Forms.Main.Instance.Console.Update("This is a recurring item with some exceptions:-", verbose: true);
                        Recurrence.Instance.CreateOutlookExceptions(ref newAi, ev);
                        Forms.Main.Instance.Console.Update("Recurring exceptions completed.", verbose: true);
                    }*/
                /*} finally {
                    newAi = (AppointmentItem)ReleaseObject(newAi);
                    mapiFolder = (MAPIFolder)ReleaseObject(mapiFolder);
                }*/
            }
        }

        private void createCalendarEntry(GcalData.Event ev, ref Microsoft.Graph.Event ai) {
            string itemSummary = Ogcs.Google.Calendar.GetEventSummary(ev, out String anonItemSummary);
            log.Debug("Processing >> " + (anonItemSummary ?? itemSummary));
            Forms.Main.Instance.Console.Update(itemSummary, anonItemSummary, Console.Markup.calendar, verbose: true);

            SettingsStore.Calendar profile = Sync.Engine.Calendar.Instance.Profile;

            ai.Start = new DateTimeTimeZone() { DateTime = ev.Start.DateTimeRaw };
            if (!String.IsNullOrEmpty(ev.Start.TimeZone)) {
                log.Fine("Has starting timezone: " + ev.Start.TimeZone);
                ai.Start.TimeZone = ev.Start.TimeZone;
            }
            ai.End = new DateTimeTimeZone() { DateTime = ev.End.DateTimeRaw };
            if (!String.IsNullOrEmpty(ev.End.TimeZone)) {
                log.Fine("Has ending timezone: " + ev.End.TimeZone);
                ai.End.TimeZone = ev.End.TimeZone;
            }
            ai.IsAllDay = ev.AllDayEvent();
            
            //Recurrence.Instance.BuildOutlookPattern(ev, ai);

            ai.Subject = Obfuscate.ApplyRegex(Obfuscate.Property.Subject, ev.Summary, null, Sync.Direction.GoogleToOutlook);
            if (profile.AddDescription && ev.Description != null) {
                ai.Body = new ItemBody() { ContentType = BodyType.Html };
                ai.Body.Content = Obfuscate.ApplyRegex(Obfuscate.Property.Description, ev.Description, null, Sync.Direction.GoogleToOutlook);
            }
            if (profile.AddLocation) ai.Location = new Location { DisplayName = Obfuscate.ApplyRegex(Obfuscate.Property.Location, ev.Location, null, Sync.Direction.GoogleToOutlook) };
            ai.Sensitivity = getPrivacy(ev.Visibility, null);
            ai.ShowAs = getAvailability(ev.Transparency, null);
            //ai.Categories = getColour(ev.ColorId, null);

            if (profile.AddAttendees && ev.Attendees != null) {
                if (ev.Attendees != null && ev.Attendees.Count > profile.MaxAttendees) {
                    log.Warn("This Google event has " + ev.Attendees.Count + " attendees, more than the user configured maximum.");
                } else {
                    List<Attendee> attendees = new List<Attendee>();
                    foreach (GcalData.EventAttendee ea in ev.Attendees) {
                        if (Settings.Instance.MSaccountEmail.ToLower() == ea.Email) continue;

                        Attendee attendee = new Attendee() {
                            EmailAddress = new Microsoft.Graph.EmailAddress() { Name = ea.DisplayName, Address = ea.Email },
                            Type = (ea.Optional ?? false ? AttendeeType.Optional : AttendeeType.Required),
                            Status = new ResponseStatus() { Response = ResponseType.None }
                        };
                        switch (ea.ResponseStatus) {
                            case "needsAction": attendee.Status.Response = ResponseType.NotResponded; break;
                            case "declined": attendee.Status.Response = ResponseType.Declined; break;
                            case "tentative": attendee.Status.Response = ResponseType.TentativelyAccepted; break;
                            case "accepted": attendee.Status.Response = ResponseType.Accepted; break;
                        }
                        attendees.Add(attendee);
                    }
                    ai.Attendees = attendees;
                }
            }

            //Reminder alert
            if (profile.AddReminders) {
                if (ev.Reminders?.Overrides?.Any(r => r.Method == "popup") ?? false) {
                    ai.IsReminderOn = true;
                    try {
                        GcalData.EventReminder reminder = ev.Reminders.Overrides.Where(r => r.Method == "popup").OrderBy(x => x.Minutes).First();
                        ai.ReminderMinutesBeforeStart = (int)reminder.Minutes;
                    } catch (System.Exception ex) {
                        ex.Analyse("Failed setting Outlook reminder for final popup Google notification.");
                    }
                } else if ((ev.Reminders?.UseDefault ?? false) && Ogcs.Google.Calendar.Instance.MinDefaultReminder != int.MinValue) {
                    ai.IsReminderOn = true;
                    ai.ReminderMinutesBeforeStart = Ogcs.Google.Calendar.Instance.MinDefaultReminder;
                } else {
                    ai.IsReminderOn = profile.UseOutlookDefaultReminder;
                }
            } else ai.IsReminderOn = profile.UseOutlookDefaultReminder;

            /*
            if (profile.AddGMeet && !String.IsNullOrEmpty(ev.HangoutLink)) {
                ai.GoogleMeet(ev.HangoutLink);
            }
            */
            //Add the Google event IDs into Outlook appointment.
            O365CustomProperty.AddGoogleIDs(ref ai, ev);
        }

        private void createCalendarEntry_save(Microsoft.Graph.Event ai, ref GcalData.Event ev) {
            SettingsStore.Calendar profile = Sync.Engine.Calendar.Instance.Profile;
            if (profile.SyncDirection.Id == Sync.Direction.Bidirectional.Id) {
                log.Debug("Saving timestamp when OGCS updated appointment.");
                O365CustomProperty.SetOGCSlastModified(ref ai);
            }

            try {
                System.Threading.Tasks.Task<Event> createThread =  GraphClient.Me.Calendars[profile.UseOutlookCalendar.Id].Events.Request().AddAsync(ai);
                Event bar = createThread.Result;
            } catch (System.AggregateException ex) {
                if (ex.InnerException is Microsoft.Graph.ServiceException) {
                    ServiceException gex = ex.InnerException as ServiceException;
                    if (gex.Error.Code == "InvalidAuthenticationToken") {
                        this.Authenticator.GetAuthenticated(true);
                    } else
                        throw ex.InnerException;
                }
                //*** Need API handling
            }

            /*if (profile.SyncDirection.Id == Sync.Direction.Bidirectional.Id || Ogcs.Google.CustomProperty.ExistAnyOutlookIDs(ev)) {
                log.Debug("Storing the Outlook appointment IDs in Google event.");
                Ogcs.Google.CustomProperty.AddOutlookIDs(ref ev, ai);
                Ogcs.Google.Calendar.Instance.UpdateCalendarEntry_save(ref ev);
            }*/
        }
        #endregion

        #region Delete
        public void DeleteCalendarEntries(List<Microsoft.Graph.Event> oAppointments) {
            for (int o = oAppointments.Count - 1; o >= 0; o--) {
                if (Sync.Engine.Instance.CancellationPending) return;

                Microsoft.Graph.Event ai = oAppointments[o];
                Boolean doDelete = false;
                try {
                    doDelete = deleteCalendarEntry(ai);
                } catch (System.Exception ex) {
                    Forms.Main.Instance.Console.UpdateWithError(GetEventSummary("Appointment deletion failed.", ai, out String anonSummary, true), ex, logEntry: anonSummary);
                    Ogcs.Exception.Analyse(ex, true);
                    if (Ogcs.Extensions.MessageBox.Show("Outlook appointment deletion failed. Continue with synchronisation?", "Sync item failed", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                        continue;
                    else
                        throw new UserCancelledSyncException("User chose not to continue sync.");
                }

                try {
                    if (doDelete) deleteCalendarEntry_save(ai);
                    else oAppointments.Remove(ai);
                } catch (System.Exception ex) {
                    if (ex is Microsoft.Graph.ServiceException) {
                        Microsoft.Graph.ServiceException gex = ex as Microsoft.Graph.ServiceException;
                        if (gex.Error != null && gex.Error.Code == "ErrorItemNotFound") { //Resource has been deleted
                            log.Fail("This event is already deleted! Ignoring failed request to delete.");
                            continue;
                        }
                    }
                    if (ex is ApplicationException) {
                        String summary = GetEventSummary("<br/>Appointment deletion skipped.<br/>" + ex.Message, ai, out String anonSummary);
                        Forms.Main.Instance.Console.Update(summary, anonSummary, Console.Markup.warning);
                        if (ex.InnerException is Microsoft.Graph.ServiceException) break;
                        continue;
                    } else {
                        String summary = GetEventSummary("<br/>Appointment deletion failed.", ai, out String anonSummary);
                        Forms.Main.Instance.Console.UpdateWithError(summary, ex, logEntry: anonSummary);
                        Ogcs.Exception.Analyse(ex, true);
                        if (Ogcs.Extensions.MessageBox.Show("Outlook appointment deletion failed. Continue with synchronisation?", "Sync item failed", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                            continue;
                        else
                            throw new UserCancelledSyncException("User chose not to continue sync.");
                    }
                }
            }
        }

        private Boolean deleteCalendarEntry(Microsoft.Graph.Event ai) {
            String eventSummary = GetEventSummary(ai, out String anonSummary);
            Boolean doDelete = true;

            if (Sync.Engine.Calendar.Instance.Profile.ConfirmOnDelete) {
                if (Ogcs.Extensions.MessageBox.Show("Delete " + eventSummary + "?", "Confirm Deletion From Outlook",
                    MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2) == DialogResult.No) {
                    doDelete = false;
                    if (Sync.Engine.Calendar.Instance.Profile.SyncDirection.Id == Sync.Direction.Bidirectional.Id && O365CustomProperty.ExistAnyGoogleIDs(ai)) {
                        O365CustomProperty.RemoveGoogleIDs(ref ai);
                        //ai.Save();
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

        private void deleteCalendarEntry_save(Microsoft.Graph.Event ai) {
            try {
                GraphClient.Me.Events[ai.Id].Request().DeleteAsync().Wait();
            } catch (System.AggregateException ex) {
                if (ex.InnerException is Microsoft.Graph.ServiceException) throw ex.InnerException;
                //*** Need API handling
            }
        }
        #endregion

        /// <summary>
        /// Determine Appointment Item's privacy setting
        /// </summary>
        /// <param name="gVisibility">Google's current setting</param>
        /// <param name="oSensitivity">Outlook's current setting</param>
        private Sensitivity getPrivacy(String gVisibility, Sensitivity? oSensitivity) {
            SettingsStore.Calendar profile = Sync.Engine.Calendar.Instance.Profile;

            if (!profile.SetEntriesPrivate)
                return (gVisibility == "private") ? Sensitivity.Private : Sensitivity.Normal;

            if (profile.SyncDirection.Id != Sync.Direction.Bidirectional.Id) {
                return (profile.PrivacyLevel == Sensitivity.Private.ToString()) ? Sensitivity.Private : Sensitivity.Normal;
            } else {
                if (profile.TargetCalendar.Id == Sync.Direction.OutlookToGoogle.Id) { //Privacy enforcement is in other direction
                    if (oSensitivity == null)
                        return (gVisibility == "private") ? Sensitivity.Private : Sensitivity.Normal;
                    else if (oSensitivity == Sensitivity.Private && gVisibility != "private") {
                        log.Fine("Source of truth for enforced privacy is already set private and target is NOT - so syncing this back.");
                        return Sensitivity.Normal;
                    } else
                        return (Sensitivity)oSensitivity;
                } else {
                    if (!profile.CreatedItemsOnly || (profile.CreatedItemsOnly && oSensitivity == null))
                        return (profile.PrivacyLevel == Sensitivity.Private.ToString()) ? Sensitivity.Private : Sensitivity.Normal;
                    else
                        return (gVisibility == "private") ? Sensitivity.Private : Sensitivity.Normal;
                }
            }
        }

        /// <summary>
        /// Determine Appointment's availability setting
        /// </summary>
        /// <param name="gTransparency">Google's current setting</param>
        /// <param name="oBusyStatus">Outlook's current setting</param>
        private FreeBusyStatus getAvailability(String gTransparency, FreeBusyStatus? oBusyStatus) {
            SettingsStore.Calendar profile = Sync.Engine.Calendar.Instance.Profile;
            List<FreeBusyStatus> persistOutlookStatus = new List<FreeBusyStatus> { FreeBusyStatus.Tentative, FreeBusyStatus.Oof, FreeBusyStatus.WorkingElsewhere };

            if (!profile.SetEntriesAvailable)
                return (gTransparency == "transparent") ? FreeBusyStatus.Free :
                    persistOutlookStatus.Contains((FreeBusyStatus)oBusyStatus) ? (FreeBusyStatus)oBusyStatus : FreeBusyStatus.Busy;

            FreeBusyStatus overrideFbStatus = FreeBusyStatus.Free;
            try {
                Enum.TryParse(profile.AvailabilityStatus, out overrideFbStatus);
            } catch (System.Exception ex) {
                ex.Analyse("Could not convert string '" + profile.AvailabilityStatus + "' to FreeBusyStatus type. Defaulting override to available.");
            }

            if (profile.SyncDirection.Id != Sync.Direction.Bidirectional.Id) {
                return overrideFbStatus;
            } else {
                if (profile.TargetCalendar.Id == Sync.Direction.OutlookToGoogle.Id) { //Availability enforcement is in other direction
                    if (oBusyStatus == null)
                        return (gTransparency == "transparent") ? FreeBusyStatus.Free :
                            persistOutlookStatus.Contains((FreeBusyStatus)oBusyStatus) ? (FreeBusyStatus)oBusyStatus : FreeBusyStatus.Busy;
                    else
                        return (FreeBusyStatus)oBusyStatus;
                } else {
                    if (!profile.CreatedItemsOnly || (profile.CreatedItemsOnly && oBusyStatus == null))
                        return overrideFbStatus;
                    else
                        return (gTransparency == "transparent") ? FreeBusyStatus.Free :
                            persistOutlookStatus.Contains((FreeBusyStatus)oBusyStatus) ? (FreeBusyStatus)oBusyStatus : FreeBusyStatus.Busy;
                }
            }
        }


        #region STATIC functions
        /// <summary>
        /// Get the anonymised summary of an appointment item, else standard summary.
        /// </summary>
        /// <param name="ai">The Graph Event item.</param>
        /// <returns>The summary, anonymised if settings dictate.</returns>
        public static String GetEventSummary(Microsoft.Graph.Event ai) {
            String eventSummary = GetEventSummary(ai, out String anonymisedSummary, false);
            return anonymisedSummary ?? eventSummary;
        }

        /// <summary>
        /// Pre/Append context to the summary of an appointment item.
        /// </summary>
        /// <param name="context">Text to add before/after the summary and anonymised summary.</param>
        /// <param name="ai">The Graph Event item.</param>
        /// <param name="eventSummaryAnonymised">The anonymised summary with context.</param>
        /// <param name="onlyIfNotVerbose">Only return if user doesn't have Verbose output on. Useful for indicating offending item during errors.</param>
        /// <param name="appendContext">If the context should be before or after.</param>
        /// <returns>The standard summary.</returns>
        public static string GetEventSummary(String context, Microsoft.Graph.Event ai, out String eventSummaryAnonymised, Boolean onlyIfNotVerbose = false, Boolean appendContext = true) {
            String eventSummary = GetEventSummary(ai, out String anonymisedSummary, onlyIfNotVerbose);
            if (appendContext) {
                eventSummary = eventSummary + context;
                eventSummaryAnonymised = anonymisedSummary + context;
            } else {
                eventSummary = context + eventSummary;
                eventSummaryAnonymised = context + anonymisedSummary;
            }
            return eventSummary;
        }

        /// <summary>
        /// Get the summary of an appointment item.
        /// </summary>
        /// <param name="ai">The appointment item</param>
        /// <param name="eventSummaryAnonymised">Anonymised version of the returned summary string value.</param>
        /// <param name="onlyIfNotVerbose">Only return if user doesn't have Verbose output on. Useful for indicating offending item during errors.</param>
        /// <returns>The standard summary.</returns>
        public static string GetEventSummary(Microsoft.Graph.Event ai, out String eventSummaryAnonymised, Boolean onlyIfNotVerbose = false) {
            String eventSummary = "";
            eventSummaryAnonymised = null;
            if (!onlyIfNotVerbose || onlyIfNotVerbose && !Settings.Instance.VerboseOutput) {
                try {
                    if (ai.IsAllDay ?? false) {
                        log.Fine("GetSummary - all day event");
                        eventSummary += ai.Start.SafeDateTime().Date.ToShortDateString();
                    } else {
                        log.Fine("GetSummary - not all day event");
                        eventSummary += ai.Start.SafeDateTime().ToShortDateString() + " " + ai.Start.SafeDateTime().ToShortTimeString();
                    }
                    eventSummary += " " + (ai.Recurrence != null ? "(R) " : "") + "=> ";

                    if (Settings.Instance.AnonymiseLogs)
                        eventSummaryAnonymised = eventSummary + '"' + Ogcs.Google.Authenticator.GetMd5(ai.Subject, silent: true) + '"' + (onlyIfNotVerbose ? "<br/>" : "");
                    eventSummary += '"' + ai.Subject + '"' + (onlyIfNotVerbose ? "<br/>" : "");

                } catch (System.Exception ex) {
                    ex.Analyse("Failed to get appointment summary: " + eventSummary, true);
                }
            }
            return eventSummary;
        }

        public static void IdentifyEventDifferences(
            ref List<GcalData.Event> google,          //need creating
            ref List<Microsoft.Graph.Event> outlook,  //need deleting
            ref Dictionary<Microsoft.Graph.Event, GcalData.Event> compare) //
        {
            log.Debug("Comparing Google events to Outlook items...");
            Forms.Main.Instance.Console.Update("Matching calendar items...", verbose: true);

            SettingsStore.Calendar profile = Sync.Engine.Calendar.Instance.Profile;

            //Order by start date (same as Outlook) for quickest matching
            google.Sort((x, y) => (x.Start.DateTimeRaw ?? x.Start.Date).CompareTo((y.Start.DateTimeRaw ?? y.Start.Date)));

            // Count backwards so that we can remove found items without affecting the order of remaining items
            int metadataEnhanced = 0;
            for (int o = outlook.Count - 1; o >= 0; o--) {
                if (Sync.Engine.Instance.CancellationPending) return;
                log.Fine("Checking " + GetEventSummary(outlook[o]));
                
                String compare_oEventID = O365CustomProperty.Get(outlook[o], O365CustomProperty.MetadataId.gEventID);
                /*
                if (!string.IsNullOrEmpty(compare_oEventID)) {
                    Boolean? googleIDmissing = null;
                    Boolean foundMatch = false;

                    for (int g = google.Count - 1; g >= 0; g--) {
                        log.UltraFine("Checking " + Ogcs.Google.Calendar.GetEventSummary(google[g]));

                        if (compare_oEventID == google[g].Id.ToString()) {
                            if (googleIDmissing == null) googleIDmissing = CustomProperty.GoogleIdMissing(outlook[o]);
                            if ((Boolean)googleIDmissing) {
                                log.Info("Enhancing appointment's metadata...");
                                AppointmentItem ai = outlook[o];
                                CustomProperty.AddGoogleIDs(ref ai, google[g]);
                                CustomProperty.Add(ref ai, CustomProperty.MetadataId.forceSave, "True");
                                outlook[o] = ai;
                                metadataEnhanced++;
                            }
                            if (ItemIDsMatch(outlook[o], google[g])) {
                                foundMatch = true;
                                compare.Add(outlook[o], google[g]);
                                outlook.Remove(outlook[o]);
                                google.Remove(google[g]);
                                break;
                            }
                        }
                    }
                    if (!foundMatch && profile.MergeItems &&
                        Outlook.CustomProperty.Get(outlook[o], CustomProperty.MetadataId.gCalendarId) != profile.UseGoogleCalendar.Id)
                        outlook.Remove(outlook[o]);

                } else if (profile.MergeItems) {
                    //Remove the non-Google item so it doesn't get deleted
                    outlook.Remove(outlook[o]);
                }
                */
            }
            if (metadataEnhanced > 0) log.Info(metadataEnhanced + " item's metadata enhanced.");
            /*
            if (profile.OnlyRespondedInvites) {
                //Check if items to be deleted have invitations not responded to
                int responseFiltered = 0;
                for (int o = outlook.Count - 1; o >= 0; o--) {
                    if (outlook[o].ResponseStatus == OlResponseStatus.olResponseNotResponded) {
                        outlook.Remove(outlook[o]);
                        responseFiltered++;
                    }
                }
                if (responseFiltered > 0) log.Info(responseFiltered + " Outlook items will not be deleted due to only syncing invites that have been responded to.");
            }

            if (outlook.Count > 0 && Ogcs.Google.Calendar.Instance.ExcludedByColour?.Count > 0 && !profile.DeleteWhenColourExcluded) {
                //Check if Outlook items to be deleted were filtered out from Google
                for (int o = outlook.Count - 1; o >= 0; o--) {
                    if (Ogcs.Google.Calendar.Instance.ExcludedByColour.ContainsValue(outlook[o].EntryID) ||
                        Ogcs.Google.Calendar.Instance.ExcludedByColour.ContainsKey(CustomProperty.Get(outlook[o], CustomProperty.MetadataId.gEventID) ?? "")) {
                        outlook.Remove(outlook[o]);
                    }
                }
            }
            if (google.Count > 0 && Instance.ExcludedByCategory?.Count > 0) {
                //Check if Google items to be created were filtered out from Outlook
                for (int g = google.Count - 1; g >= 0; g--) {
                    if (Instance.ExcludedByCategory.ContainsValue(google[g].Id) ||
                        Instance.ExcludedByCategory.ContainsKey(Ogcs.Google.CustomProperty.Get(google[g], Ogcs.Google.CustomProperty.MetadataId.oEntryId) ?? "")) {
                        google.Remove(google[g]);
                    }
                }
            }

            if (profile.DisableDelete) {
                if (outlook.Count > 0) {
                    Forms.Main.Instance.Console.Update(outlook.Count + " Outlook items would have been deleted, but you have deletions disabled.", Console.Markup.warning);
                    for (int o = 0; o < outlook.Count; o++)
                        Forms.Main.Instance.Console.Update(GetEventSummary(outlook[o], out String anonSummary), anonSummary, verbose: true);
                }
                outlook = new List<AppointmentItem>();
            }
            if (profile.SyncDirection.Id == Sync.Direction.Bidirectional.Id) {
                //Don't recreate any items that have been deleted in Outlook
                for (int g = google.Count - 1; g >= 0; g--) {
                    if (Ogcs.Google.CustomProperty.Exists(google[g], Ogcs.Google.CustomProperty.MetadataId.oEntryId))
                        google.Remove(google[g]);
                }
                //Don't delete any items that aren't yet in Google or just created in Google during this sync
                for (int o = outlook.Count - 1; o >= 0; o--) {
                    if (!CustomProperty.Exists(outlook[o], CustomProperty.MetadataId.gEventID) ||
                        CustomProperty.GetOGCSlastModified(outlook[o]) > Sync.Engine.Instance.SyncStarted)
                        outlook.Remove(outlook[o]);
                }
            }
            if (Settings.Instance.CreateCSVFiles) {
                ExportToCSV("Appointments for deletion in Outlook", "outlook_delete.csv", outlook);
                Ogcs.Google.Calendar.ExportToCSV("Events for creation in Outlook", "outlook_create.csv", google);
            }
            */
        }

        #endregion
    }
}
