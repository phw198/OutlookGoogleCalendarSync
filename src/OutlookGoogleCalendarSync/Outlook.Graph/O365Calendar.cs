using log4net;
using Microsoft.Graph;
using OutlookGoogleCalendarSync.Extensions;
using OutlookGoogleCalendarSync.GraphExtension;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using CVRB = Microsoft.Graph.Me.Calendars.Item.CalendarView.CalendarViewRequestBuilder;
using GcalData = Google.Apis.Calendar.v3.Data;
using Kiota = Microsoft.Kiota.Abstractions;
using MsGraph = Microsoft.Graph.Models;
using Ogcs = OutlookGoogleCalendarSync;

namespace OutlookGoogleCalendarSync.Outlook.Graph {
    public class Calendar {
        private static readonly ILog log = LogManager.GetLogger(typeof(Calendar));

        private static Calendar instance;
        public static Boolean IsInstanceNull { get { return instance == null; } }
        public static Calendar Instance {
            get {
                return instance ??= new Ogcs.Outlook.Graph.Calendar {
                    Authenticator = new Ogcs.Outlook.Graph.Authenticator()
                };
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

        /// <summary>
        /// Graph API v1.0 doesn't properly surface cancelled series occurrences as of Jan-2025
        /// Therefore home-brewing our own dictionary workaround.
        /// </summary>
        public Dictionary<String, List<System.DateTime>> CancelledOccurrences { get; set; }

        /// <summary>Outlook Appointment excluded through user config <Appt.EntryId></summary>
        public List<String> ExcludedByConfig { get; private set; }
        /// <summary>Outlook Appointment excluded by category through user config <Event.Id, Appt.EntryId></summary>
        public Dictionary<String, String> ExcludedByCategory { get; private set; }

        private Dictionary<String, OutlookCalendarListEntry> calendarFolders = new Dictionary<string, OutlookCalendarListEntry>();
        public Dictionary<String, OutlookCalendarListEntry> CalendarFolders {
            get { return calendarFolders; }
        }

        /// <summary>Retrieve calendar list from the cloud.</summary>
        public Dictionary<String, OutlookCalendarListEntry> GetCalendars() {
            calendarFolders = new();
            List<MsGraph.Calendar > cals = new();

            try {
                Microsoft.Graph.Me.Calendars.CalendarsRequestBuilder calendarsRequest = GraphClient.Me.Calendars;
                MsGraph.CalendarCollectionResponse calPage = calendarsRequest.GetAsync(config => {
                    config.QueryParameters.Select = new[] {
                        "id", "name", "color", "changeKey", "canShare", "canViewPrivateItems", "hexColor", "canEdit", "isTallyingResponses", "isRemovable", "owner"
                    };
                }).Result;

                cals.AddRange(calPage.Value ?? new());
                while (!String.IsNullOrEmpty(calPage.OdataNextLink)) {
                    calPage = calendarsRequest.WithUrl(calPage.OdataNextLink).GetAsync(config => { }).Result;
                    cals.AddRange(calPage.Value ?? new());
                }
            } catch (System.Exception ex) {
                switch (O365Errors.HandleAPIlimits(ref ex)) {
                    case ApiException.throwException: throw ex;
                    default: throw ex;
                }
            }

            foreach (MsGraph.Calendar cal in cals) {
                if (cal.AdditionalData.ContainsKey("isDefaultCalendar") && (Boolean)cal.AdditionalData["isDefaultCalendar"])
                    cal.Name = "Default " + cal.Name;
                log.Debug(cal.Name);
                calendarFolders.Add(cal.Name, new OutlookCalendarListEntry(cal));
            }

            return calendarFolders;
        }

        /// <summary>
        /// Retrieve specific Graph Event. Also updates cache of cancelled occurrences.
        /// </summary>
        /// <param name="eventId">Event ID to retrieve</param>
        /// <returns>The Graph Event</returns>
        public MsGraph.Event GetCalendarEntry(String eventId) {
            MsGraph.Event ai = null;
            try {
                log.Debug("Retrieving specific Graph Event with ID " + eventId);
                SettingsStore.Calendar profile = Settings.Profile.InPlay();

                ai = GraphClient.Me.Calendars[profile.UseOutlookCalendar.Id].Events[eventId].GetAsync(cfg => {
                    cfg.QueryParameters.Expand = new string[] { $"extensions($filter=Id eq '{CustomProperty.ExtensionName()}')" };
                    cfg.QueryParameters.Select = new string[] { "*" }; 
                }).Result;

                foreach (String cancelledOccurrence in (ai?.CancelledOccurrences ?? new())) {
                    System.DateTime cancelledDate = System.DateTime.ParseExact(cancelledOccurrence.Replace($"OID.{eventId}.", ""), "yyyy-MM-dd", System.Globalization.CultureInfo.InvariantCulture);
                    if (cancelledDate < profile.SyncStart.Date || cancelledDate > profile.SyncEnd.Date) {
                        log.Fine("Exception is deleted and outside date range being synced: " + cancelledDate.Date.ToString("dd/MM/yyyy"));
                        continue;
                    }
                    if (CancelledOccurrences.ContainsKey(eventId))
                        CancelledOccurrences[eventId].Add(cancelledDate);
                    else
                        CancelledOccurrences.Add(eventId, new List<System.DateTime>() { cancelledDate });
                }

                if (ai != null)
                    return ai;
                else
                    throw new System.Exception("Returned null");
            } catch (System.Exception ex) {
                ex.Analyse();
                Forms.Main.Instance.Console.Update("Failed to retrieve Graph event.", Console.Markup.error);
                return null;
            }
        }

        /// <summary>Retrieve all the occurences in a series, excluding those that have been cancelled</summary>
        /// <param name="seriesId">Series master ID</param>
        public List<MsGraph.Event> GetCalendarEntriesInRecurrence(String seriesId) {
            List<MsGraph.Event> occurrences = new();
            try {
                log.Debug("Retrieving occurrences for recurring master Graph Event with ID " + seriesId);
                SettingsStore.Calendar profile = Settings.Profile.InPlay();

                Microsoft.Graph.Me.Calendars.Item.Events.Item.Instances.InstancesRequestBuilder itemInstancesReq = GraphClient.Me.Calendars[profile.UseOutlookCalendar.Id].Events[seriesId].Instances;
                Kiota.RequestConfiguration<Microsoft.Graph.Me.Calendars.Item.Events.Item.Instances.InstancesRequestBuilder.InstancesRequestBuilderGetQueryParameters> reqCfg = new();
                reqCfg.QueryParameters.StartDateTime = profile.SyncStart.ToString("yyyy-MM-dd");
                reqCfg.QueryParameters.EndDateTime = profile.SyncEnd.ToString("yyyy-MM-dd");
                reqCfg.QueryParameters.Top = 250;
                reqCfg.QueryParameters.Select = new string[] { "*" };
                reqCfg.QueryParameters.Expand = new string[] { $"extensions($filter=Id eq '{CustomProperty.ExtensionName()}')" };
                reqCfg.QueryParameters.Orderby = new string[] { "start/dateTime" };
                reqCfg.QueryParameters.Count = true;
                
                Kiota.RequestInformation reqInfo = itemInstancesReq.ToGetRequestInformation(cfg => cfg.QueryParameters = reqCfg.QueryParameters);
                log.Fine(reqInfo.URI.ToString());
                MsGraph.EventCollectionResponse instancesPage = itemInstancesReq.GetAsync(cfg => cfg.QueryParameters = reqCfg.QueryParameters).Result;
                int pageCnt = 1;
                log.Fine($"Page {pageCnt} retrieved with {instancesPage.OdataCount} items.");
                occurrences.AddRange(instancesPage.Value ?? new());
                while (!String.IsNullOrEmpty(instancesPage.OdataNextLink)) {
                    pageCnt++;
                    instancesPage = itemInstancesReq.WithUrl(instancesPage.OdataNextLink).GetAsync().Result;
                    log.Fine($"Page {pageCnt} retrieved with {instancesPage.OdataCount} items.");
                    occurrences.AddRange(instancesPage.Value ?? new());
                }

            } catch (System.Exception ex) {
                ex.Analyse("Could not retrieve occurrences for recurring series ID " + seriesId);
            }
            log.Debug(occurrences.Count + " occurrences retrieved.");
            return occurrences;
        }

        /// <summary>
        /// Get all calendar entries within the defined date-range for sync
        /// </summary>
        /// <param name="suppressAdvisories">Don't give user feedback, eg during background Push sync</param>
        /// <returns></returns>
        public List<MsGraph.Event> GetCalendarEntriesInRange(SettingsStore.Calendar profile, Boolean suppressAdvisories) {
            List<MsGraph.Event> filtered;
            try {
                filtered = filterCalendarEntries(profile, suppressAdvisories: suppressAdvisories);
            } catch (System.Exception) {
                if (!suppressAdvisories) Forms.Main.Instance.Console.Update("Unable to access the Outlook calendar.", Console.Markup.error);
                throw;
            }
            Recurrence.SeparateOutlookExceptions(filtered);
            return filtered;
        }

        private List<MsGraph.Event> filterCalendarEntries(SettingsStore.Calendar profile, Boolean suppressAdvisories = false) {
            List<MsGraph.Event> result = new();
            ExcludedByConfig = new();
            ExcludedByCategory = new();

            profile ??= Settings.Profile.InPlay();

            System.DateTimeOffset min = System.DateTimeOffset.MinValue;
            System.DateTimeOffset max = System.DateTimeOffset.MaxValue;
            min = profile.SyncStart;
            max = profile.SyncEnd;

            try {
                // Code snippets are only available for the latest version. Current version is 5.x
                // To initialize your graphClient, see https://learn.microsoft.com/en-us/graph/sdks/create-client?from=snippets&tabs=csharp
                //.e.Events.GetAsync((requestConfiguration) =>
                //{
                //    requestConfiguration.QueryParameters.Select = new string[] { "subject", "body", "bodyPreview", "organizer", "attendees", "start", "end", "location" };
                //    requestConfiguration.Headers.Add("Prefer", "outlook.timezone=\"Pacific Standard Time\"");
                //});

                //A master series may span the sync date range but have no exceptions - this isn't returned by /calendar/events end-point.
                //To get these master series, /calendarView end-point needs to be used
                //1. Get all single instances, occurrences and exceptions within date range
                //2. Get distinct list of series IDs for which there is no master series
                //3. Get the specific missing master event(s)
                CVRB calendarViewRequest = GraphClient.Me.Calendars[profile.UseOutlookCalendar.Id].CalendarView;
                Kiota.RequestConfiguration<CVRB.CalendarViewRequestBuilderGetQueryParameters> reqCfg = new();
                reqCfg.QueryParameters.StartDateTime = min.ToPreciseString();
                reqCfg.QueryParameters.EndDateTime = max.ToPreciseString();
                reqCfg.QueryParameters.Top = 250;
                reqCfg.QueryParameters.Expand = new String[] { $"extensions($filter=Id eq '{CustomProperty.ExtensionName()}')" };
                reqCfg.QueryParameters.Select = new String[] { "*" }; //Otherwise OriginalStart is always null
                reqCfg.QueryParameters.Orderby = new String[] { "start/dateTime" };
                reqCfg.QueryParameters.Count = true;

                Kiota.RequestInformation reqInfo = calendarViewRequest.ToGetRequestInformation(cfg => { cfg.QueryParameters = reqCfg.QueryParameters; });
                log.Fine(reqInfo.URI.ToString());
                MsGraph.EventCollectionResponse eventsPage = calendarViewRequest.GetAsync(cfg => { cfg.QueryParameters = reqCfg.QueryParameters; }).Result;
                int pageCnt = 1;
                log.Fine($"Page {pageCnt} retrieved with {eventsPage.OdataCount} items.");
                result.AddRange(eventsPage.Value ?? new());
                while (!String.IsNullOrEmpty(eventsPage.OdataNextLink)) {
                    pageCnt++;
                    eventsPage = calendarViewRequest.WithUrl(eventsPage.OdataNextLink).GetAsync().Result;
                    log.Fine($"Page {pageCnt} retrieved with {eventsPage.OdataCount} items.");
                    result.AddRange(eventsPage.Value ?? new());
                }

            } catch {
                log.Fail($"Could not query '{Settings.Profile.Name(profile)}' profile calendar '{profile.UseOutlookCalendar.Name}'");
                throw;
            }

            log.Fine(result.Count + " calendar items exist in total.");

            Recurrence.GetOutlookMasterEvent(result);
            List<MsGraph.Event> seriesOccurrences = result.Where(ai => ai.Type == MsGraph.EventType.Occurrence).ToList();
            result = result.Except(seriesOccurrences).ToList();
            result.Sort((x, y) => x.Start.SafeDateTime().CompareTo(y.Start.SafeDateTime()));
            log.Fine(seriesOccurrences.Count + " standard series occurrences removed.");

            List<MsGraph.Event> endsOnSyncStart = result.Where(ai => (ai.End != null && ai.End.SafeDateTime() == min && ai.Type != MsGraph.EventType.SeriesMaster)).ToList();
            if (endsOnSyncStart.Count > 0) {
                log.Debug(endsOnSyncStart.Count + " Outlook Appointments end at midnight of the sync start date window.");
                result = result.Except(endsOnSyncStart).ToList();
            }

            List<MsGraph.Event> allExcluded = applyExclusions(ref result, profile);

            if (allExcluded.Count > 0) {
                if (!suppressAdvisories) {
                    String filterWarning = "Due to your OGCS Outlook settings, " + (result.Count == 0 ? "all" : allExcluded.Count) + " Outlook items have been filtered out" + (result.Count == 0 ? "!" : ".");
                    Forms.Main.Instance.Console.Update(filterWarning, Console.Markup.config, newLine: false, notifyBubble: (result.Count == 0));

                    filterWarning = "";
                    if (profile.SyncDirection.Id != Sync.Direction.GoogleToOutlook.Id && ExcludedByCategory.Count > 0 && profile.DeleteWhenCategoryExcluded) {
                        filterWarning = "If they exist in Google, they may get deleted. To avoid deletion, uncheck \"Delete synced items if excluded\".";
                        if (!profile.DisableDelete) {
                            filterWarning += " Recover unintentional deletions from the <a href='https://calendar.google.com/calendar/u/0/r/trash'>Google 'Bin'</a>.";
                            if (profile.ConfirmOnDelete)
                                filterWarning += "<p style='margin-top: 8px;'>If prompted to confirm deletion and you opt <i>not</i> to delete them, this will reoccur every sync. " +
                                    "Consider assigning an excluded colour to those items in Google.</p>" +
                                    "<p style='margin-top: 8px;'>See the wiki for tips if needing to <a href='https://github.com/phw198/OutlookGoogleCalendarSync/wiki/FAQs#duplicates-due-to-colourcategory-exclusion'>resolve duplicates</a>.</p>";
                        }
                    }
                    if (!String.IsNullOrEmpty(filterWarning))
                        Forms.Main.Instance.Console.Update(filterWarning, Console.Markup.warning, newLine: false);
                }

                if (profile.SyncDirection.Id == Sync.Direction.Bidirectional.Id) {
                    for (int o = 0; o < allExcluded.Count; o++) {
                        MsGraph.Event ai = allExcluded[o];
                        if (CustomProperty.ExistAnyGoogleIDs(ai)) {
                            log.Debug("Previously synced Outlook item is now excluded. Removing Google metadata.");
                            //We don't want them getting automatically deleted if brought back in scope; better to create possible duplicate
                            CustomProperty.RemoveGoogleIDs(ref ai);
                            UpdateCalendarEntry_save(ref ai);
                        }
                    }
                }
            }

            log.Debug("Filtered down to " + result.Count);
            return result;
        }

        private List<MsGraph.Event> applyExclusions(ref List<MsGraph.Event> result, SettingsStore.Calendar profile) {
            List<MsGraph.Event> allDays = new();
            List<MsGraph.Event> availability = new();
            List<MsGraph.Event> privacy = new();
            List<MsGraph.Event> subject = new();
            List<MsGraph.Event> response = new();

            /*              
                //Categories
                try {
                    if (profile.CategoriesRestrictBy == SettingsStore.Calendar.RestrictBy.Include) {
                        filtered = (profile.Categories.Count() == 0 || (ai.Categories == null && !profile.Categories.Contains("<No category assigned>")) ||
                            (ai.Categories != null && ai.Categories.Split(new[] { Categories.Delimiter }, StringSplitOptions.None).Intersect(profile.Categories).Count() == 0));

                    } else if (profile.CategoriesRestrictBy == SettingsStore.Calendar.RestrictBy.Exclude) {
                        filtered = (profile.Categories.Count() > 0 && ((ai.Categories == null && profile.Categories.Contains("<No category assigned>")) ||
                            (ai.Categories != null && ai.Categories.Split(new[] { Categories.Delimiter }, StringSplitOptions.None).Intersect(profile.Categories).Count() > 0)));
                    }
                } catch (System.Runtime.InteropServices.COMException ex) {
                    if (ex.TargetSite.Name == "get_Categories") {
                        log.Warn("Could not access Categories property for " + GetEventSummary(ai));
                        filtered = ((profile.CategoriesRestrictBy == SettingsStore.Calendar.RestrictBy.Include && !profile.Categories.Contains("<No category assigned>")) ||
                            (profile.CategoriesRestrictBy == SettingsStore.Calendar.RestrictBy.Exclude && profile.Categories.Contains("<No category assigned>")));
                    } else throw;
                }
                if (filtered) { ExcludedByCategory.Add(ai.EntryID, CustomProperty.Get(ai, CustomProperty.MetadataId.gEventID)); continue; }
            */
            //Availability, Privacy, Subject
            if (profile.SyncDirection.Id != Sync.Direction.GoogleToOutlook.Id) { //Sync direction means O->G will delete previously synced excluded items
                List<MsGraph.Event> filterable = result.Where(ai => (ai.Type == MsGraph.EventType.SingleInstance || ai.Type == MsGraph.EventType.SeriesMaster)).ToList();

                if (profile.ExcludeFree || profile.ExcludeTentative) {
                    availability = filterable.Where(ai => ai.ShowAs == MsGraph.FreeBusyStatus.Free || ai.ShowAs == MsGraph.FreeBusyStatus.Tentative).ToList();
                    if (availability.Count > 0) {
                        log.Debug(availability.Count + " Outlook Free/Tentative items excluded.");
                        result = result.Except(availability).ToList();
                    }
                }
                if (profile.ExcludeAllDays) {
                    allDays = filterable.Where(ai => ai.AllDayEvent(true) && (profile.ExcludeFreeAllDays ? ai.ShowAs == MsGraph.FreeBusyStatus.Free : true)).ToList();
                    if (allDays.Count > 0) {
                        log.Debug(allDays.Count + " Outlook all-day items excluded.");
                        result = result.Except(allDays).ToList();
                    }
                }

                if (profile.ExcludePrivate) {
                    privacy = filterable.Where(ai => ai.Sensitivity == MsGraph.Sensitivity.Private).ToList();
                    if (privacy.Count > 0) {
                        log.Debug(privacy.Count + " Outlook private items excluded.");
                        result = result.Except(privacy).ToList();
                    }
                }

                if (profile.ExcludeSubject && !String.IsNullOrEmpty(profile.ExcludeSubjectText)) {
                    Regex rgx = new Regex(profile.ExcludeSubjectText, RegexOptions.IgnoreCase);
                    subject = filterable.Where(ai => rgx.IsMatch(ai.Subject ?? "")).ToList();
                    if (subject.Count > 0) {
                        log.Debug(subject.Count + " Outlook items excluded with Subject containing '" + profile.ExcludeSubjectText + "'");
                        result = result.Except(subject).ToList();
                    }
                }
            }
            //Invitation
            if (profile.OnlyRespondedInvites) {
                //These are actually filtered out later on when identifying differences
                response = result.Where(ai => ai.ResponseStatus.Response == MsGraph.ResponseType.NotResponded).ToList();
                if (response.Count > 0) 
                    log.Debug(response.Count + " Outlook items are invites not yet responded to.");
            }            
            
            List<MsGraph.Event> allExcluded = /*colour.Concat*/(availability).Concat(allDays).Concat(privacy).Concat(subject).ToList();
            foreach (MsGraph.Event ev in allExcluded) {
                if (!ExcludedByConfig.Contains(ev.Id))
                    ExcludedByConfig.Add(ev.Id);
            }
            return allExcluded;
        }

        #region Create
        public void CreateCalendarEntries(List<GcalData.Event> events) {
            for (int g = events.Count - 1; g >= 0; g--) {
                if (Sync.Engine.Instance.CancellationPending) return;

                GcalData.Event ev = events[g];
                MsGraph.Event newAi = new();
                try {
                    createCalendarEntry(ev, ref newAi);
                } catch (System.Exception ex) {
                    events.Remove(ev);
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

                MsGraph.Event createdAi = new MsGraph.Event();
                try {
                    createdAi = createCalendarEntry_save(newAi, ref ev);
                    events[g] = ev;
                } catch (System.Exception ex) {
                    log.Debug(Newtonsoft.Json.JsonConvert.SerializeObject(newAi));
                    events.RemoveAt(g);
                    Forms.Main.Instance.Console.UpdateWithError(Ogcs.Google.Calendar.GetEventSummary("New appointment failed to save.", ev, out String anonSummary, true), ex, logEntry: anonSummary);
                    Ogcs.Exception.Analyse(ex, true);
                    if (Ogcs.Extensions.MessageBox.Show("New Outlook appointment failed to save. Continue with synchronisation?", "Sync item failed", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                        continue;
                    else
                        throw new UserCancelledSyncException("User chose not to continue sync.");
                }

                try {
                    //Add the Google event IDs into Outlook appointment.
                    //This needs to be done after the creation, else sporadic IrresolvableConflict HTTP 409 errors can occur
                    MsGraph.Event aiPatch = new MsGraph.Event() { 
                        Id = createdAi.Id,
                        Start = createdAi.Start,
                        Subject = createdAi.Subject,
                        SeriesMasterId = createdAi.SeriesMasterId,
                        Recurrence = createdAi.Recurrence
                    };
                    CustomProperty.AddGoogleIDs(ref aiPatch, ev);
                    UpdateCalendarEntry_save(ref aiPatch);
                } catch (System.Exception ex) {
                    ex.Analyse("Unable to save Extension data to newly created appointment.");
                    log.Warn("This should result in a 'reclaim' during the next sync.");
                }

                Recurrence.CreateOutlookExceptions(ev, createdAi);
            }
        }

        private void createCalendarEntry(GcalData.Event ev, ref MsGraph.Event ai) {
            string itemSummary = Ogcs.Google.Calendar.GetEventSummary(ev, out String anonItemSummary);
            log.Debug("Processing >> " + (anonItemSummary ?? itemSummary));
            Forms.Main.Instance.Console.Update(itemSummary, anonItemSummary, Console.Markup.calendar, verbose: true);

            SettingsStore.Calendar profile = Sync.Engine.Calendar.Instance.Profile;

            Int16 offset = 0;
            ai.Start = new MsGraph.DateTimeTimeZone();
            if (String.IsNullOrEmpty(ev.Start.TimeZone)) {
                log.Fine("Has no starting timezone.");
                ai.Start.TimeZone = "UTC";
            } else {
                offset = TimezoneDB.GetUtcOffset(ev.Start.TimeZone);
                log.Fine("Has starting timezone: " + ev.Start.TimeZone);
                ai.Start.TimeZone = ev.Start.TimeZone;
            }

            offset = 0;
            ai.End = new MsGraph.DateTimeTimeZone();
            if (String.IsNullOrEmpty(ev.End.TimeZone)) {
                log.Fine("Has no ending timezone.");
                ai.End.TimeZone = "UTC";
            } else {
                offset = TimezoneDB.GetUtcOffset(ev.End.TimeZone);
                log.Fine("Has ending timezone: " + ev.End.TimeZone);
                ai.End.TimeZone = ev.End.TimeZone;
            }

            if ((bool)(ai.IsAllDay = ev.AllDayEvent())) {
                ai.Start.DateTime = ev.Start.SafeDateTime().ToString("yyyy-MM-dd");
                ai.End.DateTime = ev.End.SafeDateTime().ToString("yyyy-MM-dd");
            } else {
                ai.Start.DateTime = ev.Start.SafeDateTime().AddMinutes(offset).ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ss");
                ai.End.DateTime = ev.End.SafeDateTime().AddMinutes(offset).ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ss");
            }

            ai.Recurrence = Recurrence.BuildOutlookPattern(ev);

            ai.Subject = Obfuscate.ApplyRegex(Obfuscate.Property.Subject, ev.Summary, null, Sync.Direction.GoogleToOutlook);
            if (profile.AddDescription && ev.Description != null) {
                ai.Body = new MsGraph.ItemBody() { ContentType = MsGraph.BodyType.Html };
                ai.Body.Content = Obfuscate.ApplyRegex(Obfuscate.Property.Description, ev.Description, null, Sync.Direction.GoogleToOutlook);
            }
            if (profile.AddLocation) ai.Location = new MsGraph.Location { DisplayName = Obfuscate.ApplyRegex(Obfuscate.Property.Location, ev.Location, null, Sync.Direction.GoogleToOutlook) };
            ai.Sensitivity = getPrivacy(ev.Visibility, null);
            ai.ShowAs = getAvailability(ev.Transparency, null);
            //ai.Categories = getColour(ev.ColorId, null);

            if (profile.AddAttendees && ev.Attendees != null) {
                if (ev.Attendees != null && ev.Attendees.Count > profile.MaxAttendees) {
                    log.Warn("This Google event has " + ev.Attendees.Count + " attendees, more than the user configured maximum.");
                } else {
                    List<MsGraph.Attendee> attendees = new List<MsGraph.Attendee>();
                    foreach (GcalData.EventAttendee ea in ev.Attendees) {
                        if (Settings.Instance.MSaccountEmail.ToLower() == ea.Email) continue;

                        attendees.Add(createRecipient(ea));
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
        }

        private MsGraph.Attendee createRecipient(GcalData.EventAttendee ea) {
            MsGraph.Attendee attendee = new MsGraph.Attendee() {
                EmailAddress = new MsGraph.EmailAddress() { Name = ea.DisplayName, Address = ea.Email },
                Type = (ea.Optional ?? false ? MsGraph.AttendeeType.Optional : MsGraph.AttendeeType.Required),
                Status = new MsGraph.ResponseStatus() { Response = MsGraph.ResponseType.None, Time = System.DateTime.UtcNow }
            };
            switch (ea.ResponseStatus) {
                case "needsAction": attendee.Status.Response = MsGraph.ResponseType.NotResponded; break;
                case "declined": attendee.Status.Response = MsGraph.ResponseType.Declined; break;
                case "tentative": attendee.Status.Response = MsGraph.ResponseType.TentativelyAccepted; break;
                case "accepted": attendee.Status.Response = MsGraph.ResponseType.Accepted; break;
            }
            return attendee;
        }

        private MsGraph.Event createCalendarEntry_save(MsGraph.Event ai, ref GcalData.Event ev) {
            SettingsStore.Calendar profile = Sync.Engine.Calendar.Instance.Profile;
            if (profile.SyncDirection.Id == Sync.Direction.Bidirectional.Id) {
                log.Debug("Saving timestamp when OGCS updated appointment.");
                CustomProperty.SetOGCSlastModified(ref ai);
            }

            MsGraph.Event createdAi = null;
            try {
                createdAi = GraphClient.Me.Calendars[profile.UseOutlookCalendar.Id].Events.PostAsync(ai).Result;
            } catch (System.Exception ex) {
                MsGraph.ODataErrors.ODataError oDataErr = O365Errors.GetODataError(ex);
                if (oDataErr?.Error?.Code == "InvalidAuthenticationToken") {
                    this.Authenticator.GetAuthenticated(true);
                } else {
                    switch (O365Errors.HandleAPIlimits(ref ex)) {
                        case ApiException.throwException: throw ex;
                        default: throw ex;
                    }
                }
            }

            if (createdAi != null && (profile.SyncDirection.Id == Sync.Direction.Bidirectional.Id || Ogcs.Google.CustomProperty.ExistAnyOutlookIDs(ev))) {
                log.Debug("Storing the Outlook appointment IDs in Google event.");
                Ogcs.Google.Graph.CustomProperty.AddOutlookIDs(ref ev, createdAi);
                Ogcs.Google.Calendar.Instance.UpdateCalendarEntry_save(ref ev);
            }

            return createdAi;
        }
        #endregion

        #region Update
        public void UpdateCalendarEntries(Dictionary<MsGraph.Event, GcalData.Event> entriesToBeCompared, ref int entriesUpdated) {
            foreach (KeyValuePair<MsGraph.Event, GcalData.Event> compare in entriesToBeCompared) {
                if (Sync.Engine.Instance.CancellationPending) return;

                int itemModified = 0;
                MsGraph.Event ai = compare.Key;

                Boolean aiWasRecurring = ai.Type == MsGraph.EventType.SeriesMaster;
                Boolean needsUpdating = false;
                MsGraph.Event aiPatch = new();
                try {
                    Boolean forceCompare = !aiWasRecurring && compare.Value.Recurrence != null;
                    needsUpdating = UpdateCalendarEntry(ref ai, compare.Value, ref itemModified, out aiPatch, forceCompare);
                } catch (System.Exception ex) {
                    Forms.Main.Instance.Console.UpdateWithError(Ogcs.Google.Calendar.GetEventSummary("<br/>Appointment update failed.", compare.Value, out String anonSummary), ex, logEntry: anonSummary);
                    Ogcs.Exception.Analyse(ex, true);
                    if (Ogcs.Extensions.MessageBox.Show("Outlook appointment update failed. Continue with synchronisation?", "Sync item failed", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                        continue;
                    else
                        throw new UserCancelledSyncException("User chose not to continue sync.");
                }

                if (itemModified > 0) {
                    try {
                        UpdateCalendarEntry_save(ref aiPatch);
                        ai = aiPatch;
                        entriesUpdated++;
                    } catch (System.Exception ex) {
                        Forms.Main.Instance.Console.UpdateWithError(Ogcs.Google.Calendar.GetEventSummary("Updated appointment failed to save.", compare.Value, out String anonSummary, true), ex, logEntry: anonSummary);
                        Ogcs.Exception.Analyse(ex, true);
                        if (Ogcs.Extensions.MessageBox.Show("Updated Outlook appointment failed to save. Continue with synchronisation?", "Sync item failed", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                            continue;
                        else
                            throw new UserCancelledSyncException("User chose not to continue sync.");
                    }
                    if (ai.Type == MsGraph.EventType.SeriesMaster) {
                        if (!aiWasRecurring) {
                            log.Debug("Appointment has changed from single instance to recurring.");
                            entriesUpdated += Recurrence.CreateOutlookExceptions(compare.Value, ai);
                        } else {
                            log.Debug("Recurring master appointment has been updated, so now checking if exceptions need reinstating.");
                            entriesUpdated += Recurrence.UpdateOutlookExceptions(compare.Value, ai, forceCompare: true);
                        }
                    }

                } else {
                    if (ai.Type == MsGraph.EventType.SeriesMaster && compare.Value.Recurrence != null && compare.Value.RecurringEventId == null) {
                        log.Debug(Ogcs.Google.Calendar.GetEventSummary(compare.Value));
                        entriesUpdated += Recurrence.UpdateOutlookExceptions(compare.Value, ai, forceCompare: false);

                    } else if (needsUpdating || CustomProperty.Exists(ai, CustomProperty.MetadataId.forceSave)) {
                        if (ai.LastModifiedDateTime > compare.Value.UpdatedDateTimeOffset && !CustomProperty.Exists(ai, CustomProperty.MetadataId.forceSave))
                            continue;

                        log.Debug("Doing a dummy update in order to update the last modified date.");
                        CustomProperty.SetOGCSlastModified(ref ai);
                        try {
                            UpdateCalendarEntry_save(ref ai);
                        } catch (System.Exception ex) {
                            Forms.Main.Instance.Console.UpdateWithError(Ogcs.Google.Calendar.GetEventSummary("Updated appointment failed to save.", compare.Value, out String anonSummary, true), ex, logEntry: anonSummary);
                            Ogcs.Exception.Analyse(ex, true);
                            if (Ogcs.Extensions.MessageBox.Show("Updated Outlook appointment failed to save. Continue with synchronisation?", "Sync item failed", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                                continue;
                            else
                                throw new UserCancelledSyncException("User chose not to continue sync.");
                        }
                    }
                }
            }
        }

        public Boolean UpdateCalendarEntry(ref MsGraph.Event ai, GcalData.Event ev, ref int itemModified, out MsGraph.Event aiPatch, Boolean forceCompare = false) {
            SettingsStore.Calendar profile = Sync.Engine.Calendar.Instance.Profile;
            aiPatch = new MsGraph.Event() { Id = ai.Id };

            if (!(Sync.Engine.Instance.ManualForceCompare || forceCompare)) { //Needed if the exception has just been created, but now needs updating
                if (profile.SyncDirection.Id != Sync.Direction.Bidirectional.Id) {
                    if (ai.LastModifiedDateTime > ev.UpdatedDateTimeOffset)
                        return false;
                } else {
                    if (Ogcs.Google.CustomProperty.GetOGCSlastModified(ev).AddSeconds(5) >= ev.UpdatedDateTimeOffset)
                        //Google last modified by OGCS
                        return false;
                    if (ai.LastModifiedDateTime > ev.UpdatedDateTimeOffset)
                        return false;
                }
            }

            String evSummary = Ogcs.Google.Calendar.GetEventSummary(ev, out String anonSummary);
            log.Debug("Processing >> " + (anonSummary ?? evSummary));

            StringBuilder sb = new StringBuilder();
            sb.AppendLine(evSummary);

            #region Start/End & TimeZone
            //Microsoft always convert Start/End.TimeZone to UTC and store the actual timezone in OriginalStart/EndTimeZone
            //Doesn't match their documentation at all, but hey ho.
            //https://learn.microsoft.com/en-us/graph/api/resources/event?view=graph-rest-1.0#properties

            Boolean startChange = false;
            Boolean endChange = false;
            Boolean aiAllDay = ai.AllDayEvent();
            OgcsDateTimeOffset aiStart = new(ai.Start.SafeDateTimeOffset(), aiAllDay);
            OgcsDateTimeOffset aiEnd = new(ai.End.SafeDateTimeOffset(), aiAllDay);
            if (ev.AllDayEvent()) {
                Sync.Engine.CompareAttribute("All-Day", Sync.Direction.GoogleToOutlook, true, aiAllDay, sb, ref itemModified);
                startChange = Sync.Engine.CompareAttribute("Start time", Sync.Direction.GoogleToOutlook, new OgcsDateTimeOffset(ev.Start.SafeDateTimeOffset(), true), aiStart, sb, ref itemModified);
                endChange = Sync.Engine.CompareAttribute("End time", Sync.Direction.GoogleToOutlook, new OgcsDateTimeOffset(ev.End.SafeDateTimeOffset(), true), aiEnd, sb, ref itemModified);
            } else {
                Sync.Engine.CompareAttribute("All-Day", Sync.Direction.GoogleToOutlook, false, aiAllDay, sb, ref itemModified);
                startChange = Sync.Engine.CompareAttribute("Start time", Sync.Direction.GoogleToOutlook, new OgcsDateTimeOffset(ev.Start.SafeDateTimeOffset(), false), aiStart, sb, ref itemModified);
                endChange = Sync.Engine.CompareAttribute("End time", Sync.Direction.GoogleToOutlook, new OgcsDateTimeOffset(ev.End.SafeDateTimeOffset(), false), aiEnd, sb, ref itemModified);
            }
            Boolean startTzChange = Sync.Engine.CompareAttribute("Start Timezone", Sync.Direction.GoogleToOutlook,
                string.IsNullOrEmpty(ev.Start.TimeZone) ? "UTC" : ev.Start.TimeZone, string.IsNullOrEmpty(ai.OriginalStartTimeZone) ? "UTC" : ai.OriginalStartTimeZone, sb, ref itemModified);
            Boolean endTzChange = Sync.Engine.CompareAttribute("End Timezone", Sync.Direction.GoogleToOutlook,
                string.IsNullOrEmpty(ev.End.TimeZone) ? "UTC" : ev.End.TimeZone, string.IsNullOrEmpty(ai.OriginalEndTimeZone) ? "UTC" : ai.OriginalEndTimeZone, sb, ref itemModified);

            if (startChange || startTzChange || endChange || endTzChange) {
                aiPatch.IsAllDay = ev.AllDayEvent();
                Int16 offset = 0;

                aiPatch.Start = ai.Start;
                aiPatch.End = ai.End;
                if ((bool)aiPatch.IsAllDay) {
                    aiPatch.Start.DateTime = ev.Start.SafeDateTime().ToString("yyyy-MM-dd");
                    aiPatch.End.DateTime = ev.End.SafeDateTime().ToString("yyyy-MM-dd");
                } else {
                    offset = TimezoneDB.GetUtcOffset(ev.Start.TimeZone);
                    aiPatch.Start.DateTime = ev.Start.SafeDateTime().AddMinutes(offset).ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ss");
                    offset = TimezoneDB.GetUtcOffset(ev.End.TimeZone);
                    aiPatch.End.DateTime = ev.End.SafeDateTime().AddMinutes(offset).ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ss");
                }
                aiPatch.Start.TimeZone = string.IsNullOrEmpty(ev.Start.TimeZone) ? aiPatch.Start.TimeZone : ev.Start.TimeZone;
                aiPatch.End.TimeZone = string.IsNullOrEmpty(ev.End.TimeZone) ? aiPatch.End.TimeZone : ev.End.TimeZone;
            }
            #endregion

            #region Recurrence
            aiPatch.Recurrence = ai.Recurrence;

            if (startChange || startTzChange || endTzChange) {
                if (ai.Type == MsGraph.EventType.SeriesMaster) {
                    if (startTzChange || endTzChange) {
                        aiPatch.Recurrence.Range.RecurrenceTimeZone = ai.Start.TimeZone;
                    }
                    if (startChange) {
                        aiPatch.Recurrence.Range.StartDate = ai.Start.SafeDateTime().ToGraphDate();
                    }
                }
            }

            if (ai.Type == MsGraph.EventType.SeriesMaster) {
                if (ev.Recurrence == null || ev.RecurringEventId != null) {
                    log.Debug("Converting to non-recurring appointment.");
                    aiPatch.AdditionalData = new Dictionary<String, Object>();
                    aiPatch.AdditionalData.Add("Recurrence", null);
                    sb.Append("Recurrence: => Removed.");
                    itemModified++;
                } else {
                    aiPatch.Recurrence = Recurrence.CompareOutlookPattern(ev, ai.Recurrence, Sync.Direction.GoogleToOutlook, sb, ref itemModified);
                }
            } else if (ai.Type == MsGraph.EventType.SingleInstance) {
                if (ev.Recurrence != null && ev.RecurringEventId == null) {
                    log.Debug("Converting to recurring appointment.");
                    aiPatch.Recurrence = Recurrence.BuildOutlookPattern(ev);
                    sb.Append("Recurrence: => Added");
                    itemModified++;
                }
            }
            #endregion

            String summaryObfuscated = Obfuscate.ApplyRegex(Obfuscate.Property.Subject, ev.Summary, ai.Subject, Sync.Direction.GoogleToOutlook);
            if (Sync.Engine.CompareAttribute("Subject", Sync.Direction.GoogleToOutlook, summaryObfuscated, ai.Subject, sb, ref itemModified)) {
                aiPatch.Subject = summaryObfuscated;
            }
            if (profile.AddDescription) {
                //String oGMeetUrl = CustomProperty.Get(ai, CustomProperty.MetadataId.gMeetUrl);

                if (profile.SyncDirection.Id == Sync.Direction.GoogleToOutlook.Id || !profile.AddDescription_OnlyToGoogle) {
                    String aiBody = ai.Body.BodyInnerHtml();
                    Boolean descriptionChanged = false;
                    if (!String.IsNullOrEmpty(aiBody)) {
                        /*Regex htmlDataTag = new Regex(@"<data:image.*?>");
                        aiBody = htmlDataTag.Replace(aiBody, "").Trim();
                        OlBodyFormat bodyFormat = ai.BodyFormat();
                        if (bodyFormat != OlBodyFormat.olFormatUnspecified)
                            aiBody = aiBody.Replace(GMeet.PlainInfo(oGMeetUrl, bodyFormat).RemoveLineBreaks(), "").Trim();*/
                    }
                    String bodyObfuscated = Obfuscate.ApplyRegex(Obfuscate.Property.Description, ev.Description?.RemoveNBSP(), aiBody, Sync.Direction.GoogleToOutlook);
                    if (bodyObfuscated.Length == 8 * 1024 && aiBody?.Length > 8 * 1024) {
                        log.Warn("Event description has been truncated, so will not be synced to Outlook.");
                    } else {
                        String evBodyForCompare = bodyObfuscated;
                        String aiBodyForCompare = aiBody;
                        //Remove HTML markup from Console output
                        String evTagsStripped = Regex.Replace(evBodyForCompare, "<.*?>", String.Empty);
                        String aiTagsStripped = Regex.Replace(aiBodyForCompare, "<.*?>", String.Empty);
                        
                        //Although there is a body type "Text" - MS automatically converts it to HTML!
                        //So the real way of detecting it is looking in the converted HTML...
                        Match plainText = Regex.Match(ai.Body.Content.RemoveLineBreaks(), @"<!-- converted from text -->.*</head><body>.*?<div class=\""PlainText\"">");
                        if (ai.Body.ContentType == MsGraph.BodyType.Text || plainText.Success) {
                            aiBodyForCompare = ai.BodyPreview;
                            aiTagsStripped = Regex.Replace(ai.BodyPreview, "<.*?>", String.Empty);
                        }
                        /*switch (ai.BodyFormat()) {
                            case OlBodyFormat.olFormatHTML:
                                evBodyForCompare = Regex.Replace(bodyObfuscated, "[\n]+", " "); break;
                            case OlBodyFormat.olFormatRichText:
                                evBodyForCompare = Regex.Replace(bodyObfuscated, "[\n]", ""); break;
                            case OlBodyFormat.olFormatPlain:
                                evBodyForCompare = Regex.Replace(bodyObfuscated, "[\n]", ""); break;
                        }*/
                        StringBuilder currentSB = new(sb.Capacity);
                        currentSB.Append(sb);
                        if (descriptionChanged = Sync.Engine.CompareAttribute("Description", Sync.Direction.GoogleToOutlook, evBodyForCompare, aiBodyForCompare, sb, ref itemModified)) {
                            aiPatch.Body = ai.Body;
                            aiPatch.Body.Content = evBodyForCompare;
                            String googleAttr_stub = ((evTagsStripped.Length > 50) ? evTagsStripped.Substring(0, 47) + "..." : evTagsStripped).RemoveLineBreaks();
                            String outlookAttr_stub = ((aiTagsStripped.Length > 50) ? aiTagsStripped.Substring(0, 47) + "..." : aiTagsStripped).RemoveLineBreaks();
                            sb = currentSB.AppendLine("Description" + ": " + outlookAttr_stub + " => " + googleAttr_stub);
                        }
                    }
                    /*if (profile.AddGMeet) {
                        if (descriptionChanged || Sync.Engine.CompareAttribute("Google Meet", Sync.Direction.GoogleToOutlook, ev.HangoutLink, oGMeetUrl, sb, ref itemModified)) {
                            ai.GoogleMeet(ev.HangoutLink);
                            if (String.IsNullOrEmpty(ev.HangoutLink) && !String.IsNullOrEmpty(oGMeetUrl) && !descriptionChanged) {
                                log.Debug("Removing GMeet information from body.");
                                ai.Body = bodyObfuscated;
                            }
                        }
                    }*/
                }
            }

            if (profile.AddLocation) {
                String locationObfuscated = Obfuscate.ApplyRegex(Obfuscate.Property.Description, ev.Location, ai.Location.DisplayName, Sync.Direction.GoogleToOutlook);
                if (Sync.Engine.CompareAttribute("Location", Sync.Direction.GoogleToOutlook, locationObfuscated, ai.Location.DisplayName, sb, ref itemModified)) {
                    if (ai.Locations.Count() <= 1) {
                        aiPatch.Location = ai.Location;
                        aiPatch.Location.DisplayName = locationObfuscated;
                    } else {
                        aiPatch.Locations = ai.Locations;
                        aiPatch.Locations.ElementAt(0).DisplayName = locationObfuscated;
                    }
                }
            }
            if (ai.Recurrence == null || ai.Type == MsGraph.EventType.SeriesMaster) {
                MsGraph.Sensitivity gPrivacy = getPrivacy(ev.Visibility, ai.Sensitivity);
                if (Sync.Engine.CompareAttribute("Privacy", Sync.Direction.GoogleToOutlook, gPrivacy.ToString(), ai.Sensitivity.ToString(), sb, ref itemModified)) {
                    aiPatch.Sensitivity = gPrivacy;
                }
            }
            MsGraph.FreeBusyStatus gFreeBusy = getAvailability(ev.Transparency ?? "opaque", ai.ShowAs);
            if (Sync.Engine.CompareAttribute("Free/Busy", Sync.Direction.GoogleToOutlook, gFreeBusy.ToString(), ai.ShowAs.ToString(), sb, ref itemModified)) {
                aiPatch.ShowAs = gFreeBusy;
            }

            /*
            if ((profile.AddColours || profile.SetEntriesColour) && (
                ai.RecurrenceState == OlRecurrenceState.olApptMaster ||
                ai.RecurrenceState == OlRecurrenceState.olApptNotRecurring)) //
            {
                log.Fine("Comparing colours/categories");
                List<String> aiCategories = new List<string>();
                String oCategoryName = "";
                if (!string.IsNullOrEmpty(ai.Categories)) {
                    aiCategories = ai.Categories.Split(new[] { Categories.Delimiter }, StringSplitOptions.None).ToList();
                    oCategoryName = aiCategories.FirstOrDefault();
                }
                String gCategoryName = getColour(ev.ColorId, oCategoryName ?? "");
                if (Sync.Engine.CompareAttribute("Category/Colour", Sync.Direction.GoogleToOutlook, gCategoryName, oCategoryName, sb, ref itemModified)) {
                    if (profile.SingleCategoryOnly)
                        aiCategories = new List<string>();
                    else {
                        //Only allow one OGCS category at a time (Google Events can only have one colour)
                        aiCategories.RemoveAll(x => x.StartsWith("OGCS ") || x == gCategoryName);
                    }
                    aiCategories.Insert(0, gCategoryName);
                    ai.Categories = String.Join(Categories.Delimiter, aiCategories.ToArray());
                }
            }
            */
            #region Attendees
            if (profile.AddAttendees) {
                if (ev.Attendees != null && ev.Attendees.Count > profile.MaxAttendees) {
                    log.Warn("This Google event has " + ev.Attendees.Count + " attendees, more than the user configured maximum.");
                } else if (profile.SyncDirection.Id == Sync.Direction.Bidirectional.Id &&
                        ai.Attendees.Count() > profile.MaxAttendees && (ev.Attendees == null ? 0 : ev.Attendees.Count) <= profile.MaxAttendees) {
                    log.Warn("This Outlook appointment has " + ai.Attendees.Count() + " attendees, more than the user configured maximum. They can't safely be compared.");
                } else {
                    log.Fine("Comparing meeting attendees");
                    List<MsGraph.Attendee> recipients = ai.Attendees.ToList();
                    List<GcalData.EventAttendee> addAttendees = new List<GcalData.EventAttendee>();

                    //Build a list of Google attendees. Any remaining at the end of the diff must be added.
                    if (ev.Attendees != null) {
                        addAttendees = ev.Attendees.ToList();
                    }
                    for (int o = recipients.Count() - 1; o >= 0; o--) {
                        Boolean foundAttendee = false;
                        MsGraph.Attendee recipient = recipients[o];

                        if (recipient.EmailAddress.Address == ai.Organizer.EmailAddress.Address) continue;

                        for (int g = (ev.Attendees == null ? -1 : ev.Attendees.Count - 1); g >= 0; g--) {
                            Ogcs.Google.EventAttendee attendee = new Ogcs.Google.EventAttendee(ev.Attendees[g]);
                            if (recipient.EmailAddress.Address == attendee.Email) {
                                foundAttendee = true;

                                //Optional attendee
                                bool oOptional = (recipient.Type ?? MsGraph.AttendeeType.Required) == MsGraph.AttendeeType.Optional;
                                bool gOptional = attendee.Optional ?? false;
                                if (Sync.Engine.CompareAttribute("Attendee " + (recipient.EmailAddress.Name ?? recipient.EmailAddress.Address) + " - Optional Check",
                                    Sync.Direction.GoogleToOutlook, gOptional, oOptional, sb, ref itemModified)) {
                                    recipient.Type = gOptional ? MsGraph.AttendeeType.Optional : MsGraph.AttendeeType.Required;
                                }
                                //Response status
                                MsGraph.Attendee compareRecipient = createRecipient(attendee);
                                if (Sync.Engine.CompareAttribute("Attendee " + (recipient.EmailAddress.Name ?? recipient.EmailAddress.Address) + " - Response Status",
                                    Sync.Direction.GoogleToOutlook, compareRecipient.Status.Response.Value.ToString(), recipient.Status.Response.Value.ToString(), sb, ref itemModified)) {
                                    recipient.Status = compareRecipient.Status;
                                }
                                addAttendees.Remove(ev.Attendees[g]);
                                break;
                            }
                        }
                        if (!foundAttendee) {
                            sb.AppendLine("Recipient removed: " + (recipient.EmailAddress.Name ?? recipient.EmailAddress.Address));
                            recipients.Remove(recipient);
                            itemModified++;
                        }
                    }
                    foreach (GcalData.EventAttendee gAttendee in addAttendees) {
                        Ogcs.Google.EventAttendee attendee = new Ogcs.Google.EventAttendee(gAttendee);
                        if (attendee.Email == ai.Organizer.EmailAddress.Address) continue; //Attendee in Google is owner in Outlook, so can't also be added as a recipient)

                        sb.AppendLine("Recipient added: " + (attendee.DisplayName ?? attendee.Email));
                        recipients.Add(createRecipient(gAttendee));
                        itemModified++;
                    }
                    aiPatch.Attendees = recipients;
                }
            }
            #endregion

            #region Reminders
            Boolean googleReminders = ev.Reminders?.Overrides?.Any(r => r.Method == "popup") ?? false;
            int reminderMins = int.MinValue;
            if (profile.AddReminders) {
                if (googleReminders) {
                    //Find the last popup reminder in Google
                    GcalData.EventReminder reminder = ev.Reminders.Overrides.Where(r => r.Method == "popup").OrderBy(r => r.Minutes).First();
                    reminderMins = (int)reminder.Minutes;
                } else if (ev.Reminders?.UseDefault ?? false) {
                    reminderMins = Ogcs.Google.Calendar.Instance.MinDefaultReminder;
                }

                if (reminderMins != int.MinValue) {
                    try {
                        if ((bool)ai.IsReminderOn) {
                            if (Sync.Engine.CompareAttribute("Reminder", Sync.Direction.GoogleToOutlook, reminderMins.ToString(), ai.ReminderMinutesBeforeStart.ToString(), sb, ref itemModified)) {
                                aiPatch.ReminderMinutesBeforeStart = reminderMins;
                            }
                        } else {
                            sb.AppendLine("Reminder: nothing => " + reminderMins);
                            aiPatch.IsReminderOn = true;
                            aiPatch.ReminderMinutesBeforeStart = reminderMins;
                            itemModified++;
                        }
                    } catch (System.Exception ex) {
                        ex.Analyse("Failed setting Outlook reminder for final popup Google notification.");
                    }
                }

            }
            if (!googleReminders && (!(ev.Reminders?.UseDefault ?? false) || reminderMins == int.MinValue)) {
                if ((bool)ai.IsReminderOn && !profile.UseOutlookDefaultReminder) {
                    sb.AppendLine("Reminder: " + ai.ReminderMinutesBeforeStart + " => removed");
                    aiPatch.IsReminderOn = false;
                    itemModified++;
                } else if (!(bool)ai.IsReminderOn && profile.UseOutlookDefaultReminder) {
                    sb.AppendLine("Reminder: nothing => default");
                    aiPatch.IsReminderOn = true;
                    itemModified++;
                }
            }
            #endregion

            if (itemModified > 0) {
                Forms.Main.Instance.Console.FormatEventChanges(sb, sb.ToString().Replace(evSummary, anonSummary));
                Forms.Main.Instance.Console.Update(itemModified + " attributes updated.", Console.Markup.appointmentEnd, verbose: true, newLine: false);
                System.Windows.Forms.Application.DoEvents();
            }
            return true;
        }

        public void UpdateCalendarEntry_save(ref MsGraph.Event ai) {
            ai.BackingStore.InitializationCompleted = false;

            SettingsStore.Calendar profile = Sync.Engine.Calendar.Instance.Profile;
            if (Sync.Engine.Calendar.Instance.Profile.SyncDirection.Id == Sync.Direction.Bidirectional.Id) {
                log.Debug("Saving timestamp when OGCS updated appointment.");
                CustomProperty.SetOGCSlastModified(ref ai);
            }
            CustomProperty.Remove(ref ai, CustomProperty.MetadataId.forceSave);

            try {
                MsGraph.Extension ogcsExtension = ai.OgcsExtension();
                if (ogcsExtension != null) {
                    Boolean patchExtension = false;
                    if (patchExtension = CustomProperty.Exists(ai, CustomProperty.MetadataId.requiresPatch)) {
                        ai.OgcsExtension().AdditionalData.Remove(CustomProperty.MetadataId.requiresPatch.ToString());
                    }
                    //Graph doesn't support removing properties via PATCH with null values. Have to manually delete and recreate
                    List<KeyValuePair<String, Object>> deletedProperties = ogcsExtension.AdditionalData.Where(prop => prop.Value == null).ToList();
                    if (deletedProperties.Count > 0) {
                        GraphClient.Me.Events[ai.Id].Extensions[CustomProperty.ExtensionName(true)].DeleteAsync().Wait();
                        ogcsExtension.AdditionalData = ogcsExtension.AdditionalData.Except(deletedProperties).ToDictionary(k => k.Key, k => k.Value);
                        patchExtension = true;
                    }
                    if (patchExtension) {
                        ogcsExtension = GraphClient.Me.Events[ai.Id].Extensions[CustomProperty.ExtensionName(true)].PatchAsync(ogcsExtension).Result;
                        ai.Extensions = new();
                    }
                }
                ai = GraphClient.Me.Events[ai.Id].PatchAsync(ai).Result;

                if (ogcsExtension != null)
                    ai = ai.UpdateOgcsExtension(ogcsExtension);

            } catch (System.Exception ex) {
                switch (O365Errors.HandleAPIlimits(ref ex)) {
                    case ApiException.throwException: throw ex;
                    default: throw ex;
                }
            }
        }
        #endregion

        #region Delete
        public void DeleteCalendarEntries(List<MsGraph.Event> oAppointments) {
            for (int o = oAppointments.Count - 1; o >= 0; o--) {
                if (Sync.Engine.Instance.CancellationPending) return;

                MsGraph.Event ai = oAppointments[o];
                Boolean doDelete = false;
                try {
                    doDelete = deleteCalendarEntry(ai);
                } catch (System.Exception ex) {
                    oAppointments.Remove(ai);
                    Forms.Main.Instance.Console.UpdateWithError(GetEventSummary("Appointment deletion failed.", ai, out String anonSummary, true), ex, logEntry: anonSummary);
                    Ogcs.Exception.Analyse(ex, true);
                    if (Ogcs.Extensions.MessageBox.Show("Outlook appointment deletion failed. Continue with synchronisation?", "Sync item failed", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                        continue;
                    else
                        throw new UserCancelledSyncException("User chose not to continue sync.");
                }

                try {
                    if (doDelete) DeleteCalendarEntry_save(ai);
                    else oAppointments.Remove(ai);
                } catch (System.Exception ex) {
                    oAppointments.Remove(ai);
                    if (ex is ApplicationException) {
                        String summary = GetEventSummary("<br/>Appointment deletion skipped.<br/>" + ex.Message, ai, out String anonSummary);
                        Forms.Main.Instance.Console.Update(summary, anonSummary, Console.Markup.warning);
                        if (ex.InnerException is MsGraph.ODataErrors.ODataError) break;
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

        private Boolean deleteCalendarEntry(MsGraph.Event ai) {
            String eventSummary = GetEventSummary(ai, out String anonSummary);
            Boolean doDelete = true;

            if (Sync.Engine.Calendar.Instance.Profile.ConfirmOnDelete) {
                if (Ogcs.Extensions.MessageBox.Show(
                    $"Calendar: {EmailAddress.MaskAddressWithinText(Sync.Engine.Calendar.Instance.Profile.UseOutlookCalendar.Name)}\r\nItem: {eventSummary}", "Confirm Deletion From Outlook",
                    MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2, anonSummary) == DialogResult.No
                ) { //
                    doDelete = false;
                    if (Sync.Engine.Calendar.Instance.Profile.SyncDirection.Id == Sync.Direction.Bidirectional.Id && CustomProperty.ExistAnyGoogleIDs(ai)) {
                        if (Ogcs.Google.Calendar.Instance.ExcludedByColour.ContainsKey(CustomProperty.Get(ai, CustomProperty.MetadataId.gEventID))) {
                            log.Fine("Refrained from removing Google metadata from Appointment; avoids duplication back into Google.");
                        } else {
                            CustomProperty.RemoveGoogleIDs(ref ai);
                            UpdateCalendarEntry_save(ref ai);
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

        public void DeleteCalendarEntry_save(MsGraph.Event ai) {
            try {
                GraphClient.Me.Events[ai.Id].DeleteAsync().Wait();
            } catch (System.Exception ex) {
                MsGraph.ODataErrors.ODataError oDataErr = O365Errors.GetODataError(ex);
                if (oDataErr != null) {
                    if (oDataErr.Error?.Code == "ErrorItemNotFound") { //Resource has been deleted
                        log.Fail("This event is already deleted! Ignoring failed request to delete.");
                        return;
                    }
                }
                switch (O365Errors.HandleAPIlimits(ref ex)) {
                    case ApiException.throwException: throw ex;
                    default: throw ex;
                }
            }
        }
        #endregion

        public static void ReclaimOrphanCalendarEntries(ref List<MsGraph.Event> oAppointments, ref List<GcalData.Event> gEvents) {
            SettingsStore.Calendar profile = Sync.Engine.Calendar.Instance.Profile;

            if (profile.SyncDirection.Id == Sync.Direction.OutlookToGoogle.Id) return;

            if (profile.SyncDirection.Id == Sync.Direction.GoogleToOutlook.Id)
                Forms.Main.Instance.Console.Update("Checking for orphaned Outlook items...", verbose: true);

            try {
                log.Debug("Scanning " + oAppointments.Count + " Outlook appointments for orphans to reclaim...");
                String consoleTitle = "Reclaiming Outlook calendar entries";

                //This is needed for people migrating from other tools, which do not have our GoogleID extendedProperty
                List<MsGraph.Event> unclaimedAi = new();

                for (int o = oAppointments.Count - 1; o >= 0; o--) {
                    if (Sync.Engine.Instance.CancellationPending) return;
                    MsGraph.Event ai = oAppointments[o];
                    try {
                        CustomProperty.LogProperties(ai, Program.MyFineLevel);

                        //Find entries with no Google ID
                        if (!CustomProperty.Exists(ai, CustomProperty.MetadataId.gEventID)) {
                            String sigAi = Signature(ai);
                            unclaimedAi.Add(ai);

                            for (int g = gEvents.Count - 1; g >= 0; g--) {
                                GcalData.Event ev = gEvents[g];
                                String sigEv = Ogcs.Google.Calendar.Signature(ev);
                                if (String.IsNullOrEmpty(sigEv)) {
                                    gEvents.Remove(ev);
                                    continue;
                                }

                                if (Ogcs.Google.Calendar.SignaturesMatch(sigEv, sigAi)) {
                                    CustomProperty.AddGoogleIDs(ref ai, ev);
                                    MsGraph.Event aiPatch = new() { Id = ai.Id, Extensions = ai.Extensions };
                                    Instance.UpdateCalendarEntry_save(ref aiPatch);
                                    unclaimedAi.Remove(ai);
                                    if (consoleTitle != "") Forms.Main.Instance.Console.Update("<span class='em em-reclaim'></span>" + consoleTitle, Console.Markup.h2, newLine: false, verbose: true);
                                    consoleTitle = "";
                                    Forms.Main.Instance.Console.Update(GetEventSummary("Reclaimed: ", ai, out String anonSummary, appendContext: false), anonSummary, verbose: true);
                                    oAppointments[o] = ai;

                                    if (profile.SyncDirection.Id == Sync.Direction.Bidirectional.Id || Ogcs.Google.CustomProperty.ExistAnyOutlookIDs(ev)) {
                                        log.Debug("Updating the Outlook appointment IDs in Google event.");
                                        Ogcs.Google.Graph.CustomProperty.AddOutlookIDs(ref ev, ai);
                                        Ogcs.Google.Calendar.Instance.UpdateCalendarEntry_save(ref ev);
                                        gEvents[g] = ev;
                                    }
                                    break;
                                }
                            }
                        }
                    } catch (System.Exception) {
                        Forms.Main.Instance.Console.Update(GetEventSummary("Failure processing Outlook item:-<br/>", ai, out String anonSummary, appendContext: false), anonSummary, Console.Markup.warning);
                        throw;
                    }
                    if (Sync.Engine.Instance.CancellationPending) return;
                }
                log.Debug(unclaimedAi.Count + " unclaimed.");
                if (unclaimedAi.Count > 0 &&
                    (profile.SyncDirection.Id == Sync.Direction.GoogleToOutlook.Id ||
                     profile.SyncDirection.Id == Sync.Direction.Bidirectional.Id)) //
                {
                    log.Info(unclaimedAi.Count + " unclaimed orphan appointments found.");
                    if (profile.MergeItems || profile.DisableDelete || profile.ConfirmOnDelete) {
                        log.Info("These will be kept due to configuration settings.");
                    } else if (profile.SyncDirection.Id == Sync.Direction.Bidirectional.Id) {
                        log.Debug("These 'orphaned' items must not be deleted - they need syncing up.");
                    } else {
                        if (Ogcs.Extensions.MessageBox.Show(unclaimedAi.Count + " Outlook calendar items can't be matched to Google.\r\n" +
                            "Remember, it's recommended to have a dedicated Outlook calendar to sync with, " +
                            "or you may wish to merge with unmatched events. Continue with deletions?",
                            "Delete unmatched Outlook items?", MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation) == DialogResult.No) {

                            log.Info("User has requested to keep them.");
                            foreach (MsGraph.Event ai in unclaimedAi) {
                                oAppointments.Remove(ai);
                            }
                        } else {
                            log.Info("User has opted to delete them.");
                        }
                    }
                }
            } catch (System.Exception) {
                Forms.Main.Instance.Console.Update("Unable to reclaim orphan calendar entries in Outlook calendar.", Console.Markup.error);
                throw;
            }
        }

        /// <summary>
        /// Determine Appointment Item's privacy setting
        /// </summary>
        /// <param name="gVisibility">Google's current setting</param>
        /// <param name="oSensitivity">Outlook's current setting</param>
        private MsGraph.Sensitivity getPrivacy(String gVisibility, MsGraph.Sensitivity? oSensitivity) {
            SettingsStore.Calendar profile = Sync.Engine.Calendar.Instance.Profile;

            if (!profile.SetEntriesPrivate)
                return (gVisibility == "private") ? MsGraph.Sensitivity.Private : MsGraph.Sensitivity.Normal;

            MsGraph.Sensitivity overrideSensitivity = MsGraph.Sensitivity.Normal;
            if (!Enum.TryParse(Regex.Replace(profile.PrivacyLevel, "^ol", ""), out overrideSensitivity))
                log.Error("Could not convert string '" + profile.PrivacyLevel + "' to Graph.Sensitivity type. Defaulting override to normal.");

            if (profile.TargetCalendar.Id == Sync.Direction.OutlookToGoogle.Id) { //Privacy enforcement is in other direction
                if (oSensitivity == null)
                    return (gVisibility == "private") ? MsGraph.Sensitivity.Private : MsGraph.Sensitivity.Normal;
                else if (!profile.CreatedItemsOnly && (gVisibility != ((overrideSensitivity == MsGraph.Sensitivity.Normal) ? "public" : "private"))) {
                    log.Warn("Google privacy override has been manually altered - so syncing this back.");
                    return (gVisibility == "private") ? MsGraph.Sensitivity.Private : MsGraph.Sensitivity.Normal;
                } else
                    return (MsGraph.Sensitivity)oSensitivity;
            } else {
                if (!profile.CreatedItemsOnly || (profile.CreatedItemsOnly && oSensitivity == null))
                    return overrideSensitivity;
                else {
                    if (profile.CreatedItemsOnly) return (MsGraph.Sensitivity)oSensitivity;
                    else return overrideSensitivity;
                }
            }
        }

        /// <summary>
        /// Determine Appointment's availability setting
        /// </summary>
        /// <param name="gTransparency">Google's current setting</param>
        /// <param name="oBusyStatus">Outlook's current setting</param>
        private MsGraph.FreeBusyStatus getAvailability(String gTransparency, MsGraph.FreeBusyStatus? oBusyStatus) {
            SettingsStore.Calendar profile = Sync.Engine.Calendar.Instance.Profile;
            List<MsGraph.FreeBusyStatus> persistOutlookStatus = new List<MsGraph.FreeBusyStatus> { MsGraph.FreeBusyStatus.Tentative, MsGraph.FreeBusyStatus.Oof, MsGraph.FreeBusyStatus.WorkingElsewhere };

            if (!profile.SetEntriesAvailable)
                return (gTransparency == "transparent") ? MsGraph.FreeBusyStatus.Free :
                    persistOutlookStatus.Contains(oBusyStatus ?? MsGraph.FreeBusyStatus.Busy) ? (MsGraph.FreeBusyStatus)oBusyStatus : MsGraph.FreeBusyStatus.Busy;

            MsGraph.FreeBusyStatus overrideFbStatus = MsGraph.FreeBusyStatus.Busy;
            if (!Enum.TryParse(Regex.Replace(profile.AvailabilityStatus, "^ol", "").Replace("OutOfOffice", "Oof"), out overrideFbStatus))
                log.Error("Could not convert string '" + profile.AvailabilityStatus + "' to Graph.FreeBusyStatus type. Defaulting override to busy.");

            if (profile.TargetCalendar.Id == Sync.Direction.OutlookToGoogle.Id) { //Availability enforcement is in other direction
                if (oBusyStatus == null)
                    return (gTransparency == "transparent") ? MsGraph.FreeBusyStatus.Free : MsGraph.FreeBusyStatus.Busy;

                else if (!profile.CreatedItemsOnly && (gTransparency != ((overrideFbStatus == MsGraph.FreeBusyStatus.Free) ? "transparent" : "opaque"))) {
                    log.Warn("Google availability override has been manually altered - so syncing this back.");
                    return (gTransparency == "transparent") ? MsGraph.FreeBusyStatus.Free : MsGraph.FreeBusyStatus.Busy;
                } else
                    return (MsGraph.FreeBusyStatus)oBusyStatus;
            } else {
                if (!profile.CreatedItemsOnly || (profile.CreatedItemsOnly && oBusyStatus == null))
                    return overrideFbStatus;
                else {
                    if (profile.CreatedItemsOnly || persistOutlookStatus.Contains((MsGraph.FreeBusyStatus)oBusyStatus))
                        return (MsGraph.FreeBusyStatus)oBusyStatus;
                    else
                        return overrideFbStatus;
                }
            }
        }


        #region STATIC functions
        public static string Signature(MsGraph.Event ai) {
            return (ai.Subject + ";" + ai.Start.SafeDateTimeOffset().ToPreciseString() + ";" + ai.End.SafeDateTimeOffset().ToPreciseString()).Trim();
        }

        public static void ExportToCSV(String action, String filename, List<MsGraph.Event> ais) {
            if (!Settings.Instance.CreateCSVFiles) return;

            log.Debug("CSV export: " + action);

            String fullFilename = Path.Combine(Program.UserFilePath, filename);
            try {
                if (System.IO.File.Exists(fullFilename)) {
                    String backupFilename = Path.Combine(Program.UserFilePath, Path.GetFileNameWithoutExtension(filename) + "-prev") + Path.GetExtension(filename);
                    if (System.IO.File.Exists(backupFilename)) System.IO.File.Delete(backupFilename);
                    System.IO.File.Move(fullFilename, backupFilename);
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
                    CSVheader += "Outlook GlobalID,Outlook EntryID,Outlook CalendarID,";
                    CSVheader += "Google EventID,Google CalendarID,OGCS Modified,Force Save";
                    tw.WriteLine(CSVheader);
                    foreach (MsGraph.Event ai in ais) {
                        try {
                            tw.WriteLine(exportToCSV(ai));
                        } catch (System.Exception ex) {
                            Forms.Main.Instance.Console.Update(GetEventSummary("Failed to output following Outlook appointment to CSV:-<br/>", ai, out String anonSummary, appendContext: false), anonSummary, Console.Markup.warning);
                            Ogcs.Exception.Analyse(ex, true);
                        }
                    }
                } catch (System.Exception ex) {
                    Forms.Main.Instance.Console.Update("Failed to output Outlook events to CSV.", Console.Markup.error);
                    Ogcs.Exception.Analyse(ex);
                }
            } finally {
                if (tw != null) tw.Close();
                if (stream != null) stream.Close();
            }
            log.Fine("CSV export done.");
        }
        private static string exportToCSV(MsGraph.Event ai) {
            StringBuilder csv = new StringBuilder();
            
            csv.Append(ai.Start.SafeDateTimeOffset().ToPreciseString() + ",");
            csv.Append(ai.End.SafeDateTimeOffset().ToPreciseString() + ",");
            csv.Append("\"" + ai.Subject + "\",");

            if (ai.Location == null) csv.Append(",");
            else csv.Append("\"" + ai.Location.DisplayName + "\",");

            if (ai.Body == null) csv.Append(",");
            else {
                String csvBody = ai.BodyPreview.Replace("\"", "");
                csvBody = csvBody.Replace("\r\n", " ");
                csv.Append("\"" + csvBody.Substring(0, System.Math.Min(csvBody.Length, 100)) + "\",");
            }

            csv.Append("\"" + ai.Sensitivity?.ToString() + "\",");
            csv.Append("\"" + ai.ShowAs?.ToString() + "\",");
            String[] requiredAttendees = ai.Attendees?.Where(a => a.Type == MsGraph.AttendeeType.Required).ToList().Select(a => a.EmailAddress.Name).ToArray() ;
            csv.Append("\"" + string.Join(";", requiredAttendees) + "\",");
            String[] optionalAttendees = ai.Attendees?.Where(a => a.Type == MsGraph.AttendeeType.Optional).ToList().Select(a => a.EmailAddress.Name).ToArray();
            csv.Append("\"" + string.Join(";", optionalAttendees) + "\",");
            csv.Append(ai.IsReminderOn + ",");
            csv.Append(ai.ReminderMinutesBeforeStart.ToString() + ",");
            csv.Append(ai.ICalUId + ",");
            csv.Append(ai.Id + "," + Sync.Engine.Calendar.Instance.Profile.UseOutlookCalendar.Id + ",");
            csv.Append((CustomProperty.Get(ai, CustomProperty.MetadataId.gEventID) ?? "") + ",");
            csv.Append((CustomProperty.Get(ai, CustomProperty.MetadataId.gCalendarId) ?? "") + ",");
            csv.Append((CustomProperty.Get(ai, CustomProperty.MetadataId.ogcsModified) ?? "") + ",");
            csv.Append((CustomProperty.Get(ai, CustomProperty.MetadataId.forceSave) ?? ""));

            return csv.ToString();
        }

        /// <summary>
        /// Get the anonymised summary of an appointment item, else standard summary.
        /// </summary>
        /// <param name="ai">The Graph Event item.</param>
        /// <returns>The summary, anonymised if settings dictate.</returns>
        public static String GetEventSummary(MsGraph.Event ai) {
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
        public static string GetEventSummary(String context, MsGraph.Event ai, out String eventSummaryAnonymised, Boolean onlyIfNotVerbose = false, Boolean appendContext = true) {
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
        public static string GetEventSummary(MsGraph.Event ai, out String eventSummaryAnonymised, Boolean onlyIfNotVerbose = false) {
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
                    eventSummary += " " + (ai.Recurrence != null ? "(R) " : (!string.IsNullOrEmpty(ai.SeriesMasterId) ? "(R1) " : "")) + "=> ";

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
            ref List<MsGraph.Event> outlook,          //need deleting
            ref Dictionary<MsGraph.Event, GcalData.Event> compare) //
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
                
                String compare_oEventID = CustomProperty.Get(outlook[o], CustomProperty.MetadataId.gEventID);
                if (!string.IsNullOrEmpty(compare_oEventID)) {
                    Boolean? googleIDmissing = null;
                    Boolean foundMatch = false;

                    for (int g = google.Count - 1; g >= 0; g--) {
                        log.UltraFine("Checking " + Ogcs.Google.Calendar.GetEventSummary(google[g]));

                        if (compare_oEventID == google[g].Id) {
                            googleIDmissing ??= CustomProperty.GoogleIdMissing(outlook[o]);
                            if ((Boolean)googleIDmissing) {
                                log.Info("Enhancing appointment's metadata...");
                                MsGraph.Event ai = outlook[o];
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
                        CustomProperty.Get(outlook[o], CustomProperty.MetadataId.gCalendarId) != profile.UseGoogleCalendar.Id)
                        outlook.Remove(outlook[o]);

                } else if (profile.MergeItems) {
                    //Remove the non-Google item so it doesn't get deleted
                    outlook.Remove(outlook[o]);
                }
            }
            if (metadataEnhanced > 0) log.Info(metadataEnhanced + " item's metadata enhanced.");
            
            if (profile.OnlyRespondedInvites) {
                //Check if items to be deleted have invitations not responded to
                List<MsGraph.Event> responseFiltered = new();
                responseFiltered = outlook.Where(ai => ai.ResponseStatus.Response == MsGraph.ResponseType.NotResponded).ToList();
                if (responseFiltered.Count > 0) log.Info(responseFiltered + " Outlook items will not be deleted due to only syncing invites that have been responded to.");
                outlook = outlook.Except(responseFiltered).ToList();
            }

            /*
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
            */
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
            if (profile.DisableDelete) {
                if (outlook.Count > 0) {
                    Forms.Main.Instance.Console.Update(outlook.Count + " Outlook items would have been deleted, but you have deletions disabled.", Console.Markup.warning);
                    for (int o = 0; o < outlook.Count; o++)
                        Forms.Main.Instance.Console.Update(GetEventSummary(outlook[o], out String anonSummary), anonSummary, verbose: true);
                }
                outlook = new();
            }
            if (Settings.Instance.CreateCSVFiles) {
                ExportToCSV("Appointments for deletion in Outlook", "outlook_delete.csv", outlook);
                Ogcs.Google.Calendar.ExportToCSV("Events for creation in Outlook", "outlook_create.csv", google);
            }
        }

        public static Boolean ItemIDsMatch(MsGraph.Event ai, GcalData.Event ev) {
            log.Fine("Comparing Google Event ID");
            SettingsStore.Calendar profile = Sync.Engine.Calendar.Instance.Profile;
            if (CustomProperty.Get(ai, CustomProperty.MetadataId.gEventID) == ev.Id) {
                log.Fine("Comparing Google Calendar ID");
                if (CustomProperty.Get(ai, CustomProperty.MetadataId.gCalendarId) == profile.UseGoogleCalendar.Id)
                    return true;
                else {
                    log.Warn("Could not find Google calendar ID against Outlook appointment item.");
                    return true;
                }
            } else {
                if (profile.MergeItems)
                    log.Fine("Could not find Google event ID against Outlook appointment item.");
                else
                    log.Warn("Could not find Google event ID against Outlook appointment item.");
            }
            return false;
        }

        public Boolean IsOKtoSyncReminder(MsGraph.Event ai) {
            SettingsStore.Calendar profile = Sync.Engine.Calendar.Instance.Profile;

            if (profile.ReminderDND) {
                System.DateTime alarm;
                if ((bool)ai.IsReminderOn)
                    alarm = ai.Start.SafeDateTime().AddMinutes((int)-ai.ReminderMinutesBeforeStart);
                else {
                    if (profile.UseGoogleDefaultReminder && Ogcs.Google.Calendar.Instance.MinDefaultReminder != int.MinValue) {
                        log.Fine("Using default Google reminder value: " + Ogcs.Google.Calendar.Instance.MinDefaultReminder);
                        alarm = ai.Start.SafeDateTime().AddMinutes(-Ogcs.Google.Calendar.Instance.MinDefaultReminder);
                    } else
                        return false;
                }
                return Outlook.Calendar.Instance.IsOKtoSyncReminder(alarm);
            }
            return true;
        }
        #endregion
    }
}
