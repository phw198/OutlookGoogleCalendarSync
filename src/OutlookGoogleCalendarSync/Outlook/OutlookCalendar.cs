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

namespace OutlookGoogleCalendarSync.Outlook {
    /// <summary>
    /// Class to target Outlook.Calendar for sync.
    /// </summary>
    public class Calendar {
        private static Calendar instance;
        private static readonly ILog log = LogManager.GetLogger(typeof(Calendar));
        public Interface IOutlook;

        /// <summary>Whether instance of OutlookCalendar class should connect to Outlook application</summary>
        public static Boolean InstanceConnect { get; private set; }

        public static Calendar Instance {
            get {
                try {
                    if (instance == null)
                        instance = new Calendar();
                    if (instance.Folders == null)
                        instance.IOutlook.Connect();

                } catch (System.ApplicationException) {
                    throw;
                } catch (System.Exception ex) {
                    if (ex is System.Runtime.InteropServices.COMException) Ogcs.Exception.LogAsFail(ref ex);
                    Ogcs.Exception.Analyse(ex);
                    log.Info("It appears Outlook has been closed/restarted after OGCS was started. Reconnecting...");
                    instance = new Calendar();
                    instance.IOutlook.Connect();
                }
                return instance;
            }
        }
        
        /// <summary>Force OGCS to drop Outlook connection at the end of a sync. Useful when appointments have become inaccessible.</summary>
        public static Boolean ForceClientReconnect = false;

        public static Boolean OOMsecurityInfo = false;
        private static List<String> alreadyRedirectedToWikiForComError = new List<String>();
        public const String GlobalIdPattern = "040000008200E00074C5B7101A82E008";
        public Folders Folders {
            get { return IOutlook.Folders(); }
        }
        public Dictionary<string, MAPIFolder> CalendarFolders {
            get { return IOutlook.CalendarFolders(); }
        }
        public static Outlook.Categories Categories;
        public enum Service {
            DefaultMailbox,
            AlternativeMailbox,
            SharedCalendar
        }
        public Outlook.EphemeralProperties EphemeralProperties = new EphemeralProperties();

        /// <summary>Outlook Appointment excluded through user config <Appt.EntryId, Event.Id></summary>
        public Dictionary<String, String> ExcludedByCategory { get; private set; }

        public Calendar() {
            InstanceConnect = true;
            IOutlook = Factory.GetOutlookInterface();
        }

        public void Reset() {
            log.Info("Resetting connection to Outlook.");
            if (IOutlook != null) Disconnect();
            instance = new Calendar();
            instance.IOutlook.Connect();
        }

        /// <summary>
        /// Wrapper for IOutlook.Disconnect - cannot dereference fully inside interface
        /// </summary>
        public static void Disconnect(Boolean onlyWhenNoGUI = false) {
            if (instance == null) return;

            try {
                InstanceConnect = false;
                Instance.IOutlook.Disconnect(onlyWhenNoGUI);
            } catch (System.Exception ex) {
                ex.LogAsFail().Analyse("Could not disconnect from Outlook.");
            } finally {
                GC.Collect();
                GC.WaitForPendingFinalizers();
                InstanceConnect = true;
            }
        }

        /// <summary>
        /// Get all calendar entries within the defined date-range for sync
        /// </summary>
        /// <param name="suppressAdvisories">Don't give user feedback, eg during background Push sync</param>
        /// <returns></returns>
        public List<AppointmentItem> GetCalendarEntriesInRange(SettingsStore.Calendar profile, Boolean suppressAdvisories) {
            List<AppointmentItem> filtered = new List<AppointmentItem>();
            try {
                filtered = FilterCalendarEntries(profile, suppressAdvisories: suppressAdvisories);
            } catch (System.Runtime.InteropServices.InvalidComObjectException ex) {
                if (Outlook.Errors.HandleComError(ex) == Outlook.Errors.ErrorType.ObjectSeparatedFromRcw) {
                    try { Outlook.Calendar.Instance.Reset(); } catch { }
                    ex.Data.Add("OGCS", "Failed to access the Outlook calendar. Please try again.");
                    throw;
                }
            } catch (System.Runtime.InteropServices.COMException ex) {
                if (ex.GetErrorCode(0x0000FFFF) == "0x00004005" && ex.TargetSite.Name == "get_LastModificationTime") { //You must specify a time.
                    Ogcs.Exception.LogAsFail(ref ex);
                    ex.Data.Add("OGCS", "Corrupted item(s) with no start/end date exist in your Outlook calendar that need fixing or removing before a sync can run.<br/>" +
                        "Switch the calendar folder to <i>List View</i>, sort by date and look for entries with no start and/or end date.");
                } else if (ex.GetErrorCode(0x000FFFFF) == "0x00020009") { //One or more items in the folder you synchronized do not match. 
                    Ogcs.Exception.LogAsFail(ref ex);
                    String wikiURL = "https://github.com/phw198/OutlookGoogleCalendarSync/wiki/Resolving-Outlook-Error-Messages#one-or-more-items-in-the-folder-you-synchronized-do-not-match";
                    ex.Data.Add("OGCS", ex.Message + "<br/>Please view the wiki for suggestions on " +
                        "<a href='" + wikiURL + "'>how to resolve conflicts</a> within your Outlook account.");
                    if (!suppressAdvisories && Ogcs.Extensions.MessageBox.Show("Your Outlook calendar contains conflicts that need resolving in order to sync successfully.\r\nView the wiki for advice?",
                        "Outlook conflicts exist", MessageBoxButtons.YesNo, MessageBoxIcon.Warning) == DialogResult.Yes)
                        Helper.OpenBrowser(wikiURL);
                }
                throw;

            } catch (System.ArgumentNullException ex) {
                ex.LogAsFail().Analyse("It seems that Outlook has just been closed.");
                Outlook.Calendar.Instance.Reset();
                filtered = FilterCalendarEntries(profile, suppressAdvisories: suppressAdvisories);

            } catch (System.Exception) {
                if (!suppressAdvisories) Forms.Main.Instance.Console.Update("Unable to access the Outlook calendar.", Console.Markup.error);
                throw;
            }
            return filtered;
        }

        public List<AppointmentItem> FilterCalendarEntries(SettingsStore.Calendar profile, Boolean filterBySettings = true, Boolean suppressAdvisories = false) {
            //Filtering info @ https://msdn.microsoft.com/en-us/library/cc513841%28v=office.12%29.aspx

            List<AppointmentItem> result = new List<AppointmentItem>();
            Items OutlookItems = null;
            ExcludedByCategory = new();

            if (profile is null)
                profile = Settings.Profile.InPlay();

            try {
                MAPIFolder thisUseOutlookCalendar = IOutlook.GetFolderByID(profile.UseOutlookCalendar.Id);
                OutlookItems = thisUseOutlookCalendar.Items;
            } catch {
                log.Fail("Could not open '" + Settings.Profile.Name(profile) + "' profile calendar folder with ID " + profile.UseOutlookCalendar.Id);
                throw;
            }

            if (OutlookItems != null) {
                log.Fine(OutlookItems.Count + " calendar items exist.");

                OutlookItems.Sort("[Start]", Type.Missing);
                OutlookItems.IncludeRecurrences = false;

                System.DateTime min = System.DateTime.MinValue;
                System.DateTime max = System.DateTime.MaxValue;
                min = profile.SyncStart;
                max = profile.SyncEnd;

                string filter = "[End] >= '" + min.ToString(profile.OutlookDateFormat) +
                    "' AND [Start] < '" + max.ToString(profile.OutlookDateFormat) + "'";
                log.Fine("Filter string: " + filter);

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
                        if (ai.End == min && !ai.IsRecurring) continue; //Required for midnight to midnight events 
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
                            if (filtered) {
                                try {
                                    ExcludedByCategory.Add(ai.EntryID, CustomProperty.Get(ai, CustomProperty.MetadataId.gEventID)); continue;
                                } catch (System.ArgumentException ex) {
                                    if (ex.Message == "An item with the same key has already been added.") {
                                        log.Warn("Investigating issue #2233 and duplicate EntryID...");
                                        try {
                                            List<AppointmentItem> ais = new();
                                            IOutlook.FilterItems(OutlookItems, filter).ForEach(ai => ais.Add(ai as AppointmentItem));
                                            ExportToCSV("Appointments containing possible duplicate EntryID", "outlook_appointments_duplicateEntryID.csv", ais);

                                            Dictionary<String, List<AppointmentItem>> duplicates = ais.
                                                GroupBy(ai => ai.EntryID).
                                                Where(group => group.Count() > 1).
                                                ToDictionary(group => group.Key, group => group.ToList());
                                            log.Warn($"Outlook has {duplicates.Count} duplicate EntryIDs");
                                            if (duplicates.Count > 0) {
                                                foreach (KeyValuePair<String, List<AppointmentItem>> duplicate in duplicates) {
                                                    log.Debug(duplicate.Key);
                                                    duplicate.Value.ForEach(d => {
                                                        log.Debug("   -> " + d.GlobalAppointmentID);
                                                        log.Debug("      " + GetEventSummary(d));
                                                    });
                                                }
                                            }
                                        } catch (System.Exception ex2) {
                                            ex2.Analyse("Couldn't export all the Outlook appointments, in search of duplicate EntryIDs");
                                        }
                                    }
                                    throw;
                                }
                            }

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
                                //We don't want them getting automatically deleted if brought back in scope; better to create possible duplicate
                                CustomProperty.RemoveGoogleIDs(ref ai);
                                ai.Save();
                            }
                        }

                        result.Add(ai);
                    }
                }
                if (availabilityFiltered > 0) log.Info(availabilityFiltered + " Outlook items excluded due to availability.");
                if (allDayFiltered > 0) log.Info(allDayFiltered + " Outlook all day items excluded.");
                if (ExcludedByCategory.Count > 0) log.Info(ExcludedByCategory.Count + " Outlook items contain a category that is filtered out.");
                if (subjectFiltered > 0) log.Info(subjectFiltered + " Outlook items with subject containing '" + profile.ExcludeSubjectText + "' filtered out.");
                if (responseFiltered > 0) log.Info(responseFiltered + " Outlook items are invites not yet responded to.");

                Int32 allExcluded = availabilityFiltered + allDayFiltered + ExcludedByCategory.Count + subjectFiltered + responseFiltered;
                if (allExcluded > 0 && !suppressAdvisories) {
                    String filterWarning = "Due to your OGCS Outlook settings, " + (result.Count == 0 ? "all" : allExcluded) + " Outlook items have been filtered out" + (result.Count == 0 ? "!" : ".");
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
            }
            log.Fine("Filtered down to " + result.Count);
            return result;
        }

        #region Create
        public void CreateCalendarEntries(List<Event> events) {
            for (int g = events.Count -1; g >= 0; g--) {
                if (Sync.Engine.Instance.CancellationPending) return;

                Event ev = events[g];
                AppointmentItem newAi = IOutlook.UseOutlookCalendar().Items.Add() as AppointmentItem;
                try {
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

                    try {
                        createCalendarEntry_save(newAi, ref ev);
                        events[g] = ev;
                    } catch (System.Exception ex) {
                        events.RemoveAt(g);
                        Forms.Main.Instance.Console.UpdateWithError(Ogcs.Google.Calendar.GetEventSummary("New appointment failed to save.", ev, out String anonSummary, true), ex, logEntry: anonSummary);
                        Ogcs.Exception.Analyse(ex, true);
                        if (Ogcs.Extensions.MessageBox.Show("New Outlook appointment failed to save. Continue with synchronisation?", "Sync item failed", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                            continue;
                        else
                            throw new UserCancelledSyncException("User chose not to continue sync.");
                    }

                    if (ev.Recurrence != null && ev.RecurringEventId == null && Google.Recurrence.HasExceptions(ev)) {
                        Forms.Main.Instance.Console.Update("This is a recurring item with some exceptions:-", verbose: true);
                        Recurrence.CreateOutlookExceptions(ev, ref newAi);
                        Forms.Main.Instance.Console.Update("Recurring exceptions completed.", verbose: true);
                    }
                } finally {
                    newAi = (AppointmentItem)ReleaseObject(newAi);
                }
            }
        }

        private void createCalendarEntry(Event ev, ref AppointmentItem ai) {
            string itemSummary = Ogcs.Google.Calendar.GetEventSummary(ev, out String anonItemSummary);
            log.Debug("Processing >> " + (anonItemSummary ?? itemSummary));
            Forms.Main.Instance.Console.Update(itemSummary, anonItemSummary, Console.Markup.calendar, verbose: true);

            SettingsStore.Calendar profile = Sync.Engine.Calendar.Instance.Profile;

            ai.Start = new System.DateTime();
            ai.End = new System.DateTime();
            ai.AllDayEvent = ev.AllDayEvent();
            ai = Outlook.Calendar.Instance.IOutlook.WindowsTimeZone_set(ai, ev);
            Recurrence.BuildOutlookPattern(ev, ai);

            ai.Subject = Obfuscate.ApplyRegex(Obfuscate.Property.Subject, ev.Summary, null, Sync.Direction.GoogleToOutlook);
            if (profile.AddDescription && ev.Description != null) ai.Body = Obfuscate.ApplyRegex(Obfuscate.Property.Description, ev.Description, null, Sync.Direction.GoogleToOutlook);
            if (profile.AddLocation) ai.Location = Obfuscate.ApplyRegex(Obfuscate.Property.Location, ev.Location, null, Sync.Direction.GoogleToOutlook);
            ai.Sensitivity = getPrivacy(ev.Visibility, null);
            ai.BusyStatus = getAvailability(ev.Transparency, null);
            ai.Categories = getColour(ev.ColorId, null);

            if (profile.AddAttendees && ev.Attendees != null) {
                if (ev.Attendees != null && ev.Attendees.Count > profile.MaxAttendees) {
                    log.Warn("This Google event has " + ev.Attendees.Count + " attendees, more than the user configured maximum.");
                } else {
                    foreach (EventAttendee ea in ev.Attendees) {
                        Recipients recipients = ai.Recipients;
                        createRecipient(ea, ref recipients);
                        recipients = (Recipients)ReleaseObject(recipients);
                    }
                }
            }

            //Reminder alert
            if (profile.AddReminders) {
                if (ev.Reminders?.Overrides?.Any(r => r.Method == "popup") ?? false) {
                    ai.ReminderSet = true;
                    try {
                        EventReminder reminder = ev.Reminders.Overrides.Where(r => r.Method == "popup").OrderBy(x => x.Minutes).First();
                        ai.ReminderMinutesBeforeStart = (int)reminder.Minutes;
                    } catch (System.Exception ex) {
                        ex.Analyse("Failed setting Outlook reminder for final popup Google notification.");
                    }
                } else if ((ev.Reminders?.UseDefault ?? false) && Ogcs.Google.Calendar.Instance.MinDefaultReminder != int.MinValue) {
                    ai.ReminderSet = true;
                    ai.ReminderMinutesBeforeStart = Ogcs.Google.Calendar.Instance.MinDefaultReminder;
                } else {
                    ai.ReminderSet = profile.UseOutlookDefaultReminder;
                }
            } else ai.ReminderSet = profile.UseOutlookDefaultReminder;

            if (profile.AddGMeet && !String.IsNullOrEmpty(ev.HangoutLink)) {
                ai.GoogleMeet(ev.HangoutLink);
            }

            //Add the Google event IDs into Outlook appointment.
            CustomProperty.AddGoogleIDs(ref ai, ev);
        }

        private static void createCalendarEntry_save(AppointmentItem ai, ref Event ev) {
            SettingsStore.Calendar profile = Sync.Engine.Calendar.Instance.Profile;
            if (profile.SyncDirection.Id == Sync.Direction.Bidirectional.Id) {
                log.Debug("Saving timestamp when OGCS updated appointment.");
                CustomProperty.SetOGCSlastModified(ref ai);
            }

            ai.Save();

            if (profile.SyncDirection.Id == Sync.Direction.Bidirectional.Id || Ogcs.Google.CustomProperty.ExistAnyOutlookIDs(ev)) {
                log.Debug("Storing the Outlook appointment IDs in Google event.");
                Ogcs.Google.CustomProperty.AddOutlookIDs(ref ev, ai);
                try {
                    Ogcs.Google.Calendar.Instance.UpdateCalendarEntry_save(ref ev);
                } catch (System.Exception) {
                    log.Debug(Newtonsoft.Json.JsonConvert.SerializeObject(ev));
                    throw;
                }
            }
        }
        #endregion

        #region Update
        public void UpdateCalendarEntries(Dictionary<AppointmentItem, Event> entriesToBeCompared, ref int entriesUpdated) {
            foreach (KeyValuePair<AppointmentItem, Event> compare in entriesToBeCompared) {
                if (Sync.Engine.Instance.CancellationPending) return;

                int itemModified = 0;
                AppointmentItem ai = compare.Key;
                try {
                    Boolean aiWasRecurring = ai.IsRecurring;
                    Boolean needsUpdating = false;
                    try {
                        Boolean forceCompare = !aiWasRecurring && compare.Value.Recurrence != null;
                        needsUpdating = UpdateCalendarEntry(ref ai, compare.Value, ref itemModified, forceCompare);
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
                            updateCalendarEntry_save(ref ai);
                            entriesUpdated++;
                        } catch (System.Exception ex) {
                            Forms.Main.Instance.Console.UpdateWithError(Ogcs.Google.Calendar.GetEventSummary("Updated appointment failed to save.", compare.Value, out String anonSummary, true), ex, logEntry: anonSummary);
                            Ogcs.Exception.Analyse(ex, true);
                            if (Ogcs.Extensions.MessageBox.Show("Updated Outlook appointment failed to save. Continue with synchronisation?", "Sync item failed", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                                continue;
                            else
                                throw new UserCancelledSyncException("User chose not to continue sync.");
                        }
                        if (ai.IsRecurring) {
                            if (!aiWasRecurring) {
                                log.Debug("Appointment has changed from single instance to recurring.");
                                entriesUpdated += Recurrence.CreateOutlookExceptions(compare.Value, ref ai);
                            } else {
                                log.Debug("Recurring master appointment has been updated, so now checking if exceptions need reinstating.");
                                entriesUpdated += Recurrence.UpdateOutlookExceptions(compare.Value, ref ai, forceCompare: true);
                            }
                        }

                    } else {
                        if (ai.RecurrenceState == OlRecurrenceState.olApptMaster && compare.Value.Recurrence != null && compare.Value.RecurringEventId == null) {
                            log.Debug(Ogcs.Google.Calendar.GetEventSummary(compare.Value));
                            entriesUpdated += Recurrence.UpdateOutlookExceptions(compare.Value, ref ai, forceCompare: false);

                        } else if (needsUpdating || CustomProperty.Exists(ai, CustomProperty.MetadataId.forceSave)) {
                            if (ai.LastModificationTime > compare.Value.UpdatedDateTimeOffset && !CustomProperty.Exists(ai, CustomProperty.MetadataId.forceSave))
                                continue;

                            log.Debug("Doing a dummy update in order to update the last modified date.");
                            CustomProperty.SetOGCSlastModified(ref ai);
                            updateCalendarEntry_save(ref ai);
                        }
                    }
                } finally {
                    ai = (AppointmentItem)ReleaseObject(ai);
                }
            }
        }

        public Boolean UpdateCalendarEntry(ref AppointmentItem ai, Event ev, ref int itemModified, Boolean forceCompare = false) {
            SettingsStore.Calendar profile = Sync.Engine.Calendar.Instance.Profile;

            if (!(Sync.Engine.Instance.ManualForceCompare || forceCompare)) { //Needed if the exception has just been created, but now needs updating
                if (profile.SyncDirection.Id != Sync.Direction.Bidirectional.Id) {
                    if (ai.LastModificationTime > ev.UpdatedDateTimeOffset)
                        return false;
                } else {
                    if (Ogcs.Google.CustomProperty.GetOGCSlastModified(ev).AddSeconds(5) >= ev.UpdatedDateTimeOffset)
                        //Google last modified by OGCS
                        return false;
                    if (ai.LastModificationTime > ev.UpdatedDateTimeOffset)
                        return false;
                }
            }

            String evSummary = Ogcs.Google.Calendar.GetEventSummary(ev, out String anonSummary);
            log.Debug("Processing >> " + (anonSummary ?? evSummary));

            StringBuilder sb = new StringBuilder();
            sb.AppendLine(evSummary);

            Boolean aiAllDay = ai.AllDayEvent;
            if (ai.RecurrenceState != OlRecurrenceState.olApptMaster) {
                if (Sync.Engine.CompareAttribute("All-Day", Sync.Direction.GoogleToOutlook, ev.AllDayEvent(), aiAllDay, sb, ref itemModified))
                    ai.AllDayEvent = ev.AllDayEvent();
            }

            #region TimeZone
            String currentStartTZ = "UTC";
            String currentEndTZ = "UTC";
            String newStartTZ = "UTC";
            String newEndTZ = "UTC";
            IOutlook.WindowsTimeZone_get(ai, out currentStartTZ, out currentEndTZ);
            ai = Outlook.Calendar.Instance.IOutlook.WindowsTimeZone_set(ai, ev, onlyTZattribute: true);
            IOutlook.WindowsTimeZone_get(ai, out newStartTZ, out newEndTZ);
            Boolean startTzChange = Sync.Engine.CompareAttribute("Start Timezone", Sync.Direction.GoogleToOutlook, newStartTZ, currentStartTZ, sb, ref itemModified);
            Boolean endTzChange = Sync.Engine.CompareAttribute("End Timezone", Sync.Direction.GoogleToOutlook, newEndTZ, currentEndTZ, sb, ref itemModified);
            #endregion

            #region Start/End & Recurrence
            Boolean startChange = false;
            Boolean endChange = false;
            OgcsDateTime aiStart = new(ai.Start, aiAllDay);
            OgcsDateTime aiEnd = new(ai.End, aiAllDay);
            System.DateTime evStartParsedDate = ev.Start.SafeDateTime();
            System.DateTime evEndParsedDate = ev.End.SafeDateTime();
            if (ev.AllDayEvent()) {
                startChange = Sync.Engine.CompareAttribute("Start time", Sync.Direction.GoogleToOutlook, new OgcsDateTime(evStartParsedDate, true), aiStart, sb, ref itemModified);
                endChange = Sync.Engine.CompareAttribute("End time", Sync.Direction.GoogleToOutlook, new OgcsDateTime(evEndParsedDate, true), aiEnd, sb, ref itemModified);
            } else {
                startChange = Sync.Engine.CompareAttribute("Start time", Sync.Direction.GoogleToOutlook, new OgcsDateTime(evStartParsedDate, false), aiStart, sb, ref itemModified);
                endChange = Sync.Engine.CompareAttribute("End time", Sync.Direction.GoogleToOutlook, new OgcsDateTime(evEndParsedDate, false), aiEnd, sb, ref itemModified);
            }
            RecurrencePattern oPattern = null;
            try {
                if (startChange || endChange || startTzChange || endTzChange) {
                    if (ai.RecurrenceState == OlRecurrenceState.olApptMaster) {
                        if (startTzChange || endTzChange) {
                            oPattern = (RecurrencePattern)Outlook.Calendar.ReleaseObject(oPattern);
                            ai.ClearRecurrencePattern();
                            ai = Outlook.Calendar.Instance.IOutlook.WindowsTimeZone_set(ai, ev, onlyTZattribute: false);
                            ai.Save();
                            Recurrence.BuildOutlookPattern(ev, ai);
                            ai.Save(); //Explicit save required to make ai.IsRecurring true again
                        } else {
                            oPattern = (ai.RecurrenceState == OlRecurrenceState.olApptMaster) ? ai.GetRecurrencePattern() : null;
                            if (startChange) {
                                oPattern.PatternStartDate = evStartParsedDate;
                                oPattern.StartTime = TimeZoneInfo.ConvertTime(evStartParsedDate, TimeZoneInfo.FindSystemTimeZoneById(newStartTZ));
                            }
                            if (endChange) {
                                oPattern.PatternEndDate = evEndParsedDate;
                                oPattern.EndTime = TimeZoneInfo.ConvertTime(evEndParsedDate, TimeZoneInfo.FindSystemTimeZoneById(newEndTZ));
                            }
                        }
                    } else {
                        ai = Outlook.Calendar.Instance.IOutlook.WindowsTimeZone_set(ai, ev);
                    }
                }

                if (oPattern == null)
                    oPattern = (ai.RecurrenceState == OlRecurrenceState.olApptMaster) ? ai.GetRecurrencePattern() : null;
                if (oPattern != null) {
                    oPattern.Duration = Convert.ToInt32((evEndParsedDate - evStartParsedDate).TotalMinutes);
                    Recurrence.CompareOutlookPattern(ev, ref oPattern, Sync.Direction.GoogleToOutlook, sb, ref itemModified);
                }
            } finally {
                oPattern = (RecurrencePattern)ReleaseObject(oPattern);
            }

            if (ai.RecurrenceState == OlRecurrenceState.olApptMaster) {
                if (ev.Recurrence == null || ev.RecurringEventId != null) {
                    log.Debug("Converting to non-recurring events.");
                    ai.ClearRecurrencePattern();
                    sb.Append("Recurrence: => Removed.");
                    itemModified++;
                }
            } else if (ai.RecurrenceState == OlRecurrenceState.olApptNotRecurring) {
                if (!ai.IsRecurring && ev.Recurrence != null && ev.RecurringEventId == null) {
                    log.Debug("Converting to recurring appointment.");
                    Recurrence.BuildOutlookPattern(ev, ai);
                    sb.Append("Recurrence: => Added");
                    itemModified++;
                }
            }
            #endregion

            String summaryObfuscated = Obfuscate.ApplyRegex(Obfuscate.Property.Subject, ev.Summary, ai.Subject, Sync.Direction.GoogleToOutlook);
            if (Sync.Engine.CompareAttribute("Subject", Sync.Direction.GoogleToOutlook, summaryObfuscated, ai.Subject, sb, ref itemModified)) {
                ai.Subject = summaryObfuscated;
            }
            if (profile.AddDescription) {
                String oGMeetUrl = CustomProperty.Get(ai, CustomProperty.MetadataId.gMeetUrl);

                if (profile.SyncDirection.Id == Sync.Direction.GoogleToOutlook.Id || !profile.AddDescription_OnlyToGoogle) {
                    String aiBody = ai.Body?.RemoveLineBreaks();
                    Boolean descriptionChanged = false;
                    if (!String.IsNullOrEmpty(aiBody)) {
                        Regex htmlDataTag = new Regex(@"<data:image.*?>");
                        aiBody = htmlDataTag.Replace(aiBody, "").Trim();
                        OlBodyFormat bodyFormat = ai.BodyFormat();
                        if (bodyFormat != OlBodyFormat.olFormatUnspecified)
                            aiBody = aiBody.Replace(GMeet.PlainInfo(oGMeetUrl, bodyFormat).RemoveLineBreaks(), "").Trim();
                    }
                    String bodyObfuscated = Obfuscate.ApplyRegex(Obfuscate.Property.Description, ev.Description?.RemoveNBSP(), aiBody, Sync.Direction.GoogleToOutlook);
                    if (bodyObfuscated.Length == 8 * 1024 && aiBody?.Length > 8 * 1024) {
                        log.Warn("Event description has been truncated, so will not be synced to Outlook.");
                    } else {
                        String evBodyForCompare = bodyObfuscated;
                        switch (ai.BodyFormat()) {
                            case OlBodyFormat.olFormatHTML:
                                evBodyForCompare = Regex.Replace(bodyObfuscated, "[\n]+", " "); break;
                            case OlBodyFormat.olFormatRichText:
                                evBodyForCompare = Regex.Replace(bodyObfuscated, "[\n]", ""); break;
                            case OlBodyFormat.olFormatPlain:
                                evBodyForCompare = Regex.Replace(bodyObfuscated, "[\n]", ""); break;
                        }
                        if (descriptionChanged = Sync.Engine.CompareAttribute("Description", Sync.Direction.GoogleToOutlook, evBodyForCompare, aiBody, sb, ref itemModified))
                            ai.Body = bodyObfuscated;
                    }
                    if (profile.AddGMeet) {
                        if (descriptionChanged || Sync.Engine.CompareAttribute("Google Meet", Sync.Direction.GoogleToOutlook, ev.HangoutLink, oGMeetUrl, sb, ref itemModified)) {
                            ai.GoogleMeet(ev.HangoutLink);
                            if (String.IsNullOrEmpty(ev.HangoutLink) && !String.IsNullOrEmpty(oGMeetUrl) && !descriptionChanged) {
                                log.Debug("Removing GMeet information from body.");
                                ai.Body = bodyObfuscated;
                            }
                        }
                    }
                }
            }

            if (profile.AddLocation) {
                String locationObfuscated = Obfuscate.ApplyRegex(Obfuscate.Property.Description, ev.Location, ai.Location, Sync.Direction.GoogleToOutlook);
                if (Sync.Engine.CompareAttribute("Location", Sync.Direction.GoogleToOutlook, locationObfuscated, ai.Location, sb, ref itemModified))
                    ai.Location = ev.Location;
            }

            if (ai.RecurrenceState == OlRecurrenceState.olApptMaster ||
                ai.RecurrenceState == OlRecurrenceState.olApptNotRecurring) {
                OlSensitivity gPrivacy = getPrivacy(ev.Visibility, ai.Sensitivity);
                if (Sync.Engine.CompareAttribute("Privacy", Sync.Direction.GoogleToOutlook, gPrivacy.ToString(), ai.Sensitivity.ToString(), sb, ref itemModified)) {
                    ai.Sensitivity = gPrivacy;
                }
            }
            OlBusyStatus gFreeBusy = getAvailability(ev.Transparency ?? "opaque", ai.BusyStatus);
            if (Sync.Engine.CompareAttribute("Free/Busy", Sync.Direction.GoogleToOutlook, gFreeBusy.ToString(), ai.BusyStatus.ToString(), sb, ref itemModified)) {
                ai.BusyStatus = gFreeBusy;
            }

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

            #region Attendees
            if (profile.AddAttendees) {
                if (ev.Attendees != null && ev.Attendees.Count > profile.MaxAttendees) {
                    log.Warn("This Google event has " + ev.Attendees.Count + " attendees, more than the user configured maximum.");
                } else if (profile.SyncDirection.Id == Sync.Direction.Bidirectional.Id &&
                        ai.Recipients.Count > profile.MaxAttendees && (ev.Attendees == null ? 0 : ev.Attendees.Count) <= profile.MaxAttendees) {
                    log.Warn("This Outlook appointment has " + ai.Recipients.Count + " attendees, more than the user configured maximum. They can't safely be compared.");
                } else {
                    log.Fine("Comparing meeting attendees");
                    Recipients recipients = ai.Recipients;
                    List<EventAttendee> addAttendees = new List<EventAttendee>();
                    try {
                        //Build a list of Google attendees. Any remaining at the end of the diff must be added.
                        if (ev.Attendees != null) {
                            addAttendees = ev.Attendees.ToList();
                        }
                        for (int r = 1; r <= recipients.Count; r++) {
                            Recipient recipient = null;
                            Boolean foundAttendee = false;
                            try {
                                recipient = recipients[r];
                                if (recipient.Name == ai.Organizer) continue;

                                if (!recipient.Resolved) recipient.Resolve();
                                String recipientSMTP = IOutlook.GetRecipientEmail(recipient);

                                for (int g = (ev.Attendees == null ? -1 : ev.Attendees.Count - 1); g >= 0; g--) {
                                    Ogcs.Google.EventAttendee attendee = new Ogcs.Google.EventAttendee(ev.Attendees[g]);
                                    if (recipientSMTP.ToLower() == attendee.Email.ToLower()) {
                                        foundAttendee = true;

                                        //Optional attendee
                                        bool oOptional = (ai.OptionalAttendees != null && ai.OptionalAttendees.Contains(attendee.DisplayName ?? attendee.Email));
                                        bool gOptional = (attendee.Optional == null) ? false : (bool)attendee.Optional;
                                        if (Sync.Engine.CompareAttribute("Recipient " + recipient.Name + " - Optional Check",
                                            Sync.Direction.GoogleToOutlook, gOptional, oOptional, sb, ref itemModified)) {
                                            if (gOptional) {
                                                recipient.Type = (int)OlMeetingRecipientType.olOptional;
                                            } else {
                                                recipient.Type = (int)OlMeetingRecipientType.olRequired;
                                            }
                                        }
                                        //Response is readonly in Outlook :(
                                        addAttendees.Remove(ev.Attendees[g]);
                                        break;
                                    }
                                }
                                if (!foundAttendee) {
                                    sb.AppendLine("Recipient removed: " + recipient.Name);
                                    recipient.Delete();
                                    itemModified++;
                                }
                            } finally {
                                recipient = (Recipient)Outlook.Calendar.ReleaseObject(recipient);
                            }
                        }
                        foreach (EventAttendee gAttendee in addAttendees) {
                            Ogcs.Google.EventAttendee attendee = new Ogcs.Google.EventAttendee(gAttendee);
                            if ((attendee.DisplayName ?? attendee.Email) == ai.Organizer) continue; //Attendee in Google is owner in Outlook, so can't also be added as a recipient)

                            sb.AppendLine("Recipient added: " + (attendee.DisplayName ?? attendee.Email));
                            createRecipient(attendee, ref recipients);
                            itemModified++;
                        }
                    } finally {
                        recipients = (Recipients)Outlook.Calendar.ReleaseObject(recipients);
                    }
                }
            }
            #endregion

            #region Reminders
            Boolean googleReminders = ev.Reminders?.Overrides?.Any(r => r.Method == "popup") ?? false;
            int reminderMins = int.MinValue;
            if (profile.AddReminders) {
                if (googleReminders) {
                    //Find the last popup reminder in Google
                    EventReminder reminder = ev.Reminders.Overrides.Where(r => r.Method == "popup").OrderBy(r => r.Minutes).First();
                    reminderMins = (int)reminder.Minutes;
                } else if (ev.Reminders?.UseDefault ?? false) {
                    reminderMins = Ogcs.Google.Calendar.Instance.MinDefaultReminder;
                }

                if (reminderMins != int.MinValue) {
                    try {
                        if (ai.ReminderSet) {
                            if (Sync.Engine.CompareAttribute("Reminder", Sync.Direction.GoogleToOutlook, reminderMins.ToString(), ai.ReminderMinutesBeforeStart.ToString(), sb, ref itemModified)) {
                                ai.ReminderMinutesBeforeStart = reminderMins;
                            }
                        } else {
                            sb.AppendLine("Reminder: nothing => " + reminderMins);
                            ai.ReminderSet = true;
                            ai.ReminderMinutesBeforeStart = reminderMins;
                            itemModified++;
                        }
                    } catch (System.Exception ex) {
                        ex.Analyse("Failed setting Outlook reminder for final popup Google notification.");
                    }
                }

            }
            if (!googleReminders && (!(ev.Reminders?.UseDefault ?? false) || reminderMins == int.MinValue)) {
                if (ai.ReminderSet && !profile.UseOutlookDefaultReminder) {
                    sb.AppendLine("Reminder: " + ai.ReminderMinutesBeforeStart + " => removed");
                    ai.ReminderSet = false;
                    itemModified++;
                } else if (!ai.ReminderSet && profile.UseOutlookDefaultReminder) {
                    sb.AppendLine("Reminder: nothing => default");
                    ai.ReminderSet = true;
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

        private void updateCalendarEntry_save(ref AppointmentItem ai) {
            if (Sync.Engine.Calendar.Instance.Profile.SyncDirection.Id == Sync.Direction.Bidirectional.Id) {
                log.Debug("Saving timestamp when OGCS updated appointment.");
                CustomProperty.SetOGCSlastModified(ref ai);
            }
            CustomProperty.Remove(ref ai, CustomProperty.MetadataId.forceSave);
            ai.Save();
        }
        #endregion

        #region Delete
        public void DeleteCalendarEntries(List<AppointmentItem> oAppointments) {
            for (int o = oAppointments.Count - 1; o >= 0; o--) {
                if (Sync.Engine.Instance.CancellationPending) return;

                AppointmentItem ai = oAppointments[o];
                Boolean doDelete = false;
                try {
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
                        if (doDelete) deleteCalendarEntry_save(ai);
                        else oAppointments.Remove(ai);
                    } catch (System.Exception ex) {
                        oAppointments.Remove(ai);
                        Forms.Main.Instance.Console.UpdateWithError(GetEventSummary("Deleted appointment failed to remove.", ai, out String anonSummary, true), ex, logEntry: anonSummary);
                        Ogcs.Exception.Analyse(ex, true);
                        if (Ogcs.Extensions.MessageBox.Show("Deleted Outlook appointment failed to remove. Continue with synchronisation?", "Sync item failed", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                            continue;
                        else
                            throw new UserCancelledSyncException("User chose not to continue sync.");
                    }
                } finally {
                    ai = (AppointmentItem)ReleaseObject(ai);
                }
            }
        }

        private Boolean deleteCalendarEntry(AppointmentItem ai) {
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
                            ai.Save();
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

        private void deleteCalendarEntry_save(AppointmentItem ai) {
            ai.Delete();
        }
        #endregion

        public void ReclaimOrphanCalendarEntries(ref List<AppointmentItem> oAppointments, ref List<Event> gEvents) {
            SettingsStore.Calendar profile = Sync.Engine.Calendar.Instance.Profile;

            if (profile.SyncDirection.Id == Sync.Direction.OutlookToGoogle.Id) return;

            if (profile.SyncDirection.Id == Sync.Direction.GoogleToOutlook.Id)
                Forms.Main.Instance.Console.Update("Checking for orphaned Outlook items...", verbose: true);

            try {
                log.Debug("Scanning " + oAppointments.Count + " Outlook appointments for orphans to reclaim...");
                String consoleTitle = "Reclaiming Outlook calendar entries";

                //This is needed for people migrating from other tools, which do not have our GoogleID extendedProperty
                List<AppointmentItem> unclaimedAi = new List<AppointmentItem>();

                for (int o = oAppointments.Count - 1; o >= 0; o--) {
                    if (Sync.Engine.Instance.CancellationPending) return;
                    AppointmentItem ai = oAppointments[o];
                    try {
                        CustomProperty.LogProperties(ai, Program.MyFineLevel);

                        //Find entries with no Google ID
                        if (!CustomProperty.Exists(ai, CustomProperty.MetadataId.gEventID)) {
                            String sigAi = signature(ai);
                            unclaimedAi.Add(ai);

                            for (int g = gEvents.Count - 1; g >= 0; g--) {
                                Event ev = gEvents[g];
                                String sigEv = Ogcs.Google.Calendar.signature(ev);
                                if (String.IsNullOrEmpty(sigEv)) {
                                    gEvents.Remove(ev);
                                    continue;
                                }

                                if (Ogcs.Google.Calendar.SignaturesMatch(sigEv, sigAi)) {
                                    CustomProperty.AddGoogleIDs(ref ai, ev);
                                    updateCalendarEntry_save(ref ai);
                                    unclaimedAi.Remove(ai);
                                    if (consoleTitle != "") Forms.Main.Instance.Console.Update("<span class='em em-reclaim'></span>" + consoleTitle, Console.Markup.h2, newLine: false, verbose: true);
                                    consoleTitle = "";
                                    Forms.Main.Instance.Console.Update(GetEventSummary("Reclaimed: ", ai, out String anonSummary, appendContext: false), anonSummary, verbose: true);
                                    oAppointments[o] = ai;

                                    if (profile.SyncDirection.Id == Sync.Direction.Bidirectional.Id || Ogcs.Google.CustomProperty.ExistAnyOutlookIDs(ev)) {
                                        log.Debug("Updating the Outlook appointment IDs in Google event.");
                                        Ogcs.Google.CustomProperty.AddOutlookIDs(ref ev, ai);
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
                            foreach (AppointmentItem ai in unclaimedAi) {
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

        private void createRecipient(EventAttendee gea, ref Recipients recipients) {
            Ogcs.Google.EventAttendee ea = new Ogcs.Google.EventAttendee(gea);
            if (IOutlook.CurrentUserSMTP().ToLower() != ea.Email) {
                Recipient recipient = null;
                try {
                    recipient = recipients.Add(ea.DisplayName + "<" + ea.Email + ">");
                    try {
                        recipient.Resolve();
                    } catch (System.Runtime.InteropServices.COMException ex) {
                        ex.LogAsFail().Analyse("Unable to resolve recipient against address book.");
                        log.Debug($"Resolved: {recipient.Resolved.ToString()}; Address: {recipient.Address}; Name: {recipient.Name};");
                    }
                    //ReadOnly: recipient.Type = (int)((bool)ea.Organizer ? OlMeetingRecipientType.olOrganizer : OlMeetingRecipientType.olRequired);
                    recipient.Type = (int)(ea.Optional == null ? OlMeetingRecipientType.olRequired : ((bool)ea.Optional ? OlMeetingRecipientType.olOptional : OlMeetingRecipientType.olRequired));
                    //ReadOnly: ea.ResponseStatus
                } finally {
                    recipient = (Recipient)Outlook.Calendar.ReleaseObject(recipient);
                }
            }
        }

        /// <summary>
        /// Determine Appointment Item's privacy setting
        /// </summary>
        /// <param name="gVisibility">Google's current setting</param>
        /// <param name="oSensitivity">Outlook's current setting</param>
        private OlSensitivity getPrivacy(String gVisibility, OlSensitivity? oSensitivity) {
            SettingsStore.Calendar profile = Sync.Engine.Calendar.Instance.Profile;

            if (!profile.SetEntriesPrivate)
                return (gVisibility == "private") ? OlSensitivity.olPrivate : OlSensitivity.olNormal;

            OlSensitivity overrideSensitivity = OlSensitivity.olNormal;
            try {
                Enum.TryParse(profile.PrivacyLevel, out overrideSensitivity);
            } catch (System.Exception ex) {
                ex.Analyse("Could not convert string '" + profile.PrivacyLevel + "' to OlSensitivity type. Defaulting override to normal.");
            }

            if (profile.TargetCalendar.Id == Sync.Direction.OutlookToGoogle.Id) { //Privacy enforcement is in other direction
                if (oSensitivity == null)
                    return (gVisibility == "private") ? OlSensitivity.olPrivate : OlSensitivity.olNormal;
                else
                    return (OlSensitivity)oSensitivity;
            } else {
                if (!profile.CreatedItemsOnly || (profile.CreatedItemsOnly && oSensitivity == null))
                    return overrideSensitivity;
                else {
                    if (profile.CreatedItemsOnly) return (OlSensitivity)oSensitivity;
                    else return overrideSensitivity;
                }
            }
        }

        /// <summary>
        /// Determine Appointment's availability setting
        /// </summary>
        /// <param name="gTransparency">Google's current setting</param>
        /// <param name="oBusyStatus">Outlook's current setting</param>
        private OlBusyStatus getAvailability(String gTransparency, OlBusyStatus? oBusyStatus) {
            SettingsStore.Calendar profile = Sync.Engine.Calendar.Instance.Profile;
            List<String> persistOutlookStatus = new List<String> { OlBusyStatus.olTentative.ToString(), OlBusyStatus.olOutOfOffice.ToString(), "olWorkingElsewhere" };

            if (!profile.SetEntriesAvailable)
                return (gTransparency == "transparent") ? OlBusyStatus.olFree :
                    persistOutlookStatus.Contains(oBusyStatus.ToString()) ? (OlBusyStatus)oBusyStatus : OlBusyStatus.olBusy;

            OlBusyStatus overrideFbStatus = OlBusyStatus.olBusy;
            try {
                Enum.TryParse(profile.AvailabilityStatus, out overrideFbStatus);
            } catch (System.Exception ex) {
                ex.Analyse("Could not convert string '" + profile.AvailabilityStatus + "' to OlBusyStatus type. Defaulting override to busy.");
            }

            if (profile.TargetCalendar.Id == Sync.Direction.OutlookToGoogle.Id) { //Availability enforcement is in other direction
                if (oBusyStatus == null)
                    return (gTransparency == "transparent") ? OlBusyStatus.olFree : OlBusyStatus.olBusy;
                else
                    return (OlBusyStatus)oBusyStatus;
            } else {
                if (!profile.CreatedItemsOnly || (profile.CreatedItemsOnly && oBusyStatus == null))
                    return overrideFbStatus;
                else {
                    if (profile.CreatedItemsOnly || persistOutlookStatus.Contains(oBusyStatus.ToString()))
                        return (OlBusyStatus)oBusyStatus;
                    else
                        return overrideFbStatus;
                }
            }
        }

        /// <summary>
        /// Get the Outlook category colour name from a Google colour ID
        /// </summary>
        /// <param name="gColourId">The Google colour ID</param>
        /// <param name="oColour">The Outlook category, if already assigned to appointment</param>
        /// <returns>Outlook category name</returns>
        private String getColour(String gColourId, String oColour) {
            SettingsStore.Calendar profile = Sync.Engine.Calendar.Instance.Profile;

            if (!profile.AddColours && !profile.SetEntriesColour) return "";

            OlCategoryColor outlookColour = Ogcs.Outlook.Categories.Map.Colours.Where(c => c.Key.ToString() == profile.SetEntriesColourValue).FirstOrDefault().Key;
            String overrideColour = Categories.FindName(outlookColour, profile.SetEntriesColourName);

            if (profile.SetEntriesColour) {
                if (profile.TargetCalendar.Id == Sync.Direction.OutlookToGoogle.Id) { //Colour forced to sync in other direction
                    if (oColour == null) //Creating item
                        return "";
                    else return oColour;

                } else {
                    if (!profile.CreatedItemsOnly || (profile.CreatedItemsOnly && oColour == null))
                        return overrideColour;
                    else {
                        if (profile.CreatedItemsOnly) return oColour;
                        else return overrideColour;
                    }
                }

            } else {
                return GetCategoryColour(gColourId ?? "0");
            }
        }
        public String GetCategoryColour(String gColourId, Boolean createMissingCategory = true) {
            OlCategoryColor? outlookColour = null;

            SettingsStore.Calendar profile = Settings.Profile.InPlay();
            if (profile.ColourMaps.Count > 0) {
                KeyValuePair<String, String> kvp = profile.ColourMaps.FirstOrDefault(cm => cm.Value == gColourId);
                if (kvp.Key != null) {
                    outlookColour = Outlook.Calendar.Categories.OutlookColour(kvp.Key);
                    if (outlookColour != null) {
                        log.Debug("Colour mapping used: " + kvp.Value + ":" + Ogcs.Google.Calendar.Instance.ColourPalette.GetColour(gColourId).Name + " => " + kvp.Key);
                        return kvp.Key;
                    }
                }
            }

            //Algorithmic closest colour matching
            Ogcs.Google.EventColour.Palette pallete = Ogcs.Google.Calendar.Instance.ColourPalette.GetColour(gColourId);
            if (pallete == Ogcs.Google.EventColour.Palette.NullPalette) return null;

            outlookColour = Categories.Map.GetClosestCategory(pallete);
            return Categories.FindName(outlookColour, createMissingCategory: createMissingCategory);
        }

        #region STATIC functions
        public static void AttachToOutlook(ref Microsoft.Office.Interop.Outlook.Application oApp, Boolean openOutlookOnFail = true, Boolean withSystemCall = false) {
            if (System.Diagnostics.Process.GetProcessesByName("OUTLOOK").Count() > 0) {
                log.Info("Attaching to the already running Outlook process.");
                try {
                    oApp = System.Runtime.InteropServices.Marshal.GetActiveObject("Outlook.Application") as Microsoft.Office.Interop.Outlook.Application;
                    if (oApp == null)
                        throw new ApplicationException("GetActiveObject() returned NULL without throwing an error.");
                } catch (System.Exception ex) {
                    if (Outlook.Errors.HandleComError(ex) == Outlook.Errors.ErrorType.Unavailable) { //MK_E_UNAVAILABLE
                        log.Warn("Attachment failed - Outlook is running without GUI for programmatic access.");
                    } else {
                        log.Warn("Attachment failed.");
                        Ogcs.Exception.Analyse(ex);
                    }
                    if (openOutlookOnFail) openOutlookHandler(ref oApp, withSystemCall);
                }
            } else {
                log.Warn("No Outlook process available to attach to.");
                if (openOutlookOnFail) openOutlookHandler(ref oApp, withSystemCall);
            }
        }

        private static void openOutlookHandler(ref Microsoft.Office.Interop.Outlook.Application oApp, Boolean withSystemCall = false) {
            int openAttempts = 1;
            int maxAttempts = 3;
            while (openAttempts <= maxAttempts) {
                try {
                    openOutlook(ref oApp, withSystemCall);
                    openAttempts = maxAttempts;
                } catch (ApplicationException aex) {
                    if (aex.Message == "Outlook is busy.") {
                        log.Warn(aex.Message + " Attempt " + openAttempts + "/" + maxAttempts);
                        if (openAttempts == maxAttempts) {
                            String message = "Outlook has been unresponsive for " + maxAttempts * 10 + " seconds.\n" +
                                "Please try running OGCS again later" +
                                (Settings.Instance.StartOnStartup ? " or " + ((Settings.Instance.StartupDelay == 0) ? "set a" : "increase the") + " delay on startup." : ".");

                            if (aex.InnerException.Message.Contains("CO_E_SERVER_EXEC_FAILURE"))
                                message += "\nAlso check that one of OGCS and Outlook are not running 'as Administrator' or if Outlook's stuck loading, eg. waiting for an Outlook profile to be chosen.";

                            throw new ApplicationException(message);
                        }
                        System.Threading.Thread.Sleep(10000);
                    } else {
                        throw;
                    }
                }
                openAttempts++;
            }
        }
        private static void openOutlook(ref Microsoft.Office.Interop.Outlook.Application oApp, Boolean withSystemCall = false) {
            log.Info("Starting a new instance of Outlook.");
            try {
                if (!withSystemCall)
                    oApp = new Microsoft.Office.Interop.Outlook.Application();
                else {
                    System.Diagnostics.Process oProcess = new System.Diagnostics.Process();
                    oProcess.StartInfo.FileName = "outlook";
                    oProcess.StartInfo.Arguments = "/recycle";
                    oProcess.Start();

                    int maxWaits = 8;
                    while (maxWaits > 0 && oApp == null) {
                        if (maxWaits % 2 == 0) log.Info("Waiting for Outlook to start...");
                        oProcess.WaitForInputIdle(15000);
                        Outlook.Calendar.AttachToOutlook(ref oApp, openOutlookOnFail: false);
                        if (oApp == null) {
                            log.Debug("Reattempting starting Outlook without system call.");
                            try { oApp = new Microsoft.Office.Interop.Outlook.Application(); } catch (System.Exception ex) { log.Debug("Errored with: " + ex.Message); }
                        }
                        maxWaits--;
                    }
                    if (oApp == null) {
                        log.Error("Giving up waiting for Outlook to open!");
                        throw new System.ApplicationException("Could not establish a connection with Outlook.");
                    }
                }
            } catch (System.Exception ex) {
                oApp = null;
                PoorlyOfficeInstall(ex);
            }
        }

        /// <summary>
        /// An exception handler for COM errors etc when attaching to/accessing Outlook
        /// </summary>
        public static void PoorlyOfficeInstall(System.Exception caughtException) {
            try {
                throw caughtException;
            } catch (System.Runtime.InteropServices.COMException ex) {
                Outlook.Errors.ErrorType error = Outlook.Errors.HandleComError(ex, out String hResult);

                if (error == Outlook.Errors.ErrorType.RpcRejected ||
                    error == Outlook.Errors.ErrorType.PermissionFailure ||
                    error == Outlook.Errors.ErrorType.RpcServerUnavailable ||
                    error == Outlook.Errors.ErrorType.RpcFailed) //
                {
                    log.Warn(ex.Message);
                    throw new ApplicationException("Outlook is busy.", ex);

                } else if (ex.GetErrorCode(0x000FFFFF) == "0x00040115") {
                    log.Warn(ex.Message);
                    log.Debug("OGCS is not able to run as Outlook is not properly connected to the Exchange server?");
                    throw new ApplicationException("Outlook is busy.", ex);

                } else if (ex.GetErrorCode(0x000FFFFF) == "0x000702E4") {
                    log.Warn(ex.Message);
                    throw new ApplicationException("Outlook and OGCS are running in different security elevations.\n" +
                        "Both must be running in Standard or Administrator mode.");

                } else if (!comErrorInWiki(ex)) {
                    ex.Analyse("COM error not in wiki.");
                    if (!alreadyRedirectedToWikiForComError.Contains(hResult)) {
                        Helper.OpenBrowser("https://github.com/phw198/OutlookGoogleCalendarSync/wiki/FAQs---COM-Errors");
                        alreadyRedirectedToWikiForComError.Add(hResult);
                    }
                    throw new ApplicationException("COM error " + hResult + " encountered.\r\n" +
                        "Please check if there is a published solution on the OGCS wiki.");
                }

            } catch (System.InvalidCastException ex) {
                if (!comErrorInWiki(ex)) throw;

            } catch (System.UnauthorizedAccessException ex) {
                if (ex.GetErrorCode() == "0x80070005") { // E_ACCESSDENIED
                    log.Warn(ex.Message);
                    throw new ApplicationException("OGCS was not permitted to start Outlook.\r\n" +
                        "Please manually start Outlook and then restart OGCS again.");
                }

            } catch (System.Exception ex) {
                ex.Analyse("Early binding to Outlook appears to have failed.", true);
                log.Debug("Could try late binding??");
                //System.Type oAppType = System.Type.GetTypeFromProgID("Outlook.Application");
                //ApplicationClass oAppClass = System.Activator.CreateInstance(oAppType) as ApplicationClass;
                //oApp = oAppClass.CreateObject("Outlook.Application") as Microsoft.Office.Interop.Outlook.Application;
                throw;
            }
        }

        private static Boolean comErrorInWiki(System.Exception ex) {
            String hResult = ex.GetErrorCode();
            String wikiUrl = "";
            Regex rgx;

            new Telemetry.GA4Event.Event(Telemetry.GA4Event.Event.Name.ogcs_error)
                .AddParameter("com_object", hResult)
                .AddParameter(GA4.General.sync_count, Settings.Instance.CompletedSyncs)
                .Send();

            if (hResult == "0x80040154") {
                String regkey = @"SOFTWARE\Classes\Local Settings\Software\Microsoft\Windows\CurrentVersion\AppModel\Repository\Packages";
                try {
                    Microsoft.Win32.RegistryKey openedKey = Microsoft.Win32.Registry.CurrentUser.OpenSubKey(regkey, false);
                    if (openedKey != null) {
                        String[] subkeys = openedKey.GetSubKeyNames();
                        if (subkeys.Where(k => k.StartsWith("Microsoft.OutlookForWindows_")).Count() > 0) {
                            Helper.OpenBrowser("https://github.com/phw198/OutlookGoogleCalendarSync/discussions/1888");
                            throw new ApplicationException("This version of OGCS requires the classic Outlook client to be installed.\r\n\r\n" +
                                "The next major release of OGCS does not require an Outlook client - further details have opened in your browser.");
                        } else
                            log.Debug("Found " + subkeys.Count() + " subkeys, but none started with 'Microsoft.OutlookForWindows_'");
                    } else {
                        log.Warn("Could not open registry key: " + regkey);
                    }
                } catch (System.ApplicationException) {
                    throw;
                } catch (System.Exception reg) {
                    reg.Analyse("Unable to check if New Outlook is installed.");
                }
            }

            if (hResult == "0x80004002" && (ex is System.InvalidCastException || ex is System.Runtime.InteropServices.COMException)) {
                log.Warn(ex.Message);
                log.Debug("Extracting specific COM error code from Exception error message.");
                try {
                    rgx = new Regex(@"HRESULT\s*: (0x[\dA-F]{8})", RegexOptions.IgnoreCase);
                    MatchCollection matches = rgx.Matches(ex.Message);
                    if (matches.Count == 0) {
                        log.Error("Could not regex HRESULT out of the error message");
                        hResult = "";
                    } else {
                        hResult = matches[0].Groups[1].Value;
                    }
                } catch (System.Exception ex2) {
                    ex2.Analyse("Parsing error message with regex failed.");
                }
            }

            if (!string.IsNullOrEmpty(hResult)) {
                try {
                    String html = "";
                    try {
                        html = new Extensions.OgcsWebClient().DownloadString("https://github.com/phw198/OutlookGoogleCalendarSync/wiki/FAQs---COM-Errors");
                    } catch (System.Exception) {
                        log.Fail("Could not download wiki HTML.");
                        throw;
                    }
                    if (!string.IsNullOrEmpty(html)) {
                        html = html.Replace("\n", "");
                        rgx = new Regex(@"<h2.*?><a.*?href=\""(#" + hResult + ".*?)\"", RegexOptions.IgnoreCase);
                        MatchCollection sourceAnchors = rgx.Matches(html);
                        if (sourceAnchors.Count == 0) {
                            log.Debug("Could not find the COM error " + hResult + " in the wiki.");
                        } else {
                            wikiUrl = "https://github.com/phw198/OutlookGoogleCalendarSync/wiki/FAQs---COM-Errors" + sourceAnchors[0].Groups[1].Value;
                        }
                    }

                } catch (System.Exception htmlEx) {
                    htmlEx.Analyse("Could not parse Wiki for existance of COM error.");
                }
            }

            if (string.IsNullOrEmpty(wikiUrl)) {
                log.Warn("Did not find COM error in Wiki, so now checking for hard-coded URLs.");
                if (ex.Message.Contains("0x80004002 (E_NOINTERFACE)")) {
                    wikiUrl = "https://github.com/phw198/OutlookGoogleCalendarSync/wiki/FAQs---COM-Errors#0x80004002---e_nointerface";

                } else if (ex.Message.Contains("0x8002801D (TYPE_E_LIBNOTREGISTERED)")) {
                    wikiUrl = "https://github.com/phw198/OutlookGoogleCalendarSync/wiki/FAQs---COM-Errors#0x8002801d---type_e_libnotregistered";

                } else if (ex.Message.Contains("0x80029C4A (TYPE_E_CANTLOADLIBRARY)")) {
                    wikiUrl = "https://github.com/phw198/OutlookGoogleCalendarSync/wiki/FAQs---COM-Errors#0x80029c4a---type__e__cantloadlibrary";

                } else if (ex.Message.Contains("0x800401F3 (CO_E_CLASSSTRING)")) {
                    wikiUrl = "https://github.com/phw198/OutlookGoogleCalendarSync/wiki/FAQs---COM-Errors#0x800401f3---co_e_classstring";

                } else if (ex.Message.Contains("0x80040154 (REGDB_E_CLASSNOTREG)")) {
                    wikiUrl = "https://github.com/phw198/OutlookGoogleCalendarSync/wiki/FAQs---COM-Errors#0x80040154---regdb_e_classnotreg";

                } else if (ex.Message.Contains("0x80040155")) {
                    wikiUrl = "https://github.com/phw198/OutlookGoogleCalendarSync/wiki/FAQs---COM-Errors#0x80040155---interface-not-registered";
                }
            }

            if (string.IsNullOrEmpty(wikiUrl)) return false;

            log.Warn(ex.Message);
            if (!alreadyRedirectedToWikiForComError.Contains(hResult)) {
                Helper.OpenBrowser(wikiUrl);
                alreadyRedirectedToWikiForComError.Add(hResult);
            }
            throw new ApplicationException("A problem was encountered with your Office install.\r\n" +
                "Please see the wiki for a solution. [" + hResult + "]");
        }

        public static string signature(AppointmentItem ai) {
            return (ai.Subject + ";" + ((DateTimeOffset)ai.Start).ToPreciseString() + ";" + ((DateTimeOffset)ai.End).ToPreciseString()).Trim();
        }

        public static void ExportToCSV(String action, String filename, List<AppointmentItem> ais) {
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
                    CSVheader += "Outlook GlobalID,Outlook EntryID,Outlook CalendarID,";
                    CSVheader += "Google EventID,Google CalendarID,OGCS Modified,Force Save,Copied";
                    tw.WriteLine(CSVheader);
                    foreach (AppointmentItem ai in ais) {
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
        private static string exportToCSV(AppointmentItem ai) {
            StringBuilder csv = new StringBuilder();

            csv.Append(((DateTimeOffset)ai.Start).ToPreciseString() + ",");
            csv.Append(((DateTimeOffset)ai.End).ToPreciseString() + ",");
            csv.Append("\"" + ai.Subject + "\",");

            if (ai.Location == null) csv.Append(",");
            else csv.Append("\"" + ai.Location + "\",");

            if (ai.Body == null) csv.Append(",");
            else {
                String csvBody = ai.Body.Replace("\"", "");
                csvBody = csvBody.Replace("\r\n", " ");
                csv.Append("\"" + csvBody.Substring(0, System.Math.Min(csvBody.Length, 100)) + "\",");
            }

            csv.Append("\"" + ai.Sensitivity.ToString() + "\",");
            csv.Append("\"" + ai.BusyStatus.ToString() + "\",");
            csv.Append("\"" + (ai.RequiredAttendees == null ? "" : ai.RequiredAttendees) + "\",");
            csv.Append("\"" + (ai.OptionalAttendees == null ? "" : ai.OptionalAttendees) + "\",");
            csv.Append(ai.ReminderSet + ",");
            csv.Append(ai.ReminderMinutesBeforeStart.ToString() + ",");
            csv.Append(Outlook.Calendar.Instance.IOutlook.GetGlobalApptID(ai) + ",");
            csv.Append(ai.EntryID + "," + Sync.Engine.Calendar.Instance.Profile.UseOutlookCalendar.Id + ",");
            csv.Append((CustomProperty.Get(ai, CustomProperty.MetadataId.gEventID) ?? "") + ",");
            csv.Append((CustomProperty.Get(ai, CustomProperty.MetadataId.gCalendarId) ?? "") + ",");
            csv.Append(CustomProperty.GetOGCSlastModified(ai).ToString() + ",");
            csv.Append((CustomProperty.Get(ai, CustomProperty.MetadataId.forceSave) ?? "") + ",");
            csv.Append(CustomProperty.Get(ai, CustomProperty.MetadataId.locallyCopied) ?? "");

            return csv.ToString();
        }

        /// <summary>
        /// Get the anonymised summary of an appointment item, else standard summary.
        /// </summary>
        /// <param name="ai">The appointment item.</param>
        /// <returns>The summary, anonymised if settings dictate.</returns>
        public static String GetEventSummary(AppointmentItem ai) {
            String eventSummary = GetEventSummary(ai, out String anonymisedSummary, false);
            return anonymisedSummary ?? eventSummary;
        }

        /// <summary>
        /// Pre/Append context to the summary of an appointment item.
        /// </summary>
        /// <param name="context">Text to add before/after the summary and anonymised summary.</param>
        /// <param name="ai">The appointment item.</param>
        /// <param name="eventSummaryAnonymised">The anonymised summary with context.</param>
        /// <param name="onlyIfNotVerbose">Only return if user doesn't have Verbose output on. Useful for indicating offending item during errors.</param>
        /// <param name="appendContext">If the context should be before or after.</param>
        /// <returns>The standard summary.</returns>
        public static string GetEventSummary(String context, AppointmentItem ai, out String eventSummaryAnonymised, Boolean onlyIfNotVerbose = false, Boolean appendContext = true) {
            String eventSummary = GetEventSummary(ai, out String anonymisedSummary, onlyIfNotVerbose);
            if (appendContext) {
                eventSummary = eventSummary + context;
                eventSummaryAnonymised = string.IsNullOrEmpty(anonymisedSummary) ? null : (anonymisedSummary + context);
            } else {
                eventSummary = context + eventSummary;
                eventSummaryAnonymised = string.IsNullOrEmpty(anonymisedSummary) ? null : (context + anonymisedSummary);
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
        public static string GetEventSummary(AppointmentItem ai, out String eventSummaryAnonymised, Boolean onlyIfNotVerbose = false) {
            String eventSummary = "";
            eventSummaryAnonymised = null;
            if (!onlyIfNotVerbose || onlyIfNotVerbose && !Settings.Instance.VerboseOutput) {
                try {
                    if (ai.AllDayEvent) {
                        log.Fine("GetSummary - all day event");
                        eventSummary += ai.Start.Date.ToShortDateString();
                    } else {
                        log.Fine("GetSummary - not all day event");
                        eventSummary += ai.Start.ToShortDateString() + " " + ai.Start.ToShortTimeString();
                    }
                    if (ai.IsRecurring) {
                        if (ai.RecurrenceState == OlRecurrenceState.olApptException) eventSummary += " (R1)";
                        else eventSummary += " (R)";
                    }
                    eventSummary += " => ";

                    if (Settings.Instance.AnonymiseLogs)
                        eventSummaryAnonymised = eventSummary + '"' + Ogcs.Google.Authenticator.GetMd5(ai.Subject, silent: true) + '"' + (onlyIfNotVerbose ? "<br/>" : "");
                    eventSummary += '"' + ai.Subject + '"' + (onlyIfNotVerbose ? "<br/>" : "");

                } catch (System.Runtime.InteropServices.COMException ex) {
                    if (ex.GetErrorCode() == "0x8004010F")
                        throw new System.Exception("Cannot access Outlook OST/PST file. Try restarting Outlook.", ex);
                    else
                        ex.Analyse("Failed to get appointment summary: " + eventSummary, true);

                } catch (System.Exception ex) {
                    ex.Analyse("Failed to get appointment summary: " + eventSummary, true);
                }
            }
            return eventSummary;
        }

        public static void IdentifyEventDifferences(
            ref List<Event> google,             //need creating
            ref List<AppointmentItem> outlook,  //need deleting
            ref Dictionary<AppointmentItem, Event> compare) //
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

                        if (compare_oEventID == google[g].Id.ToString()) {
                            if (googleIDmissing == null) googleIDmissing = CustomProperty.GoogleIdMissing(outlook[o]);
                            if ((Boolean)googleIDmissing) {
                                log.Info("Enhancing appointment's metadata...");
                                AppointmentItem ai = outlook[o];
                                CustomProperty.AddGoogleIDs(ref ai, google[g]);
                                CustomProperty.Add(ref ai, CustomProperty.MetadataId.forceSave, true.ToString());
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
            }
            if (metadataEnhanced > 0) log.Info(metadataEnhanced + " item's metadata enhanced.");

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

            if (profile.SyncDirection.Id == Sync.Direction.Bidirectional.Id) {
                //Don't recreate any items that have been deleted in Outlook
                for (int g = google.Count - 1; g >= 0; g--) {
                    if (Ogcs.Google.CustomProperty.Exists(google[g], Ogcs.Google.CustomProperty.MetadataId.oEntryId))
                        google.Remove(google[g]);
                }
                //Don't delete any items that aren't yet in Google or just created in Google during this sync
                for (int o = outlook.Count - 1; o >= 0; o--) {
                    if (!CustomProperty.Exists(outlook[o], CustomProperty.MetadataId.gEventID) ||
                        outlook[o].LastModificationTime > Sync.Engine.Instance.SyncStarted)
                        outlook.Remove(outlook[o]);
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
            if (Settings.Instance.CreateCSVFiles) {
                ExportToCSV("Appointments for deletion in Outlook", "outlook_delete.csv", outlook);
                Ogcs.Google.Calendar.ExportToCSV("Events for creation in Outlook", "outlook_create.csv", google);
            }
        }

        public static Boolean ItemIDsMatch(AppointmentItem ai, Event ev) {
            //For format of Entry ID : https://msdn.microsoft.com/en-us/library/ee201952(v=exchg.80).aspx
            //For format of Global ID: https://msdn.microsoft.com/en-us/library/ee157690%28v=exchg.80%29.aspx

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

        public static object ReleaseObject(object obj) {
            try {
                if (obj != null && System.Runtime.InteropServices.Marshal.IsComObject(obj)) {
                    while (System.Runtime.InteropServices.Marshal.ReleaseComObject(obj) > 0)
                        System.Windows.Forms.Application.DoEvents();
                }
            } catch (System.Exception ex) {
                Ogcs.Exception.Analyse(ex, true);
            }
            GC.Collect();
            return null;
        }

        public Boolean IsOKtoSyncReminder(AppointmentItem ai) {
            SettingsStore.Calendar profile = Sync.Engine.Calendar.Instance.Profile;

            if (profile.ReminderDND) {
                System.DateTime alarm;
                if (ai.ReminderSet)
                    alarm = ai.Start.AddMinutes(-ai.ReminderMinutesBeforeStart);
                else {
                    if (profile.UseGoogleDefaultReminder && Ogcs.Google.Calendar.Instance.MinDefaultReminder != int.MinValue) {
                        log.Fine("Using default Google reminder value: " + Ogcs.Google.Calendar.Instance.MinDefaultReminder);
                        alarm = ai.Start.AddMinutes(-Ogcs.Google.Calendar.Instance.MinDefaultReminder);
                    } else
                        return false;
                }
                return isOKtoSyncReminder(alarm);
            }
            return true;
        }
        private Boolean isOKtoSyncReminder(System.DateTime alarm) {
            SettingsStore.Calendar profile = Sync.Engine.Calendar.Instance.Profile;

            if (profile.ReminderDNDstart.TimeOfDay > profile.ReminderDNDend.TimeOfDay) {
                //eg 22:00 to 06:00
                //Make sure end time is the day following the start time
                int shiftDay = 0;
                System.DateTime dndStart;
                System.DateTime dndEnd;
                if (alarm.TimeOfDay < profile.ReminderDNDstart.TimeOfDay) shiftDay = -1;
                dndStart = alarm.Date.AddDays(shiftDay).Add(profile.ReminderDNDstart.TimeOfDay);
                dndEnd = alarm.Date.AddDays(shiftDay + 1).Add(profile.ReminderDNDend.TimeOfDay);
                if (alarm > dndStart && alarm < dndEnd) {
                    log.Debug("Reminder (@" + alarm.ToString("HH:mm") + ") falls in DND range - not synced.");
                    return false;
                } else
                    return true;

            } else {
                //eg 01:00 to 06:00
                if (alarm.TimeOfDay < profile.ReminderDNDstart.TimeOfDay ||
                    alarm.TimeOfDay > profile.ReminderDNDend.TimeOfDay) {
                    return true;
                } else {
                    log.Debug("Reminder (@" + alarm.ToString("HH:mm") + ") falls in DND range - not synced.");
                    return false;
                }
            }
        }
        #endregion
    }
}
