using Google.Apis.Calendar.v3.Data;
using log4net;
using Microsoft.Office.Interop.Outlook;
using System;
using System.Collections.Generic;
using System.IO;
using System.Windows.Forms;

namespace OutlookGoogleCalendarSync {
    /// <summary>
    /// Description of OutlookCalendar.
    /// </summary>
    public class OutlookCalendar {
        private static OutlookCalendar instance;
        private static readonly ILog log = LogManager.GetLogger(typeof(OutlookCalendar));
        public OutlookInterface IOutlook;
        
        public static OutlookCalendar Instance {
            get {
                try {
                    if (instance == null || instance.Accounts == null) instance = new OutlookCalendar();
                } catch {
                    log.Info("It appears Outlook has been restarted after OGCS was started. Reconnecting...");
                    instance = new OutlookCalendar();
                }
                return instance;
            }
        }
        private String currentUserSMTP {
            get { return IOutlook.CurrentUserSMTP(); }
        }
        public String CurrentUserName {
            get { return IOutlook.CurrentUserName(); }
        }
        public MAPIFolder UseOutlookCalendar {
            get { return IOutlook.UseOutlookCalendar(); }
            set {
                IOutlook.UseOutlookCalendar(value);
                Settings.Instance.UseOutlookCalendar = new MyOutlookCalendarListEntry(value);
            }
        }
        public List<String> Accounts {
            get { return IOutlook.Accounts(); }
        }
        public Dictionary<string, MAPIFolder> CalendarFolders {
            get { return IOutlook.CalendarFolders(); }
        }
        public enum Service {
            DefaultMailbox,
            AlternativeMailbox,
            EWS
        }
        public const String gEventID = "googleEventID";

        public OutlookCalendar() {
            IOutlook = OutlookFactory.getOutlookInterface();
            IOutlook.Connect();
        }

        public void Reset() {
            instance = new OutlookCalendar();
        }
        
        #region Push Sync
        public void RegisterForAutoSync() {
            log.Info("Registering for Outlook appointment change events...");
            try {
                UseOutlookCalendar.Items.ItemAdd -= new ItemsEvents_ItemAddEventHandler(appointmentItem_Add);
                UseOutlookCalendar.Items.ItemChange -= new ItemsEvents_ItemChangeEventHandler(appointmentItem_Change);
                UseOutlookCalendar.Items.ItemRemove -= new ItemsEvents_ItemRemoveEventHandler(appointmentItem_Remove);
            } catch { }
            if (Settings.Instance.SyncDirection != SyncDirection.GoogleToOutlook) {
                UseOutlookCalendar.Items.ItemAdd += new ItemsEvents_ItemAddEventHandler(appointmentItem_Add);
                UseOutlookCalendar.Items.ItemChange += new ItemsEvents_ItemChangeEventHandler(appointmentItem_Change);
                UseOutlookCalendar.Items.ItemRemove += new ItemsEvents_ItemRemoveEventHandler(appointmentItem_Remove);

                log.Debug("Create the timer for the push synchronisation");
                MainForm.Instance.OgcsPushTimer = new Timer();
                MainForm.Instance.OgcsPushTimer.Tick += new EventHandler(MainForm.Instance.OgcsPushTimer_Tick);
                if (!MainForm.Instance.OgcsPushTimer.Enabled) {
                    MainForm.Instance.OgcsPushTimer.Interval = 2 * 60000;
                    MainForm.Instance.OgcsPushTimer.Tag = "PushTimer";
                    MainForm.Instance.OgcsPushTimer.Start();
                }
            }
        }

        public void DeregisterForAutoSync() {
            log.Info("Deregistering from Outlook appointment change events...");
            try {
                UseOutlookCalendar.Items.ItemAdd -= new ItemsEvents_ItemAddEventHandler(appointmentItem_Add);
                UseOutlookCalendar.Items.ItemChange -= new ItemsEvents_ItemChangeEventHandler(appointmentItem_Change);
                UseOutlookCalendar.Items.ItemRemove -= new ItemsEvents_ItemRemoveEventHandler(appointmentItem_Remove);
            } catch {
                log.Debug("No event handlers set.");
            } 
            if (MainForm.Instance.OgcsPushTimer != null && MainForm.Instance.OgcsPushTimer.Enabled) 
            MainForm.Instance.OgcsPushTimer.Stop();
        }

        private void appointmentItem_Add(object Item) {
            if (Settings.Instance.SyncDirection == SyncDirection.GoogleToOutlook) return;

            log.Debug("Detected Outlook item added.");
            AppointmentItem ai = Item as AppointmentItem;
            
            DateTime syncMin = DateTime.Today.AddDays(-Settings.Instance.DaysInThePast);
            DateTime syncMax = DateTime.Today.AddDays(+Settings.Instance.DaysInTheFuture + 1);
            if (ai.Start >= syncMin && ai.End <= syncMax) {
                log.Debug(GetEventSummary(ai));
                log.Debug("Item is in sync range, so push sync flagged for Go.");
                int pushFlag = Convert.ToInt16(MainForm.Instance.GetControlPropertyThreadSafe(MainForm.Instance.bSyncNow, "Tag"));
                pushFlag++;
                log.Info(pushFlag + " items changed since last sync.");
                MainForm.Instance.SetControlPropertyThreadSafe(MainForm.Instance.bSyncNow, "Tag", pushFlag);
            }
        }
        private void appointmentItem_Change(object Item) {
            if (Settings.Instance.SyncDirection == SyncDirection.GoogleToOutlook) return;

            log.Debug("Detected Outlook item changed.");
            AppointmentItem ai = Item as AppointmentItem;
            
            DateTime syncMin = DateTime.Today.AddDays(-Settings.Instance.DaysInThePast);
            DateTime syncMax = DateTime.Today.AddDays(+Settings.Instance.DaysInTheFuture + 1);
            if (ai.Start >= syncMin && ai.End <= syncMax) {
                log.Debug(GetEventSummary(ai));
                log.Debug("Item is in sync range, so push sync flagged for Go.");
                int pushFlag = Convert.ToInt16(MainForm.Instance.GetControlPropertyThreadSafe(MainForm.Instance.bSyncNow, "Tag"));
                pushFlag++;
                log.Info(pushFlag + " items changed since last sync.");
                MainForm.Instance.SetControlPropertyThreadSafe(MainForm.Instance.bSyncNow, "Tag", pushFlag);
            }
        }
        private void appointmentItem_Remove() {
            if (Settings.Instance.SyncDirection == SyncDirection.GoogleToOutlook) return;

            log.Debug("Detected Outlook item removed, so push sync flagged for Go.");
            int pushFlag = Convert.ToInt16(MainForm.Instance.GetControlPropertyThreadSafe(MainForm.Instance.bSyncNow, "Tag"));
            pushFlag++;
            log.Info(pushFlag + " items changed since last sync.");
            MainForm.Instance.SetControlPropertyThreadSafe(MainForm.Instance.bSyncNow, "Tag", pushFlag);
        }
        #endregion

        public List<AppointmentItem> getCalendarEntriesInRange() {
            List<AppointmentItem> filtered = new List<AppointmentItem>();
            filtered = filterCalendarEntries(UseOutlookCalendar.Items);

            if (Settings.Instance.CreateCSVFiles) {
                log.Debug("Outputting CSV files...");
                TextWriter tw = new StreamWriter(Path.Combine(Program.UserFilePath,"outlook_appointments.csv"));
                String CSVheader = "Start Time,Finish Time,Subject,Location,Description,Privacy,FreeBusy,";
                CSVheader += "Required Attendees,Optional Attendees,Reminder Set,Reminder Minutes,Outlook ID,Google ID";
                tw.WriteLine(CSVheader);
                foreach (AppointmentItem ai in filtered) {
                    try {
                        tw.WriteLine(exportToCSV(ai));
                    } catch (System.Exception ex) {
                        MainForm.Instance.Logboxout("Failed to output following Outlook appointment to CSV:-");
                        MainForm.Instance.Logboxout(GetEventSummary(ai));
                        log.Error(ex.Message);
                    }
                }
                tw.Close();
                log.Debug("Done.");
            }
            return filtered;
        }

        public List<AppointmentItem> filterCalendarEntries(Items OutlookItems) {
            List<AppointmentItem> result = new List<AppointmentItem>();
            log.Fine(OutlookItems.Count + " calendar items exist.");

            //OutlookItems.Sort("[Start]", Type.Missing);
            OutlookItems.IncludeRecurrences = false;

            if (OutlookItems != null) {
                DateTime min = Settings.Instance.SyncStart;
                DateTime max = Settings.Instance.SyncEnd;
                
                string filter = "[End] >= '" + min.ToString("g") + "' AND [Start] < '" + max.ToString("g") + "'";
                log.Fine("Filter string: " + filter);
                foreach (AppointmentItem ai in OutlookItems.Restrict(filter)) {
                    if (ai.End == min) continue; //Required for midnight to midnight events 
                    result.Add(ai);
                }
                log.Fine("Filtered down to " + result.Count);
                result = new List<AppointmentItem>();
                
                //Outlook can't handle dates or times formatted with a . delimeter!
                string format = System.Globalization.CultureInfo.CurrentCulture.DateTimeFormat.ShortDatePattern;
                switch (format) {
                    case "yyyy.MMdd": format = "yyyy-MM-dd"; break;
                    default: break;
                }
                format += " " + System.Globalization.CultureInfo.CurrentCulture.DateTimeFormat.ShortTimePattern.Replace(".", ":");
                filter = "[End] >= '" + min.ToString(format) + "' AND [Start] < '" + max.ToString(format) + "'";
                log.Fine("Filter string: " + filter);
                foreach (AppointmentItem ai in OutlookItems.Restrict(filter)) {
                    if (ai.End == min) continue; //Required for midnight to midnight events 
                    result.Add(ai);
                }
            }
            log.Fine("Filtered down to "+ result.Count);
            return result;
        }

        #region Create
        public void CreateCalendarEntries(List<Event> events) {
            foreach (Event ev in events) {
                AppointmentItem newAi = IOutlook.UseOutlookCalendar().Items.Add() as AppointmentItem;
                try {
                    newAi = createCalendarEntry(ev);
                } catch (System.Exception ex) {
                    if (!Settings.Instance.VerboseOutput) MainForm.Instance.Logboxout(GoogleCalendar.GetEventSummary(ev));
                    MainForm.Instance.Logboxout("WARNING: Appointment creation failed.\r\n" + ex.Message);
                    log.Error(ex.StackTrace);
                    if (MessageBox.Show("Outlook appointment creation failed. Continue with synchronisation?", "Sync item failed", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                        continue;
                    else {
                        newAi = (AppointmentItem)ReleaseObject(newAi);
                        throw new UserCancelledSyncException("User chose not to continue sync.");
                    }
                }

                try {
                    createCalendarEntry_save(newAi, ev);
                } catch (System.Exception ex) {
                    MainForm.Instance.Logboxout("WARNING: New appointment failed to save.\r\n" + ex.Message);
                    log.Error(ex.StackTrace);
                    if (MessageBox.Show("New Outlook appointment failed to save. Continue with synchronisation?", "Sync item failed", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                        continue;
                    else {
                        newAi = (AppointmentItem)ReleaseObject(newAi);
                        throw new UserCancelledSyncException("User chose not to continue sync.");
                    }
                }

                if (ev.Recurrence != null && ev.RecurringEventId == null && Recurrence.Instance.HasExceptions(ev)) {
                    MainForm.Instance.Logboxout("This is a recurring item with some exceptions:-");
                    Recurrence.Instance.CreateOutlookExceptions(newAi, ev);
                    MainForm.Instance.Logboxout("Recurring exceptions completed.");
                }
                newAi = (AppointmentItem)ReleaseObject(newAi);
            }
        }
        
        private AppointmentItem createCalendarEntry(Event ev) {
            string itemSummary = GoogleCalendar.GetEventSummary(ev);
            log.Debug("Processing >> " + itemSummary);
            MainForm.Instance.Logboxout(itemSummary, verbose: true);

            AppointmentItem ai = IOutlook.UseOutlookCalendar().Items.Add() as AppointmentItem;

            //Add the Google event ID into Outlook appointment.
            AddOGCSproperty(ref ai, gEventID, ev.Id);

            ai.Start = new DateTime();
            ai.End = new DateTime();
            ai.AllDayEvent = (ev.Start.Date != null);
            ai = OutlookCalendar.Instance.IOutlook.WindowsTimeZone_set(ai, ev);
            Recurrence.Instance.BuildOutlookPattern(ev, ai);
            
            ai.Subject = Obfuscate.ApplyRegex(ev.Summary, SyncDirection.GoogleToOutlook);
            if (Settings.Instance.AddDescription && ev.Description != null) ai.Body = ev.Description;
            ai.Location = ev.Location;
            ai.Sensitivity = (ev.Visibility == "private") ? OlSensitivity.olPrivate : OlSensitivity.olNormal;
            ai.BusyStatus = (ev.Transparency == "transparent") ? OlBusyStatus.olFree : OlBusyStatus.olBusy;

            if (Settings.Instance.AddAttendees && ev.Attendees != null) {
                foreach (EventAttendee ea in ev.Attendees) {
                    createRecipient(ea, ai);
                }
            }

            //Reminder alert
            if (Settings.Instance.AddReminders && ev.Reminders != null && ev.Reminders.Overrides != null) {
                foreach (EventReminder reminder in ev.Reminders.Overrides) {
                    if (reminder.Method == "popup") {
                        ai.ReminderSet = true;
                        ai.ReminderMinutesBeforeStart = (int)reminder.Minutes;
                    }
                }
            }
            return ai;
        }

        private static void createCalendarEntry_save(AppointmentItem ai, Event ev) {
            if (Settings.Instance.SyncDirection == SyncDirection.Bidirectional) {
                log.Debug("Saving timestamp when OGCS updated appointment.");
                AddOGCSproperty(ref ai, Program.OGCSmodified, DateTime.Now);
            }
            
            ai.Save();

            Boolean oKeyExists = false;
            try {
                oKeyExists = ev.ExtendedProperties.Private.ContainsKey(GoogleCalendar.oEntryID);
            } catch {}
            if (Settings.Instance.SyncDirection == SyncDirection.Bidirectional || oKeyExists) {
                log.Debug("Storing the Outlook appointment ID in Google event.");
                GoogleCalendar.AddOutlookID(ref ev, ai);
                GoogleCalendar.Instance.UpdateCalendarEntry_save(ev);
            }
        }
        #endregion

        #region Update
        public void UpdateCalendarEntries(Dictionary<AppointmentItem, Event> entriesToBeCompared, ref int entriesUpdated) {
            entriesUpdated = 0;
            foreach (KeyValuePair<AppointmentItem, Event> compare in entriesToBeCompared) {
                int itemModified = 0;
                AppointmentItem ai = IOutlook.UseOutlookCalendar().Items.Add() as AppointmentItem;
                Boolean aiWasRecurring = compare.Key.IsRecurring;
                try {
                    ai = UpdateCalendarEntry(compare.Key, compare.Value, ref itemModified);
                } catch (System.Exception ex) {
                    if (!Settings.Instance.VerboseOutput) MainForm.Instance.Logboxout(GoogleCalendar.GetEventSummary(compare.Value));
                    MainForm.Instance.Logboxout("WARNING: Appointment update failed.\r\n" + ex.Message);
                    log.Error(ex.StackTrace);
                    if (MessageBox.Show("Outlook appointment update failed. Continue with synchronisation?", "Sync item failed", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                        continue;
                    else {
                        ai = (AppointmentItem)ReleaseObject(ai);
                        throw new UserCancelledSyncException("User chose not to continue sync.");
                    }
                }

                if (itemModified > 0) {
                    try {
                        updateCalendarEntry_save(ai);
                        entriesUpdated++;
                    } catch (System.Exception ex) {
                        MainForm.Instance.Logboxout("WARNING: Updated appointment failed to save.\r\n" + ex.Message);
                        log.Error(ex.StackTrace);
                        if (MessageBox.Show("Updated Outlook appointment failed to save. Continue with synchronisation?", "Sync item failed", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                            continue;
                        else {
                            ai = (AppointmentItem)ReleaseObject(ai); 
                            throw new UserCancelledSyncException("User chose not to continue sync.");
                        }
                    }
                    if (!aiWasRecurring && ai.IsRecurring) {
                        log.Debug("Appointment has changed from single instance to recurring, so exceptions may need processing.");
                        Recurrence.Instance.UpdateOutlookExceptions(ai, compare.Value);
                    }
                } else if (ai != null && ai.RecurrenceState != OlRecurrenceState.olApptMaster) { //Master events are always compared anyway
                    log.Debug("Doing a dummy update in order to update the last modified date.");
                    AddOGCSproperty(ref ai, Program.OGCSmodified, DateTime.Now);
                    updateCalendarEntry_save(ai);
                }
                ai = (AppointmentItem)ReleaseObject(ai);
            }
        }

        public AppointmentItem UpdateCalendarEntry(AppointmentItem ai, Event ev, ref int itemModified, Boolean forceCompare = false) {
            if (ai.RecurrenceState == OlRecurrenceState.olApptMaster) { //The exception child objects might have changed
                log.Debug("Processing recurring master appointment.");
            } else {
                if (!forceCompare) { //Needed if the exception has just been created, but now needs updating
                    if (Settings.Instance.SyncDirection != SyncDirection.Bidirectional) {
                        if (DateTime.Parse(GoogleCalendar.GoogleTimeFrom(ai.LastModificationTime)) > DateTime.Parse(ev.Updated))
                            return null;
                    } else {
                        if (GoogleCalendar.OGCSlastModified(ev).AddSeconds(5) >= DateTime.Parse(ev.Updated))
                            //Google last modified by OGCS
                            return null;
                        if (DateTime.Parse(GoogleCalendar.GoogleTimeFrom(ai.LastModificationTime)) > DateTime.Parse(ev.Updated))
                            return null;
                    }
                }
            }

            String evSummary = GoogleCalendar.GetEventSummary(ev);
            log.Debug("Processing >> " + evSummary);

            System.Text.StringBuilder sb = new System.Text.StringBuilder();
            sb.AppendLine(evSummary);

            if (ev.Start.Date != null) {
                RecurrencePattern oPattern = ai.GetRecurrencePattern();
                if (ai.RecurrenceState != OlRecurrenceState.olApptMaster) ai.AllDayEvent = true;
                if (MainForm.CompareAttribute("Start time", SyncDirection.GoogleToOutlook, ev.Start.Date, ai.Start.ToString("yyyy-MM-dd"), sb, ref itemModified)) {
                    if (ai.RecurrenceState == OlRecurrenceState.olApptMaster) oPattern.PatternStartDate = DateTime.Parse(ev.Start.Date);
                    else ai.Start = DateTime.Parse(ev.Start.Date);
                }
                if (MainForm.CompareAttribute("End time", SyncDirection.GoogleToOutlook, ev.End.Date, ai.End.ToString("yyyy-MM-dd"), sb, ref itemModified)) {
                    if (ai.RecurrenceState == OlRecurrenceState.olApptMaster) oPattern.PatternEndDate = DateTime.Parse(ev.End.Date);
                    else ai.End = DateTime.Parse(ev.End.Date);
                }
                oPattern = (RecurrencePattern)ReleaseObject(oPattern);
            } else {
                RecurrencePattern oPattern = ai.GetRecurrencePattern();
                if (ai.RecurrenceState != OlRecurrenceState.olApptMaster) ai.AllDayEvent = false;
                if (MainForm.CompareAttribute("Start time",
                    SyncDirection.GoogleToOutlook,
                    GoogleCalendar.GoogleTimeFrom(DateTime.Parse(ev.Start.DateTime)),
                    GoogleCalendar.GoogleTimeFrom(ai.Start), sb, ref itemModified)) 
                {
                    if (ai.RecurrenceState == OlRecurrenceState.olApptMaster) oPattern.PatternStartDate = DateTime.Parse(ev.Start.DateTime);
                    else ai.Start = DateTime.Parse(ev.Start.DateTime);
                }
                if (MainForm.CompareAttribute("End time",
                    SyncDirection.GoogleToOutlook,
                    GoogleCalendar.GoogleTimeFrom(DateTime.Parse(ev.End.DateTime)),
                    GoogleCalendar.GoogleTimeFrom(ai.End), sb, ref itemModified)) 
                {
                    if (ai.RecurrenceState == OlRecurrenceState.olApptMaster) oPattern.PatternEndDate = DateTime.Parse(ev.End.DateTime);
                    else ai.End = DateTime.Parse(ev.End.DateTime);
                }
                oPattern = (RecurrencePattern)ReleaseObject(oPattern);
            }

            if (ai.RecurrenceState == OlRecurrenceState.olApptMaster) {
                if (ev.Recurrence == null || ev.RecurringEventId != null) {
                    log.Debug("Converting to non-recurring events.");
                    ai.ClearRecurrencePattern();
                    itemModified++;
                } else {
                    Recurrence.Instance.CompareOutlookPattern(ev, ai, sb, ref itemModified);
                    Recurrence.Instance.UpdateOutlookExceptions(ai, ev);
                }
            } else if (ai.RecurrenceState == OlRecurrenceState.olApptNotRecurring) {
                if (!ai.IsRecurring && ev.Recurrence != null && ev.RecurringEventId == null) {
                    log.Debug("Converting to recurring appointment.");
                    Recurrence.Instance.CreateOutlookExceptions(ai, ev);
                    itemModified++;
                }
            }

            String summaryObfuscated = Obfuscate.ApplyRegex(ev.Summary, SyncDirection.GoogleToOutlook);
            if (MainForm.CompareAttribute("Subject", SyncDirection.GoogleToOutlook, summaryObfuscated, ai.Subject, sb, ref itemModified)) {
                ai.Subject = summaryObfuscated;
            }
            if (!Settings.Instance.AddDescription) ev.Description = "";
            if (Settings.Instance.SyncDirection == SyncDirection.GoogleToOutlook || !Settings.Instance.AddDescription_OnlyToGoogle) {
                if (MainForm.CompareAttribute("Description", SyncDirection.GoogleToOutlook, ev.Description, ai.Body, sb, ref itemModified))
                    ai.Body = ev.Description;
            }

            if (MainForm.CompareAttribute("Location", SyncDirection.GoogleToOutlook, ev.Location, ai.Location, sb, ref itemModified))
                ai.Location = ev.Location;

            String oPrivacy = (ai.Sensitivity == OlSensitivity.olNormal) ? "default" : "private";
            String gPrivacy = (ev.Visibility == null ? "default" : ev.Visibility);
            if (MainForm.CompareAttribute("Private", SyncDirection.GoogleToOutlook, gPrivacy, oPrivacy, sb, ref itemModified)) {
                ai.Sensitivity = (ev.Visibility != null && ev.Visibility == "private") ? OlSensitivity.olPrivate : OlSensitivity.olNormal;
            }
            String oFreeBusy = (ai.BusyStatus == OlBusyStatus.olFree) ? "transparent" : "opaque";
            String gFreeBusy = (ev.Transparency == null ? "opaque" : ev.Transparency);
            if (MainForm.CompareAttribute("Free/Busy", SyncDirection.GoogleToOutlook, gFreeBusy, oFreeBusy, sb, ref itemModified)) {
                ai.BusyStatus = (ev.Transparency != null && ev.Transparency == "transparent") ? OlBusyStatus.olFree : OlBusyStatus.olBusy;
            }

            if (Settings.Instance.AddAttendees) {
                if (ev.Description != null && ev.Description.Contains("===--- Attendees ---===")) {
                    //Protect against <v1.2.4 where attendees were stored as text
                    log.Info("This event still has attendee information in the description - cannot sync them.");
                } else if (Settings.Instance.SyncDirection == SyncDirection.Bidirectional &&
                    ev.Attendees != null && ev.Attendees.Count == 0 && ai.Recipients.Count > 150) {
                        log.Info("Attendees not being synced - there are too many ("+ ai.Recipients.Count +") for Google.");
                } else {
                    //Build a list of Outlook attendees. Any remaining at the end of the diff must be deleted.
                    List<Recipient> removeRecipient = new List<Recipient>();
                    if (ai.Recipients != null) {
                        foreach (Recipient recipient in ai.Recipients) {
                            if (recipient.Name != ai.Organizer)
                                removeRecipient.Add(recipient);
                        }
                    }
                    if (ev.Attendees != null) {
                        for (int g = ev.Attendees.Count - 1; g >= 0; g--) {
                            bool foundRecipient = false;
                            EventAttendee attendee = ev.Attendees[g];
                            
                            foreach (Recipient recipient in ai.Recipients) {
                                if (!recipient.Resolved) recipient.Resolve();
                                String recipientSMTP = IOutlook.GetRecipientEmail(recipient);
                                if (recipientSMTP.ToLower() == attendee.Email.ToLower()) {
                                    foundRecipient = true;
                                    removeRecipient.Remove(recipient);

                                    //Optional attendee
                                    bool oOptional = (ai.OptionalAttendees != null && ai.OptionalAttendees.Contains(attendee.DisplayName ?? attendee.Email));
                                    bool gOptional = (attendee.Optional == null) ? false : (bool)attendee.Optional;
                                    if (MainForm.CompareAttribute("Recipient " + recipient.Name + " - Optional Check",
                                        SyncDirection.GoogleToOutlook, gOptional, oOptional, sb, ref itemModified)) {
                                        if (gOptional) {
                                            recipient.Type = (int)OlMeetingRecipientType.olOptional;
                                        } else {
                                            recipient.Type = (int)OlMeetingRecipientType.olRequired;
                                        }
                                    }
                                    //Response is readonly in Outlook :(
                                    break;
                                }
                            }
                            if (!foundRecipient &&
                                (attendee.DisplayName != ai.Organizer)) //Attendee in Google is owner in Outlook, so can't also be added as a recipient)
                                {
                                sb.AppendLine("Recipient added: " + (attendee.DisplayName ?? attendee.Email));
                                createRecipient(attendee, ai);
                                itemModified++;
                            }
                        }
                    }

                    foreach (Recipient recipient in removeRecipient) {
                        sb.AppendLine("Recipient removed: " + recipient.Name);
                        recipient.Delete();
                        itemModified++;
                    }
                }
            }
            //Reminders
            if (Settings.Instance.AddReminders) {
                if (ev.Reminders.Overrides != null) {
                    //Find the popup reminder in Google
                    for (int r = ev.Reminders.Overrides.Count - 1; r >= 0; r--) {
                        EventReminder reminder = ev.Reminders.Overrides[r];
                        if (reminder.Method == "popup") {
                            if (ai.ReminderSet) {
                                if (MainForm.CompareAttribute("Reminder", SyncDirection.GoogleToOutlook, reminder.Minutes.ToString(), ai.ReminderMinutesBeforeStart.ToString(), sb, ref itemModified)) {
                                    ai.ReminderMinutesBeforeStart = (int)reminder.Minutes;
                                }
                            } else {
                                sb.AppendLine("Reminder: nothing => " + reminder.Minutes);
                                ai.ReminderSet = true;
                                ai.ReminderMinutesBeforeStart = (int)reminder.Minutes;
                                itemModified++;
                            } //if Outlook reminders set
                        } //if google reminder found
                    } //foreach reminder

                } else { //no google reminders set
                    if (ai.ReminderSet) {
                        sb.AppendLine("Reminder: " + ai.ReminderMinutesBeforeStart + " => removed");
                        ai.ReminderSet = false;
                        itemModified++;
                    }
                }
            }
            if (itemModified > 0) {
                MainForm.Instance.Logboxout(sb.ToString(), false, verbose: true);
                MainForm.Instance.Logboxout(itemModified + " attributes updated.", verbose: true);
                System.Windows.Forms.Application.DoEvents();
            }
            return ai;
        }

        private void updateCalendarEntry_save(AppointmentItem ai) {
            if (Settings.Instance.SyncDirection == SyncDirection.Bidirectional) {
                log.Debug("Saving timestamp when OGCS updated appointment.");
                AddOGCSproperty(ref ai, Program.OGCSmodified, DateTime.Now);
            }
            ai.Save();
        }
        #endregion

        #region Delete
        public void DeleteCalendarEntries(List<AppointmentItem> oAppointments) {
            for (int o = oAppointments.Count - 1; o >= 0; o--) {
                AppointmentItem ai = oAppointments[o];
                Boolean doDelete = false;
                try {
                    doDelete = deleteCalendarEntry(ai);
                } catch (System.Exception ex) {
                    if (!Settings.Instance.VerboseOutput) MainForm.Instance.Logboxout(OutlookCalendar.GetEventSummary(ai));
                    MainForm.Instance.Logboxout("WARNING: Appointment deletion failed.\r\n" + ex.Message);
                    log.Error(ex.StackTrace);
                    if (MessageBox.Show("Outlook appointment deletion failed. Continue with synchronisation?", "Sync item failed", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                        continue;
                    else {
                        ai = (AppointmentItem)ReleaseObject(ai);
                        throw new UserCancelledSyncException("User chose not to continue sync.");
                    }
                }

                try {
                    if (doDelete) deleteCalendarEntry_save(ai);
                    else oAppointments.Remove(ai);
                } catch (System.Exception ex) {
                    MainForm.Instance.Logboxout("WARNING: Deleted appointment failed to remove.\r\n" + ex.Message);
                    log.Error(ex.StackTrace);
                    if (MessageBox.Show("Deleted Outlook appointment failed to remove. Continue with synchronisation?", "Sync item failed", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                        continue;
                    else {
                        throw new UserCancelledSyncException("User chose not to continue sync.");
                    }
                } finally {
                    ai = (AppointmentItem)ReleaseObject(ai);
                }
            }
        }
        
        private Boolean deleteCalendarEntry(AppointmentItem ai) {
            String eventSummary = GetEventSummary(ai);
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
                MainForm.Instance.Logboxout(eventSummary, verbose: true);
            }
            return doDelete;
        }

        private void deleteCalendarEntry_save(AppointmentItem ai) {
            ai.Delete();
        }
        #endregion
        
        public void ReclaimOrphanCalendarEntries(ref List<AppointmentItem> oAppointments, ref List<Event> gEvents) {
            log.Debug("Looking for orphaned items to reclaim...");

            //This is needed for people migrating from other tools, which do not have our GoogleID extendedProperty
            List<AppointmentItem> unclaimedAi = new List<AppointmentItem>();

            for (int o = oAppointments.Count-1; o>=0; o--){
                AppointmentItem ai = oAppointments[o];
                //Find entries with no Google ID
                if (ai.UserProperties[gEventID] == null) {
                    unclaimedAi.Add(ai);
                    foreach (Event ev in gEvents) {
                        //Use simple matching on start,end,subject,location to pair events
                        String sigAi = signature(ai);
                        String sigEv = GoogleCalendar.signature(ev);
                        if (Settings.Instance.Obfuscation.Enabled) {
                            if (Settings.Instance.Obfuscation.Direction == SyncDirection.OutlookToGoogle)
                                sigAi = Obfuscate.ApplyRegex(sigAi, SyncDirection.OutlookToGoogle);
                            else
                                sigEv = Obfuscate.ApplyRegex(sigEv, SyncDirection.GoogleToOutlook);
                        }
                        if (sigAi == sigEv) {
                            AddOGCSproperty(ref ai, gEventID, ev.Id);
                            updateCalendarEntry_save(ai);
                            unclaimedAi.Remove(ai);
                            MainForm.Instance.Logboxout("Reclaimed: " + GetEventSummary(ai), verbose: true);
                            break;
                        }
                    }
                }
                ai = (AppointmentItem)ReleaseObject(ai);
            }
            if ((Settings.Instance.SyncDirection == SyncDirection.GoogleToOutlook ||
                    Settings.Instance.SyncDirection == SyncDirection.Bidirectional) &&
                unclaimedAi.Count > 0 &&
                !Settings.Instance.MergeItems && !Settings.Instance.DisableDelete && !Settings.Instance.ConfirmOnDelete) {
                    
                if (MessageBox.Show(unclaimedAi.Count + " Outlook calendar items can't be matched to Google.\r\n" +
                    "Remember, it's recommended to have a dedicated Outlook calendar to sync with, " +
                    "or you may wish to merge with unmatched events. Continue with deletions?",
                    "Delete unmatched Outlook items?", MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation) == DialogResult.No) {

                    foreach (AppointmentItem ai in unclaimedAi) {
                        oAppointments.Remove(ai);
                    }
                }
            }
        }

        private void createRecipient(EventAttendee ea, AppointmentItem ai) {
            if (IOutlook.CurrentUserSMTP().ToLower() != ea.Email) {
                Recipient recipient = ai.Recipients.Add(ea.DisplayName + "<" + ea.Email + ">");
                recipient.Resolve();
                //ReadOnly: recipient.Type = (int)((bool)ea.Organizer ? OlMeetingRecipientType.olOrganizer : OlMeetingRecipientType.olRequired);
                recipient.Type = (int)(ea.Optional == null ? OlMeetingRecipientType.olRequired : ((bool)ea.Optional ? OlMeetingRecipientType.olOptional : OlMeetingRecipientType.olRequired));
                //ReadOnly: ea.ResponseStatus
            }
        }
        
        #region STATIC functions
        public static string signature(AppointmentItem ai) {
            return (GoogleCalendar.GoogleTimeFrom(ai.Start) + ";" + GoogleCalendar.GoogleTimeFrom(ai.End) + ";" + ai.Subject).Trim();
        }
        
        private static string exportToCSV(AppointmentItem ai) {
            System.Text.StringBuilder csv = new System.Text.StringBuilder();
            
            csv.Append(GoogleCalendar.GoogleTimeFrom(ai.Start) + ",");
            csv.Append(GoogleCalendar.GoogleTimeFrom(ai.End) + ",");
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
            csv.Append("\"" + (ai.RequiredAttendees==null?"":ai.RequiredAttendees) + "\",");
            csv.Append("\"" + (ai.OptionalAttendees==null?"":ai.OptionalAttendees) + "\",");
            csv.Append(ai.ReminderSet + ",");
            csv.Append(ai.ReminderMinutesBeforeStart.ToString() + ",");
            csv.Append(ai.EntryID + ",");
            if (ai.UserProperties[gEventID] != null)
                csv.Append(ai.UserProperties[gEventID].Value.ToString());

            return csv.ToString();
        }

        public static string GetEventSummary(AppointmentItem ai) {
            String eventSummary = "";
            if (ai.AllDayEvent) {
                log.Fine("GetSummary - all day event");
                eventSummary += ai.Start.Date.ToShortDateString();
            } else {
                log.Fine("GetSummary - not all day event");
                eventSummary += ai.Start.ToShortDateString() + " " + ai.Start.ToShortTimeString();
            }
            eventSummary += " " + (ai.IsRecurring ? "(R) " : "") + "=> ";
            eventSummary += '"' + ai.Subject + '"';
            return eventSummary;
        }

        public static void IdentifyEventDifferences(
            ref List<Event> google,             //need creating
            ref List<AppointmentItem> outlook,  //need deleting
            Dictionary<AppointmentItem, Event> compare) {
            log.Debug("Comparing Google events to Outlook items...");

            // Count backwards so that we can remove found items without affecting the order of remaining items
            for (int g = google.Count - 1; g >= 0; g--) {
                for (int o = outlook.Count - 1; o >= 0; o--) {
                    if (outlook[o].UserProperties[gEventID] != null &&
                        outlook[o].UserProperties[gEventID].Value.ToString() == google[g].Id.ToString()) {

                        compare.Add(outlook[o], google[g]);
                        outlook.Remove(outlook[o]);
                        google.Remove(google[g]);
                        break;

                    } else if (Settings.Instance.MergeItems) {
                        //Remove the non-Google item so it doesn't get deleted
                        outlook.Remove(outlook[o]);
                    }
                }
            }

            if (Settings.Instance.DisableDelete) {
                outlook = new List<AppointmentItem>();
            }
            if (Settings.Instance.SyncDirection == SyncDirection.Bidirectional) {
                //Don't recreate any items that have been deleted in Outlook
                for (int g = google.Count - 1; g >= 0; g--) {
                    if (google[g].ExtendedProperties != null &&
                        google[g].ExtendedProperties.Private != null &&
                        google[g].ExtendedProperties.Private.ContainsKey(GoogleCalendar.oEntryID))
                        google.Remove(google[g]);
                }
                //Don't delete any items that aren't yet in Google
                for (int o = outlook.Count - 1; o >= 0; o--) {
                    if (outlook[o].UserProperties[OutlookCalendar.gEventID] == null ||
                        outlook[o].LastModificationTime > Settings.Instance.LastSyncDate)
                        outlook.Remove(outlook[o]);
                }
            }
            if (Settings.Instance.CreateCSVFiles) {
                //Outlook Deletions
                log.Debug("Outputting items for deletion to CSV...");
                TextWriter tw = new StreamWriter(Path.Combine(Program.UserFilePath,"outlook_delete.csv"));
                foreach (AppointmentItem ai in outlook) {
                    tw.WriteLine(exportToCSV(ai));
                }
                tw.Close();

                //Outlook Creations
                log.Debug("Outputting items for creation to CSV...");
                tw = new StreamWriter(Path.Combine(Program.UserFilePath,"outlook_create.csv"));
                foreach (AppointmentItem ai in outlook) {
                    tw.WriteLine(OutlookCalendar.signature(ai));
                }
                tw.Close();
                log.Debug("Done.");
            }
        }

        public static object ReleaseObject(object obj) {
            try {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
            } catch { }
            return null;
        }
        #region OGCS Outlook properties
        public static void AddOGCSproperty(ref AppointmentItem ai, String key, String value) {
            UserProperty prop = ai.UserProperties.Find(key);
            if (prop == null) 
                ai.UserProperties.Add(key, OlUserPropertyType.olText);
            ai.UserProperties[key].Value = value;
        }

        public static void AddOGCSproperty(ref AppointmentItem ai, String key, DateTime value) {
            UserProperty prop = ai.UserProperties.Find(key);
            if (prop == null)
                ai.UserProperties.Add(key, OlUserPropertyType.olDateTime);
            ai.UserProperties[key].Value = value;
        }

        public static DateTime OGCSlastModified(AppointmentItem ai) {
            if (ai.UserProperties.Find(Program.OGCSmodified) == null)
                return new DateTime();
            return (DateTime)ai.UserProperties[Program.OGCSmodified].Value;
        }
        #endregion
        #endregion

    }
}
