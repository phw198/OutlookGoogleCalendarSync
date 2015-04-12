using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Google.Apis.Calendar.v3.Data;
using log4net;
using Microsoft.Office.Interop.Outlook;

namespace OutlookGoogleCalendarSync {
    /// <summary>
    /// Description of OutlookCalendar.
    /// </summary>
    public class OutlookCalendar {
        private static OutlookCalendar instance;
        private Dictionary<string, AppointmentItem> changeQueue = new Dictionary<string, AppointmentItem>();
        
        private static readonly ILog log = LogManager.GetLogger(typeof(OutlookCalendar));
        public OutlookInterface IOutlook;
        
        public static OutlookCalendar Instance {
            get {
                if (instance == null) instance = new OutlookCalendar();
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
        private const String gEventID = "googleEventID";

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

            OutlookItems.Sort("[Start]", Type.Missing);
            OutlookItems.IncludeRecurrences = true;

            if (OutlookItems != null) {
                DateTime min = DateTime.Today.AddDays(-Settings.Instance.DaysInThePast);
                DateTime max = DateTime.Today.AddDays(+Settings.Instance.DaysInTheFuture + 1);
                
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

        private static void addCalendarEntry(AppointmentItem ai) {
            ai.Save();
        }

        private void updateCalendarEntry(AppointmentItem ai) {
            ai.Save();
        }

        private void deleteCalendarEntry(AppointmentItem ai) {
            ai.Delete();
        }

        public void CreateCalendarEntries(List<Event> events) {
            foreach (Event ev in events) {
                log.Fine("Processing >> " + GoogleCalendar.GetEventSummary(ev));
                AppointmentItem ai = IOutlook.UseOutlookCalendar().Items.Add() as AppointmentItem;

                //Add the Google event ID into Outlook appointment.
                ai.UserProperties.Add(gEventID, OlUserPropertyType.olText);
                ai.UserProperties[gEventID].Value = ev.Id;

                ai.Start = new DateTime();
                ai.End = new DateTime();

                if (ev.Start.Date != null) {
                    ai.AllDayEvent = true;
                    ai.Start = DateTime.Parse(ev.Start.Date);
                    ai.End = DateTime.Parse(ev.End.Date);
                } else {
                    ai.AllDayEvent = false;
                    ai.Start = DateTime.Parse(ev.Start.DateTime);
                    ai.End = DateTime.Parse(ev.End.DateTime);
                }
                ai.Subject = ev.Summary;
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

                MainForm.Instance.Logboxout(OutlookCalendar.GetEventSummary(ai), verbose: true);
                OutlookCalendar.addCalendarEntry(ai);
            }
        }

        public void UpdateCalendarEntries(Dictionary<AppointmentItem, Event> entriesToBeCompared, ref int entriesUpdated) {
            entriesUpdated = 0;

            foreach (KeyValuePair<AppointmentItem, Event> compare in entriesToBeCompared) {
                AppointmentItem ai = compare.Key;
                Event ev = compare.Value;
                if (DateTime.Parse(ev.Updated) < DateTime.Parse(GoogleCalendar.GoogleTimeFrom(ai.LastModificationTime))) continue;

                int itemModified = 0;
                String evSummary = GoogleCalendar.GetEventSummary(ev);
                log.Fine("Processing >> " + evSummary);

                System.Text.StringBuilder sb = new System.Text.StringBuilder();
                sb.AppendLine(evSummary);

                if (ev.Start.Date != null) {
                    ai.AllDayEvent = true;
                    if (MainForm.CompareAttribute("Start time", SyncDirection.GoogleToOutlook, ev.Start.Date, ai.Start.ToString("yyyy-MM-dd"), sb, ref itemModified)) {
                        ai.Start = DateTime.Parse(ev.Start.Date);
                    }
                    if (MainForm.CompareAttribute("End time", SyncDirection.GoogleToOutlook, ev.End.Date, ai.End.ToString("yyyy-MM-dd"), sb, ref itemModified)) {
                        ai.End = DateTime.Parse(ev.End.Date);
                    }
                } else {
                    ai.AllDayEvent = false;
                    if (MainForm.CompareAttribute("Start time",
                        SyncDirection.GoogleToOutlook,
                        GoogleCalendar.GoogleTimeFrom(DateTime.Parse(ev.Start.DateTime)),
                        GoogleCalendar.GoogleTimeFrom(ai.Start), sb, ref itemModified)) {
                        ai.Start = DateTime.Parse(ev.Start.DateTime);
                    }
                    if (MainForm.CompareAttribute("End time",
                        SyncDirection.GoogleToOutlook,
                        GoogleCalendar.GoogleTimeFrom(DateTime.Parse(ev.End.DateTime)),
                        GoogleCalendar.GoogleTimeFrom(ai.End), sb, ref itemModified)) {
                        ai.End = DateTime.Parse(ev.End.DateTime);
                    }
                }
                if (MainForm.CompareAttribute("Subject", SyncDirection.GoogleToOutlook, ev.Summary, ai.Subject, sb, ref itemModified)) {
                    ai.Subject = ev.Summary;
                }
                if (!Settings.Instance.AddDescription) ev.Description = "";
                if (MainForm.CompareAttribute("Description", SyncDirection.GoogleToOutlook, ev.Description, ai.Body, sb, ref itemModified))
                    ai.Body = ev.Description;

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
                    } else {
                        //Build a list of Outlook attendees. Any remaining at the end of the diff must be deleted.
                        List<Recipient> removeRecipient = new List<Recipient>();
                        if (ai.Recipients != null) {
                            foreach (Recipient recipient in ai.Recipients) {
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
                                        bool oOptional = (ai.OptionalAttendees != null && ai.OptionalAttendees.Contains(attendee.DisplayName));
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
                                if (!foundRecipient) {
                                    sb.AppendLine("Recipient added: " + attendee.DisplayName);
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

                    OutlookCalendar.Instance.updateCalendarEntry(ai);
                    entriesUpdated++;
                } else {
                    //Do a dummy update in order to update the last modified date
                    ai.Subject += " ";
                    OutlookCalendar.Instance.updateCalendarEntry(ai);
                }
            }
        }

        public void DeleteCalendarEntries(List<AppointmentItem> oAppointments) {
            foreach (AppointmentItem ai in oAppointments) {
                String eventSummary = GetEventSummary(ai);
                Boolean delete = true;

                if (Settings.Instance.ConfirmOnDelete) {
                    if (MessageBox.Show("Delete " + eventSummary + "?", "Deletion Confirmation",
                        MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No) {
                        delete = false;
                        MainForm.Instance.Logboxout("Not deleted: " + eventSummary);
                    }
                } else {
                    MainForm.Instance.Logboxout(eventSummary, verbose: true);
                }
                if (delete) {
                    OutlookCalendar.Instance.deleteCalendarEntry(ai);
                    if (Settings.Instance.ConfirmOnDelete) MainForm.Instance.Logboxout("Deleted: " + eventSummary);
                }
            }
        }

        public void ReclaimOrphanCalendarEntries(ref List<AppointmentItem> oAppointments, ref List<Event> gEvents) {
            log.Debug("Looking for orphaned items to reclaim...");

            //This is needed for people migrating from other tools, which do not have our GoogleID extendedProperty
            int unclaimed = 0;
            List<AppointmentItem> unclaimedAi = new List<AppointmentItem>();

            foreach (AppointmentItem ai in oAppointments) {
                //Find entries with no Google ID
                if (ai.UserProperties[gEventID] == null) {
                    unclaimedAi.Add(ai);
                    foreach (Event ev in gEvents) {
                        //Use simple matching on start,end,subject,location to pair events
                        if (signature(ai) == GoogleCalendar.signature(ev)) {
                            ai.UserProperties.Add(gEventID, OlUserPropertyType.olText).Value = ev.Id;
                            updateCalendarEntry(ai);
                            unclaimedAi.Remove(ai);
                            MainForm.Instance.Logboxout("Reclaimed: " + GetEventSummary(ai), verbose: true);
                            break;
                        }
                    }
                }
            }
            if ((Settings.Instance.SyncDirection == SyncDirection.GoogleToOutlook ||
                    Settings.Instance.SyncDirection == SyncDirection.Bidirectional) &&
                unclaimedAi.Count > 0 &&
                !Settings.Instance.MergeItems && !Settings.Instance.DisableDelete && !Settings.Instance.ConfirmOnDelete) {

                if (MessageBox.Show(unclaimed + " Outlook calendar items can't be matched to Google.\r\n" +
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
                    if (recipientSMTP.IndexOf("<") > 0) {
                        recipientSMTP = recipientSMTP.Substring(recipientSMTP.IndexOf("<") + 1);
                        recipientSMTP = recipientSMTP.TrimEnd(Convert.ToChar(">"));
                    }
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
                log.Fine("Attendee removed: " + (ea.DisplayName == null ? ea.Email : ea.DisplayName));
                sb.AppendLine("Attendee removed: " + (ea.DisplayName==null ? ea.Email : ea.DisplayName));
                ev.Attendees.Remove(ea);
                itemModified++;
            }
            return (itemModified > 0);
        }

        #region STATIC functions
        public static string signature(AppointmentItem ai) {
            return (GoogleCalendar.GoogleTimeFrom(ai.Start) + ";" + GoogleCalendar.GoogleTimeFrom(ai.End) + ";" + ai.Subject + ";" + ai.Location).Trim();
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
                eventSummary += DateTime.Parse(GoogleCalendar.GoogleTimeFrom(ai.Start)).ToString("dd/MM/yyyy");
            } else {
                log.Fine("GetSummary - not all day event");
                eventSummary += DateTime.Parse(GoogleCalendar.GoogleTimeFrom(ai.Start)).ToString("dd/MM/yyyy HH:mm");
            }
            eventSummary += " => ";
            eventSummary += '"' + ai.Subject + '"';
            return eventSummary;
        }

        public static void IdentifyEventDifferences(
            ref List<Event> google,
            ref List<AppointmentItem> outlook,
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

                    } else if (Settings.Instance.MergeItems && !Settings.Instance.DisableDelete) {
                        //Remove the non-Google item so it doesn't get deleted
                        outlook.Remove(outlook[o]);
                    }
                }
            }

            if (Settings.Instance.DisableDelete) {
                outlook = new List<AppointmentItem>();
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

        public static String BuildFakeEmailAddress(String recipientName) {
            String buildFakeEmail = recipientName.Replace(",", "");
            buildFakeEmail = buildFakeEmail.Replace(" ", "");
            buildFakeEmail += "@unknownemail.com";
            log.Debug("Built a fake email for them: " + buildFakeEmail);
            return buildFakeEmail;
        }

        #endregion

    }
}
