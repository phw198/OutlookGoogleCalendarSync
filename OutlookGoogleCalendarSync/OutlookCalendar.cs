using System;
using System.Windows.Forms;
using System.Collections.Generic;
using Microsoft.Office.Interop.Outlook;
using System.IO;
using Google.Apis.Calendar.v3.Data;

namespace OutlookGoogleCalendarSync {
    /// <summary>
    /// Description of OutlookCalendar.
    /// </summary>
    public class OutlookCalendar {
        public const String PR_SMTP_ADDRESS = "http://schemas.microsoft.com/mapi/proptag/0x39FE001E";

        private Microsoft.Office.Interop.Outlook.Application oApp;
        private static OutlookCalendar instance;
        private String currentUserSMTP;  //SMTP of account owner that has Outlook open
        private String currentUserName;  //Name of account owner - used to determine if attendee is "self"
        private MAPIFolder useOutlookCalendar;
        private Accounts accounts;
        private Dictionary<string, MAPIFolder> calendarFolders = new Dictionary<string, MAPIFolder>();

        public static OutlookCalendar Instance {
            get {
                if (instance == null) instance = new OutlookCalendar();
                return instance;
            }
        }
        public String CurrentUserSMTP {
            get { return currentUserSMTP; }
        }
        public String CurrentUserName {
            get { return currentUserName; }
        }
        public MAPIFolder UseOutlookCalendar {
            get { return useOutlookCalendar; }
            set {
                useOutlookCalendar = value;
                Settings.Instance.UseOutlookCalendar = new MyOutlookCalendarListEntry(value);
            }
        }
        public Accounts Accounts {
            get { return accounts; }
        }
        public Dictionary<string, MAPIFolder> CalendarFolders {
            get { return calendarFolders; }
        }
        public enum Service {
            DefaultMailbox,
            AlternativeMailbox,
            EWS
        }
        private const String gEventID = "googleEventID";

        public OutlookCalendar() {

            // Create the Outlook application.
            oApp = new Microsoft.Office.Interop.Outlook.Application();

            // Get the NameSpace and Logon information.
            NameSpace oNS = oApp.GetNamespace("mapi");

            //Log on by using a dialog box to choose the profile.
            oNS.Logon("", "", true, true);
            currentUserSMTP = ((oNS.CurrentUser as Recipient).PropertyAccessor as PropertyAccessor).GetProperty(OutlookCalendar.PR_SMTP_ADDRESS).ToString().ToLower();
            currentUserName = oNS.CurrentUser.Name;

            //Alternate logon method that uses a specific profile.
            // If you use this logon method, 
            // change the profile name to an appropriate value.
            //oNS.Logon("YourValidProfile", Missing.Value, false, true); 

            //Get the accounts configured in Outlook
            accounts = oNS.Accounts;

            // Get the Default Calendar folder
            if (Settings.Instance.OutlookService == Service.AlternativeMailbox && Settings.Instance.MailboxName != "") {
                useOutlookCalendar = oNS.Folders[Settings.Instance.MailboxName].Folders["Calendar"];
            } else {
                // Use the logged in user's Calendar folder.
                useOutlookCalendar = oNS.GetDefaultFolder(OlDefaultFolders.olFolderCalendar);
            }
            calendarFolders.Add("Default " + useOutlookCalendar.Name, useOutlookCalendar);
            //Get any subfolders - note, this isn't recursive
            foreach (MAPIFolder calendar in useOutlookCalendar.Folders) {
                if (calendar.DefaultItemType == OlItemType.olAppointmentItem) {
                    calendarFolders.Add(calendar.Name, calendar);
                }
            }

            // Done. Log off.
            oNS.Logoff();
        }

        public void Reset() {
            instance = new OutlookCalendar();
        }

        public List<AppointmentItem> getCalendarEntriesInRange() {
            List<AppointmentItem> result = new List<AppointmentItem>();

            Items OutlookItems = UseOutlookCalendar.Items;
            OutlookItems.Sort("[Start]", Type.Missing);
            OutlookItems.IncludeRecurrences = true;

            if (OutlookItems != null) {
                DateTime min = DateTime.Today.AddDays(-Settings.Instance.DaysInThePast);
                DateTime max = DateTime.Today.AddDays(+Settings.Instance.DaysInTheFuture + 1);
                string filter = "[End] >= '" + min.ToString("g") + "' AND [Start] < '" + max.ToString("g") + "'";

                foreach (AppointmentItem ai in OutlookItems.Restrict(filter)) {
                    if (ai.End == min) continue; //Required for midnight to midnight events 
                    result.Add(ai);
                }
            }

            if (Settings.Instance.CreateCSVFiles) {
                TextWriter tw = new StreamWriter("outlook_appointments.csv");
                String CSVheader = "Start Time,Finish Time,Subject,Location,Description,Privacy,FreeBusy,";
                CSVheader += "Required Attendees,Optional Attendees,Reminder Set,Reminder Minutes,Outlook ID,Google ID";
                tw.WriteLine(CSVheader);
                foreach (AppointmentItem ai in result) {
                    try {
                        tw.WriteLine(exportToCSV(ai));
                    } catch {
                        MainForm.Instance.Logboxout("Failed to output following Outlook appointment to CSV:-");
                        MainForm.Instance.Logboxout(getEventSummary(ai));
                    }
                }
                tw.Close();
            }

            return result;
        }

        private void addCalendarEntry(AppointmentItem ai) {
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
                AppointmentItem ai = UseOutlookCalendar.Items.Add() as AppointmentItem; 
                
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
                
                Boolean foundCurrentUser = false;
                if (Settings.Instance.AddAttendees && ev.Attendees != null) {
                    foreach (EventAttendee ea in ev.Attendees) {
                        if (ea.DisplayName == CurrentUserName) foundCurrentUser = true;
                        ai.Recipients.Add(ea.Email);
                        bool gOptional = (ea.Optional == null) ? false : (bool)ea.Optional;
                        if (gOptional) {
                            ai.OptionalAttendees += "; "+ ea.Email ;
                            ai.RequiredAttendees = ai.RequiredAttendees.Replace(ea.Email, "");
                        } 
                    }
                }
                if (!foundCurrentUser) ai.Recipients.Add(CurrentUserSMTP);
                ai.Recipients.ResolveAll();
                
                //Reminder alert
                if (Settings.Instance.AddReminders && ev.Reminders != null && ev.Reminders.Overrides != null) {
                    foreach (EventReminder reminder in ev.Reminders.Overrides) {
                        if (reminder.Method == "popup") {
                            ai.ReminderSet = true;
                            ai.ReminderMinutesBeforeStart = (int)reminder.Minutes;
                        }
                    }
                }
                
                MainForm.Instance.Logboxout(getEventSummary(ai), verbose: true);
                addCalendarEntry(ai);
            }
        }

        public void UpdateCalendarEntries(Dictionary<AppointmentItem, Event> entriesToBeCompared, ref int entriesUpdated) {
            entriesUpdated = 0;
            foreach (KeyValuePair<AppointmentItem, Event> compare in entriesToBeCompared) {
                AppointmentItem ai = compare.Key;
                Event ev = compare.Value;
                if (DateTime.Parse(ev.Updated) < DateTime.Parse(GoogleCalendar.GoogleTimeFrom(ai.LastModificationTime))) continue;

                int itemModified = 0;
                System.Text.StringBuilder sb = new System.Text.StringBuilder();
                sb.AppendLine(GoogleCalendar.getEventSummary(ev));
                
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
                if (Settings.Instance.AddDescription) {
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
                    //Build a list of Outlook attendees. Any remaining at the end of the diff must be deleted.
                    List<Recipient> removeRecipient = new List<Recipient>();
                    if (ai.Recipients != null) {
                        foreach (Recipient recipient in ai.Recipients) {
                            removeRecipient.Add(recipient);
                        }
                    }
                    if (ev.Attendees != null && ev.Attendees.Count > 1) {
                        for (int g = ev.Attendees.Count - 1; g >= 0; g--) {
                            bool foundRecipient = false;
                            EventAttendee attendee = ev.Attendees[g];

                            if (ai.Recipients == null) break;
                            for (int o = removeRecipient.Count - 1; o >= 0; o--) {
                                Recipient recipient = removeRecipient[o];
                                recipient.Resolve();
                                Microsoft.Office.Interop.Outlook.PropertyAccessor pa = recipient.PropertyAccessor;
                                String recipientSMTP = pa.GetProperty(PR_SMTP_ADDRESS).ToString();
                                if (recipientSMTP.ToLower() == attendee.Email.ToLower()) {
                                    foundRecipient = true;
                                    removeRecipient.RemoveAt(o);

                                    //Optional attendee
                                    bool oOptional = (ai.OptionalAttendees != null && ai.OptionalAttendees.Contains(recipient.Name));
                                    bool gOptional = (attendee.Optional == null) ? false : (bool)attendee.Optional;
                                    if (MainForm.CompareAttribute("Recipient " + recipient.Name + " - Optional",
                                        SyncDirection.GoogleToOutlook, gOptional, oOptional, sb, ref itemModified)) {
                                        if (gOptional) {
                                            ai.OptionalAttendees += "; " + recipient.Name;
                                            ai.RequiredAttendees = ai.RequiredAttendees.Replace(recipient.Name, "");
                                        } else {
                                            ai.RequiredAttendees += "; " + recipient.Name;
                                            ai.OptionalAttendees = ai.OptionalAttendees.Replace(recipient.Name, "");
                                        }
                                    }
                                    //Response is readonly in Outlook :(
                                    break;
                                }
                            }
                            if (!foundRecipient) {
                                sb.AppendLine("Recipient added: " + attendee.DisplayName);
                                ai.Recipients.Add(attendee.Email).Resolve();
                                if (attendee.Optional != null && (bool)attendee.Optional) {
                                    ai.OptionalAttendees += ";" + attendee.Email;
                                } else {
                                    ai.RequiredAttendees += ";" + attendee.Email;
                                }
                                itemModified++;
                            }
                        }
                    } //more than just 1 (me) recipients
                    
                    foreach (Recipient recipient in removeRecipient) {
                        if (recipient.Name != this.CurrentUserName) { 
                            //Outlook must have current user as recipient, Google doesn't (organiser doesn't have to be an attendee)
                            sb.AppendLine("Recipient removed: " + recipient.Name);
                            recipient.Delete();
                            itemModified++;
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
                }
                if (itemModified > 0) {
                    MainForm.Instance.Logboxout(sb.ToString(), false, verbose: true);
                    MainForm.Instance.Logboxout(itemModified + " attributes updated.", verbose: true);
                    System.Windows.Forms.Application.DoEvents();

                    OutlookCalendar.Instance.updateCalendarEntry(ai);
                    entriesUpdated++;
                }
            }
        }

        public void DeleteCalendarEntries(List<AppointmentItem> oAppointments) {
            foreach (AppointmentItem ai in oAppointments) {
                String eventSummary = getEventSummary(ai);
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
                            MainForm.Instance.Logboxout("Reclaimed: " + getEventSummary(ai), verbose: true);
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

        public static string getEventSummary(AppointmentItem ai) {
            String eventSummary = "";
            if (ai.AllDayEvent)
                eventSummary += DateTime.Parse(GoogleCalendar.GoogleTimeFrom(ai.Start)).ToString("dd/MM/yyyy");
            else
                eventSummary += DateTime.Parse(GoogleCalendar.GoogleTimeFrom(ai.Start)).ToString("dd/MM/yyyy HH:mm");
            eventSummary += " => ";
            eventSummary += '"' + ai.Subject + '"';
            return eventSummary;
        }

        public static void IdentifyEventDifferences(
            ref List<Event> google,
            ref List<AppointmentItem> outlook,
            Dictionary<AppointmentItem, Event> compare) {
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
                TextWriter tw = new StreamWriter("outlook_delete.csv");
                foreach (AppointmentItem ai in outlook) {
                    tw.WriteLine(exportToCSV(ai));
                }
                tw.Close();

                //Outlook Creations
                tw = new StreamWriter("outlook_create.csv");
                foreach (AppointmentItem ai in outlook) {
                    tw.WriteLine(OutlookCalendar.signature(ai));
                }
                tw.Close();
            }
        }
        #endregion
    }
}
