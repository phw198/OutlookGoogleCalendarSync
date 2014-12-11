using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using log4net;
using Microsoft.Office.Interop.Outlook;
using Google.Apis.Calendar.v3.Data;

namespace OutlookGoogleCalendarSync {
    class OutlookOld : OutlookInterface {
        private static readonly ILog log = LogManager.GetLogger(typeof(OutlookOld));

        private Application oApp;
        private String currentUserSMTP;  //SMTP of account owner that has Outlook open
        private String currentUserName;  //Name of account owner - used to determine if attendee is "self"
        private MAPIFolder useOutlookCalendar;
        private Dictionary<string, MAPIFolder> calendarFolders = new Dictionary<string, MAPIFolder>();

        public void Connect() {
            log.Debug("Setting up Outlook connection.");

            // Create the Outlook application.
            oApp = new Application();

            // Get the NameSpace and Logon information.
            NameSpace oNS = oApp.GetNamespace("mapi");
            
            //Log on by using a dialog box to choose the profile.
            oNS.Logon("", "", true, true);
            currentUserSMTP = (oNS.CurrentUser as Recipient).Address.ToString().ToLower();
            currentUserName = oNS.CurrentUser.Name;

            //Alternate logon method that uses a specific profile.
            // If you use this logon method, 
            // change the profile name to an appropriate value.
            //oNS.Logon("YourValidProfile", Missing.Value, false, true); 

            // Get the Default Calendar folder
            if (Settings.Instance.OutlookService == OutlookCalendar.Service.AlternativeMailbox && Settings.Instance.MailboxName != "") {
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

        public List<String> Accounts() {
            List<String> accs = new List<String>();
            accs.Add(currentUserSMTP.ToLower());
            return accs;
        }
        public Dictionary<string, MAPIFolder> CalendarFolders() {
            return calendarFolders;
        }
        public MAPIFolder UseOutlookCalendar() {
            return useOutlookCalendar;
        }
        public void UseOutlookCalendar(MAPIFolder set) {
            useOutlookCalendar = set;
        }
        public String CurrentUserSMTP() {
            return currentUserSMTP;
        }
        public String CurrentUserName() {
            return currentUserName;
        }
        public Boolean Offline() {
            return oApp.GetNamespace("mapi").Offline;
        }

        private const String gEventID = "googleEventID";
        private const String attendeeSeparator = "\r\n===--- Attendees ---===";

        public void CreateCalendarEntries(List<Event> events) {
            foreach (Event ev in events) {
                AppointmentItem ai = useOutlookCalendar.Items.Add() as AppointmentItem;

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
                        if (ea.DisplayName == currentUserName) foundCurrentUser = true;
                        ai.Recipients.Add(ea.DisplayName);
                        bool gOptional = (ea.Optional == null) ? false : (bool)ea.Optional;
                        if (gOptional) {
                            ai.OptionalAttendees += "; " + ea.Email;
                            ai.RequiredAttendees = ai.RequiredAttendees.Replace(ea.Email, "");
                        }
                    }
                }
                if (!foundCurrentUser) ai.Recipients.Add(currentUserSMTP);
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

                MainForm.Instance.Logboxout(OutlookCalendar.GetEventSummary(ai), verbose: true);
                OutlookCalendar.AddCalendarEntry(ai);
            }
        }

        public void UpdateCalendarEntries(Dictionary<AppointmentItem, Event> entriesToBeCompared, ref int entriesUpdated) {
            foreach (KeyValuePair<AppointmentItem, Event> compare in entriesToBeCompared) {
                AppointmentItem ai = compare.Key;
                Event ev = compare.Value;
                if (DateTime.Parse(ev.Updated) < DateTime.Parse(GoogleCalendar.GoogleTimeFrom(ai.LastModificationTime))) continue;

                int itemModified = 0;
                System.Text.StringBuilder sb = new System.Text.StringBuilder();
                sb.AppendLine(GoogleCalendar.GetEventSummary(ev));

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

                Dictionary<String, Boolean> attendees = GetAttendeesFromDescription(ev.Description);
                ev.Description = GetDescription(ev.Description);
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
                                String recipientSMTP = ""; // pa.GetProperty(PR_SMTP_ADDRESS).ToString();
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
                        if (recipient.Name != currentUserName) {
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

                    OutlookCalendar.Instance.UpdateCalendarEntry(ai);
                    entriesUpdated++;
                }
            }
        }

        public String GetRecipientEmail(Recipient recipient) {
            return "";
        }

        public Event AddGoogleAttendee(EventAttendee ea, Event ev) {
            if (ev.Description==null || !ev.Description.Contains(attendeeSeparator)) {
                ev.Description += attendeeSeparator;
            }
            ev.Description += "\r\n" + ea.DisplayName;
            if ((bool)ea.Optional)
                ev.Description += " >> Optional";
            return ev;
        }

        public static String GetDescription(String description) {
            if (description.Contains(attendeeSeparator)) {
                string[] splitter = new string[] { attendeeSeparator };
                string[] result = description.Split(splitter,2,StringSplitOptions.None);
                return result[0];
            } 
            return description;
        }

        public static Dictionary<String,Boolean> GetAttendeesFromDescription(String description) {
            Dictionary<String,Boolean> attendeeList = new Dictionary<String,Boolean>();
            if (description.Contains(attendeeSeparator)) {
                string[] splitter = new string[] { attendeeSeparator };
                string[] result = description.Split(splitter, 2, StringSplitOptions.None);
                if (result.Count() == 2) {
                    splitter = new string[] { "\r\n" };
                    string[] attendees = result[1].Split(splitter, 100, StringSplitOptions.RemoveEmptyEntries);
                    foreach (String attendee in attendees) {
                        String attendeeName = attendee;
                        Boolean optional = false;
                        if (attendee.Contains(" >> Optional")) {
                            optional = true;
                            attendeeName = attendee.Replace(" >> Optional", "");
                        }
                        attendeeList.Add(attendeeName, optional);
                    }
                }
            }
            return attendeeList;
        }

        public Boolean CompareRecipientsToAttendees(AppointmentItem ai, Event ev, 
            Dictionary<String, Boolean> attendeesFromDescription, 
            StringBuilder sb, ref int itemModified
        ) {
            //Build a list of Google attendees. Any remaining at the end of the diff must be deleted.
            Dictionary<String, Boolean> removeAttendee = attendeesFromDescription;
            
            if (ai.Recipients.Count > 1) {
                for (int o = ai.Recipients.Count; o > 0; o--) {
                    bool foundAttendee = false;
                    Recipient recipient = ai.Recipients[o];

                    for (int g = removeAttendee.Count - 1; g >= 0; g--) {
                        String attendeeName = removeAttendee.Keys.ElementAt(g);
                        if (recipient.Name == attendeeName) {
                            foundAttendee = true;

                            //Optional attendee
                            bool oOptional = (ai.OptionalAttendees != null && ai.OptionalAttendees.Contains(recipient.Name));
                            bool gOptional = removeAttendee.Values.ElementAt(g);
                            if (MainForm.CompareAttribute("Attendee " + attendeeName + " - Optional",
                                SyncDirection.OutlookToGoogle, gOptional, oOptional, sb, ref itemModified)) {
                                    itemModified++;
                            }
                            removeAttendee.Remove(attendeeName);
                        }
                    }
                    if (!foundAttendee) {
                        sb.AppendLine("Attendee added: " + recipient.Name);
                        EventAttendee ea = GoogleCalendar.CreateAttendee(recipient, ai);
                        attendeesFromDescription.Add(ea.DisplayName, (bool)ea.Optional);
                        removeAttendee.Remove(recipient.Name);
                        itemModified++;
                    }
                }
            } //more than just 1 (me) recipients
            
            foreach (KeyValuePair<String,Boolean> attendee in removeAttendee) {
                sb.AppendLine("Attendee removed: " + attendee.Key);
                itemModified++;
            }
            return (itemModified > 0);
        }

    }
}
