using System;
using System.Windows.Forms;
using System.Collections.Generic;
using Microsoft.Office.Interop.Outlook;
using System.IO;
using Google.Apis.Calendar.v3.Data;
using log4net;

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
            UseOutlookCalendar.Items.ItemAdd += new ItemsEvents_ItemAddEventHandler(appointmentItem_Add);
            UseOutlookCalendar.Items.ItemChange += new ItemsEvents_ItemChangeEventHandler(appointmentItem_Change);
            //useOutlookCalendar.Items.ItemRemove += new ItemsEvents_ItemRemoveEventHandler(appointmentItem_Remove);
        }

        public void DeregisterForAutoSync() {
            log.Info("Deregistering from Outlook appointment change events...");
            UseOutlookCalendar.Items.ItemAdd -= new ItemsEvents_ItemAddEventHandler(appointmentItem_Add);
            UseOutlookCalendar.Items.ItemChange -= new ItemsEvents_ItemChangeEventHandler(appointmentItem_Change);
            //Can't do removes easily as the event doesn't tell us which item was removed.
            //useOutlookCalendar.Items.ItemRemove -= new ItemsEvents_ItemRemoveEventHandler(appointmentItem_Remove);
        }

        private void appointmentItem_Add(object Item) {
            //We could deregister the event when syncing, but then we'd have to handle 
            //turning it back on if there's an exception. Also, no need to turn off if only going from Outlook to Google
            if (MainForm.Instance.SyncingNow && Settings.Instance.SyncDirection != SyncDirection.OutlookToGoogle) return;
            log.Debug("Outlook item added.");
            changeQueue_Add(Item as AppointmentItem);
        }
        private void appointmentItem_Change(object Item) {
            if (MainForm.Instance.SyncingNow && Settings.Instance.SyncDirection != SyncDirection.OutlookToGoogle) return;
            log.Debug("Outlook item changed.");
            changeQueue_Add(Item as AppointmentItem);
        }
        //void appointmentItem_Remove() {
        //    log.Debug("Outlook item removed.");
        //}
        private void changeQueue_Add(AppointmentItem ai) {
            if (changeQueue.ContainsKey(ai.EntryID)) 
                changeQueue.Remove(ai.EntryID);
            changeQueue.Add(ai.EntryID, ai);
            //***Is this item in the right date range?
            //***Trigger another thread in 30secs.
            //That thread also needs to check if already syncing and maybe retry
        }
        #endregion

        public List<AppointmentItem> getCalendarEntriesInRange() {
            return filterCalendarEntries(UseOutlookCalendar.Items);
        }

        public List<AppointmentItem> getCalendarEntriesInRange(Dictionary<String, AppointmentItem> changeQueue) {
            Items OutlookItems = null;
            foreach (KeyValuePair<String,AppointmentItem> qItem in changeQueue) {
                OutlookItems.Add(qItem.Value);
            }
            return filterCalendarEntries(OutlookItems);
        }

        public List<AppointmentItem> filterCalendarEntries(Items OutlookItems) {
            List<AppointmentItem> result = new List<AppointmentItem>();

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
                log.Debug("Outputting CSV files...");
                TextWriter tw = new StreamWriter("outlook_appointments.csv");
                String CSVheader = "Start Time,Finish Time,Subject,Location,Description,Privacy,FreeBusy,";
                CSVheader += "Required Attendees,Optional Attendees,Reminder Set,Reminder Minutes,Outlook ID,Google ID";
                tw.WriteLine(CSVheader);
                foreach (AppointmentItem ai in result) {
                    try {
                        tw.WriteLine(exportToCSV(ai));
                    } catch {
                        MainForm.Instance.Logboxout("Failed to output following Outlook appointment to CSV:-");
                        MainForm.Instance.Logboxout(GetEventSummary(ai));
                    }
                }
                tw.Close();
                log.Debug("Done.");
            }

            return result;
        }

        public static void AddCalendarEntry(AppointmentItem ai) {
            ai.Save();
        }

        public void UpdateCalendarEntry(AppointmentItem ai) {
            ai.Save();
        }

        private void deleteCalendarEntry(AppointmentItem ai) {
            ai.Delete();
        }

        public void CreateCalendarEntries(List<Event> events) {
            instance.IOutlook.CreateCalendarEntries(events);
        }

        public void UpdateCalendarEntries(Dictionary<AppointmentItem, Event> entriesToBeCompared, ref int entriesUpdated) {
            entriesUpdated = 0;
            instance.IOutlook.UpdateCalendarEntries(entriesToBeCompared, ref entriesUpdated);
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
                            UpdateCalendarEntry(ai);
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
                TextWriter tw = new StreamWriter("outlook_delete.csv");
                foreach (AppointmentItem ai in outlook) {
                    tw.WriteLine(exportToCSV(ai));
                }
                tw.Close();

                //Outlook Creations
                log.Debug("Outputting items for creation to CSV...");
                tw = new StreamWriter("outlook_create.csv");
                foreach (AppointmentItem ai in outlook) {
                    tw.WriteLine(OutlookCalendar.signature(ai));
                }
                tw.Close();
                log.Debug("Done.");
            }
        }
        #endregion
    }
}
