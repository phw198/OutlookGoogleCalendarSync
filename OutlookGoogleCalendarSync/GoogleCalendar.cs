using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using DotNetOpenAuth.OAuth2;
using Google.Apis.Authentication.OAuth2;
using Google.Apis.Authentication.OAuth2.DotNetOpenAuth;
using Google.Apis.Calendar.v3;
using Google.Apis.Calendar.v3.Data;
using Google.Apis.Util;
using log4net;
using Microsoft.Office.Interop.Outlook;

namespace OutlookGoogleCalendarSync {
    /// <summary>
    /// Description of GoogleCalendar.
    /// </summary>
    public class GoogleCalendar {
        private static readonly ILog log = LogManager.GetLogger(typeof(GoogleCalendar));

        private static GoogleCalendar instance;
        public static GoogleCalendar Instance {
            get {
                if (instance == null) instance = new GoogleCalendar();
                return instance;
            }
        }

        private CalendarService service;
        public const String oEntryID = "outlook_EntryID";
        public static Boolean APIlimitReached_attendee = false;

        public GoogleCalendar() {
            var provider = new NativeApplicationClient(GoogleAuthenticationServer.Description);
            provider.ClientIdentifier = "653617509806-2nq341ol8ejgqhh2ku4j45m7q2bgdimv.apps.googleusercontent.com";
            provider.ClientSecret = "tAi-gZLWtasS58i8CcCwVwsq";
            service = new CalendarService(new OAuth2Authenticator<NativeApplicationClient>(provider, getAuthentication));
        }

        public void Reset() {
            instance = new GoogleCalendar();
            Settings.Instance.RefreshToken = "";
        }

        public List<MyGoogleCalendarListEntry> GetCalendars() {
            CalendarList request = null;
            try {
                request = service.CalendarList.List().Fetch();
            } catch (ApplicationException ex) {
                throw ex;
            }

            if (request != null) {

                List<MyGoogleCalendarListEntry> result = new List<MyGoogleCalendarListEntry>();
                foreach (CalendarListEntry cle in request.Items) {
                    result.Add(new MyGoogleCalendarListEntry(cle));
                }
                return result;
            } else {
                log.Error("Handshaking with the Google calendar service failed.");
            }
            return null;
        }
        
        public List<Event> GetCalendarEntriesInRecurrence(String recurringEventId) {
            List<Event> result = new List<Event>();
            Events request = null;
            String pageToken = null;

            try {
                log.Debug("Retrieving all recurring event instances from Google.");
                do {
                    EventsResource.InstancesRequest ir = service.Events.Instances(Settings.Instance.UseGoogleCalendar.Id, recurringEventId);
                    ir.ShowDeleted = true;
                    ir.PageToken = pageToken;
                    request = ir.Fetch();
                    pageToken = request.NextPageToken;

                    if (request != null)
                        if (request.Items != null) result.AddRange(request.Items);
                } while (pageToken != null);
                return result;
            } catch (System.Exception ex) {
                MainForm.Instance.Logboxout("ERROR: Failed to retrieve recurring events");
                log.Error(ex.Message);
                return null;
            }
        }

        public Event GetCalendarEntry(String eventId) {
            Event request = null;

            try {
                log.Debug("Retrieving specific Event with ID " + eventId);
                EventsResource.GetRequest gr = service.Events.Get(Settings.Instance.UseGoogleCalendar.Id, eventId);
                request = gr.Fetch();

                if (request != null)
                    return request;
                else
                    throw new System.Exception("Returned null");
            } catch (System.Exception ex) {
                MainForm.Instance.Logboxout("ERROR: Failed to retrieve Google event");
                log.Error(ex.Message);
                return null;
            }
        }

        public List<Event> GetCalendarEntriesInRange() {
            return GetCalendarEntriesInRange(Settings.Instance.SyncStart, Settings.Instance.SyncEnd);
        }

        public List<Event> GetCalendarEntriesInRange(DateTime from, DateTime to) {
            List<Event> result = new List<Event>();
            Events request = null;
            String pageToken = null;

            do {
                log.Debug("Retrieving all events from Google: "+ from.ToShortDateString() +" -> "+ to.ToShortDateString());
                EventsResource.ListRequest lr = service.Events.List(Settings.Instance.UseGoogleCalendar.Id);

                lr.TimeMin = GoogleTimeFrom(from);
                lr.TimeMax = GoogleTimeFrom(to);
                lr.PageToken = pageToken;
                lr.SingleEvents = false;
                
                request = lr.Fetch();
                pageToken = request.NextPageToken;

                if (request != null) {
                    if (request.Items != null) result.AddRange(request.Items);
                }
            } while (pageToken != null);

            if (Settings.Instance.CreateCSVFiles) {
                log.Debug("Outputting CSV files...");
                TextWriter tw = new StreamWriter(Path.Combine(Program.UserFilePath,"google_events.csv"));
                String CSVheader = "Start Time,Finish Time,Subject,Location,Description,Privacy,FreeBusy,";
                CSVheader += "Required Attendees,Optional Attendees,Reminder Set,Reminder Minutes,Google ID,Outlook ID";
                tw.WriteLine(CSVheader);
                foreach (Event ev in result) {
                    try {
                        tw.WriteLine(exportToCSV(ev));
                    } catch {
                        MainForm.Instance.Logboxout("Failed to output following Google event to CSV:-");
                        MainForm.Instance.Logboxout(GetEventSummary(ev));
                    }
                }
                tw.Close();
                log.Debug("Done.");
            }

            return result;
        }

        #region Create
        public void CreateCalendarEntries(List<AppointmentItem> appointments) {
            foreach (AppointmentItem ai in appointments) {
                Event newEvent = new Event();
                try {
                    newEvent = createCalendarEntry(ai);
                } catch (System.Exception ex) {
                    if (!Settings.Instance.VerboseOutput) MainForm.Instance.Logboxout(OutlookCalendar.GetEventSummary(ai));
                    MainForm.Instance.Logboxout("WARNING: Event creation failed.\r\n" + ex.Message);
                    log.Error(ex.StackTrace);
                    if (MessageBox.Show("Google event creation failed. Continue with synchronisation?", "Sync item failed", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                        continue;
                    else {
                        throw new UserCancelledSyncException("User chose not to continue sync.");
                    }
                }

                Event createdEvent = new Event();
                try {
                    createdEvent = createCalendarEntry_save(newEvent, ai);
                    if (Settings.Instance.AddAttendees && Settings.Instance.APIlimit_inEffect) {
                        log.Info("API limit for attendee sync lifted :-)");
                        Settings.Instance.APIlimit_inEffect = false;
                    }
                } catch (System.Exception ex) {
                    if (handleAPIlimits(ex, newEvent, ai))
                        createdEvent = createCalendarEntry_save(newEvent, ai);
                    else {
                        MainForm.Instance.Logboxout("WARNING: New event failed to save.\r\n" + ex.Message);
                        log.Error(ex.StackTrace);
                        if (MessageBox.Show("New Google event failed to save. Continue with synchronisation?", "Sync item failed", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                            continue;
                        else {
                            throw new UserCancelledSyncException("User chose not to continue sync.");
                        }
                    }
                }

                if (ai.IsRecurring && Recurrence.HasExceptions(ai) && createdEvent != null) {
                    MainForm.Instance.Logboxout("This is a recurring item with some exceptions:-");
                    Recurrence.CreateGoogleExceptions(ai, createdEvent.Id);
                    MainForm.Instance.Logboxout("Recurring exceptions completed.");
                }
            }
        }

        private Event createCalendarEntry(AppointmentItem ai) {
            string itemSummary = OutlookCalendar.GetEventSummary(ai);
            log.Debug("Processing >> " + itemSummary);
            MainForm.Instance.Logboxout(itemSummary, verbose: true);

            Event ev = new Event(); 
            //Add the Outlook appointment ID into Google event
            AddOutlookID(ref ev, ai);
                
            ev.Start = new EventDateTime();
            ev.End = new EventDateTime();

            ev.Recurrence = Recurrence.Instance.BuildGooglePattern(ai, ev);
            if (ev.Recurrence != null) {
                ev = OutlookCalendar.Instance.IOutlook.IANAtimezone_set(ev, ai);
            }
                            
            if (ai.AllDayEvent) {
                ev.Start.Date = ai.Start.ToString("yyyy-MM-dd");
                ev.End.Date = ai.End.ToString("yyyy-MM-dd");
            } else {
                ev.Start.DateTime = GoogleCalendar.GoogleTimeFrom(ai.Start);
                ev.End.DateTime = GoogleCalendar.GoogleTimeFrom(ai.End);
            }
            ev.Summary = Obfuscate.ApplyRegex(ai.Subject, SyncDirection.OutlookToGoogle);
            if (Settings.Instance.AddDescription) ev.Description = ai.Body;
            ev.Location = ai.Location;
            ev.Visibility = (ai.Sensitivity == OlSensitivity.olNormal) ? "default" : "private";
            ev.Transparency = (ai.BusyStatus == OlBusyStatus.olFree) ? "transparent" : "opaque";

            ev.Attendees = new List<EventAttendee>();
            if (Settings.Instance.AddAttendees && ai.Recipients.Count > 1 && !APIlimitReached_attendee) { //Don't add attendees if there's only 1 (me)
                if (ai.Recipients.Count >= 200) {
                    MainForm.Instance.Logboxout("ALERT: Attendees will not be synced for this meeting as it has " +
                        "more than 200, which Google does not allow.");
                } else {
                    foreach (Microsoft.Office.Interop.Outlook.Recipient recipient in ai.Recipients) {
                        EventAttendee ea = GoogleCalendar.CreateAttendee(recipient);
                        ev.Attendees.Add(ea);
                    }
                }
            }

            //Reminder alert
            if (Settings.Instance.AddReminders && ai.ReminderSet) {
                ev.Reminders = new Event.RemindersData();
                ev.Reminders.UseDefault = false;
                EventReminder reminder = new EventReminder();
                reminder.Method = "popup";
                reminder.Minutes = ai.ReminderMinutesBeforeStart;
                ev.Reminders.Overrides = new List<EventReminder>();
                ev.Reminders.Overrides.Add(reminder);
            }
            return ev;
        }

        private Event createCalendarEntry_save(Event ev, AppointmentItem ai) {
            if (Settings.Instance.SyncDirection == SyncDirection.Bidirectional) {
                log.Debug("Saving timestamp when OGCS updated event.");
                addOGCSproperty(ref ev, Program.OGCSmodified, DateTime.Now);
            }
            
            Event createdEvent = service.Events.Insert(ev, Settings.Instance.UseGoogleCalendar.Id).Fetch();
            
            Boolean gKeyExists = false;
            try {
                UserProperty up = ai.UserProperties.Find(OutlookCalendar.gEventID);
                gKeyExists = (up != null);
            } catch { }
            if (Settings.Instance.SyncDirection == SyncDirection.Bidirectional || gKeyExists) {
                log.Debug("Storing the Google event ID in Outlook appointment.");
                OutlookCalendar.AddOGCSproperty(ref ai, OutlookCalendar.gEventID, createdEvent.Id);
                ai.Save();
            }
            //DOS ourself by triggering API limit
            //for (int i = 1; i <= 30; i++) {
            //    MainForm.Instance.Logboxout("Add #" + i, verbose:true);
            //    Event result = service.Events.Insert(e, Settings.Instance.UseGoogleCalendar.Id).Fetch();
            //    System.Threading.Thread.Sleep(300);
            //    GoogleCalendar.Instance.deleteCalendarEntry(result);
            //    System.Threading.Thread.Sleep(300);
            //}
            return createdEvent;
        }
        #endregion

        #region Update
        public void UpdateCalendarEntries(Dictionary<AppointmentItem, Event> entriesToBeCompared, ref int entriesUpdated) {
            entriesUpdated = 0;
            foreach (KeyValuePair<AppointmentItem, Event> compare in entriesToBeCompared) {
                int itemModified = 0;
                Event ev = new Event();
                try {
                    ev = UpdateCalendarEntry(compare.Key, compare.Value, ref itemModified);
                } catch (System.Exception ex) {
                    if (!Settings.Instance.VerboseOutput) MainForm.Instance.Logboxout(OutlookCalendar.GetEventSummary(compare.Key));
                    MainForm.Instance.Logboxout("WARNING: Event update failed.\r\n" + ex.Message);
                    log.Error(ex.StackTrace);
                    if (MessageBox.Show("Google event update failed. Continue with synchronisation?", "Sync item failed", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                        continue;
                    else {
                        throw new UserCancelledSyncException("User chose not to continue sync.");
                    }
                }

                if (itemModified > 0) {
                    try {
                        UpdateCalendarEntry_save(ev);
                        entriesUpdated++;
                        if (Settings.Instance.AddAttendees && Settings.Instance.APIlimit_inEffect) {
                            log.Info("API limit for attendee sync lifted :-)");
                            Settings.Instance.APIlimit_inEffect = false;
                        }
                    } catch (System.Exception ex) {
                        if (handleAPIlimits(ex, ev, compare.Key)) {
                            UpdateCalendarEntry_save(ev);
                            entriesUpdated++;
                        } else {
                            MainForm.Instance.Logboxout("WARNING: Updated event failed to save.\r\n" + ex.Message);
                            log.Error(ex.StackTrace);
                            if (MessageBox.Show("Updated Google event failed to save. Continue with synchronisation?", "Sync item failed", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                                continue;
                            else {
                                throw new UserCancelledSyncException("User chose not to continue sync.");
                            }
                        }
                    }
                } else if (ev != null) {
                    log.Debug("Doing a dummy update in order to update the last modified date.");
                    addOGCSproperty(ref ev, Program.OGCSmodified, DateTime.Now);
                    UpdateCalendarEntry_save(ev);
                }

                Recurrence.UpdateGoogleExceptions(compare.Key, ev ?? compare.Value);  
            }
        }

        public Event UpdateCalendarEntry(AppointmentItem ai, Event ev, ref int itemModified, Boolean forceCompare = false) {
            if (!Settings.Instance.APIlimit_inEffect) {
                //if (!forceCompare) { //Compare/process recurring items
                    if (Settings.Instance.SyncDirection != SyncDirection.Bidirectional) {
                        if (DateTime.Parse(ev.Updated) > DateTime.Parse(GoogleCalendar.GoogleTimeFrom(ai.LastModificationTime)))
                            return null;
                    } else {
                        if (OutlookCalendar.OGCSlastModified(ai).AddSeconds(5) >= ai.LastModificationTime)
                            //Outlook last modified by OGCS
                            return null;
                        if (DateTime.Parse(ev.Updated) > DateTime.Parse(GoogleCalendar.GoogleTimeFrom(ai.LastModificationTime)))
                            return null;
                    }
                //}
            } else { //Settings.Instance.APIlimit_inEffect
                if (!(OGCSlastModified(ev).AddSeconds(5) < Settings.Instance.APIlimit_lastHit &&
                    DateTime.Parse(ev.Updated) > Settings.Instance.APIlimit_lastHit))
                    //Not been synced without attendees
                    return null;
            }
                
            String aiSummary = OutlookCalendar.GetEventSummary(ai);
            log.Debug("Processing >> " + aiSummary);

            System.Text.StringBuilder sb = new System.Text.StringBuilder();
            sb.AppendLine(aiSummary);
                
            //Handle an event's all-day attribute being toggled
            String evStart = (ev.Start.DateTime == null) ? ev.Start.Date : ev.Start.DateTime;
            String evEnd = (ev.End.DateTime == null) ? ev.End.Date : ev.End.DateTime;
            if (ai.AllDayEvent) {
                ev.Start.DateTime = null;
                ev.End.DateTime = null;
                if (MainForm.CompareAttribute("Start time", SyncDirection.OutlookToGoogle, evStart, ai.Start.ToString("yyyy-MM-dd"), sb, ref itemModified)) {
                    ev.Start.Date = ai.Start.ToString("yyyy-MM-dd");
                }
                if (MainForm.CompareAttribute("End time", SyncDirection.OutlookToGoogle, evEnd, ai.End.ToString("yyyy-MM-dd"), sb, ref itemModified)) {
                    ev.End.Date = ai.End.ToString("yyyy-MM-dd");
                }
            } else {
                ev.Start.Date = null;
                ev.End.Date = null;
                if (MainForm.CompareAttribute("Start time", SyncDirection.OutlookToGoogle,
                    GoogleCalendar.GoogleTimeFrom(DateTime.Parse(evStart)), GoogleCalendar.GoogleTimeFrom(ai.Start), sb, ref itemModified)) {
                    ev.Start.DateTime = GoogleCalendar.GoogleTimeFrom(ai.Start);
                }
                if (MainForm.CompareAttribute("End time", SyncDirection.OutlookToGoogle, 
                    GoogleCalendar.GoogleTimeFrom(DateTime.Parse(evEnd)), GoogleCalendar.GoogleTimeFrom(ai.End), sb, ref itemModified)) {
                    ev.End.DateTime = GoogleCalendar.GoogleTimeFrom(ai.End);
                }
            }

            List<String> oRrules = Recurrence.Instance.BuildGooglePattern(ai, ev);
            if (ev.Recurrence != null) {
                for (int r = 0; r < ev.Recurrence.Count; r++) {
                    String rrule = ev.Recurrence[r];
                    if (rrule.StartsWith("RRULE:")) {
                        if (oRrules != null) {
                            String[] oRrule_bits = oRrules.First().TrimStart("RRULE:".ToCharArray()).Split(';');
                            foreach (String oRrule_bit in oRrule_bits) {
                                if (!rrule.Contains(oRrule_bit)) {
                                    if (MainForm.CompareAttribute("Recurrence", SyncDirection.OutlookToGoogle, rrule, oRrules.First(), sb, ref itemModified)) {
                                        ev.Recurrence[r] = oRrules.First();
                                    }
                                }
                            }
                        } else {
                            log.Debug("Converting to non-recurring event.");
                            MainForm.CompareAttribute("Recurrence", SyncDirection.OutlookToGoogle, rrule, null, sb, ref itemModified);
                            ev.Recurrence[r] = null;
                        }
                        break;
                    }
                }
            } else {
                if (oRrules != null && ev.RecurringEventId == null) {
                    log.Debug("Converting to recurring event.");
                    MainForm.CompareAttribute("Recurrence", SyncDirection.OutlookToGoogle, null, oRrules.First(), sb, ref itemModified);
                    ev.Recurrence = oRrules;
                }
            }
            if (ev.Recurrence != null && ev.RecurringEventId == null) {
                ev = OutlookCalendar.Instance.IOutlook.IANAtimezone_set(ev, ai);
            }
            
            String subjectObfuscated = Obfuscate.ApplyRegex(ai.Subject, SyncDirection.OutlookToGoogle);
            if (MainForm.CompareAttribute("Subject", SyncDirection.OutlookToGoogle, ev.Summary, subjectObfuscated, sb, ref itemModified)) {
                ev.Summary = subjectObfuscated;
            }
            if (!Settings.Instance.AddDescription) ai.Body = "";
            String outlookBody = ai.Body;
            //Check for Google description truncated @ 8Kb
            if (!string.IsNullOrEmpty(ai.Body) && !string.IsNullOrEmpty(ev.Description)
                && ev.Description.Length == 8 * 1024  
                && ai.Body.Length > 8 * 1024) {
                outlookBody = ai.Body.Substring(0, 8 * 1024);
            }
            if (MainForm.CompareAttribute("Description", SyncDirection.OutlookToGoogle, ev.Description, outlookBody, sb, ref itemModified)) 
                ev.Description = outlookBody;
            
            if (MainForm.CompareAttribute("Location", SyncDirection.OutlookToGoogle, ev.Location, ai.Location, sb, ref itemModified)) 
                ev.Location = ai.Location;

            String oPrivacy = (ai.Sensitivity == OlSensitivity.olNormal) ? "default" : "private";
            String gPrivacy = (ev.Visibility == null ? "default" : ev.Visibility);
            if (MainForm.CompareAttribute("Private", SyncDirection.OutlookToGoogle, gPrivacy, oPrivacy, sb, ref itemModified)) {
                ev.Visibility = oPrivacy;
            }
            String oFreeBusy = (ai.BusyStatus == OlBusyStatus.olFree) ? "transparent" : "opaque";
            String gFreeBusy = (ev.Transparency == null ? "opaque" : ev.Transparency);
            if (MainForm.CompareAttribute("Free/Busy", SyncDirection.OutlookToGoogle, gFreeBusy, oFreeBusy, sb, ref itemModified)) {
                ev.Transparency = oFreeBusy;
            }

            if (Settings.Instance.AddAttendees && ai.Recipients.Count > 1 && !APIlimitReached_attendee) {
                if (ai.Recipients.Count >= 200) {
                    MainForm.Instance.Logboxout("ALERT: Attendees will not be synced for this meeting as it has " +
                        "more than 200, which Google does not allow.");
                    ev.Attendees = new List<EventAttendee>();
                } else {
                    try {
                        CompareRecipientsToAttendees(ai, ev, sb, ref itemModified);
                    } catch (System.Exception ex) {
                        if (ex.Message.Contains("An error occurred while performing the operation") &&
                            OutlookCalendar.Instance.IOutlook.ExchangeConnectionMode().ToString().Contains("Disconnected")) {
                            MainForm.Instance.Logboxout("Outlook is currently disconnected from Exchange, so it's not possible to sync attendees.");
                            MainForm.Instance.Logboxout("Please reconnect or do not sync attendees.");
                            throw new System.Exception("Outlook has disconnected from Exchange.");
                        } else {
                            MainForm.Instance.Logboxout("WARNING: Unable to sync attendees.\n" + ex.Message);
                        }
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
                                if (MainForm.CompareAttribute("Reminder", SyncDirection.OutlookToGoogle, reminder.Minutes.ToString(), ai.ReminderMinutesBeforeStart.ToString(), sb, ref itemModified)) {
                                    reminder.Minutes = ai.ReminderMinutesBeforeStart;
                                }
                            } else {
                                sb.AppendLine("Reminder: " + reminder.Minutes + " => removed");
                                ev.Reminders.Overrides.Remove(reminder);
                                if (ev.Reminders.Overrides == null || ev.Reminders.Overrides.Count == 0) {
                                    ev.Reminders.UseDefault = true;
                                }
                                itemModified++;
                            } //if Outlook reminders set
                        } //if google reminder found
                    } //foreach reminder

                } else { //no google reminders set
                    if (ai.ReminderSet) {
                        sb.AppendLine("Reminder: nothing => " + ai.ReminderMinutesBeforeStart);
                        ev.Reminders.UseDefault = false;
                        EventReminder newReminder = new EventReminder();
                        newReminder.Method = "popup";
                        newReminder.Minutes = ai.ReminderMinutesBeforeStart;
                        ev.Reminders.Overrides = new List<EventReminder>();
                        ev.Reminders.Overrides.Add(newReminder);
                        itemModified++;
                    }
                }
            }
            if (itemModified > 0) {
                MainForm.Instance.Logboxout(sb.ToString(), false, verbose: true);
                MainForm.Instance.Logboxout(itemModified + " attributes updated.", verbose: true);
                System.Windows.Forms.Application.DoEvents();
            }
            return ev;
        }

        public void UpdateCalendarEntry_save(Event ev) {
            if (Settings.Instance.SyncDirection == SyncDirection.Bidirectional) {
                log.Debug("Saving timestamp when OGCS updated event.");
                addOGCSproperty(ref ev, Program.OGCSmodified, DateTime.Now);
            } 
            var request = service.Events.Update(ev, Settings.Instance.UseGoogleCalendar.Id, ev.Id).Fetch();
        }
        #endregion

        #region Delete
        public void DeleteCalendarEntries(List<Event> events) {
            for (int g=events.Count-1; g>=0; g--) {
                Event ev = events[g];
                Boolean doDelete = false;
                try {
                    doDelete = deleteCalendarEntry(ev);
                } catch (System.Exception ex) {
                    if (!Settings.Instance.VerboseOutput) MainForm.Instance.Logboxout(GoogleCalendar.GetEventSummary(ev));
                    MainForm.Instance.Logboxout("WARNING: Event deletion failed.\r\n" + ex.Message);
                    log.Error(ex.StackTrace);
                    if (MessageBox.Show("Google event deletion failed. Continue with synchronisation?", "Sync item failed", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                        continue;
                    else {
                        throw new UserCancelledSyncException("User chose not to continue sync.");
                    }
                }

                try {
                    if (doDelete) deleteCalendarEntry_save(ev);
                    else events.Remove(ev);
                } catch (System.Exception ex) {
                    MainForm.Instance.Logboxout("WARNING: Deleted event failed to remove.\r\n" + ex.Message);
                    log.Error(ex.StackTrace);
                    if (MessageBox.Show("Deleted Google event failed to remove. Continue with synchronisation?", "Sync item failed", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                        continue;
                    else {
                        throw new UserCancelledSyncException("User chose not to continue sync.");
                    }
                }
            }
        }
        
        private Boolean deleteCalendarEntry(Event ev) {
            String eventSummary = GetEventSummary(ev);
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
                MainForm.Instance.Logboxout(eventSummary, verbose:true);
            }
            return doDelete;
        }

        private void deleteCalendarEntry_save(Event e) {
            string request = service.Events.Delete(Settings.Instance.UseGoogleCalendar.Id, e.Id).Fetch();
        }
        #endregion

        public void ReclaimOrphanCalendarEntries(ref List<Event> gEvents, ref List<AppointmentItem> oAppointments, Boolean neverDelete = false) {
            log.Debug("Looking for orphaned events to reclaim...");

            //This is needed for people migrating from other tools, which do not have our OutlookID extendedProperty
            List<Event> unclaimedEvents = new List<Event>();

            for (int g = gEvents.Count-1; g >= 0; g--) {
                Event ev = gEvents[g];
                    
                //Find entries with no Outlook ID
                if (!GetOGCSproperty(ev, oEntryID)) {
                    unclaimedEvents.Add(ev);
                    foreach (AppointmentItem ai in oAppointments) {
                        //Use simple matching on start,end,subject,location to pair events
                        String sigEv = signature(ev);
                        String sigAi = OutlookCalendar.signature(ai);

                        if (Settings.Instance.Obfuscation.Enabled) {
                            if (Settings.Instance.Obfuscation.Direction == SyncDirection.OutlookToGoogle)
                                sigAi = Obfuscate.ApplyRegex(sigAi, SyncDirection.OutlookToGoogle);
                            else
                                sigEv = Obfuscate.ApplyRegex(sigEv, SyncDirection.GoogleToOutlook);
                        }
                        if (sigEv == sigAi) {                            
                            AddOutlookID(ref ev, ai);
                            UpdateCalendarEntry_save(ev);
                            unclaimedEvents.Remove(ev);
                            MainForm.Instance.Logboxout("Reclaimed: " + GetEventSummary(ev), verbose: true);
                            break;
                        }
                    }
                }
            }
            log.Debug(unclaimedEvents.Count + " unclaimed.");
            if (!neverDelete && unclaimedEvents.Count > 0 &&
                (Settings.Instance.SyncDirection == SyncDirection.OutlookToGoogle ||
                 Settings.Instance.SyncDirection == SyncDirection.Bidirectional )) 
            {
                log.Info(unclaimedEvents.Count +" unclaimed orphan events found.");
                if (Settings.Instance.MergeItems || Settings.Instance.DisableDelete || Settings.Instance.ConfirmOnDelete) {
                    log.Info("These will be kept due to configuration settings.");
                } else {
                    if (MessageBox.Show(unclaimedEvents.Count + " Google calendar events can't be matched to Outlook.\r\n" +
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
        }

        //<summary>New logic for comparing Outlook and Google events works as follows:
        //      1.  Scan through both lists looking for duplicates
        //      2.  Remove found duplicates from both lists
        //      3.  Items remaining in Outlook list are new and need to be created
        //      4.  Items remaining in Google list need to be deleted
        //</summary>
        public void IdentifyEventDifferences(
            ref List<AppointmentItem> outlook,  //need creating
            ref List<Event> google,             //need deleting
            Dictionary<AppointmentItem, Event> compare) {
            log.Debug("Comparing Outlook items to Google events...");

            // Count backwards so that we can remove found items without affecting the order of remaining items
            String compare_gEntryID;
            int migrated = 0;
            for (int g = google.Count - 1; g >= 0; g--) {
                if (GetOGCSproperty(google[g], oEntryID, out compare_gEntryID)) {
                    for (int o = outlook.Count - 1; o >= 0; o--) {
                        if (outlook[o].EntryID == compare_gEntryID) {
                            if (migrated == 0) log.Info("Migrating Events from EntryID to GlobalAppointmentID...");
                            //Should have used AppointmentItem Global ID, not Object ID (which changes when accepting invites!)
                            //To prevent existing users from having everything deleted and recreated, need to migrate them.
                            //This change was effective from v2.0.2.3
                            Event ev = google[g];
                            AddOutlookID(ref ev, outlook[o]);
                            UpdateCalendarEntry_save(ev);
                            google[g] = ev;
                            compare_gEntryID = outlook[o].GlobalAppointmentID;
                            migrated++;
                            //There could be a lot of these, so let's not batter Google too hard - it doesn't like >4/sec
                            System.Threading.Thread.Sleep(250);
                        }
                        if (outlook[o].GlobalAppointmentID == compare_gEntryID) {
                            compare.Add(outlook[o], google[g]);
                            outlook.Remove(outlook[o]);
                            google.Remove(google[g]);
                            break;
                        }
                    }
                } else if (Settings.Instance.MergeItems) {
                    //Remove the non-Outlook item so it doesn't get deleted
                    google.Remove(google[g]);
                }
            }
            if (migrated > 0) log.Info(migrated + " migrated.");

            if (Settings.Instance.DisableDelete) {
                google = new List<Event>();
            }
            if (Settings.Instance.SyncDirection == SyncDirection.Bidirectional) {
                //Don't recreate any items that have been deleted in Google
                for (int o = outlook.Count - 1; o >= 0; o--) {
                    if (outlook[o].UserProperties[OutlookCalendar.gEventID] != null)
                        outlook.Remove(outlook[o]);
                }
                //Don't delete any items that aren't yet in Outlook or just created in Outlook during this sync
                for (int g = google.Count - 1; g >= 0; g--) {
                    if (!GetOGCSproperty(google[g], oEntryID) ||
                        DateTime.Parse(google[g].Updated) > Settings.Instance.LastSyncDate)
                        google.Remove(google[g]);
                }
            }
            if (Settings.Instance.CreateCSVFiles) {
                //Google Deletions
                log.Debug("Outputting items for deletion to CSV...");
                TextWriter tw = new StreamWriter(Path.Combine(Program.UserFilePath, "google_delete.csv"));
                foreach (Event ev in google) {
                    tw.WriteLine(signature(ev));
                }
                tw.Close();

                //Google Creations
                log.Debug("Outputting items for creation to CSV...");
                tw = new StreamWriter(Path.Combine(Program.UserFilePath, "google_create.csv"));
                foreach (AppointmentItem ai in outlook) {
                    tw.WriteLine(OutlookCalendar.signature(ai));
                }
                tw.Close();
                log.Debug("Done.");
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
                sb.AppendLine("Attendee removed: " + (ea.DisplayName == null ? ea.Email : ea.DisplayName));
                ev.Attendees.Remove(ea);
                itemModified++;
            }
            return (itemModified > 0);
        }

        public void GetSetting(string setting) {
            try {
                SettingsResource.GetRequest sr = service.Settings.Get(setting);
                sr.FetchAsync();
            } catch { }
        }

        #region STATIC FUNCTIONS
        private static IAuthorizationState getAuthentication(NativeApplicationClient arg) {
            log.Debug("Authenticating with Google calendar service...");
            // Get the auth URL:
            IAuthorizationState state = new AuthorizationState(new[] { CalendarService.Scopes.Calendar.GetStringValue() });
            state.Callback = new Uri(NativeApplicationClient.OutOfBandCallbackUrl);
            state.RefreshToken = Settings.Instance.RefreshToken;
            Uri authUri = arg.RequestUserAuthorization(state);

            IAuthorizationState result = null;

            if (state.RefreshToken == "") {
                log.Info("No refresh token available - need user authorisation.");

                // Request authorization from the user (by opening a browser window):
                Process.Start(authUri.ToString());

                frmGoogleAuthorizationCode eac = new frmGoogleAuthorizationCode();
                if (eac.ShowDialog() == DialogResult.OK) {
                    if (string.IsNullOrEmpty(eac.authcode))
                        log.Debug("User continued but did not provide a code! This isn't going to work...");
                    else
                        log.Debug("User has provided authentication code.");

                    // Retrieve the access/refresh tokens by using the authorization code:
                    result = arg.ProcessUserAuthorization(eac.authcode, state);

                    //save the refresh token for future use
                    Settings.Instance.RefreshToken = result.RefreshToken;
                    Settings.Instance.Save();
                    log.Info("Refresh and Access token successfully retrieved.");

                    return result;
                } else {
                    log.Info("User declined to provide authorisation code. Sync will not be able to work.");
                    String noAuth = "Sorry, but this application will not work if you don't give it access to your Google Calendar :(";
                    MessageBox.Show(noAuth, "Authorisation not given", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    throw new System.ApplicationException(noAuth);
                }
            } else {
                arg.RefreshToken(state, null);
                if (string.IsNullOrEmpty(state.AccessToken))
                    log.Error("Failed to retrieve Access token.");
                else
                    log.Debug("Access token refreshed - expires " + ((DateTime)state.AccessTokenExpirationUtc).ToLocalTime().ToString());
                result = state;
                return result;
            }

        }

        //returns the Google Time Format String of a given .Net DateTime value
        //Google Time Format = "2012-08-20T00:00:00+02:00"
        public static string GoogleTimeFrom(DateTime dt) {
            return dt.ToString("yyyy-MM-ddTHH:mm:sszzz");
        }
        
        public static string signature(Event ev) {
            String signature = "";
            if (ev.RecurringEventId != null && ev.Status == "cancelled" && ev.OriginalStartTime != null) {
                signature += GoogleTimeFrom(DateTime.Parse(ev.OriginalStartTime.Date ?? ev.OriginalStartTime.DateTime));
                signature += ";" + (ev.Summary ?? "[cancelled]");
            } else {
                signature += GoogleTimeFrom(DateTime.Parse(ev.Start.Date ?? ev.Start.DateTime));
                signature += ";" + GoogleTimeFrom(DateTime.Parse(ev.End.Date ?? ev.End.DateTime));
                signature += ";" + ev.Summary;
            }
            return signature.Trim();
        }

        private static string exportToCSV(Event ev) {
            System.Text.StringBuilder csv = new System.Text.StringBuilder();

            if (ev.Start.Date == null) {
                csv.Append(ev.Start.DateTime + ",");
            } else {
                csv.Append(ev.Start.Date + ",");
            }
            if (ev.End.Date == null) {
                csv.Append(ev.End.DateTime + ",");
            } else {
                csv.Append(ev.End.Date + ",");
            }
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
                foreach (EventAttendee ea in ev.Attendees) {
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
                        csv.Append("true," + er.Minutes +",");
                        foundReminder = true;
                    }
                    break;
                }
            }
            if (!foundReminder) csv.Append(",,");

            csv.Append(ev.Id + ",");
            String gEntryID;
            if (GetOGCSproperty(ev, oEntryID, out gEntryID)) {
                csv.Append(gEntryID);
            }

            return csv.ToString();
        }

        public static string GetEventSummary(Event ev) {
            String eventSummary = "";
            if (ev.Start.DateTime != null) {
                DateTime gDate = DateTime.Parse(ev.Start.DateTime);
                eventSummary += gDate.ToShortDateString() + " " + gDate.ToShortTimeString();
            } else
                eventSummary += DateTime.Parse(ev.Start.Date).ToShortDateString();
            if ((ev.Recurrence != null && ev.RecurringEventId == null) || ev.RecurringEventId != null)
                eventSummary += " (R)";
            eventSummary += " => \"" + ev.Summary + "\"";
            return eventSummary;
        }

        public static EventAttendee CreateAttendee(Recipient recipient) {
            EventAttendee ea = new EventAttendee();
            log.Fine("Creating attendee " + recipient.Name);
            ea.DisplayName = recipient.Name;
            ea.Email = OutlookCalendar.Instance.IOutlook.GetRecipientEmail(recipient);
            ea.Optional = (recipient.Type == (int)OlMeetingRecipientType.olOptional);
            //Readonly: ea.Organizer = (ai.Organizer == recipient.Name);
            switch (recipient.MeetingResponseStatus) {
                case OlResponseStatus.olResponseNone: ea.ResponseStatus = "needsAction"; break;
                case OlResponseStatus.olResponseAccepted: ea.ResponseStatus = "accepted"; break;
                case OlResponseStatus.olResponseDeclined: ea.ResponseStatus = "declined"; break;
                case OlResponseStatus.olResponseTentative: ea.ResponseStatus = "tentative"; break;
            }
            return ea;
        }

        private static Boolean handleAPIlimits(System.Exception ex, Event ev, AppointmentItem ai) {
            //https://developers.google.com/analytics/devguides/reporting/core/v3/coreErrors

            if (Settings.Instance.AddAttendees && ex.Message.Contains("Calendar usage limits exceeded. [403]")) {
                //"Google.Apis.Requests.RequestError\r\nCalendar usage limits exceeded. [403]\r\nErrors [\r\n\tMessage[Calendar usage limits exceeded.] Location[ - ] Reason[quotaExceeded] Domain[usageLimits]\r\n]\r\n"
                //This happens because too many attendees have been added in a short period of time.
                //See https://support.google.com/a/answer/2905486?hl=en-uk&hlrm=en

                MainForm.Instance.Logboxout("ALERT: You have added enough meeting attendees to have reached the Google API limit.");
                MainForm.Instance.Logboxout("Don't worry, this only lasts for an hour or two, but until then attendees will not be synced.");
                
                APIlimitReached_attendee = true;
                Settings.Instance.APIlimit_inEffect = true;
                Settings.Instance.APIlimit_lastHit = ai.LastModificationTime;

                ev.Attendees = new List<EventAttendee>();
                return true;

            } else if (ex.Message.Contains("Rate Limit Exceeded")) {
                log.Debug("Exceeded API call rate/second/user - pausing a second...");
                System.Threading.Thread.Sleep(1000);
                return true;

            } else {
                return false;
            }
        }

        #region OGCS event properties
        public static void AddOutlookID(ref Event ev, AppointmentItem ai) {
            //Add the Outlook appointment ID into Google event.
            //This will make comparison more efficient and set the scene for 2-way sync.
            
            addOGCSproperty(ref ev, oEntryID, ai.GlobalAppointmentID);
        }

        private static void addOGCSproperty(ref Event ev, String key, String value) {
            if (ev.ExtendedProperties == null) ev.ExtendedProperties = new Event.ExtendedPropertiesData();
            if (ev.ExtendedProperties.Private == null) ev.ExtendedProperties.Private = new Event.ExtendedPropertiesData.PrivateData();
            if (ev.ExtendedProperties.Private.ContainsKey(key))
                ev.ExtendedProperties.Private[key] = value;
            else
                ev.ExtendedProperties.Private.Add(key, value);
        }

        private static void addOGCSproperty(ref Event ev, String key, DateTime value) {
            if (key == Program.OGCSmodified && Settings.Instance.APIlimit_inEffect) return;
            addOGCSproperty(ref ev, key, value.ToString());
        }

        public static Boolean GetOGCSproperty(Event ev, String key) {
            String throwAway;
            return GoogleCalendar.GetOGCSproperty(ev, key, out throwAway);
        }
        public static Boolean GetOGCSproperty(Event ev, String key, out String value) {
            if (ev.ExtendedProperties != null &&
                ev.ExtendedProperties.Private != null &&
                ev.ExtendedProperties.Private.ContainsKey(key)) {
                value = ev.ExtendedProperties.Private[key];
                return true;
            } else {
                value = null;
                return false;
            }
        }

        public static DateTime OGCSlastModified(Event ev) {
            String lastModded = null;
            if (GetOGCSproperty(ev, Program.OGCSmodified, out lastModded)) {
                return DateTime.Parse(lastModded);
            } else {
                return new DateTime();
            }
        }
        #endregion
        #endregion
    }
}
