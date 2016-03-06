using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Google.Apis.Calendar.v3.Data;
using log4net;
using Microsoft.Office.Interop.Outlook;
using NodaTime;
using System.Windows.Forms;

namespace OutlookGoogleCalendarSync {
    class Recurrence {
        /*
         * Recurrence rule standards for iCalendar: http://www.ietf.org/rfc/rfc2445 
         */
        private static readonly ILog log = LogManager.GetLogger(typeof(Recurrence));

        private static Recurrence instance;
        public static Recurrence Instance {
            get {
                if (instance == null) instance = new Recurrence();
                return instance;
            }
        }

        public Recurrence() { }

        #region iCalendar Functions
        private Dictionary<String, String> rrule = new Dictionary<String, String>();
        public List<String> BuildGooglePattern(AppointmentItem ai, Event ev) {
            if (!ai.IsRecurring || ai.RecurrenceState != OlRecurrenceState.olApptMaster) return null;

            log.Debug("Creating Google iCalendar definition for recurring event.");
            List<String> gPattern = new List<String>();
            gPattern.Add("RRULE:" + buildRrule(ai.GetRecurrencePattern()));

            log.Debug(string.Join("\r\n", gPattern.ToArray()));
            return gPattern;
        }

        public void BuildOutlookPattern(Event ev, AppointmentItem ai) {
            RecurrencePattern ignore;
            BuildOutlookPattern(ev, ai, out ignore);
            ignore = (RecurrencePattern)OutlookCalendar.ReleaseObject(ignore);
        }

        public void BuildOutlookPattern(Event ev, AppointmentItem ai, out RecurrencePattern oPattern) {
            if (ev.Recurrence == null) { oPattern = null; return; }

            Dictionary<String, String> ruleBook = explodeRrule(ev.Recurrence);
            log.Fine("Building Outlook recurrence pattern");
            oPattern = ai.GetRecurrencePattern();
            #region RECURRENCE PATTERN
            //RRULE:FREQ=WEEKLY;UNTIL=20150906T000000Z;BYDAY=SA

            switch (ruleBook["FREQ"]) {
                case "DAILY": {
                        oPattern.RecurrenceType = OlRecurrenceType.olRecursDaily;
                        break;
                    }
                case "WEEKLY": {
                        oPattern.RecurrenceType = OlRecurrenceType.olRecursWeekly;
                        // Need to work out dayMask from "BY" pattern
                        // Eg "BYDAY=MO,TU,WE,TH,FR"
                        OlDaysOfWeek dowMask = getDOWmask(ruleBook);
                        if (dowMask != 0) {
                            oPattern.DayOfWeekMask = dowMask;
                        }
                        break;
                    }
                case "MONTHLY": {
                        oPattern.RecurrenceType = OlRecurrenceType.olRecursMonthly;
                        if (ruleBook.ContainsKey("BYSETPOS")) {
                            oPattern.RecurrenceType = OlRecurrenceType.olRecursMonthNth;
                            int gInstance = Convert.ToInt16(ruleBook["BYSETPOS"]);
                            oPattern.Instance = (gInstance == -1) ? 5 : gInstance;
                            oPattern.DayOfWeekMask = getDOWmask(ruleBook);
                            if (oPattern.DayOfWeekMask == (OlDaysOfWeek)127 && gInstance == -1 &&
                                DateTime.Parse(ev.Start.DateTime ?? ev.Start.Date).Day > 28) {
                                //In Outlook this is simply a monthly recurring
                                oPattern.RecurrenceType = OlRecurrenceType.olRecursMonthly;
                                oPattern.DayOfWeekMask = 0;
                            }
                        }
                        if (ruleBook.ContainsKey("BYDAY")) {
                            if (ruleBook["BYDAY"].StartsWith("-1")) {
                                oPattern.RecurrenceType = OlRecurrenceType.olRecursMonthNth;
                                oPattern.Instance = 5;
                                oPattern.DayOfWeekMask = getDOWmask(ruleBook["BYDAY"].TrimStart("-1".ToCharArray()));
                            } else if ("1,2,3,4".Contains(ruleBook["BYDAY"].Substring(0, 1))) {
                                oPattern.RecurrenceType = OlRecurrenceType.olRecursMonthNth;
                                oPattern.Instance = Convert.ToInt16(ruleBook["BYDAY"].Substring(0, 1));
                                oPattern.DayOfWeekMask = getDOWmask(ruleBook["BYDAY"].TrimStart(oPattern.Instance.ToString().ToCharArray()));
                            }
                        }
                        break;
                    }

                case "YEARLY": {
                        oPattern.RecurrenceType = OlRecurrenceType.olRecursYearly;
                        //Google interval is years, Outlook is months
                        if (ruleBook.ContainsKey("INTERVAL"))
                            oPattern.Interval = Convert.ToInt16(ruleBook["INTERVAL"]) * 12;
                        if (ruleBook.ContainsKey("BYSETPOS")) {
                            oPattern.RecurrenceType = OlRecurrenceType.olRecursYearNth;
                            int gInstance = Convert.ToInt16(ruleBook["BYSETPOS"]);
                            oPattern.Instance = (gInstance == -1) ? 5 : gInstance;

                            oPattern.DayOfWeekMask = getDOWmask(ruleBook);
                            if (ruleBook.ContainsKey("BYMONTH")) {
                                oPattern.MonthOfYear = Convert.ToInt16(ruleBook["BYMONTH"]);
                            }
                        }
                        break;
                    }
            }
            #endregion

            #region RANGE
            if (ruleBook.ContainsKey("INTERVAL") && Convert.ToInt16(ruleBook["INTERVAL"]) > 1 && ruleBook["FREQ"] != "YEARLY")
                oPattern.Interval = Convert.ToInt16(ruleBook["INTERVAL"]);
            if (ruleBook.ContainsKey("COUNT"))
                oPattern.Occurrences = Convert.ToInt16(ruleBook["COUNT"]);
            if (ruleBook.ContainsKey("UNTIL")) {
                if (ruleBook["UNTIL"].Length == 8) {
                    oPattern.PatternEndDate = DateTime.ParseExact(ruleBook["UNTIL"], "yyyyMMdd", System.Globalization.CultureInfo.InvariantCulture);
                } else {
                    oPattern.PatternEndDate = DateTime.ParseExact(ruleBook["UNTIL"], "yyyyMMddTHHmmssZ", System.Globalization.CultureInfo.InvariantCulture);
                }
            }
            #endregion

            ai = OutlookCalendar.Instance.IOutlook.WindowsTimeZone_set(ai, ev);
        }

        public void CompareOutlookPattern(Event ev, AppointmentItem ai, System.Text.StringBuilder sb, ref int itemModified) {
            if (ev.Recurrence == null) return;

            log.Fine("Building a temporary recurrent Appointment generated from Event");
            AppointmentItem evAI = OutlookCalendar.Instance.IOutlook.UseOutlookCalendar().Items.Add() as AppointmentItem;
            evAI.Start = DateTime.Parse(ev.Start.Date ?? ev.Start.DateTime); 
            RecurrencePattern evOpattern;
            RecurrencePattern aiOpattern = ai.GetRecurrencePattern();
            BuildOutlookPattern(ev, evAI, out evOpattern);
            log.Fine("Comparing Google recurrence to Outlook equivalent");

            if (MainForm.CompareAttribute("Recurrence Type", Settings.Instance.SyncDirection,
                evOpattern.RecurrenceType.ToString(), aiOpattern.RecurrenceType.ToString(), sb, ref itemModified)) {
                aiOpattern.RecurrenceType = evOpattern.RecurrenceType;
            }
            if (MainForm.CompareAttribute("Recurrence occurences", Settings.Instance.SyncDirection,
                evOpattern.Occurrences.ToString(), aiOpattern.Occurrences.ToString(), sb, ref itemModified)) {
                aiOpattern.Occurrences = evOpattern.Occurrences;
            }
            if (MainForm.CompareAttribute("Recurrence Interval", Settings.Instance.SyncDirection,
                evOpattern.Interval.ToString(), aiOpattern.Interval.ToString(), sb, ref itemModified)) {
                aiOpattern.Interval = evOpattern.Interval;
            }
            if (MainForm.CompareAttribute("Recurrence Instance", Settings.Instance.SyncDirection,
                evOpattern.Instance.ToString(), aiOpattern.Instance.ToString(), sb, ref itemModified)) {
                aiOpattern.Instance= evOpattern.Instance;
            }
            if (MainForm.CompareAttribute("Recurrence DoW", Settings.Instance.SyncDirection,
                evOpattern.DayOfWeekMask.ToString(), aiOpattern.DayOfWeekMask.ToString(), sb, ref itemModified)) {
                aiOpattern.DayOfWeekMask = evOpattern.DayOfWeekMask;
            }
            if (MainForm.CompareAttribute("Recurrence MoY", Settings.Instance.SyncDirection,
                evOpattern.MonthOfYear.ToString(), aiOpattern.MonthOfYear.ToString(), sb, ref itemModified)) {
                aiOpattern.MonthOfYear = evOpattern.MonthOfYear;
            }
            aiOpattern = (RecurrencePattern)OutlookCalendar.ReleaseObject(aiOpattern);
            evOpattern = (RecurrencePattern)OutlookCalendar.ReleaseObject(evOpattern);
            evAI = (AppointmentItem)OutlookCalendar.ReleaseObject(evAI);
        }

        private String buildRrule(RecurrencePattern oPattern) {
            log.Fine("Building RRULE");
            rrule = new Dictionary<String, String>();
            #region RECURRENCE PATTERN
            log.Fine("Determining pattern for frequency " + oPattern.RecurrenceType.ToString() + ".");

            switch (oPattern.RecurrenceType) {
                case OlRecurrenceType.olRecursDaily: {
                        addRule(rrule, "FREQ", "DAILY");
                        setInterval(oPattern.Interval);
                        break;
                    }

                case OlRecurrenceType.olRecursWeekly: {
                        addRule(rrule, "FREQ", "WEEKLY");
                        setInterval(oPattern.Interval);
                        if ((oPattern.DayOfWeekMask & (oPattern.DayOfWeekMask-1)) != 0) { //is not a power of 2 (i.e. not just a single day) 
                            // Need to work out "BY" pattern
                            // Eg "BYDAY=MO,TU,WE,TH,FR"
                            addRule(rrule, "BYDAY", string.Join(",", getByDay(oPattern.DayOfWeekMask).ToArray()));
                        }
                        break;
                    }

                case OlRecurrenceType.olRecursMonthly: {
                        addRule(rrule, "FREQ", "MONTHLY");
                        setInterval(oPattern.Interval);
                        //Outlook runs on last day of month if day doesn't exist; Google doesn't run at all - so fix
                        if (oPattern.PatternStartDate.Day > 28) {
                            addRule(rrule, "BYDAY", "SU,MO,TU,WE,TH,FR,SA");
                            addRule(rrule, "BYSETPOS", "-1");
                        }
                        break;
                    }

                case OlRecurrenceType.olRecursMonthNth: {
                        addRule(rrule, "FREQ", "MONTHLY");
                        setInterval(oPattern.Interval);
                        addRule(rrule, "BYSETPOS", (oPattern.Instance == 5) ? "-1" : oPattern.Instance.ToString());
                        if (oPattern.DayOfWeekMask != (OlDaysOfWeek)127) { //If not every day of week, define which ones
                            addRule(rrule, "BYDAY", string.Join(",", getByDay(oPattern.DayOfWeekMask).ToArray()));
                        }
                        break;
                    }

                case OlRecurrenceType.olRecursYearly: {
                        addRule(rrule, "FREQ", "YEARLY");
                        //Google interval is years, Outlook is months
                        addRule(rrule, "INTERVAL", (oPattern.Interval / 12).ToString());
                        break;
                    }

                case OlRecurrenceType.olRecursYearNth: {
                        addRule(rrule, "FREQ", "YEARLY");
                        addRule(rrule, "BYSETPOS", (oPattern.Instance == 5) ? "-1" : oPattern.Instance.ToString());
                        if (oPattern.DayOfWeekMask != (OlDaysOfWeek)127) { //If not every day of week, define which ones
                            addRule(rrule, "BYDAY", string.Join(",", getByDay(oPattern.DayOfWeekMask).ToArray()));
                        }
                        addRule(rrule, "BYMONTH", oPattern.MonthOfYear.ToString());
                        break;
                    }
            }
            #endregion

            #region RECURRENCE RANGE
            if (!oPattern.NoEndDate) {
                log.Fine("Checking end date.");
                addRule(rrule, "UNTIL", Recurrence.IANAdate(oPattern.PatternEndDate));
            }
            #endregion
            return string.Join(";", rrule.Select(x => x.Key + "=" + x.Value).ToArray());
        }

        private Dictionary<String, String> explodeRrule(IList<String> allRules) {
            log.Fine("Analysing Event RRULEs...");
            foreach (String rrule in allRules) {
                if (rrule.StartsWith("RRULE:")) {
                    log.Debug("Converting " + rrule);
                    String[] rrules = rrule.TrimStart("RRULE:".ToCharArray()).Split(';');
                    Dictionary<String, String> rules = new Dictionary<String, String>();
                    for (int r = 0; r < rrules.Count(); r++) {
                        String[] ruleKVPs = rrules[r].Split('=');
                        rules.Add(ruleKVPs[0], ruleKVPs[1]);
                    }
                    return rules;
                }
            }
            return null;
        }

        private void addRule(Dictionary<string, string> ruleBook, string key, string value) {
            ruleBook.Add(key, value);
            log.Fine(ruleBook.Last().Value);
        }

        private void setInterval(int interval) {
            if (interval > 1) addRule(rrule, "INTERVAL", interval.ToString());
        }

        private static List<String> getByDay(OlDaysOfWeek dowMask) {
            log.Fine("DayOfWeekMask = " + dowMask);
            List<String> byDay = new List<String>();
            byDay.Add(((dowMask & OlDaysOfWeek.olMonday) != 0) ? "MO" : "");
            byDay.Add(((dowMask & OlDaysOfWeek.olTuesday) != 0) ? "TU" : "");
            byDay.Add(((dowMask & OlDaysOfWeek.olWednesday) != 0) ? "WE" : "");
            byDay.Add(((dowMask & OlDaysOfWeek.olThursday) != 0) ? "TH" : "");
            byDay.Add(((dowMask & OlDaysOfWeek.olFriday) != 0) ? "FR" : "");
            byDay.Add(((dowMask & OlDaysOfWeek.olSaturday) != 0) ? "SA" : "");
            byDay.Add(((dowMask & OlDaysOfWeek.olSunday) != 0) ? "SU" : "");
            byDay = byDay.Where(s => !string.IsNullOrEmpty(s)).ToList();
            return byDay;
        }
        private static OlDaysOfWeek getDOWmask(Dictionary<String, String> ruleBook) {
            OlDaysOfWeek dowMask = 0;
            if (ruleBook.ContainsKey("BYDAY")) dowMask = getDOWmask(ruleBook["BYDAY"]);
            return dowMask;
        }
        private static OlDaysOfWeek getDOWmask(String byDay) {
            OlDaysOfWeek dowMask = 0;
            if (byDay.Contains("MO")) dowMask |= OlDaysOfWeek.olMonday;
            if (byDay.Contains("TU")) dowMask |= OlDaysOfWeek.olTuesday;
            if (byDay.Contains("WE")) dowMask |= OlDaysOfWeek.olWednesday;
            if (byDay.Contains("TH")) dowMask |= OlDaysOfWeek.olThursday;
            if (byDay.Contains("FR")) dowMask |= OlDaysOfWeek.olFriday;
            if (byDay.Contains("SA")) dowMask |= OlDaysOfWeek.olSaturday;
            if (byDay.Contains("SU")) dowMask |= OlDaysOfWeek.olSunday;
            return dowMask;
        }

        public static String IANAdate(DateTime dt) {
            //Eg "RRULE:FREQ=DAILY;UNTIL=20150826T093000Z"
            //Need to add a day to date else Google is a day short compared to Outlook
            return dt.AddDays(1).ToString("yyyyMMddTHHmmssZ");
        }
        #endregion

        #region Exceptions

        #region Google
        private List<Event> googleExceptions; 
        public List<Event> GoogleExceptions {
            get { return googleExceptions; }
        }
        
        public Boolean HasExceptions(Event ev, Boolean checkLocalCacheOnly = false) {
            if (ev.Recurrence == null) return false;

            //There's currently no good way to know if a Google event is an exception or not.
            //If it's a change in date, the sequence number increments. However, if it's a different field (eg Subject), no increment.
            //Therefore, you'd have to GetCalendarEntriesInRange() with no date range, then filter to the recurringEventId - not efficient!
            //So...will make it only sync exceptions within the sync date range, which we have cached already
            //if (!checkLocalCacheOnly) {
            //    List<Event> gInstances = GoogleCalendar.Instance.GetCalendarEntriesInRecurrence(ev.RecurringEventId ?? ev.Id);
            //    //Add any new exceptions to local cache
            //    googleExceptions = googleExceptions.Union(gInstances.Where(exp => exp.Sequence > 0)).ToList();
            //}
            int exceptionCount = Recurrence.Instance.googleExceptions.Where(exp => exp.RecurringEventId == ev.Id).Count();
            if (exceptionCount > 0) {
                log.Debug("This is a recurring Google event with " + exceptionCount + " exceptions in the sync date range.");
                return true;
            } else
                return false;
        }

        public void SeparateGoogleExceptions(List<Event> allEvents) {
            googleExceptions = new List<Event>();
            if (allEvents.Count == 0) return;
            log.Debug("Identifying exceptions in recurring Google events.");
            googleExceptions = new List<Event>();
            for (int g = allEvents.Count - 1; g >= 0; g--) {
                if (!string.IsNullOrEmpty(allEvents[g].RecurringEventId)) {
                    googleExceptions.Add(allEvents[g]);
                    allEvents.Remove(allEvents[g]);
                }
            }
            log.Debug("Found " + googleExceptions.Count + " exceptions.");
        }

        public Event GetGoogleInstance(Microsoft.Office.Interop.Outlook.Exception oExcp, String gRecurringEventID, String oEntryID) {
            Boolean oIsDeleted = exceptionIsDeleted(oExcp);
            log.Debug("Finding Google instance for " + (oIsDeleted ? "deleted " : "") + "Outlook exception:-");
            log.Debug("  Original date: " + oExcp.OriginalDate.ToString("dd/MM/yyyy"));
            if (!oIsDeleted) log.Debug("  Current  date: " + oExcp.AppointmentItem.Start.ToString("dd/MM/yyyy"));
            foreach (Event gExcp in googleExceptions) {
                if (gExcp.RecurringEventId == gRecurringEventID) {
                    if ((!oIsDeleted &&
                        GoogleCalendar.GoogleTimeFrom(oExcp.OriginalDate) == (gExcp.OriginalStartTime.Date ?? gExcp.OriginalStartTime.DateTime)
                        ) ||
                        (oIsDeleted &&
                        GoogleCalendar.GoogleTimeFrom(oExcp.OriginalDate) == (gExcp.OriginalStartTime.Date ?? GoogleCalendar.GoogleTimeFrom(DateTime.Parse(gExcp.OriginalStartTime.DateTime).Date))
                        )) {
                        return gExcp;
                    }
                }
            }
            log.Debug("Google exception event is not cached. Retrieving all recurring instances...");
            List<Event> gInstances = GoogleCalendar.Instance.GetCalendarEntriesInRecurrence(gRecurringEventID);
            //Add any new exceptions to local cache
            googleExceptions = googleExceptions.Union(gInstances.Where(ev => !String.IsNullOrEmpty(ev.RecurringEventId))).ToList();
            foreach (Event gInst in gInstances) {
                if (gInst.RecurringEventId == gRecurringEventID) {
                    if ((!oIsDeleted &&
                        GoogleCalendar.GoogleTimeFrom(oExcp.OriginalDate) == GoogleCalendar.GoogleTimeFrom(DateTime.Parse(gInst.OriginalStartTime.Date ?? gInst.OriginalStartTime.DateTime))
                        ) ||
                        (oIsDeleted &&
                        GoogleCalendar.GoogleTimeFrom(oExcp.OriginalDate) == GoogleCalendar.GoogleTimeFrom(DateTime.Parse(gInst.OriginalStartTime.Date ?? gInst.OriginalStartTime.DateTime).Date)
                        )) {
                        return gInst;
                    }
                }
            }
            return null;
        }

        public Event GetGoogleMasterEvent(AppointmentItem ai) {
            log.Fine("Found a master Outlook recurring item outside sync date range: " + OutlookCalendar.GetEventSummary(ai));
            List<Event> events = new List<Event>();
            if (ai.UserProperties[OutlookCalendar.gEventID] == null) {
                events = GoogleCalendar.Instance.GetCalendarEntriesInRange(ai.Start.Date, ai.Start.Date.AddDays(1));
                List<AppointmentItem> ais = new List<AppointmentItem>();
                ais.Add(ai);
                GoogleCalendar.Instance.ReclaimOrphanCalendarEntries(ref events, ref ais, neverDelete: true);
            } else {
                Event ev = GoogleCalendar.Instance.GetCalendarEntry(ai.UserProperties[OutlookCalendar.gEventID].Value.ToString());
                if (ev != null) events.Add(ev);
            }
            for (int g = 0; g < events.Count(); g++) { //Event ev in events) {
                String gEntryID;
                Event ev = events[g];
                if (GoogleCalendar.GetOGCSproperty(ev, GoogleCalendar.oEntryID, out gEntryID)) {
                    if (gEntryID == ai.EntryID) {
                        log.Info("Migrating Master Event from EntryID to GlobalAppointmentID...");
                        GoogleCalendar.AddOutlookID(ref ev, ai);
                        GoogleCalendar.Instance.UpdateCalendarEntry_save(ev);
                        return ev;
                    } else if (gEntryID == OutlookCalendar.Instance.IOutlook.GetGlobalApptID(ai)) {
                        log.Fine("Found master event.");
                        return ev;
                    }
                }
            }
            log.Warn("Failed to find master Google event for: " + OutlookCalendar.GetEventSummary(ai));
            return null;
        }

        public static void CreateGoogleExceptions(AppointmentItem ai, String recurringEventId) {
            if (!ai.IsRecurring) return;

            log.Debug("Creating Google recurrence exceptions.");
            List<Event> gRecurrences = GoogleCalendar.Instance.GetCalendarEntriesInRecurrence(recurringEventId);
            if (gRecurrences != null) {
                Microsoft.Office.Interop.Outlook.Exceptions exps = ai.GetRecurrencePattern().Exceptions;
                foreach (Microsoft.Office.Interop.Outlook.Exception oExcp in exps) {
                    foreach (Event ev in gRecurrences) {
                        String gDate = ev.OriginalStartTime.DateTime ?? ev.OriginalStartTime.Date;
                        Boolean isDeleted = exceptionIsDeleted(oExcp);
                        if (isDeleted && !ai.AllDayEvent) { //Deleted items get truncated?!
                            gDate = GoogleCalendar.GoogleTimeFrom(DateTime.Parse(gDate).Date);
                        }
                        if (oExcp.OriginalDate == DateTime.Parse(gDate)) {
                            if (isDeleted) {
                                MainForm.Instance.Logboxout(GoogleCalendar.GetEventSummary(ev));
                                MainForm.Instance.Logboxout("Recurrence deleted.");
                                ev.Status = "cancelled";
                                GoogleCalendar.Instance.UpdateCalendarEntry_save(ev);
                            } else {
                                int exceptionItemsModified = 0;
                                Event modifiedEv = GoogleCalendar.Instance.UpdateCalendarEntry(oExcp.AppointmentItem, ev, ref exceptionItemsModified, forceCompare:true);
                                if (exceptionItemsModified > 0) {
                                    GoogleCalendar.Instance.UpdateCalendarEntry_save(modifiedEv);
                                }
                            }
                            break;
                        }
                    }
                }
            }
        }

        public static void UpdateGoogleExceptions(AppointmentItem ai, Event ev) {
            if (ai.IsRecurring) {
                RecurrencePattern recurrence = ai.GetRecurrencePattern();
                if (recurrence.Exceptions.Count > 0) {
                    log.Debug(OutlookCalendar.GetEventSummary(ai));
                    log.Debug("This is a recurring appointment with " + recurrence.Exceptions.Count + " exceptions that will now be iteratively compared.");
                    foreach (Microsoft.Office.Interop.Outlook.Exception oExcp in recurrence.Exceptions) {
                        int excp_itemModified = 0;

                        //Check the exception falls in the date range being synced
                        Boolean oIsDeleted = exceptionIsDeleted(oExcp);
                        String logDeleted = oIsDeleted ? " deleted and" : "";
                        DateTime oExcp_currDate = oIsDeleted ? oExcp.OriginalDate : oExcp.AppointmentItem.Start;
                        if (oExcp_currDate < Settings.Instance.SyncStart.Date || oExcp_currDate > Settings.Instance.SyncEnd.Date) {
                            log.Fine("Exception is" + logDeleted + " outside date range being synced: " + oExcp_currDate.Date.ToString("dd/MM/yyyy"));
                            continue;
                        }

                        Event gExcp = Recurrence.Instance.GetGoogleInstance(oExcp, ev.RecurringEventId ?? ev.Id, OutlookCalendar.Instance.IOutlook.GetGlobalApptID(ai));
                        if (gExcp != null) {
                            log.Debug("Matching Google Event recurrence found.");
                            if (gExcp.Status == "cancelled") {
                                log.Debug("It is deleted in Google, so cannot compare items.");
                                if (!oIsDeleted) log.Warn("Outlook is NOT deleted though - a mismatch has occurred somehow");
                                continue;
                            }
                            try {
                                GoogleCalendar.Instance.UpdateCalendarEntry(oExcp.AppointmentItem, gExcp, ref excp_itemModified);
                            } catch (System.Exception ex) {
                                if (oIsDeleted) {
                                    if (gExcp.Status != "cancelled") {
                                        gExcp.Status = "cancelled";
                                        excp_itemModified++;
                                    }
                                } else {
                                    log.Error(ex.Message);
                                    log.Error(ex.StackTrace);
                                    recurrence = (RecurrencePattern)OutlookCalendar.ReleaseObject(recurrence);
                                    throw ex;
                                }
                            }
                            if (excp_itemModified > 0) {
                                try {
                                    GoogleCalendar.Instance.UpdateCalendarEntry_save(gExcp);
                                } catch (System.Exception ex) {
                                    MainForm.Instance.Logboxout("WARNING: Updated event exception failed to save.\r\n" + ex.Message);
                                    log.Error(ex.StackTrace);
                                    if (MessageBox.Show("Updated Google event exception failed to save. Continue with synchronisation?", "Sync item failed", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                                        continue;
                                    else {
                                        throw new UserCancelledSyncException("User chose not to continue sync.");
                                    }
                                } finally {
                                    recurrence = (RecurrencePattern)OutlookCalendar.ReleaseObject(recurrence);
                                }
                            }
                        } else {
                            log.Debug("No matching Google Event recurrence found.");
                            if (oIsDeleted) log.Debug("The Outlook appointment is deleted, so not a problem.");
                        }
                    }
                }
                recurrence = (RecurrencePattern)OutlookCalendar.ReleaseObject(recurrence);
            }
        }
        #endregion

        #region Outlook
        public static Boolean HasExceptions(AppointmentItem ai) {
            return ai.GetRecurrencePattern().Exceptions.Count != 0;
        }

        private static Boolean exceptionIsDeleted(Microsoft.Office.Interop.Outlook.Exception oExcp) {
            if (oExcp.Deleted) return true;
            try {
                AppointmentItem ai = oExcp.AppointmentItem;
                return false;
            } catch (System.Exception ex) {
                if (ex.Message == "You changed one of the recurrences of this item, and this instance no longer exists. Close any open items and try again.") {
                    //log.Debug("This Outlook recurrence instance has been deleted, but the API is reporting it incorrectly due to caching");
                    return true;
                } else {
                    log.Warn("Error when determining if Outlook recurrence is deleted or not.\r\n" + ex.Message);
                    return true;
                }
            }
        }

        public void CreateOutlookExceptions(AppointmentItem ai, Event ev) {
            processOutlookExceptions(ai, ev, forceCompare: true);
        }
        public void UpdateOutlookExceptions(AppointmentItem ai, Event ev) {
            processOutlookExceptions(ai, ev, forceCompare: false);
        }
        
        private void processOutlookExceptions(AppointmentItem ai, Event ev, Boolean forceCompare) {
            if (!HasExceptions(ev, checkLocalCacheOnly: true)) return;

            if (!ai.Saved) ai.Save();

            RecurrencePattern oPattern = ai.GetRecurrencePattern();
            foreach (Event gExcp in Recurrence.Instance.googleExceptions.Where(exp => exp.RecurringEventId == ev.Id)) {
                log.Fine("Found Google exception for " + (gExcp.OriginalStartTime.DateTime ?? gExcp.OriginalStartTime.Date));

                DateTime oExcpDate = DateTime.Parse(gExcp.OriginalStartTime.DateTime ?? gExcp.OriginalStartTime.Date);
                AppointmentItem newAiExcp = getOutlookInstance(oPattern, oExcpDate);
                if (newAiExcp == null) continue;

                if (gExcp.Status != "cancelled") {
                    int itemModified = 0;
                    newAiExcp = OutlookCalendar.Instance.UpdateCalendarEntry(newAiExcp, gExcp, ref itemModified, forceCompare);
                    if (itemModified > 0) newAiExcp.Save();
                } else {
                    MainForm.Instance.Logboxout(OutlookCalendar.GetEventSummary(ai) +"\r\nDeleted.");
                    newAiExcp.Delete();
                }
                newAiExcp = (AppointmentItem)OutlookCalendar.ReleaseObject(newAiExcp);
            }
            if (!ai.Saved) ai.Save();
            oPattern = (RecurrencePattern)OutlookCalendar.ReleaseObject(oPattern);
        }

        private static AppointmentItem getOutlookInstance(RecurrencePattern oPattern, DateTime instanceDate) {
            //First check if this is not yet an exception
            AppointmentItem ai = null;
            try {
                ai = oPattern.GetOccurrence(instanceDate);
            } catch { }
            if (ai == null) {
                //The Outlook API is rubbish as the date argument is how it exists NOW (not OriginalDate). 
                //If this has changed >1 in Google then there's no way of knowing what it might be!
                
                foreach (Microsoft.Office.Interop.Outlook.Exception oExp in oPattern.Exceptions) {
                    if (oExp.OriginalDate.Date == instanceDate.Date) {
                        try {
                            log.Debug("Found Outlook exception for " + instanceDate);
                            if (exceptionIsDeleted(oExp)) {
                                log.Debug("This exception is deleted.");
                                return null;
                            } else {
                                return oExp.AppointmentItem;
                            }
                        } catch (System.Exception ex) {
                            MainForm.Instance.Logboxout(ex.Message);
                            MainForm.Instance.Logboxout("If this keeps happening, please restart OGCS.");
                            break;
                        }
                    }
                }
                if (ai == null) log.Warn("Unable to find Outlook exception for " + instanceDate);
            }
            return ai;
        }
        #endregion
        
        #endregion
    }
}
