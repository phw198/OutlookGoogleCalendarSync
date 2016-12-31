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
            RecurrencePattern rp = null;
            try {
                rp = ai.GetRecurrencePattern();
                gPattern.Add("RRULE:" + buildRrule(rp));
            } finally {
                rp = (RecurrencePattern)OutlookCalendar.ReleaseObject(rp);
            }
            log.Debug(string.Join("\r\n", gPattern.ToArray()));
            return gPattern;
        }

        public void BuildOutlookPattern(Event ev, AppointmentItem ai) {
            RecurrencePattern ignore;
            buildOutlookPattern(ev, ai, out ignore);
            ignore = (RecurrencePattern)OutlookCalendar.ReleaseObject(ignore);
        }

        private void buildOutlookPattern(Event ev, AppointmentItem ai, out RecurrencePattern oPattern) {
            if (ev.Recurrence == null) { oPattern = null; return; }

            Dictionary<String, String> ruleBook = explodeRrule(ev.Recurrence);
            if (ruleBook == null) {
                throw new ApplicationException("WARNING: The recurrence pattern is not compatble with Outlook. This event cannot be synced.");
            }
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
            ai = OutlookCalendar.Instance.IOutlook.WindowsTimeZone_set(ai, ev);
            oPattern.PatternStartDate = DateTime.Parse(ev.Start.Date ?? ev.Start.DateTime);
            if (ruleBook.ContainsKey("INTERVAL") && Convert.ToInt16(ruleBook["INTERVAL"]) > 1 && ruleBook["FREQ"] != "YEARLY")
                oPattern.Interval = Convert.ToInt16(ruleBook["INTERVAL"]);
            if (ruleBook.ContainsKey("COUNT"))
                oPattern.Occurrences = Convert.ToInt16(ruleBook["COUNT"]);
            if (ruleBook.ContainsKey("UNTIL")) {
                //if (ruleBook["UNTIL"].Length == 8) {
                    oPattern.PatternEndDate = DateTime.ParseExact(ruleBook["UNTIL"].ToString().Substring(0,8), "yyyyMMdd", System.Globalization.CultureInfo.InvariantCulture).Date;
                //} else {
                    //if (ruleBook["UNTIL"].ToString().Substring(8) == "T000000Z" && ev.Start.DateTime != null)
                    //    oPattern.PatternEndDate = DateTime.ParseExact(ruleBook["UNTIL"], "yyyyMMddTHHmmssZ", System.Globalization.CultureInfo.InvariantCulture).AddDays(-1);
                    //else
                        //oPattern.PatternEndDate = DateTime.ParseExact(ruleBook["UNTIL"], "yyyyMMddTHHmmssZ", System.Globalization.CultureInfo.InvariantCulture).Date;
                //}
            }
            if (!ruleBook.ContainsKey("COUNT") && !ruleBook.ContainsKey("UNTIL")) {
                oPattern.NoEndDate = true;
            }
            #endregion
        }

        public void CompareOutlookPattern(Event ev, ref RecurrencePattern aiOpattern, SyncDirection syncDirection, System.Text.StringBuilder sb, ref int itemModified) {
            if (ev.Recurrence == null) return;
            
            log.Fine("Building a temporary recurrent Appointment generated from Event");
            AppointmentItem evAI = OutlookCalendar.Instance.IOutlook.UseOutlookCalendar().Items.Add() as AppointmentItem;
            evAI.Start = DateTime.Parse(ev.Start.Date ?? ev.Start.DateTime); 

            RecurrencePattern evOpattern = null;
            try {
                buildOutlookPattern(ev, evAI, out evOpattern);
                log.Fine("Comparing Google recurrence to Outlook equivalent");

                if (MainForm.CompareAttribute("Recurrence Type", syncDirection,
                    evOpattern.RecurrenceType.ToString(), aiOpattern.RecurrenceType.ToString(), sb, ref itemModified)) {
                    aiOpattern.RecurrenceType = evOpattern.RecurrenceType;
                }
                if (MainForm.CompareAttribute("Recurrence Interval", syncDirection,
                    evOpattern.Interval.ToString(), aiOpattern.Interval.ToString(), sb, ref itemModified)) {
                    aiOpattern.Interval = evOpattern.Interval;
                }
                if (MainForm.CompareAttribute("Recurrence Instance", syncDirection,
                    evOpattern.Instance.ToString(), aiOpattern.Instance.ToString(), sb, ref itemModified)) {
                    aiOpattern.Instance = evOpattern.Instance;
                }
                if (MainForm.CompareAttribute("Recurrence DoW", syncDirection,
                    evOpattern.DayOfWeekMask.ToString(), aiOpattern.DayOfWeekMask.ToString(), sb, ref itemModified)) {
                    aiOpattern.DayOfWeekMask = evOpattern.DayOfWeekMask;
                }
                if (MainForm.CompareAttribute("Recurrence MoY", syncDirection,
                    evOpattern.MonthOfYear.ToString(), aiOpattern.MonthOfYear.ToString(), sb, ref itemModified)) {
                    aiOpattern.MonthOfYear = evOpattern.MonthOfYear;
                }
                if (MainForm.CompareAttribute("Recurrence NoEndDate", syncDirection,
                    evOpattern.NoEndDate, aiOpattern.NoEndDate, sb, ref itemModified)) {
                    aiOpattern.NoEndDate = evOpattern.NoEndDate;
                }
                if (MainForm.CompareAttribute("Recurrence Occurences", syncDirection,
                    evOpattern.Occurrences.ToString(), aiOpattern.Occurrences.ToString(), sb, ref itemModified)) {
                    aiOpattern.Occurrences = evOpattern.Occurrences;
                }
            } finally {
                evOpattern = (RecurrencePattern)OutlookCalendar.ReleaseObject(evOpattern);
                evAI.Delete();
                evAI = (AppointmentItem)OutlookCalendar.ReleaseObject(evAI);
            }
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
                addRule(rrule, "UNTIL", Recurrence.IANAdate(oPattern.PatternEndDate + oPattern.StartTime.TimeOfDay));
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
            log.Warn("There aren't any RRULEs present. Outlook doesn't support this: https://support.microsoft.com/en-gb/kb/2643084");
            foreach (String rule in allRules) {
                log.Debug("rule: " + rule);
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
            return dt.ToString("yyyyMMddTHHmmssZ");
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
        
        private Event getGoogleInstance(ref Microsoft.Office.Interop.Outlook.Exception oExcp, String gRecurringEventID, String oEntryID, Boolean dirtyCache) {
            Boolean oIsDeleted = exceptionIsDeleted(oExcp);
            log.Debug("Finding Google instance for " + (oIsDeleted ? "deleted " : "") + "Outlook exception:-");
            log.Debug("  Original date: " + oExcp.OriginalDate.ToString("dd/MM/yyyy"));
            if (!oIsDeleted) {
                AppointmentItem ai = null;
                try {
                    ai = oExcp.AppointmentItem;
                    log.Debug("  Current  date: " + ai.Start.ToString("dd/MM/yyyy"));
                } finally {
                    ai = (AppointmentItem)OutlookCalendar.ReleaseObject(ai);
                }
            }
            if (dirtyCache) {
                log.Debug("Google exception cache not being used. Retrieving all recurring instances afresh...");
                //Remove dirty items
                googleExceptions.RemoveAll(ev => ev.RecurringEventId == gRecurringEventID);
            } else {
                foreach (Event gExcp in googleExceptions) {
                    if (gExcp.RecurringEventId == gRecurringEventID) {
                        if ((!oIsDeleted &&
                            oExcp.OriginalDate == DateTime.Parse(gExcp.OriginalStartTime.Date ?? gExcp.OriginalStartTime.DateTime)
                            ) ||
                            (oIsDeleted &&
                            oExcp.OriginalDate == DateTime.Parse(gExcp.OriginalStartTime.Date ?? gExcp.OriginalStartTime.DateTime).Date
                            )) {
                            return gExcp;
                        }
                    }
                }
                log.Debug("Google exception event is not cached. Retrieving all recurring instances...");
            }
            List<Event> gInstances = GoogleCalendar.Instance.GetCalendarEntriesInRecurrence(gRecurringEventID);
            //Add any new exceptions to local cache
            googleExceptions = googleExceptions.Union(gInstances.Where(ev => !String.IsNullOrEmpty(ev.RecurringEventId))).ToList();
            foreach (Event gInst in gInstances) {
                if (gInst.RecurringEventId == gRecurringEventID) {
                    if (((!oIsDeleted || (oIsDeleted && !oExcp.Deleted)) /* Weirdness when exception is cancelled by organiser but not yet deleted/accepted by recipient */
                        && oExcp.OriginalDate == DateTime.Parse(gInst.OriginalStartTime.Date ?? gInst.OriginalStartTime.DateTime)
                        ) ||
                        (oIsDeleted &&
                        oExcp.OriginalDate == DateTime.Parse(gInst.OriginalStartTime.Date ?? gInst.OriginalStartTime.DateTime).Date
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
            for (int g = 0; g < events.Count(); g++) {
                String gEntryID;
                Event ev = events[g];
                if (GoogleCalendar.GetOGCSproperty(ev, GoogleCalendar.oEntryID, out gEntryID)) {
                    if (gEntryID == ai.EntryID) {
                        log.Info("Migrating Master Event from EntryID to GlobalAppointmentID...");
                        GoogleCalendar.AddOutlookID(ref ev, ai);
                        GoogleCalendar.Instance.UpdateCalendarEntry_save(ref ev);
                        return ev;
                    } else if (MainForm.ItemIDsMatch(gEntryID, OutlookCalendar.Instance.IOutlook.GetGlobalApptID(ai))) {
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
                RecurrencePattern rp = null;
                Exceptions excps = null;
                try {
                    rp = ai.GetRecurrencePattern();
                    excps = rp.Exceptions;
                    for (int e = 1; e <= excps.Count; e++) {
                        Microsoft.Office.Interop.Outlook.Exception oExcp = null;
                        try {
                            oExcp = excps[e];
                            for (int g = 0; g < gRecurrences.Count; g++) {
                                Event ev = gRecurrences[g];
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
                                        GoogleCalendar.Instance.UpdateCalendarEntry_save(ref ev);
                                    } else {
                                        int exceptionItemsModified = 0;
                                        Event modifiedEv = GoogleCalendar.Instance.UpdateCalendarEntry(oExcp.AppointmentItem, ev, ref exceptionItemsModified, forceCompare: true);
                                        if (exceptionItemsModified > 0) {
                                            GoogleCalendar.Instance.UpdateCalendarEntry_save(ref modifiedEv);
                                        }
                                    }
                                    break;
                                }
                            }
                        } finally {
                            oExcp = (Microsoft.Office.Interop.Outlook.Exception)OutlookCalendar.ReleaseObject(oExcp);
                        }
                    }
                } finally {
                    excps = (Exceptions)OutlookCalendar.ReleaseObject(excps);
                    rp = (RecurrencePattern)OutlookCalendar.ReleaseObject(rp);
                }
            }
        }

        public static void UpdateGoogleExceptions(AppointmentItem ai, Event ev, Boolean dirtyCache) {
            if (ai.IsRecurring) {
                RecurrencePattern rp = null;
                Exceptions excps = null;
                try {
                    rp = ai.GetRecurrencePattern();
                    excps = rp.Exceptions;
                    if (excps.Count > 0) {
                        log.Debug(OutlookCalendar.GetEventSummary(ai));
                        log.Debug("This is a recurring appointment with " + excps.Count + " exceptions that will now be iteratively compared.");
                        for (int e = 1; e <= excps.Count; e++) {
                            Microsoft.Office.Interop.Outlook.Exception oExcp = null;
                            AppointmentItem aiExcp = null;
                            try {
                                oExcp = excps[e];
                                int excp_itemModified = 0;

                                //Check the exception falls in the date range being synced
                                Boolean oIsDeleted = exceptionIsDeleted(oExcp);
                                String logDeleted = oIsDeleted ? " deleted and" : "";
                                DateTime oExcp_currDate;
                                if (oIsDeleted)
                                    oExcp_currDate = oExcp.OriginalDate;
                                else {
                                    aiExcp = oExcp.AppointmentItem;
                                    oExcp_currDate = aiExcp.Start;
                                    aiExcp = (AppointmentItem)OutlookCalendar.ReleaseObject(aiExcp);
                                }
                            
                                if (oExcp_currDate < Settings.Instance.SyncStart.Date || oExcp_currDate > Settings.Instance.SyncEnd.Date) {
                                    log.Fine("Exception is" + logDeleted + " outside date range being synced: " + oExcp_currDate.Date.ToString("dd/MM/yyyy"));
                                    continue;
                                }

                                Event gExcp = Recurrence.Instance.getGoogleInstance(ref oExcp, ev.RecurringEventId ?? ev.Id, OutlookCalendar.Instance.IOutlook.GetGlobalApptID(ai), dirtyCache);
                                if (gExcp != null) {
                                    log.Debug("Matching Google Event recurrence found.");
                                    if (gExcp.Status == "cancelled") {
                                        log.Debug("It is deleted in Google, so cannot compare items.");
                                        if (!oIsDeleted) log.Warn("Outlook is NOT deleted though - a mismatch has occurred somehow!");
                                        continue;
                                    } else if (oIsDeleted && gExcp.Status != "cancelled") {
                                        gExcp.Status = "cancelled";
                                        excp_itemModified++;
                                    } else {
                                        try {
                                            aiExcp = oExcp.AppointmentItem;
                                            GoogleCalendar.Instance.UpdateCalendarEntry(aiExcp, gExcp, ref excp_itemModified);
                                        } catch (System.Exception ex) {
                                            log.Error(ex.Message);
                                            log.Error(ex.StackTrace);
                                            throw ex;
                                        }
                                    }
                                    if (excp_itemModified > 0) {
                                        try {
                                            GoogleCalendar.Instance.UpdateCalendarEntry_save(ref gExcp);
                                        } catch (System.Exception ex) {
                                            MainForm.Instance.Logboxout("WARNING: Updated event exception failed to save.\r\n" + ex.Message);
                                            log.Error(ex.StackTrace);
                                            if (MessageBox.Show("Updated Google event exception failed to save. Continue with synchronisation?", "Sync item failed", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                                                continue;
                                            else {
                                                throw new UserCancelledSyncException("User chose not to continue sync.");
                                            }
                                        }
                                    }
                                } else {
                                    log.Debug("No matching Google Event recurrence found.");
                                    if (oIsDeleted) log.Debug("The Outlook appointment is deleted, so not a problem.");
                                }
                            } finally {
                                aiExcp = (AppointmentItem)OutlookCalendar.ReleaseObject(aiExcp);
                                oExcp = (Microsoft.Office.Interop.Outlook.Exception)OutlookCalendar.ReleaseObject(oExcp);
                            }
                        }
                    }
                } finally {
                    excps = (Exceptions)OutlookCalendar.ReleaseObject(excps);
                    rp = (RecurrencePattern)OutlookCalendar.ReleaseObject(rp);
                }
            }
        }
        #endregion

        #region Outlook
        public static Boolean HasExceptions(AppointmentItem ai) {
            RecurrencePattern rp = null;
            Exceptions excps = null;
            try {
                rp = ai.GetRecurrencePattern();
                excps = rp.Exceptions;
                return excps.Count != 0;
            } finally {
                excps = (Exceptions)OutlookCalendar.ReleaseObject(excps);
                rp = (RecurrencePattern)OutlookCalendar.ReleaseObject(rp);
            }
        }

        private static Boolean exceptionIsDeleted(Microsoft.Office.Interop.Outlook.Exception oExcp) {
            if (oExcp.Deleted) return true;
            AppointmentItem ai = null;
            try {
                ai = oExcp.AppointmentItem;
                return false;
            } catch (System.Exception ex) {
                OGCSexception.Analyse(ex);
                if (ex.Message == "You changed one of the recurrences of this item, and this instance no longer exists. Close any open items and try again.") {
                    log.Warn("This Outlook recurrence instance has become inaccessible, probably due to caching");
                    return true;
                } else {
                    log.Warn("Error when determining if Outlook recurrence is deleted or not.\r\n" + ex.Message);
                    return true;
                }
            } finally {
                ai = (AppointmentItem)OutlookCalendar.ReleaseObject(ai);
            }
        }

        public void CreateOutlookExceptions(ref AppointmentItem ai, Event ev) {
            processOutlookExceptions(ref ai, ev, forceCompare: true);
        }
        public void UpdateOutlookExceptions(ref AppointmentItem ai, Event ev) {
            processOutlookExceptions(ref ai, ev, forceCompare: false);
        }

        private void processOutlookExceptions(ref AppointmentItem ai, Event ev, Boolean forceCompare) {
            if (!HasExceptions(ev, checkLocalCacheOnly: true)) return;

            RecurrencePattern oPattern = null;
            try {
                oPattern = ai.GetRecurrencePattern();
                foreach (Event gExcp in Recurrence.Instance.googleExceptions.Where(exp => exp.RecurringEventId == ev.Id)) {
                    log.Fine("Found Google exception for " + (gExcp.OriginalStartTime.DateTime ?? gExcp.OriginalStartTime.Date));

                    DateTime oExcpDate = DateTime.Parse(gExcp.OriginalStartTime.DateTime ?? gExcp.OriginalStartTime.Date);
                    AppointmentItem newAiExcp = null;
                    try {
                        getOutlookInstance(oPattern, oExcpDate, ref newAiExcp);
                        if (newAiExcp == null) continue;

                        if (gExcp.Status != "cancelled") {
                            int itemModified = 0;
                            OutlookCalendar.Instance.UpdateCalendarEntry(ref newAiExcp, gExcp, ref itemModified, forceCompare);
                            if (itemModified > 0) {
                                try {
                                    newAiExcp.Save();
                                } catch (System.Exception ex) {
                                    OGCSexception.Analyse(ex);
                                    if (ex.Message == "Cannot save this item.") {
                                        MainForm.Instance.Logboxout("Uh oh! Outlook wasn't able to save this recurrence exception! " +
                                            "You may have two occurences on the same day, which it doesn't allow.");
                                    }
                                }
                            }
                        } else {
                            MainForm.Instance.Logboxout(OutlookCalendar.GetEventSummary(ai) + "\r\nDeleted.");
                            newAiExcp.Delete();
                        }
                    } finally {
                        newAiExcp = (AppointmentItem)OutlookCalendar.ReleaseObject(newAiExcp);
                    }
                }
            } finally {
                oPattern = (RecurrencePattern)OutlookCalendar.ReleaseObject(oPattern);
            }
        }

        private static void getOutlookInstance(RecurrencePattern oPattern, DateTime instanceDate, ref AppointmentItem ai) {
            //First check if this is not yet an exception
            try {
                ai = oPattern.GetOccurrence(instanceDate);
            } catch { }
            if (ai == null) {
                //The Outlook API is rubbish as the date argument is how it exists NOW (not OriginalDate). 
                //If this has changed >1 in Google then there's no way of knowing what it might be!

                Exceptions oExcps = null;
                try {
                    oExcps = oPattern.Exceptions;
                    for (int e = 1; e <= oExcps.Count; e++) {
                        Microsoft.Office.Interop.Outlook.Exception oExcp = null;
                        try {
                            oExcp = oExcps[e];
                            if (oExcp.OriginalDate.Date == instanceDate.Date) {
                                try {
                                    log.Debug("Found Outlook exception for " + instanceDate);
                                    if (exceptionIsDeleted(oExcp)) {
                                        log.Debug("This exception is deleted.");
                                        break;
                                    } else {
                                        ai = oExcp.AppointmentItem;
                                        break;
                                    }
                                } catch (System.Exception ex) {
                                    MainForm.Instance.Logboxout(ex.Message);
                                    MainForm.Instance.Logboxout("If this keeps happening, please restart OGCS.");
                                    break;
                                } finally {
                                    OutlookCalendar.ReleaseObject(oExcp);
                                }
                            }
                        } finally {
                            oExcp = (Microsoft.Office.Interop.Outlook.Exception)OutlookCalendar.ReleaseObject(oExcp);
                        }
                    }
                } finally {
                    oExcps = (Exceptions)OutlookCalendar.ReleaseObject(oExcps);
                }
                if (ai == null) log.Warn("Unable to find Outlook exception for " + instanceDate);
            }
        }
        #endregion
        
        #endregion
    }
}
