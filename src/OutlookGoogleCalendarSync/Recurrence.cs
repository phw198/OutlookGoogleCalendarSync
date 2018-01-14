using Google.Apis.Calendar.v3.Data;
using log4net;
using Microsoft.Office.Interop.Outlook;
using System;
using System.Collections.Generic;
using System.Linq;
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
                rp = (RecurrencePattern)OutlookOgcs.Calendar.ReleaseObject(rp);
            }
            log.Debug(string.Join("\r\n", gPattern.ToArray()));
            return gPattern;
        }

        public void BuildOutlookPattern(Event ev, AppointmentItem ai) {
            RecurrencePattern ignore;
            buildOutlookPattern(ev, ai, out ignore);
            ignore = (RecurrencePattern)OutlookOgcs.Calendar.ReleaseObject(ignore);
        }

        private void buildOutlookPattern(Event ev, AppointmentItem ai, out RecurrencePattern oPattern) {
            if (ev.Recurrence == null) { oPattern = null; return; }

            Dictionary<String, String> ruleBook = explodeRrule(ev.Recurrence);
            if (ruleBook == null) {
                throw new ApplicationException("WARNING: The recurrence pattern is not compatible with Outlook. This event cannot be synced.");
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
                        if (ruleBook.ContainsKey("INTERVAL") && Convert.ToInt16(ruleBook["INTERVAL"]) > 1)
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
            ai = OutlookOgcs.Calendar.Instance.IOutlook.WindowsTimeZone_set(ai, ev);
            oPattern.PatternStartDate = DateTime.Parse(ev.Start.Date ?? ev.Start.DateTime);
            if (ruleBook.ContainsKey("INTERVAL") && Convert.ToInt16(ruleBook["INTERVAL"]) > 1 && ruleBook["FREQ"] != "YEARLY")
                oPattern.Interval = Convert.ToInt16(ruleBook["INTERVAL"]);
            if (ruleBook.ContainsKey("COUNT"))
                oPattern.Occurrences = Convert.ToInt16(ruleBook["COUNT"]);
            if (ruleBook.ContainsKey("UNTIL")) {
                if (ruleBook["UNTIL"].StartsWith("4500")) {
                    log.Warn("Outlook can't handle end dates this far in the future. Converting to no end date.");
                    oPattern.NoEndDate = true;
                } else {
                    oPattern.PatternEndDate = DateTime.ParseExact(ruleBook["UNTIL"].ToString().Substring(0, 8), "yyyyMMdd", System.Globalization.CultureInfo.InvariantCulture).Date;
                }
            }
            if (!ruleBook.ContainsKey("COUNT") && !ruleBook.ContainsKey("UNTIL")) {
                oPattern.NoEndDate = true;
            }
            #endregion
        }

        public void CompareOutlookPattern(Event ev, ref RecurrencePattern aiOpattern, SyncDirection syncDirection, System.Text.StringBuilder sb, ref int itemModified) {
            if (ev.Recurrence == null) return;
            
            log.Fine("Building a temporary recurrent Appointment generated from Event");
            AppointmentItem evAI = OutlookOgcs.Calendar.Instance.IOutlook.UseOutlookCalendar().Items.Add() as AppointmentItem;
            evAI.Start = DateTime.Parse(ev.Start.Date ?? ev.Start.DateTime); 

            RecurrencePattern evOpattern = null;
            try {
                buildOutlookPattern(ev, evAI, out evOpattern);
                log.Fine("Comparing Google recurrence to Outlook equivalent");

                //Some versions of Outlook are erroring when 2-way syncing weekday recurring series.
                //Even though Outlook has Interval of zero, which is illegal, when this is updated, it won't save. Issue #398
                Boolean skipIntervalCheck = false;
                if (aiOpattern.RecurrenceType == OlRecurrenceType.olRecursWeekly && aiOpattern.DayOfWeekMask == getDOWmask("BYDAY=MO,TU,WE,TH,FR") && aiOpattern.Interval == 0 &&
                    evOpattern.RecurrenceType == aiOpattern.RecurrenceType && evOpattern.DayOfWeekMask == aiOpattern.DayOfWeekMask && evOpattern.Interval == 1)
                    skipIntervalCheck = true;

                if (Forms.Main.CompareAttribute("Recurrence Type", syncDirection,
                    evOpattern.RecurrenceType.ToString(), aiOpattern.RecurrenceType.ToString(), sb, ref itemModified)) {
                    aiOpattern.RecurrenceType = evOpattern.RecurrenceType;
                }
                if (!skipIntervalCheck && Forms.Main.CompareAttribute("Recurrence Interval", syncDirection,
                    evOpattern.Interval.ToString(), aiOpattern.Interval.ToString(), sb, ref itemModified)) {
                    aiOpattern.Interval = evOpattern.Interval;
                }
                if (Forms.Main.CompareAttribute("Recurrence Instance", syncDirection,
                    evOpattern.Instance.ToString(), aiOpattern.Instance.ToString(), sb, ref itemModified)) {
                    aiOpattern.Instance = evOpattern.Instance;
                }
                if (Forms.Main.CompareAttribute("Recurrence DoW", syncDirection,
                    evOpattern.DayOfWeekMask.ToString(), aiOpattern.DayOfWeekMask.ToString(), sb, ref itemModified)) {
                    aiOpattern.DayOfWeekMask = evOpattern.DayOfWeekMask;
                }
                if (Forms.Main.CompareAttribute("Recurrence MoY", syncDirection,
                    evOpattern.MonthOfYear.ToString(), aiOpattern.MonthOfYear.ToString(), sb, ref itemModified)) {
                    aiOpattern.MonthOfYear = evOpattern.MonthOfYear;
                }
                if (Forms.Main.CompareAttribute("Recurrence NoEndDate", syncDirection,
                    evOpattern.NoEndDate, aiOpattern.NoEndDate, sb, ref itemModified)) {
                    aiOpattern.NoEndDate = evOpattern.NoEndDate;
                }
                if (Forms.Main.CompareAttribute("Recurrence Occurences", syncDirection,
                    evOpattern.Occurrences.ToString(), aiOpattern.Occurrences.ToString(), sb, ref itemModified)) {
                    aiOpattern.Occurrences = evOpattern.Occurrences;
                }
            } finally {
                evOpattern = (RecurrencePattern)OutlookOgcs.Calendar.ReleaseObject(evOpattern);
                evAI.Delete();
                evAI = (AppointmentItem)OutlookOgcs.Calendar.ReleaseObject(evAI);
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
                        addRule(rrule, "BYDAY", string.Join(",", getByDay(oPattern.DayOfWeekMask).ToArray()));
                        addRule(rrule, "BYSETPOS", (oPattern.Instance == 5) ? "-1" : oPattern.Instance.ToString());
                        break;
                    }

                case OlRecurrenceType.olRecursYearly: {
                        addRule(rrule, "FREQ", "YEARLY");
                        //Google interval is years, Outlook is months
                        if (oPattern.Interval != 12)
                            addRule(rrule, "INTERVAL", (oPattern.Interval / 12).ToString());
                        break;
                    }

                case OlRecurrenceType.olRecursYearNth: {
                        addRule(rrule, "FREQ", "YEARLY");
                        if (oPattern.DayOfWeekMask != (OlDaysOfWeek)127) { //If not every day of week, define which ones
                            addRule(rrule, "BYDAY", string.Join(",", getByDay(oPattern.DayOfWeekMask).ToArray()));
                        }
                        addRule(rrule, "BYMONTH", oPattern.MonthOfYear.ToString());
                        addRule(rrule, "BYSETPOS", (oPattern.Instance == 5) ? "-1" : oPattern.Instance.ToString());
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
            return dt.ToUniversalTime().ToString("yyyyMMddTHHmmssZ");
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
            //    List<Event> gInstances = GoogleOgcs.Calendar.Instance.GetCalendarEntriesInRecurrence(ev.RecurringEventId ?? ev.Id);
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
                    ai = (AppointmentItem)OutlookOgcs.Calendar.ReleaseObject(ai);
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
            List<Event> gInstances = GoogleOgcs.Calendar.Instance.GetCalendarEntriesInRecurrence(gRecurringEventID);
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
            log.Fine("Found a master Outlook recurring item outside sync date range: " + OutlookOgcs.Calendar.GetEventSummary(ai));
            List<Event> events = new List<Event>();
            String googleIdValue;
            Boolean haveMatchingEv = false;
            if (OutlookOgcs.Calendar.GetOGCSproperty(ai, OutlookOgcs.Calendar.MetadataId.gEventID, out googleIdValue)) {
                Event ev = GoogleOgcs.Calendar.Instance.GetCalendarEntry(googleIdValue);
                if (ev != null) {
                    events.Add(ev);
                    haveMatchingEv = true;
                    log.Fine("Found single hard-matched Event.");
                }
            }
            if (!haveMatchingEv) {
                events = GoogleOgcs.Calendar.Instance.GetCalendarEntriesInRange(ai.Start.Date, ai.Start.Date.AddDays(1));
                List<AppointmentItem> ais = new List<AppointmentItem>();
                ais.Add(ai);
                GoogleOgcs.Calendar.Instance.ReclaimOrphanCalendarEntries(ref events, ref ais, neverDelete: true);
            }
            for (int g = 0; g < events.Count(); g++) {
                String gEntryID = null;
                Event ev = events[g];
                if (haveMatchingEv || GoogleOgcs.Calendar.GetOGCSproperty(ev, GoogleOgcs.Calendar.MetadataId.oEntryId, out gEntryID)) {
                    if (GoogleOgcs.Calendar.OutlookIdMissing(ev)) {
                        String compare_oID;
                        if (gEntryID != null && gEntryID.StartsWith("040000008200E00074C5B7101A82E008")) { //We got a Global ID, not Entry ID
                            compare_oID = OutlookOgcs.Calendar.Instance.IOutlook.GetGlobalApptID(ai);
                        } else {
                            compare_oID = ai.EntryID;
                        }
                        if (haveMatchingEv || gEntryID == compare_oID) {
                            log.Info("Adding Outlook IDs to Master Google Event...");
                            GoogleOgcs.Calendar.AddOutlookIDs(ref ev, ai);
                            try {
                                GoogleOgcs.Calendar.Instance.UpdateCalendarEntry_save(ref ev);
                            } catch (System.Exception ex) {
                                log.Error("Failed saving Outlook IDs to Google Event.");
                                OGCSexception.Analyse(ex, true);
                            }
                            return ev;
                        }
                    } else if (GoogleOgcs.Calendar.ItemIDsMatch(ref ev, ai)) {
                        log.Fine("Found master event.");
                        return ev;
                    }
                } else {
                    log.Debug("Event \"" + ev.Summary + "\" did not have Outlook EntryID stored.");
                }
            }
            log.Warn("Failed to find master Google event for: " + OutlookOgcs.Calendar.GetEventSummary(ai));
            return null;
        }

        public static void CreateGoogleExceptions(AppointmentItem ai, String recurringEventId) {
            if (!ai.IsRecurring) return;

            log.Debug("Creating Google recurrence exceptions.");
            List<Event> gRecurrences = GoogleOgcs.Calendar.Instance.GetCalendarEntriesInRecurrence(recurringEventId);
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
                                    gDate = GoogleOgcs.Calendar.GoogleTimeFrom(DateTime.Parse(gDate).Date);
                                }
                                if (oExcp.OriginalDate == DateTime.Parse(gDate)) {
                                    if (isDeleted) {
                                        Forms.Main.Instance.Console.Update(GoogleOgcs.Calendar.GetEventSummary(ev), Console.Markup.calendar);
                                        Forms.Main.Instance.Console.Update("Recurrence deleted.");
                                        ev.Status = "cancelled";
                                        GoogleOgcs.Calendar.Instance.UpdateCalendarEntry_save(ref ev);
                                    } else {
                                        int exceptionItemsModified = 0;
                                        Event modifiedEv = GoogleOgcs.Calendar.Instance.UpdateCalendarEntry(oExcp.AppointmentItem, ev, ref exceptionItemsModified, forceCompare: true);
                                        if (exceptionItemsModified > 0) {
                                            GoogleOgcs.Calendar.Instance.UpdateCalendarEntry_save(ref modifiedEv);
                                        }
                                    }
                                    break;
                                }
                            }
                        } finally {
                            oExcp = (Microsoft.Office.Interop.Outlook.Exception)OutlookOgcs.Calendar.ReleaseObject(oExcp);
                        }
                    }
                } finally {
                    excps = (Exceptions)OutlookOgcs.Calendar.ReleaseObject(excps);
                    rp = (RecurrencePattern)OutlookOgcs.Calendar.ReleaseObject(rp);
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
                        log.Debug(OutlookOgcs.Calendar.GetEventSummary(ai));
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
                                    aiExcp = (AppointmentItem)OutlookOgcs.Calendar.ReleaseObject(aiExcp);
                                }
                            
                                if (oExcp_currDate < Settings.Instance.SyncStart.Date || oExcp_currDate > Settings.Instance.SyncEnd.Date) {
                                    log.Fine("Exception is" + logDeleted + " outside date range being synced: " + oExcp_currDate.Date.ToString("dd/MM/yyyy"));
                                    continue;
                                }

                                Event gExcp = Recurrence.Instance.getGoogleInstance(ref oExcp, ev.RecurringEventId ?? ev.Id, OutlookOgcs.Calendar.Instance.IOutlook.GetGlobalApptID(ai), dirtyCache);
                                if (gExcp != null) {
                                    log.Debug("Matching Google Event recurrence found.");
                                    if (gExcp.Status == "cancelled") {
                                        log.Debug("It is deleted in Google, so cannot compare items.");
                                        if (!oIsDeleted) log.Warn("Outlook is NOT deleted though - a mismatch has occurred somehow!");
                                        continue;
                                    } else if (oIsDeleted && gExcp.Status != "cancelled") {
                                        gExcp.Status = "cancelled";
                                        log.Debug("Exception deleted.");
                                        excp_itemModified++;
                                    } else {
                                        try {
                                            aiExcp = oExcp.AppointmentItem;
                                            GoogleOgcs.Calendar.Instance.UpdateCalendarEntry(aiExcp, gExcp, ref excp_itemModified);
                                        } catch (System.Exception ex) {
                                            log.Error(ex.Message);
                                            log.Error(ex.StackTrace);
                                            throw ex;
                                        }
                                    }
                                    if (excp_itemModified > 0) {
                                        try {
                                            GoogleOgcs.Calendar.Instance.UpdateCalendarEntry_save(ref gExcp);
                                        } catch (System.Exception ex) {
                                            Forms.Main.Instance.Console.Update("Updated event exception failed to save.<br/>" + ex.Message, Console.Markup.error);
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
                                aiExcp = (AppointmentItem)OutlookOgcs.Calendar.ReleaseObject(aiExcp);
                                oExcp = (Microsoft.Office.Interop.Outlook.Exception)OutlookOgcs.Calendar.ReleaseObject(oExcp);
                            }
                        }
                    }
                } finally {
                    excps = (Exceptions)OutlookOgcs.Calendar.ReleaseObject(excps);
                    rp = (RecurrencePattern)OutlookOgcs.Calendar.ReleaseObject(rp);
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
                excps = (Exceptions)OutlookOgcs.Calendar.ReleaseObject(excps);
                rp = (RecurrencePattern)OutlookOgcs.Calendar.ReleaseObject(rp);
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
                ai = (AppointmentItem)OutlookOgcs.Calendar.ReleaseObject(ai);
            }
        }

        public void CreateOutlookExceptions(ref AppointmentItem ai, Event ev) {
            processOutlookExceptions(ref ai, ev, forceCompare: true);
        }
        public void UpdateOutlookExceptions(ref AppointmentItem ai, Event ev, Boolean forceCompare) {
            processOutlookExceptions(ref ai, ev, forceCompare);
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
                            OutlookOgcs.Calendar.Instance.UpdateCalendarEntry(ref newAiExcp, gExcp, ref itemModified, forceCompare);
                            if (itemModified > 0) {
                                try {
                                    newAiExcp.Save();
                                } catch (System.Exception ex) {
                                    OGCSexception.Analyse(ex);
                                    if (ex.Message == "Cannot save this item.") {
                                        Forms.Main.Instance.Console.Update("Uh oh! Outlook wasn't able to save this recurrence exception! " +
                                            "You may have two occurences on the same day, which it doesn't allow.", Console.Markup.warning);
                                    }
                                }
                            }
                        } else {
                            Forms.Main.Instance.Console.Update(OutlookOgcs.Calendar.GetEventSummary(newAiExcp) + "<br/>Deleted.", Console.Markup.calendar);
                            newAiExcp.Delete();
                        }
                    } finally {
                        newAiExcp = (AppointmentItem)OutlookOgcs.Calendar.ReleaseObject(newAiExcp);
                    }
                }
            } finally {
                oPattern = (RecurrencePattern)OutlookOgcs.Calendar.ReleaseObject(oPattern);
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
                                    Forms.Main.Instance.Console.Update(ex.Message + "<br/>If this keeps happening, please restart OGCS.", Console.Markup.error);
                                    break;
                                } finally {
                                    OutlookOgcs.Calendar.ReleaseObject(oExcp);
                                }
                            }
                        } finally {
                            oExcp = (Microsoft.Office.Interop.Outlook.Exception)OutlookOgcs.Calendar.ReleaseObject(oExcp);
                        }
                    }
                } finally {
                    oExcps = (Exceptions)OutlookOgcs.Calendar.ReleaseObject(oExcps);
                }
                if (ai == null) log.Warn("Unable to find Outlook exception for " + instanceDate);
            }
        }
        #endregion
        
        #endregion
    }
}
