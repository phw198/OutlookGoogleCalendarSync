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
                DateTime utcEnd;
                if (ai.AllDayEvent)
                    utcEnd = rp.PatternEndDate;
                else {
                    DateTime localEnd = rp.PatternEndDate + ai.EndInEndTimeZone.TimeOfDay;
                    utcEnd = TimeZoneInfo.ConvertTimeToUtc(localEnd, TimeZoneInfo.FindSystemTimeZoneById(ai.EndTimeZone.ID));
                }
                gPattern.Add("RRULE:" + buildRrule(rp, utcEnd));
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
                                (ev.Start.DateTime ?? DateTime.Parse(ev.Start.Date)).Day > 28) {
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
            oPattern.PatternStartDate = ev.Start.DateTime ?? DateTime.Parse(ev.Start.Date);
            if (ruleBook.ContainsKey("INTERVAL") && Convert.ToInt16(ruleBook["INTERVAL"]) > 1 && ruleBook["FREQ"] != "YEARLY")
                oPattern.Interval = Convert.ToInt16(ruleBook["INTERVAL"]);
            if (ruleBook.ContainsKey("COUNT"))
                oPattern.Occurrences = Convert.ToInt16(ruleBook["COUNT"]);
            if (ruleBook.ContainsKey("UNTIL")) {
                if (ruleBook["UNTIL"].StartsWith("4500")) {
                    log.Warn("Outlook can't handle end dates this far in the future. Converting to no end date.");
                    oPattern.NoEndDate = true;
                } else {
                    DateTime endDate;
                    if (ruleBook["UNTIL"].Length == 8 && !ruleBook["UNTIL"].EndsWith("Z"))
                        endDate = DateTime.ParseExact(ruleBook["UNTIL"], "yyyyMMdd", System.Globalization.CultureInfo.InvariantCulture).Date;
                    else {
                        endDate = DateTime.ParseExact(ruleBook["UNTIL"], "yyyyMMddTHHmmssZ", System.Globalization.CultureInfo.InvariantCulture, System.Globalization.DateTimeStyles.AdjustToUniversal);
                        endDate = endDate.AddHours(TimezoneDB.GetUtcOffset(ev.End.TimeZone)).Date;
                    }
                    if (endDate < oPattern.PatternStartDate) {
                    log.Debug("PatternStartDate: " + oPattern.PatternStartDate.ToString("yyyyMMddHHmmss"));
                        log.Debug("PatternEndDate:   " + ruleBook["UNTIL"].ToString());
                        String summary = GoogleOgcs.Calendar.GetEventSummary(ev, onlyIfNotVerbose: true);
                        Forms.Main.Instance.Console.Update(summary + "The recurring Google event has an end date <i>before</i> the start date, which Outlook doesn't allow.<br/>" +
                            "The synced Outlook recurrence has been changed to a single occurrence.", Console.Markup.warning);
                        oPattern.Occurrences = 1;
                    } else
                        oPattern.PatternEndDate = endDate;
                }
            }
            if (!ruleBook.ContainsKey("COUNT") && !ruleBook.ContainsKey("UNTIL")) {
                oPattern.NoEndDate = true;
            }
            #endregion
        }

        public void CompareOutlookPattern(Event ev, ref RecurrencePattern aiOpattern, Sync.Direction syncDirection, System.Text.StringBuilder sb, ref int itemModified) {
            if (ev.Recurrence == null) return;
            
            log.Fine("Building a temporary recurrent Appointment generated from Event");
            AppointmentItem evAI = OutlookOgcs.Calendar.Instance.IOutlook.UseOutlookCalendar().Items.Add() as AppointmentItem;
            evAI.Start = ev.Start.DateTime ?? DateTime.Parse(ev.Start.Date); 

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

                if (Sync.Engine.CompareAttribute("Recurrence Type", syncDirection,
                    evOpattern.RecurrenceType.ToString(), aiOpattern.RecurrenceType.ToString(), sb, ref itemModified)) {
                    aiOpattern.RecurrenceType = evOpattern.RecurrenceType;
                }
                if (!skipIntervalCheck && Sync.Engine.CompareAttribute("Recurrence Interval", syncDirection,
                    evOpattern.Interval.ToString(), aiOpattern.Interval.ToString(), sb, ref itemModified)) {
                    aiOpattern.Interval = evOpattern.Interval;
                }
                if (Sync.Engine.CompareAttribute("Recurrence Instance", syncDirection,
                    evOpattern.Instance.ToString(), aiOpattern.Instance.ToString(), sb, ref itemModified)) {
                    aiOpattern.Instance = evOpattern.Instance;
                }
                if (Sync.Engine.CompareAttribute("Recurrence DoW", syncDirection,
                    evOpattern.DayOfWeekMask.ToString(), aiOpattern.DayOfWeekMask.ToString(), sb, ref itemModified)) {
                    aiOpattern.DayOfWeekMask = evOpattern.DayOfWeekMask;
                }
                if (Sync.Engine.CompareAttribute("Recurrence MoY", syncDirection,
                    evOpattern.MonthOfYear.ToString(), aiOpattern.MonthOfYear.ToString(), sb, ref itemModified)) {
                    aiOpattern.MonthOfYear = evOpattern.MonthOfYear;
                }
                if (Sync.Engine.CompareAttribute("Recurrence NoEndDate", syncDirection,
                    evOpattern.NoEndDate, aiOpattern.NoEndDate, sb, ref itemModified)) {
                    aiOpattern.NoEndDate = evOpattern.NoEndDate;
                }
                if (Sync.Engine.CompareAttribute("Recurrence Occurences", syncDirection,
                    evOpattern.Occurrences.ToString(), aiOpattern.Occurrences.ToString(), sb, ref itemModified)) {
                    aiOpattern.Occurrences = evOpattern.Occurrences;
                }
            } finally {
                evOpattern = (RecurrencePattern)OutlookOgcs.Calendar.ReleaseObject(evOpattern);
                evAI.Delete();
                evAI = (AppointmentItem)OutlookOgcs.Calendar.ReleaseObject(evAI);
            }
        }

        private String buildRrule(RecurrencePattern oPattern, DateTime recurrenceEndUtc) {
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
                        //Issue 445: Outlook incorrectly surfaces 12 monthly recurrences as olRecursYearNth, so we'll undo that.
                        //In addition, many apps, indeed even the Google webapp, doesn't display a yearly recurrence rule properly 
                        //despite actually showing the events on the right dates.
                        //So to make OGCS work better with apps that aren't providing full iCal functionality, we'll translate this 
                        //into a monthly recurrence instead.
                        addRule(rrule, "FREQ", "MONTHLY");
                        addRule(rrule, "INTERVAL", oPattern.Interval.ToString());
                        
                        /*Strictly, what we /should/ be doing is:
                        addRule(rrule, "FREQ", "YEARLY");
                        if (oPattern.Interval != 12)
                            addRule(rrule, "INTERVAL", (oPattern.Interval / 12).ToString());
                        addRule(rrule, "BYMONTH", oPattern.MonthOfYear.ToString());
                        */
                        if (oPattern.DayOfWeekMask != (OlDaysOfWeek)127) { //If not every day of week, define which ones
                            addRule(rrule, "BYDAY", string.Join(",", getByDay(oPattern.DayOfWeekMask).ToArray()));
                        }
                        addRule(rrule, "BYSETPOS", (oPattern.Instance == 5) ? "-1" : oPattern.Instance.ToString());
                        break;
                    }
            }
            #endregion

            #region RECURRENCE RANGE
            if (!oPattern.NoEndDate) {
                log.Fine("Checking end date.");
                addRule(rrule, "UNTIL", Recurrence.IANAdate(recurrenceEndUtc));
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
        
        public enum DeletionState {
            Inaccessible,
            Deleted,
            NotDeleted
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
        
        private Event getGoogleInstance(Microsoft.Office.Interop.Outlook.Exception oExcp, String gRecurringEventID, String oEntryID, Boolean dirtyCache) {
            DeletionState oIsDeleted = exceptionIsDeleted(oExcp);
            if (oIsDeleted == DeletionState.Inaccessible) {
                log.Warn("Abandoning fetch of Google instance for inaccessible Outlook exception.");
                return null;
            }
            log.Debug("Finding Google instance for " + (oIsDeleted == DeletionState.Deleted ? "deleted " : "") + "Outlook exception:-");
            log.Debug("  Original date: " + oExcp.OriginalDate.ToString("dd/MM/yyyy"));
            if (oIsDeleted == DeletionState.NotDeleted ) {
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
                        if ((oIsDeleted == DeletionState.NotDeleted &&
                            oExcp.OriginalDate == (gExcp.OriginalStartTime.DateTime ?? DateTime.Parse(gExcp.OriginalStartTime.Date))
                            ) ||
                            (oIsDeleted == DeletionState.Deleted &&
                            oExcp.OriginalDate == (gExcp.OriginalStartTime.DateTime ?? DateTime.Parse(gExcp.OriginalStartTime.Date)).Date
                            )) {
                            return gExcp;
                        }
                    }
                }
                log.Debug("Google exception event is not cached. Retrieving all recurring instances...");
            }
            List<Event> gInstances = GoogleOgcs.Calendar.Instance.GetCalendarEntriesInRecurrence(gRecurringEventID);
            if (gInstances == null) return null;

            //Add any new exceptions to local cache
            googleExceptions = googleExceptions.Union(gInstances.Where(ev => !String.IsNullOrEmpty(ev.RecurringEventId))).ToList();
            foreach (Event gInst in gInstances) {
                if (gInst.RecurringEventId == gRecurringEventID) {
                    if (((oIsDeleted == DeletionState.NotDeleted || (oIsDeleted == DeletionState.Deleted && !oExcp.Deleted)) /* Weirdness when exception is cancelled by organiser but not yet deleted/accepted by recipient */
                        && oExcp.OriginalDate == (gInst.OriginalStartTime.DateTime ?? DateTime.Parse(gInst.OriginalStartTime.Date))
                        ) ||
                        (oIsDeleted == DeletionState.Deleted &&
                        oExcp.OriginalDate == (gInst.OriginalStartTime.DateTime ?? DateTime.Parse(gInst.OriginalStartTime.Date)).Date
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
            Boolean haveMatchingEv = false;
            if (OutlookOgcs.CustomProperty.Exists(ai, OutlookOgcs.CustomProperty.MetadataId.gEventID)) {
                String googleIdValue = OutlookOgcs.CustomProperty.Get(ai, OutlookOgcs.CustomProperty.MetadataId.gEventID);
                String googleCalValue = OutlookOgcs.CustomProperty.Get(ai, OutlookOgcs.CustomProperty.MetadataId.gCalendarId);
                if (googleCalValue == null || googleCalValue == Sync.Engine.Calendar.Instance.Profile.UseGoogleCalendar.Id) {
                    Event ev = GoogleOgcs.Calendar.Instance.GetCalendarEntry(googleIdValue);
                    if (ev != null) {
                        events.Add(ev);
                        haveMatchingEv = true;
                        log.Fine("Found single hard-matched Event.");
                    }
                }
            }
            if (!haveMatchingEv) {
                events = GoogleOgcs.Calendar.Instance.GetCalendarEntriesInRange(ai.Start.Date, ai.Start.Date.AddDays(1));
                if (Sync.Engine.Calendar.Instance.Profile.SyncDirection.Id != Sync.Direction.GoogleToOutlook.Id) {
                    List<AppointmentItem> ais = new List<AppointmentItem>();
                    ais.Add(ai);
                    GoogleOgcs.Calendar.Instance.ReclaimOrphanCalendarEntries(ref events, ref ais, neverDelete: true);
                }
            }
            for (int g = 0; g < events.Count(); g++) {
                Event ev = events[g];
                String gEntryID = GoogleOgcs.CustomProperty.Get(ev, GoogleOgcs.CustomProperty.MetadataId.oEntryId);
                if (haveMatchingEv || !string.IsNullOrEmpty(gEntryID)) {
                    if (haveMatchingEv && string.IsNullOrEmpty(gEntryID)) {
                        return ev;
                    }
                    if (GoogleOgcs.CustomProperty.OutlookIdMissing(ev)) {
                        String compare_oID;
                        if (!string.IsNullOrEmpty(gEntryID) && gEntryID.StartsWith("040000008200E00074C5B7101A82E008")) { //We got a Global ID, not Entry ID
                            compare_oID = OutlookOgcs.Calendar.Instance.IOutlook.GetGlobalApptID(ai);
                        } else {
                            compare_oID = ai.EntryID;
                        }
                        if (haveMatchingEv || gEntryID == compare_oID) {
                            log.Info("Adding Outlook IDs to Master Google Event...");
                            GoogleOgcs.CustomProperty.AddOutlookIDs(ref ev, ai);
                            try {
                                GoogleOgcs.Calendar.Instance.UpdateCalendarEntry_save(ref ev);
                            } catch (System.Exception ex) {
                                OGCSexception.Analyse("Failed saving Outlook IDs to Google Event.", ex, true);
                            }
                            return ev;
                        }
                    } else if (GoogleOgcs.Calendar.ItemIDsMatch(ref ev, ai)) {
                        log.Fine("Found master event.");
                        return ev;
                    }
                } else {
                    log.Debug("Event \"" + ev.Summary + "\" does not have Outlook EntryID stored.");
                    if (GoogleOgcs.Calendar.SignaturesMatch(GoogleOgcs.Calendar.signature(ev), OutlookOgcs.Calendar.signature(ai))) {
                        log.Debug("Master event matched on simple signatures.");
                        return ev;
                    }
                }
            }
            log.Warn("Failed to find master Google event for: " + OutlookOgcs.Calendar.GetEventSummary(ai));
            return null;
        }

        public static void CreateGoogleExceptions(AppointmentItem ai, String recurringEventId) {
            if (!ai.IsRecurring) return;

            log.Debug("Creating Google recurrence exceptions.");
            List<Event> gRecurrences = GoogleOgcs.Calendar.Instance.GetCalendarEntriesInRecurrence(recurringEventId);
            if (gRecurrences == null) return;

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
                            DateTime gDate = ev.OriginalStartTime.DateTime ?? DateTime.Parse(ev.OriginalStartTime.Date);
                            DeletionState isDeleted = exceptionIsDeleted(oExcp);
                            if (isDeleted == DeletionState.Inaccessible) {
                                log.Warn("Abandoning creation of Google recurrence exception as Outlook exception is inaccessible.");
                                return;
                            }
                            if (isDeleted == DeletionState.Deleted && !ai.AllDayEvent) { //Deleted items get truncated?!
                                gDate = gDate.Date;
                            }
                            if (oExcp.OriginalDate == gDate) {
                                if (isDeleted == DeletionState.Deleted) {
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
                                DateTime oExcp_currDate;

                                //Check the exception falls in the date range being synced
                                DeletionState oIsDeleted = exceptionIsDeleted(oExcp);
                                String logDeleted = "";
                                if (oIsDeleted != DeletionState.NotDeleted) {
                                    logDeleted = " " + oIsDeleted.ToString().ToLower() + " and";
                                    oExcp_currDate = oExcp.OriginalDate;
                                } else {
                                    aiExcp = oExcp.AppointmentItem;
                                    oExcp_currDate = aiExcp.Start;
                                    aiExcp = (AppointmentItem)OutlookOgcs.Calendar.ReleaseObject(aiExcp);
                                }
                                if (oExcp_currDate < Sync.Engine.Calendar.Instance.Profile.SyncStart.Date || oExcp_currDate > Sync.Engine.Calendar.Instance.Profile.SyncEnd.Date) {
                                    log.Fine("Exception is" + logDeleted + " outside date range being synced: " + oExcp_currDate.Date.ToString("dd/MM/yyyy"));
                                    continue;
                                } else if (oIsDeleted == DeletionState.Inaccessible) {
                                    log.Warn("Exception is" + logDeleted + " cannot be synced: " + oExcp_currDate.Date.ToString("dd/MM/yyyy"));
                                    continue;
                                }

                                Event gExcp = Recurrence.Instance.getGoogleInstance(oExcp, ev.RecurringEventId ?? ev.Id, OutlookOgcs.Calendar.Instance.IOutlook.GetGlobalApptID(ai), dirtyCache);
                                if (gExcp != null) {
                                    log.Debug("Matching Google Event recurrence found.");
                                    if (gExcp.Status == "cancelled") {
                                        log.Debug("It is deleted in Google, which " + (oIsDeleted == DeletionState.Deleted ? "matches" : "does not match") + " Outlook.");
                                        if (oIsDeleted == DeletionState.NotDeleted) {
                                            log.Warn("Outlook is NOT deleted though - a mismatch has occurred somehow!");
                                            String syncDirectionTip = (Sync.Engine.Calendar.Instance.Profile.SyncDirection.Id == Sync.Direction.Bidirectional.Id) ? "<br/><i>Ensure you <b>first</b> set OGCS to one-way sync O->G.</i>" : "";
                                            Forms.Main.Instance.Console.Update(OutlookOgcs.Calendar.GetEventSummary(ai) + "<br/>" +
                                                "The occurrence on " + oExcp.OriginalDate.ToShortDateString() + " does not exist in Google, but does in Outlook.<br/>" +
                                                "This can happen if, for example, you declined the occurrence (which is synced to Google) and proposed a new time that is subsequently accepted by the organiser.<br/>" +
                                                "<u>Suggested fix</u>: delete the entire series in Google and let OGCS recreate it." + syncDirectionTip, Console.Markup.warning);
                                        }
                                        continue;
                                    } else if (oIsDeleted == DeletionState.Deleted && gExcp.Status != "cancelled") {
                                        gExcp.Status = "cancelled";
                                        log.Debug("Exception deleted.");
                                        excp_itemModified++;
                                    } else {
                                        try {
                                            aiExcp = oExcp.AppointmentItem;
                                            //Force a compare of the exception if both G and O have been modified in last 24 hours
                                            TimeSpan modifiedDiff = (TimeSpan)(gExcp.Updated - aiExcp.LastModificationTime);
                                            log.Fine("Modification time difference (in days) between G and O exception: " + modifiedDiff);
                                            Boolean forceCompare = modifiedDiff < TimeSpan.FromDays(1);
                                            GoogleOgcs.Calendar.Instance.UpdateCalendarEntry(aiExcp, gExcp, ref excp_itemModified, forceCompare);
                                            if (forceCompare && excp_itemModified == 0 && DateTime.Now > aiExcp.LastModificationTime.AddDays(1)) {
                                                GoogleOgcs.CustomProperty.SetOGCSlastModified(ref gExcp);
                                                try {
                                                    log.Debug("Doing a dummy update in order to update the last modified date of Google recurring series exception.");
                                                    GoogleOgcs.Calendar.Instance.UpdateCalendarEntry_save(ref gExcp);
                                                } catch (System.Exception ex) {
                                                    OGCSexception.Analyse("Dummy update of unchanged exception for Google recurring series failed.", ex);
                                                }
                                                continue;
                                            }
                                        } catch (System.Exception ex) {
                                            log.Error(ex.Message);
                                            log.Error(ex.StackTrace);
                                            throw;
                                        }
                                    }
                                    if (excp_itemModified > 0) {
                                        try {
                                            GoogleOgcs.Calendar.Instance.UpdateCalendarEntry_save(ref gExcp);
                                        } catch (System.Exception ex) {
                                            Forms.Main.Instance.Console.UpdateWithError("Updated event exception failed to save.", ex);
                                            OGCSexception.Analyse(ex, true);
                                            if (OgcsMessageBox.Show("Updated Google event exception failed to save. Continue with synchronisation?", "Sync item failed", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                                                continue;
                                            else {
                                                throw new UserCancelledSyncException("User chose not to continue sync.");
                                            }
                                        }
                                    }
                                } else {
                                    log.Warn("No matching Google Event recurrence found.");
                                    if (oIsDeleted == DeletionState.Deleted) log.Debug("The Outlook appointment is deleted, so not a problem.");
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

        private static DeletionState exceptionIsDeleted(Microsoft.Office.Interop.Outlook.Exception oExcp) {
            if (oExcp.Deleted) return DeletionState.Deleted;
            AppointmentItem ai = null;
            try {
                ai = oExcp.AppointmentItem;
                return DeletionState.NotDeleted;
            } catch (System.Exception ex) {
                OGCSexception.LogAsFail(ref ex);
                String originalDate = oExcp.OriginalDate.ToString("dd/MM/yyyy");
                if (ex.Message == "You changed one of the recurrences of this item, and this instance no longer exists. Close any open items and try again.") {
                    OGCSexception.Analyse("This Outlook recurrence instance on " + originalDate + " has become inaccessible, probably due to caching", ex);
                } else {
                    OGCSexception.Analyse("Error when determining if Outlook recurrence on " + originalDate + " is deleted or not.", ex);
                }
                return DeletionState.Inaccessible;
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
                    DateTime oExcpDate = gExcp.OriginalStartTime.DateTime ?? DateTime.Parse(gExcp.OriginalStartTime.Date);
                    log.Fine("Found Google exception for " + oExcpDate.ToString());

                    AppointmentItem newAiExcp = null;
                    try {
                        getOutlookInstance(oPattern, oExcpDate, ref newAiExcp);
                        if (newAiExcp == null) continue;

                        if (gExcp.Status == "cancelled") {
                            Forms.Main.Instance.Console.Update(OutlookOgcs.Calendar.GetEventSummary(newAiExcp) + "<br/>Deleted.", Console.Markup.calendar);
                            newAiExcp.Delete();

                        } else if (Sync.Engine.Calendar.Instance.Profile.ExcludeDeclinedInvites && gExcp.Attendees != null && gExcp.Attendees.Count(a => a.Self == true && a.ResponseStatus == "declined") == 1) {
                            Forms.Main.Instance.Console.Update(OutlookOgcs.Calendar.GetEventSummary(newAiExcp) + "<br/>Declined.", Console.Markup.calendar);
                            newAiExcp.Delete();

                        } else {
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
                                    DeletionState isDeleted = exceptionIsDeleted(oExcp);
                                    if (isDeleted == DeletionState.Inaccessible) {
                                        log.Warn("This exception is inaccessible.");
                                        return;
                                    } else if (isDeleted == DeletionState.NotDeleted) {
                                        log.Debug("This exception is deleted.");
                                        return;
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
