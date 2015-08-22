using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Google.Apis.Calendar.v3.Data;
using log4net;
using Microsoft.Office.Interop.Outlook;
using NodaTime;

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
            if (!ai.IsRecurring) return null;

            log.Debug("Creating Google iCalendar definition for recurring event.");
            List<String> gPattern = new List<String>();
            gPattern.Add("RRULE:" + buildRrule(ai.GetRecurrencePattern()));

            log.Debug(string.Join("\r\n", gPattern.ToArray()));
            return gPattern;
        }

        public void BuildOutlookPattern(ref Event ev, AppointmentItem ai) {
            if (ev.Recurrence == null) return;

            Dictionary<String, String> ruleBook = explodeRrule(ev.Recurrence);
            log.Fine("Building Outlook recurrence pattern");
            RecurrencePattern oPattern = ai.GetRecurrencePattern();
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
                        oPattern.DayOfWeekMask = getDOWmask(ruleBook);

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
            if (ruleBook.ContainsKey("UNTIL"))
                oPattern.PatternEndDate = DateTime.ParseExact(ruleBook["UNTIL"], "yyyyMMddTHHmmssZ", System.Globalization.CultureInfo.InvariantCulture);
            #endregion

            ai.StartTimeZone = WindowsTimeZone(ev.Start.TimeZone);
            ai.EndTimeZone = WindowsTimeZone(ev.End.TimeZone);
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
                        if (oPattern.Interval == 1) {
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
        public static String IANAtimezone(Microsoft.Office.Interop.Outlook.TimeZone oTZ) {
            //Convert from Windows Timezone to Iana
            //Eg "(UTC) Dublin, Edinburgh, Lisbon, London" => "Europe/London"
            //http://unicode.org/repos/cldr/trunk/common/supplemental/windowsZones.xml
            if (oTZ.ID.Equals("UTC", StringComparison.OrdinalIgnoreCase)) {
                log.Fine("Timezone \"" + oTZ.Name + "\" mapped to \"Etc/UTC\"");
                return "Etc/UTC";
            }

            NodaTime.TimeZones.TzdbDateTimeZoneSource tzDBsource = NodaTime.TimeZones.TzdbDateTimeZoneSource.Default;
            TimeZoneInfo tzi = TimeZoneInfo.FindSystemTimeZoneById(oTZ.ID);
            String tzID = tzDBsource.MapTimeZoneId(tzi);
            log.Fine("Timezone \"" + oTZ.Name + "\" mapped to \"" + tzDBsource.CanonicalIdMap[tzID] + "\"");
            return tzDBsource.CanonicalIdMap[tzID];
        }
        public static Microsoft.Office.Interop.Outlook.TimeZone WindowsTimeZone(string ianaZoneId) {
            Microsoft.Office.Interop.Outlook.TimeZones tzs = OutlookCalendar.Instance.IOutlook.GetTimeZones();
            var utcZones = new[] { "Etc/UTC", "Etc/UCT" };
            if (utcZones.Contains(ianaZoneId, StringComparer.OrdinalIgnoreCase)) {
                log.Fine("Timezone \"" + ianaZoneId + "\" mapped to \"UTC\"");
                return tzs["UTC"];
            }

            var tzdbSource = NodaTime.TimeZones.TzdbDateTimeZoneSource.Default;
            // resolve any link, since the CLDR doesn't necessarily use canonical IDs
            var links = tzdbSource.CanonicalIdMap
              .Where(x => x.Value.Equals(ianaZoneId, StringComparison.OrdinalIgnoreCase))
              .Select(x => x.Key);
            var mappings = tzdbSource.WindowsMapping.MapZones;
            var item = mappings.FirstOrDefault(x => x.TzdbIds.Any(links.Contains));
            if (item == null) {

                log.Warn("Timezone \"" + ianaZoneId + "\" could not find a mapping");
                return null;
            }
            log.Fine("Timezone \"" + ianaZoneId + "\" mapped to \"" + item.WindowsId + "\"");

            return tzs[item.WindowsId];
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
            try {
                log.Debug("Finding Google instance for " + oExcp.AppointmentItem.Start.ToString("dd/MM/yyyy"));
            } catch (System.Exception ex) {
                if (ex.Message == "You changed one of the recurrences of this item, and this instance no longer exists. Close any open items and try again.") {
                    log.Debug("This Outlook recurrence instance has been deleted, but the API is reporting it incorrectly!");
                    log.Debug("Finding Google instance for " + oExcp.OriginalDate.ToString("dd/MM/yyyy"));
                }
            }
            foreach (Event gExcp in googleExceptions) {
                if (gExcp.Status != "cancelled" &&
                    oEntryID == gExcp.ExtendedProperties.Private[GoogleCalendar.oEntryID] &&
                    GoogleCalendar.GoogleTimeFrom(oExcp.OriginalDate) == (gExcp.OriginalStartTime.DateTime ?? gExcp.OriginalStartTime.Date)) {
                    return gExcp;
                }
            }
            log.Debug("Google exception event is not cached. Retrieving all recurring instances...");
            List<Event> gInstances = GoogleCalendar.Instance.GetCalendarEntriesInRecurrence(gRecurringEventID);
            //Add any new exceptions to local cache
            googleExceptions = googleExceptions.Union(gInstances.Where(ev => ev.Sequence > 0)).ToList();

            foreach (Event gInst in gInstances) {
                if (gInst.Status != "cancelled" &&
                    (gInst.OriginalStartTime.DateTime ?? gInst.OriginalStartTime.Date) == GoogleCalendar.GoogleTimeFrom(oExcp.OriginalDate)) {
                    return gInst;
                }
            }
            return null;
        }

        public static void CreateGoogleExceptions(AppointmentItem ai, String recurringEventId) {
            if (!ai.IsRecurring) return;

            log.Debug("Creating Google recurrence exceptions.");
            List<Event> gRecurrences = GoogleCalendar.Instance.GetCalendarEntriesInRecurrence(recurringEventId);
            if (gRecurrences != null) {
                Microsoft.Office.Interop.Outlook.Exceptions exps = ai.GetRecurrencePattern().Exceptions;
                foreach (Microsoft.Office.Interop.Outlook.Exception exp in exps) {
                    String oDate = GoogleCalendar.GoogleTimeFrom(exp.OriginalDate);
                    foreach (Event ev in gRecurrences) {
                        String gDate = ev.OriginalStartTime.DateTime ?? ev.OriginalStartTime.Date;
                        if (exp.Deleted && !ai.AllDayEvent) { //Deleted items get truncated?!
                            gDate = GoogleCalendar.GoogleTimeFrom(DateTime.Parse(gDate).Date);
                        }
                        if (oDate == gDate) {
                            if (exp.Deleted) {
                                MainForm.Instance.Logboxout(GoogleCalendar.GetEventSummary(ev));
                                MainForm.Instance.Logboxout("Recurrence deleted.");
                                ev.Status = "cancelled";
                                GoogleCalendar.Instance.UpdateCalendarEntry_save(ev);
                            } else {
                                int exceptionItemsModified = 0;
                                Event modifiedEv = GoogleCalendar.Instance.UpdateCalendarEntry(exp.AppointmentItem, ev, ref exceptionItemsModified);
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
                    log.Debug("This is a recurring appointment with " + recurrence.Exceptions.Count + " exceptions that will now be iteratively compared.");
                    foreach (Microsoft.Office.Interop.Outlook.Exception oExcp in recurrence.Exceptions) {
                        int excp_itemModified = 0;

                        //Check the exception fall in the date range being synced
                        DateTime oExcp_date = oExcp.OriginalDate;
                        try {
                            oExcp_date = oExcp.Deleted ? oExcp.OriginalDate : oExcp.AppointmentItem.Start;
                        } catch { }
                        if (oExcp_date < DateTime.Today.AddDays(-Settings.Instance.DaysInThePast).Date ||
                            oExcp_date > DateTime.Today.AddDays(+Settings.Instance.DaysInTheFuture).Date) {
                            log.Fine("Exception is outside date range being synced: " + oExcp_date.Date.ToString("dd/MM/yyyy"));
                            continue;
                        }

                        Event gExcp = Recurrence.Instance.GetGoogleInstance(oExcp, ev.RecurringEventId ?? ev.Id, ai.EntryID);
                        if (gExcp != null) {
                            log.Debug("Matching Google Event recurrence found.");
                            try {
                                GoogleCalendar.Instance.UpdateCalendarEntry(oExcp.AppointmentItem, gExcp, ref excp_itemModified, forceCompare: true);
                            } catch (System.Exception ex) {
                                if (oExcp.Deleted || ex.Message == "You changed one of the recurrences of this item, and this instance no longer exists. Close any open items and try again.") {
                                    gExcp.Status = "cancelled";
                                    excp_itemModified++;
                                }
                            }
                            if (excp_itemModified > 0) GoogleCalendar.Instance.UpdateCalendarEntry_save(gExcp);

                        } else {
                            log.Debug("No matching Google Event recurrence found.");
                            log.Debug("This may be because the appointment is changing from single instance to recurring, or it has been deleted.");
                            log.Debug("Either way, don't worry...!");
                        }
                    }
                }
            }
        }
        #endregion

        #region Outlook
        public static Boolean HasExceptions(AppointmentItem ai) {
            return ai.GetRecurrencePattern().Exceptions.Count != 0;
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
                AppointmentItem newAiExcp = getOutlookInstance(oPattern, (gExcp.Status == "cancelled" ? oExcpDate.Date : oExcpDate));
                if (newAiExcp == null) continue;

                if (gExcp.Status != "cancelled") {
                    int itemModified = 0;
                    newAiExcp = OutlookCalendar.Instance.UpdateCalendarEntry(newAiExcp, gExcp, ref itemModified, forceCompare);
                    if (itemModified > 0) newAiExcp.Save();
                } else {
                    MainForm.Instance.Logboxout(OutlookCalendar.GetEventSummary(ai) +"\r\nDeleted.");
                    newAiExcp.Delete();
                }
            }
            if (!ai.Saved) ai.Save();
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
                    if (oExp.OriginalDate == instanceDate) {
                        try {
                            log.Debug("Found Outlook exception for " + instanceDate);
                            if (oExp.Deleted) {
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
