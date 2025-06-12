using Google.Apis.Calendar.v3.Data;
using log4net;
using Microsoft.Office.Interop.Outlook;
using OutlookGoogleCalendarSync.Extensions;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;
using Ogcs = OutlookGoogleCalendarSync;

namespace OutlookGoogleCalendarSync.Google {
    public class Recurrence {
        /*
         * Recurrence rule standards for iCalendar: http://www.ietf.org/rfc/rfc2445 
         */
        private static readonly ILog log = LogManager.GetLogger(typeof(Recurrence));

        public static List<String> BuildGooglePattern(AppointmentItem ai, Event ev) {
            if (!ai.IsRecurring || ai.RecurrenceState != OlRecurrenceState.olApptMaster) return null;

            log.Debug("Creating Google iCalendar definition for recurring event.");
            List<String> gPattern = new List<String>();
            RecurrencePattern rp = null;
            try {
                rp = ai.GetRecurrencePattern();
                System.DateTime utcEnd;
                if (ai.AllDayEvent)
                    utcEnd = rp.PatternEndDate;
                else {
                    System.DateTime localEnd = rp.PatternEndDate + Outlook.Calendar.Instance.IOutlook.GetEndInEndTimeZone(ai).TimeOfDay;
                    utcEnd = TimeZoneInfo.ConvertTimeToUtc(localEnd, TimeZoneInfo.FindSystemTimeZoneById(Outlook.Calendar.Instance.IOutlook.GetEndTimeZoneID(ai)));
                }
                gPattern.Add("RRULE:" + buildRrule(rp, utcEnd));
            } finally {
                rp = (RecurrencePattern)Outlook.Calendar.ReleaseObject(rp);
            }
            log.Debug(string.Join("\r\n", gPattern.ToArray()));
            return gPattern;
        }

        private static String buildRrule(RecurrencePattern oPattern, System.DateTime recurrenceEndUtc) {
            log.Fine("Building RRULE");
            Dictionary<String, String> rrule = new Dictionary<String, String>();
            
            #region RECURRENCE PATTERN
            log.Fine("Determining pattern for frequency " + oPattern.RecurrenceType.ToString() + ".");

            switch (oPattern.RecurrenceType) {
                case OlRecurrenceType.olRecursDaily: {
                        addRule(rrule, "FREQ", "DAILY");
                        setInterval(rrule, oPattern.Interval);
                        break;
                    }

                case OlRecurrenceType.olRecursWeekly: {
                        addRule(rrule, "FREQ", "WEEKLY");
                        setInterval(rrule, oPattern.Interval);
                        if ((oPattern.DayOfWeekMask & (oPattern.DayOfWeekMask - 1)) != 0) { //is not a power of 2 (i.e. not just a single day) 
                            // Need to work out "BY" pattern
                            // Eg "BYDAY=MO,TU,WE,TH,FR"
                            addRule(rrule, "BYDAY", string.Join(",", getByDay(oPattern.DayOfWeekMask)));
                        }
                        break;
                    }

                case OlRecurrenceType.olRecursMonthly: {
                        addRule(rrule, "FREQ", "MONTHLY");
                        setInterval(rrule, oPattern.Interval);
                        //Outlook and Google interpret days of month that don't alway exist, eg 31st, differently - though it's not explicitly defined
                        //Outlook: Picks last day of month (SKIP=BACKWARD); Google: Skips that month (SKIP=OMIT)
                        //We'll adopt Outlook's definition
                        if (oPattern.PatternStartDate.Day > 28) {
                            Google.Recurrence.addRule(rrule, "RSCALE", "GREGORIAN");
                            Google.Recurrence.addRule(rrule, "BYMONTHDAY", oPattern.PatternStartDate.Day.ToString());
                            Google.Recurrence.addRule(rrule, "SKIP", "BACKWARD");
                        }
                        break;
                    }

                case OlRecurrenceType.olRecursMonthNth: {
                        addRule(rrule, "FREQ", "MONTHLY");
                        setInterval(rrule, oPattern.Interval);
                        List<String> byDay = getByDay(oPattern.DayOfWeekMask);
                        if (byDay.Count == 1) {
                            String byDayRelative = (oPattern.Instance == 5) ? "-1" : oPattern.Instance.ToString();
                            addRule(rrule, "BYDAY", byDayRelative + string.Join(",", byDay));
                        } else {
                            addRule(rrule, "BYDAY", string.Join(",", byDay));
                            addRule(rrule, "BYSETPOS", (oPattern.Instance == 5) ? "-1" : oPattern.Instance.ToString());
                        }
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
                        List<String> byDay = getByDay(oPattern.DayOfWeekMask);
                        if (byDay.Count == 1) {
                            String byDayRelative = (oPattern.Instance == 5) ? "-1" : oPattern.Instance.ToString();
                            addRule(rrule, "BYDAY", byDayRelative + string.Join(",", byDay));
                        } else {
                            if (byDay.Count != 7) //If not every day of week, define which ones
                                addRule(rrule, "BYDAY", string.Join(",", byDay));
                            addRule(rrule, "BYSETPOS", (oPattern.Instance == 5) ? "-1" : oPattern.Instance.ToString());
                        }
                        break;
                    }
            }
            #endregion

            #region RECURRENCE RANGE
            if (!oPattern.NoEndDate) {
                log.Fine("Checking end date.");
                addRule(rrule, "UNTIL", IANAdate(recurrenceEndUtc));
            }
            //Outlook converts numbered occurrences to an end date, so there is never a need to sync a COUNT RRule.
            #endregion

            return string.Join(";", rrule.Select(x => x.Key + "=" + x.Value).ToArray());
        }

        public static void CompareGooglePattern(List<String> oRrules, Event ev, System.Text.StringBuilder sb, ref int itemModified) {
            if (ev.Recurrence != null) {
                for (int r = 0; r < ev.Recurrence.Count; r++) {
                    String rrule = ev.Recurrence[r];
                    if (rrule.StartsWith("RRULE:")) {
                        log.Fine("Google recurrence = " + rrule);
                        if (oRrules != null) {
                            String[] gRrule_bits = rrule.TrimStart("RRULE:".ToCharArray()).Split(';');
                            String[] oRrule_bits = oRrules.First().TrimStart("RRULE:".ToCharArray()).Split(';');
                            if (gRrule_bits.Count() != oRrule_bits.Count()) {
                                if (Sync.Engine.CompareAttribute("Recurrence", Sync.Direction.OutlookToGoogle, rrule, oRrules.First(), sb, ref itemModified)) {
                                    ev.Recurrence[r] = oRrules.First();
                                    if (gRrule_bits.Contains("FREQ=YEARLY") && gRrule_bits.Contains("INTERVAL=1")) {
                                        //Some applications can put in superflous yearly interval, which when removed does not save, resulting in repeated "updates"
                                        //Workaround is to convert to 12 monthly; subquent sync would then revert back to yearly without unnecessary interval
                                        ev.Recurrence[r] = ev.Recurrence[r].Replace("YEARLY", "MONTHLY") + ";INTERVAL=12";
                                    }
                                    break;
                                }
                            }
                            foreach (String oRrule_bit in oRrule_bits) {
                                if (!rrule.Contains(oRrule_bit)) {
                                    if (Sync.Engine.CompareAttribute("Recurrence", Sync.Direction.OutlookToGoogle, rrule, oRrules.First(), sb, ref itemModified)) {
                                        ev.Recurrence[r] = oRrules.First();
                                        break;
                                    }
                                }
                            }
                        } else {
                            log.Debug("Converting to non-recurring event.");
                            Sync.Engine.CompareAttribute("Recurrence", Sync.Direction.OutlookToGoogle, rrule, null, sb, ref itemModified);
                            ev.Recurrence[r] = null;
                        }
                        break;
                    }
                }
            } else {
                if (oRrules != null && ev.RecurringEventId == null) {
                    if (!(ev.Creator.Self ?? (ev.Creator.Email == Settings.Instance.GaccountEmail))) {
                        log.Warn("Cannot convert Event organised by another to a recurring series.");
                    } else {
                        log.Debug("Converting to recurring event.");
                        Sync.Engine.CompareAttribute("Recurrence", Sync.Direction.OutlookToGoogle, null, oRrules.First(), sb, ref itemModified);
                        ev.Recurrence = oRrules;
                    }
                }
            }
        }

        public static Dictionary<String, String> ExplodeRrule(IList<String> allRules) {
            log.Fine("Analysing Event RRULEs...");
            foreach (String aRule in allRules) {
                String rrule = null;
                System.Text.RegularExpressions.Regex rgx = new System.Text.RegularExpressions.Regex(@"^RRULE(;([\w\-])+=.+?):", System.Text.RegularExpressions.RegexOptions.IgnoreCase);
                System.Text.RegularExpressions.MatchCollection matches = rgx.Matches(aRule);
                if (matches.Count > 0) {
                    log.Warn("The following custom RRULE parameter(s) have been ignored: " + matches[0].Result("$1"));
                    rrule = aRule.Replace(matches[0].Result("$1"), String.Empty);
                } else {
                    rrule = aRule;
                }
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

        internal static void addRule(Dictionary<string, string> ruleBook, string key, string value) {
            ruleBook.Add(key, value);
            log.Fine(ruleBook.Last().Value);
        }

        internal static void setInterval(Dictionary<String, String> rrule, int interval) {
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

        internal static String IANAdate(System.DateTime dt) {
            return dt.ToString("yyyyMMddTHHmmssZ");
        }

        #region Exceptions
        private static List<Event> googleExceptions;
        public static List<Event> GoogleExceptions {
            get { return googleExceptions; }
        }
        public static void GoogleExceptionsReset() {
            googleExceptions = new List<Event>();
        }

        public static Boolean HasExceptions(Event ev, Boolean checkLocalCacheOnly = false) {
            log.Debug($"Id:{ev.Id}; RecurrenceIsNull:{(ev.Recurrence == null).ToString()};");
            if (ev.Recurrence == null) return false;

            //There's currently no good way to know if a Google event is an exception or not.
            //If it's a change in date, the sequence number increments. However, if it's a different field (eg Subject), no increment.
            //Therefore, you'd have to GetCalendarEntriesInRange() with no date range, then filter to the recurringEventId - not efficient!
            //So...will make it only sync exceptions within the sync date range, which we have cached already
            //if (!checkLocalCacheOnly) {
            //    List<Event> gInstances = Ogcs.Google.Calendar.Instance.GetCalendarEntriesInRecurrence(ev.RecurringEventId ?? ev.Id);
            //    //Add any new exceptions to local cache
            //    googleExceptions = googleExceptions.Union(gInstances.Where(exp => exp.Sequence > 0)).ToList();
            //}
            int exceptionCount = googleExceptions.Where(exp => exp.RecurringEventId == ev.Id).Count();
            if (exceptionCount > 0) {
                log.Debug("This is a recurring Google event with " + exceptionCount + " exceptions in the sync date range.");
                return true;
            } else
                return false;
        }

        public static void SeparateGoogleExceptions(List<Event> allEvents) {
            if (allEvents.Count == 0) return;
            
            log.Debug("Identifying exceptions in recurring Google events.");
            for (int g = allEvents.Count - 1; g >= 0; g--) {
                if (!string.IsNullOrEmpty(allEvents[g].RecurringEventId)) {
                    googleExceptions.Add(allEvents[g]);
                    allEvents.Remove(allEvents[g]);
                }
            }
            log.Debug("Found " + googleExceptions.Count + " exceptions.");
            googleExceptions.ForEach(ge => log.Debug($"RecurringEventId:{ge.RecurringEventId}; Start:{ge.Start.SafeDateTime().ToString()};"));
        }

        /// <summary>
        /// Search cached exceptions for occurrence that originally started on a particular date
        /// </summary>
        /// <param name="recurringEventId">The recurring series to search within</param>
        /// <param name="originalInstanceDate">The date to search for</param>
        /// <returns></returns>
        private static Event getGoogleInstance(String recurringEventId, System.DateTime originalInstanceDate) {
            return googleExceptions.FirstOrDefault(g => g.RecurringEventId == recurringEventId && g.OriginalStartTime.SafeDateTime().Date == originalInstanceDate);
        }

        /// <summary>
        /// Get occurrence that is Outlook exception equivalent
        /// </summary>
        /// <param name="oExcp">Outlook exception to search Google for equivalent</param>
        /// <param name="gRecurringEventID">The ID for the Google series</param
        /// <param name="dirtyCache">Don't used cached items; retrieve them from the cloud</param>
        /// <returns></returns>
        private static Event getGoogleInstance(Microsoft.Office.Interop.Outlook.Exception oExcp, String gRecurringEventID, Boolean dirtyCache) {
            Outlook.Recurrence.DeletionState oIsDeleted = Outlook.Recurrence.ExceptionIsDeleted(oExcp);
            if (oIsDeleted == Outlook.Recurrence.DeletionState.Inaccessible) {
                log.Warn("Abandoning fetch of Google instance for inaccessible Outlook exception.");
                return null;
            }
            log.Debug("Finding Google instance for " + (oIsDeleted == Outlook.Recurrence.DeletionState.Deleted ? "deleted " : "") + "Outlook exception:-");
            log.Debug("  Original date: " + oExcp.OriginalDate.ToString("dd/MM/yyyy"));
            if (oIsDeleted == Outlook.Recurrence.DeletionState.NotDeleted) {
                AppointmentItem ai = null;
                try {
                    ai = oExcp.AppointmentItem;
                    log.Debug("  Current  date: " + ai.Start.ToString("dd/MM/yyyy"));
                } finally {
                    ai = (AppointmentItem)Outlook.Calendar.ReleaseObject(ai);
                }
            }
            if (dirtyCache) {
                log.Debug("Google exception cache not being used. Retrieving all recurring exceptions afresh...");
                //Remove dirty items
                googleExceptions.RemoveAll(ev => ev.RecurringEventId == gRecurringEventID);
                Google.Calendar.Instance.GetCalendarEntriesInRange(gRecurringEventID);
            }
            foreach (Event gExcp in googleExceptions) {
                if (gExcp.RecurringEventId == gRecurringEventID) {
                    if (((oIsDeleted == Outlook.Recurrence.DeletionState.NotDeleted || (oIsDeleted == Outlook.Recurrence.DeletionState.Deleted && !oExcp.Deleted)) /* Weirdness when exception is cancelled by organiser but not yet deleted/accepted by recipient */
                        && oExcp.OriginalDate == gExcp.OriginalStartTime.SafeDateTime()
                        ) ||
                        (oIsDeleted == Outlook.Recurrence.DeletionState.Deleted &&
                        oExcp.OriginalDate.Date == gExcp.OriginalStartTime.SafeDateTime().Date
                        )) {
                        return gExcp;
                    }
                }
            }
            Boolean withinSyncWindow = oExcp.OriginalDate >= Sync.Engine.Calendar.Instance.Profile.SyncStart && oExcp.OriginalDate <= Sync.Engine.Calendar.Instance.Profile.SyncEnd;
            log.Debug("Google exception event is not cached. Retrieving " + (withinSyncWindow ? "recurring instances within sync window" : "all recurring instances") + "...");
            List<Event> gInstances = Ogcs.Google.Calendar.Instance.GetCalendarEntriesInRecurrence(gRecurringEventID, withinSyncWindow);
            if (gInstances == null) return null;

            foreach (Event gInst in gInstances) {
                if (gInst.RecurringEventId == gRecurringEventID) {
                    if (((oIsDeleted == Outlook.Recurrence.DeletionState.NotDeleted || (oIsDeleted == Outlook.Recurrence.DeletionState.Deleted && !oExcp.Deleted)) /* Weirdness when exception is cancelled by organiser but not yet deleted/accepted by recipient */
                        && oExcp.OriginalDate == gInst.OriginalStartTime.SafeDateTime()
                        ) ||
                        (oIsDeleted == Outlook.Recurrence.DeletionState.Deleted &&
                        oExcp.OriginalDate.Date == gInst.OriginalStartTime.SafeDateTime().Date
                        )) {
                        return gInst;
                    }
                }
            }
            return null;
        }

        public static Event GetGoogleMasterEvent(AppointmentItem ai) {
            log.Fine("Found a master Outlook recurring item outside sync date range: " + Outlook.Calendar.GetEventSummary(ai));
            List<Event> events = new List<Event>();
            Boolean haveMatchingEv = false;
            if (Outlook.CustomProperty.Exists(ai, Outlook.CustomProperty.MetadataId.gEventID)) {
                String googleIdValue = Outlook.CustomProperty.Get(ai, Outlook.CustomProperty.MetadataId.gEventID);
                String googleCalValue = Outlook.CustomProperty.Get(ai, Outlook.CustomProperty.MetadataId.gCalendarId);
                if (googleCalValue == null || googleCalValue == Sync.Engine.Calendar.Instance.Profile.UseGoogleCalendar.Id) {
                    Event ev = Ogcs.Google.Calendar.Instance.GetCalendarEntry(googleIdValue);
                    if (ev != null) {
                        events.Add(ev);
                        haveMatchingEv = true;
                        log.Fine("Found single hard-matched Event.");
                    } else if (Ogcs.Google.Calendar.Instance.ExcludedByConfig.Contains(googleIdValue)) {
                        log.Debug("The master Google Event has been excluded by config.");
                        return null;
                    }
                }
            }
            if (!haveMatchingEv) {
                events = Ogcs.Google.Calendar.Instance.GetCalendarEntriesInRange(ai.Start.Date, ai.Start.Date.AddDays(1), true);
                if (Sync.Engine.Calendar.Instance.Profile.SyncDirection.Id != Sync.Direction.GoogleToOutlook.Id) {
                    List<AppointmentItem> ais = new List<AppointmentItem>();
                    ais.Add(ai);
                    Ogcs.Google.Calendar.Instance.ReclaimOrphanCalendarEntries(ref events, ref ais, neverDelete: true);
                }
            }
            for (int g = 0; g < events.Count(); g++) {
                Event ev = events[g];
                String gEntryID = Ogcs.Google.CustomProperty.Get(ev, Ogcs.Google.CustomProperty.MetadataId.oEntryId);
                if (haveMatchingEv || !string.IsNullOrEmpty(gEntryID)) {
                    if (haveMatchingEv && string.IsNullOrEmpty(gEntryID)) {
                        return ev;
                    }
                    if (Ogcs.Google.CustomProperty.OutlookIdMissing(ev)) {
                        String compare_oID;
                        if (!string.IsNullOrEmpty(gEntryID) && gEntryID.StartsWith(Outlook.Calendar.GlobalIdPattern)) { //We got a Global ID, not Entry ID
                            compare_oID = Outlook.Calendar.Instance.IOutlook.GetGlobalApptID(ai);
                        } else {
                            compare_oID = ai.EntryID;
                        }
                        if (haveMatchingEv || gEntryID == compare_oID) {
                            log.Info("Adding Outlook IDs to Master Google Event...");
                            Ogcs.Google.CustomProperty.AddOutlookIDs(ref ev, ai);
                            try {
                                Ogcs.Google.Calendar.Instance.UpdateCalendarEntry_save(ref ev);
                            } catch (System.Exception ex) {
                                ex.Analyse("Failed saving Outlook IDs to Google Event.", true);
                            }
                            return ev;
                        }
                    } else if (Ogcs.Google.Calendar.ItemIDsMatch(ref ev, ai)) {
                        log.Fine("Found master event.");
                        return ev;
                    }
                } else {
                    log.Debug("Event \"" + ev.Summary + "\" does not have Outlook EntryID stored.");
                    if (Ogcs.Google.Calendar.SignaturesMatch(Ogcs.Google.Calendar.signature(ev), Outlook.Calendar.signature(ai))) {
                        log.Debug("Master event matched on simple signatures.");
                        return ev;
                    }
                }
            }
            log.Warn("Failed to find master Google event for: " + Outlook.Calendar.GetEventSummary(ai));
            return null;
        }

        public static void CreateGoogleExceptions(AppointmentItem ai, String recurringEventId) {
            if (!ai.IsRecurring) return;

            log.Debug("Creating Google recurrence exceptions.");
            //Sync all exceptions regardless whether within synced date range; otherwise exceptions that come in scope later but have not been recently modified will not be compared
            List<Event> gRecurrences = Ogcs.Google.Calendar.Instance.GetCalendarEntriesInRecurrence(recurringEventId);
            if (gRecurrences == null) return;

            RecurrencePattern rp = null;
            Exceptions excps = null;
            try {
                rp = ai.GetRecurrencePattern();
                excps = rp.Exceptions;
                log.Debug(excps.Count + " recurring exceptions to be created.");
                for (int e = 1; e <= excps.Count; e++) {
                    Microsoft.Office.Interop.Outlook.Exception oExcp = null;
                    try {
                        oExcp = excps[e];
                        for (int g = 0; g < gRecurrences.Count; g++) {
                            Event ev = gRecurrences[g];
                            System.DateTime gDate = ev.OriginalStartTime.SafeDateTime();
                            Outlook.Recurrence.DeletionState isDeleted = Outlook.Recurrence.ExceptionIsDeleted(oExcp);
                            if (isDeleted == Outlook.Recurrence.DeletionState.Inaccessible) {
                                log.Warn("Abandoning creation of Google recurrence exception as Outlook exception is inaccessible.");
                                return;
                            }
                            if (isDeleted == Outlook.Recurrence.DeletionState.Deleted && !ai.AllDayEvent) { //Deleted items get truncated?!
                                gDate = gDate.Date;
                            }
                            if (oExcp.OriginalDate == gDate) {
                                if (isDeleted == Outlook.Recurrence.DeletionState.Deleted) {
                                    log.Fine("Checking if there are other exceptions that were originally on " + oExcp.OriginalDate.ToString("dd-MMM-yyyy") + " and moved.");
                                    Boolean skipDelete = false;
                                    for (int a = 1; a <= excps.Count; a++) {
                                        Microsoft.Office.Interop.Outlook.Exception oExcp2 = null;
                                        try {
                                            oExcp2 = excps[a];
                                            if (!oExcp2.Deleted) {
                                                Microsoft.Office.Interop.Outlook.AppointmentItem ai2 = null;
                                                try {
                                                    ai2 = oExcp2.AppointmentItem;
                                                    if (oExcp.OriginalDate.Date == oExcp2.OriginalDate.Date && oExcp.OriginalDate.Date != ai2.Start.Date) {
                                                        //It's an additional exception which has the same original start date, but was moved
                                                        log.Warn(Ogcs.Google.Calendar.GetEventSummary(ev));
                                                        log.Warn("This item is not really deleted, but moved to another date in Outlook on " + ai2.Start.Date.ToString("dd-MMM-yyyy"));
                                                        skipDelete = true;
                                                        log.Fine("Now checking if there is a Google item on that date - we don't want a duplicate.");
                                                        Event duplicate = gRecurrences.FirstOrDefault(g => ai2.Start.Date == g.OriginalStartTime.SafeDateTime().Date);
                                                        if (duplicate != null) {
                                                            log.Warn("Determined a 'duplicate' exists on that date - this will be deleted.");
                                                            duplicate.Status = "cancelled";
                                                            Ogcs.Google.Calendar.Instance.UpdateCalendarEntry_save(ref duplicate);
                                                        }
                                                        break;
                                                    }
                                                } catch (System.Exception ex) {
                                                    Ogcs.Exception.Analyse(ex);
                                                } finally {
                                                    ai2 = (Microsoft.Office.Interop.Outlook.AppointmentItem)Outlook.Calendar.ReleaseObject(ai2);
                                                }
                                            }
                                        } catch (System.Exception ex) {
                                            ex.Analyse("Could not check if there are other exceptions with the same original start date.");
                                        }
                                    }
                                    if (!skipDelete) {
                                        log.Fine("None found.");
                                        Forms.Main.Instance.Console.Update(Ogcs.Google.Calendar.GetEventSummary("<br/>Occurrence deleted.", ev, out String anonSummary), anonSummary, Console.Markup.calendar, verbose: true);
                                        ev.Status = "cancelled";
                                        Ogcs.Google.Calendar.Instance.UpdateCalendarEntry_save(ref ev);
                                    }
                                } else {
                                    int exceptionItemsModified = 0;
                                    Event modifiedEv = Ogcs.Google.Calendar.Instance.UpdateCalendarEntry(oExcp.AppointmentItem, ev, ref exceptionItemsModified, forceCompare: true);
                                    if (exceptionItemsModified > 0) {
                                        Ogcs.Google.Calendar.Instance.UpdateCalendarEntry_save(ref modifiedEv);
                                        if (oExcp.OriginalDate.Date != oExcp.AppointmentItem.Start.Date) {
                                            log.Fine("Double checking there is no other Google item on " + oExcp.AppointmentItem.Start.Date.ToString("dd-MMM-yyyy") + " that " + oExcp.OriginalDate.Date.ToString("dd-MMM-yyyy") + " was moved to - we don't want a duplicate.");
                                            Event duplicate = gRecurrences.FirstOrDefault(g => oExcp.AppointmentItem.Start.Date == g.OriginalStartTime.SafeDateTime().Date);
                                            if (duplicate != null && duplicate.Status != "cancelled") {
                                                log.Warn("Determined a 'duplicate' exists on that date - this will be deleted.");
                                                duplicate.Status = "cancelled";
                                                Ogcs.Google.Calendar.Instance.UpdateCalendarEntry_save(ref duplicate);
                                            }
                                        }
                                    }
                                }
                                break;
                            }
                        }
                    } finally {
                        oExcp = (Microsoft.Office.Interop.Outlook.Exception)Outlook.Calendar.ReleaseObject(oExcp);
                    }
                }
            } finally {
                for (int e = 1; e <= excps.Count; e++) {
                    Microsoft.Office.Interop.Outlook.Exception garbage = (Microsoft.Office.Interop.Outlook.Exception)Outlook.Calendar.ReleaseObject(excps[e]);
                }
                excps = (Exceptions)Outlook.Calendar.ReleaseObject(excps);
                rp = (RecurrencePattern)Outlook.Calendar.ReleaseObject(rp);
            }
        }

        public static int UpdateGoogleExceptions(AppointmentItem ai, Event ev, Boolean dirtyCache) {
            int updatesMade = 0;

            if (ai.IsRecurring) {
                RecurrencePattern rp = null;
                Exceptions excps = null;
                try {
                    rp = ai.GetRecurrencePattern();
                    excps = rp.Exceptions;
                    if (excps.Count > 0) {
                        log.Debug(Outlook.Calendar.GetEventSummary(ai));
                        log.Debug("This is a recurring appointment with " + excps.Count + " exceptions that will now be iteratively compared, if inside synced date range.");
                        for (int e = 1; e <= excps.Count; e++) {
                            Microsoft.Office.Interop.Outlook.Exception oExcp = null;
                            AppointmentItem aiExcp = null;
                            try {
                                oExcp = excps[e];
                                int excp_itemModified = 0;
                                System.DateTime oExcp_currDate;

                                //Check the exception falls in the date range being synced
                                Outlook.Recurrence.DeletionState oIsDeleted = Outlook.Recurrence.ExceptionIsDeleted(oExcp);
                                String logDeleted = "";
                                if (oIsDeleted != Outlook.Recurrence.DeletionState.NotDeleted) {
                                    logDeleted = " " + oIsDeleted.ToString().ToLower() + " and";
                                    oExcp_currDate = oExcp.OriginalDate;
                                } else {
                                    aiExcp = oExcp.AppointmentItem;
                                    oExcp_currDate = aiExcp.Start;
                                    aiExcp = (AppointmentItem)Outlook.Calendar.ReleaseObject(aiExcp);
                                }
                                if (oExcp_currDate < Sync.Engine.Calendar.Instance.Profile.SyncStart.Date || oExcp_currDate > Sync.Engine.Calendar.Instance.Profile.SyncEnd.Date) {
                                    log.Fine("Exception is" + logDeleted + " outside date range being synced: " + oExcp_currDate.Date.ToString("dd/MM/yyyy"));
                                    continue;
                                } else if (oIsDeleted == Outlook.Recurrence.DeletionState.Inaccessible) {
                                    log.Warn("Exception is" + logDeleted + " cannot be synced: " + oExcp_currDate.Date.ToString("dd/MM/yyyy"));
                                    continue;
                                }

                                Event gExcp = getGoogleInstance(oExcp, ev.RecurringEventId ?? ev.Id, dirtyCache);
                                if (gExcp != null) {
                                    log.Debug("Matching Google Event recurrence found.");
                                    if (gExcp.Status == "cancelled") {
                                        log.Debug("It is deleted in Google, which " + (oIsDeleted == Outlook.Recurrence.DeletionState.Deleted ? "matches" : "does not match") + " Outlook.");
                                        if (oIsDeleted == Outlook.Recurrence.DeletionState.NotDeleted) {
                                            log.Warn("Outlook is NOT deleted though - a mismatch has occurred somehow!");
                                            String syncDirectionTip = (Sync.Engine.Calendar.Instance.Profile.SyncDirection.Id == Sync.Direction.Bidirectional.Id) ? "<br/><i>Ensure you <b>first</b> set OGCS to one-way sync O->G.</i>" : "";
                                            Forms.Main.Instance.Console.Update(
                                                Outlook.Calendar.GetEventSummary("<br/>" +
                                                    "The occurrence on " + oExcp.OriginalDate.ToShortDateString() + " does not exist in Google, but does in Outlook.<br/>" +
                                                    "This can happen if, for example, you declined the occurrence (which is synced to Google) and proposed a new time that is subsequently accepted by the organiser.<br/>" +
                                                    "<u>Suggested fix</u>: delete the entire series in Google and let OGCS recreate it." + syncDirectionTip, ai, out String anonSummary)
                                                , anonSummary, Console.Markup.warning);
                                        }
                                        continue;
                                    } else if (oIsDeleted == Outlook.Recurrence.DeletionState.Deleted && gExcp.Status != "cancelled") {
                                        System.DateTime movedToStartDate = gExcp.Start.SafeDateTime().Date;
                                        log.Fine("Checking if we have another Google instance that /is/ cancelled on " + movedToStartDate.ToString("dd-MMM-yyyy") + " that this one has been moved to.");
                                        Event duplicate = getGoogleInstance(gExcp.RecurringEventId, movedToStartDate);
                                        DialogResult dr = DialogResult.Yes;
                                        String summary = Outlook.Calendar.GetEventSummary(ai, out String anonSummary);
                                        if (duplicate?.Status == "cancelled") {
                                            log.Warn("Another deleted occurrence on the same date " + movedToStartDate.ToString("dd-MMM-yyyy") + " found, so this Google item that has moved to that date cannot be safely deleted automatically.");
                                            String msg = summary + "\r\n\r\nAn occurrence on " + movedToStartDate.ToString("dd-MMM-yyyy") + " was previously deleted, before another occurrence on " + oExcp.OriginalDate.ToString("dd-MMM-yyyy") +
                                                " was rescheduled to the same date and then deleted again. " +
                                                "Please confirm the Google occurrence, currently on " + movedToStartDate.ToString("dd-MMM-yyyy") + ", should be deleted?";
                                            dr = Ogcs.Extensions.MessageBox.Show(msg, "Confirm deletion of recurring series occurrence", MessageBoxButtons.YesNo, MessageBoxIcon.Question, msg.Replace(summary, anonSummary));
                                        }
                                        if (dr == DialogResult.Yes) {
                                            Forms.Main.Instance.Console.Update(Ogcs.Google.Calendar.GetEventSummary("<br/>Occurrence deleted.", gExcp, out String anonSummary2), anonSummary2, Console.Markup.calendar, verbose: true);
                                            gExcp.Status = "cancelled";
                                            log.Debug("Exception deleted.");
                                            excp_itemModified++;
                                        }
                                    } else {
                                        try {
                                            aiExcp = oExcp.AppointmentItem;
                                            //Force a compare of the exception if both G and O have been modified in last 24 hours
                                            TimeSpan modifiedDiff = (TimeSpan)(gExcp.UpdatedDateTimeOffset - aiExcp.LastModificationTime);
                                            log.Fine("Modification time difference (in days) between G and O exception: " + modifiedDiff);
                                            Boolean forceCompare = modifiedDiff < TimeSpan.FromDays(1);
                                            Ogcs.Google.Calendar.Instance.UpdateCalendarEntry(aiExcp, gExcp, ref excp_itemModified, forceCompare);
                                            if (forceCompare && excp_itemModified == 0 && System.DateTime.Now > aiExcp.LastModificationTime.AddDays(1)) {
                                                Ogcs.Google.CustomProperty.SetOGCSlastModified(ref gExcp);
                                                try {
                                                    log.Debug("Doing a dummy update in order to update the last modified date of Google recurring series exception.");
                                                    Ogcs.Google.Calendar.Instance.UpdateCalendarEntry_save(ref gExcp);
                                                } catch (System.Exception ex) {
                                                    ex.Analyse("Dummy update of unchanged exception for Google recurring series failed.");
                                                }
                                                continue;
                                            }
                                        } catch (System.Exception ex) {
                                            Ogcs.Exception.Analyse(ex, true);
                                            throw;
                                        }
                                    }
                                    if (excp_itemModified > 0) {
                                        try {
                                            Ogcs.Google.Calendar.Instance.UpdateCalendarEntry_save(ref gExcp);
                                            updatesMade++;
                                        } catch (System.Exception ex) {
                                            Forms.Main.Instance.Console.UpdateWithError(Ogcs.Google.Calendar.GetEventSummary("Updated event exception failed to save.", gExcp, out String anonSummary, true), ex, logEntry: anonSummary);
                                            log.Debug(Newtonsoft.Json.JsonConvert.SerializeObject(gExcp));
                                            Ogcs.Exception.Analyse(ex, true);
                                            if (Ogcs.Extensions.MessageBox.Show("Updated Google event exception failed to save. Continue with synchronisation?", "Sync item failed", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                                                continue;
                                            else {
                                                throw new UserCancelledSyncException("User chose not to continue sync.");
                                            }
                                        }
                                    }
                                } else {
                                    log.Warn("No matching Google Event recurrence found.");
                                    if (oIsDeleted == Outlook.Recurrence.DeletionState.Deleted) log.Debug("The Outlook appointment is deleted, so not a problem.");
                                }
                            } finally {
                                aiExcp = (AppointmentItem)Outlook.Calendar.ReleaseObject(aiExcp);
                                oExcp = (Microsoft.Office.Interop.Outlook.Exception)Outlook.Calendar.ReleaseObject(oExcp);
                            }
                        }
                    }
                } finally {
                    excps = (Exceptions)Outlook.Calendar.ReleaseObject(excps);
                    rp = (RecurrencePattern)Outlook.Calendar.ReleaseObject(rp);
                }
            }
            return updatesMade;
        }
        #endregion
    }
}
