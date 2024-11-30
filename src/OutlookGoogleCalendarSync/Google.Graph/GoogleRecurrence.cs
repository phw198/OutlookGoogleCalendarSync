using GcalData = Google.Apis.Calendar.v3.Data;
using log4net;
using OutlookGoogleCalendarSync.Extensions;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;
using Ogcs = OutlookGoogleCalendarSync;
using Microsoft.Graph;

namespace OutlookGoogleCalendarSync.Google.Graph {
    public class Recurrence {
        private static readonly ILog log = LogManager.GetLogger(typeof(Recurrence));

        public static List<String> BuildGooglePattern(Microsoft.Graph.Event ai, GcalData.Event ev) {
            if (ai.Recurrence == null) return null;

            log.Debug("Creating Google iCalendar definition for recurring event.");
            List<String> gPattern = new List<String>();
            System.DateTime? utcEnd = null;
            if (ai.Recurrence.Range.Type == Microsoft.Graph.RecurrenceRangeType.EndDate) {
                utcEnd = System.DateTime.Parse(ai.Recurrence.Range.EndDate.ToString());
                if (!(ai.IsAllDay ?? false)) {
                    System.DateTime localEnd = (System.DateTime)utcEnd + ai.End.SafeDateTime().TimeOfDay;
                    utcEnd = localEnd.ToUniversalTime();
                }
            } else if (ai.Recurrence.Range.Type != Microsoft.Graph.RecurrenceRangeType.NoEnd) {
                log.Warn($"Series range type of '{ai.Recurrence.Range.Type}' is not handled.");
            }
            gPattern.Add("RRULE:" + buildRrule(ai.Recurrence, utcEnd));

            log.Debug(string.Join("\r\n", gPattern.ToArray()));
            return gPattern;
        }

        private static String buildRrule(PatternedRecurrence oPattern, System.DateTime? recurrenceEndUtc) {
            log.Fine("Building RRULE");
            Dictionary<String, String> rrule = new Dictionary<String, String>();

            #region RECURRENCE PATTERN
            log.Fine($"Determining pattern for frequency '{oPattern.Pattern.Type?.ToString()}'.");

            switch (oPattern.Pattern.Type) {
                case RecurrencePatternType.Daily: {
                        Google.Recurrence.addRule(rrule, "FREQ", "DAILY");
                        Google.Recurrence.setInterval(rrule, oPattern.Pattern.Interval ?? 0);
                        break;
                    }
                case RecurrencePatternType.Weekly: {
                        Google.Recurrence.addRule(rrule, "FREQ", "WEEKLY");
                        Google.Recurrence.setInterval(rrule, oPattern.Pattern.Interval ?? 0);
                        // Need to work out "BY" pattern
                        // Eg "BYDAY=MO,TU,WE,TH,FR"
                        Google.Recurrence.addRule(rrule, "BYDAY", string.Join(",", getByDay(oPattern.Pattern.DaysOfWeek.ToList())));
                        break;
                    }

                case RecurrencePatternType.AbsoluteMonthly: {
                        Google.Recurrence.addRule(rrule, "FREQ", "MONTHLY");
                        Google.Recurrence.setInterval(rrule, oPattern.Pattern.Interval ?? 0);
                        //Outlook and Google interpret days of month that don't alway exist, eg 31st, differently - though it's not explicitly defined
                        //Outlook: Picks last day of month (SKIP=BACKWARD); Google: Skips that month (SKIP=OMIT)
                        //We'll adopt Outlook's definition
                        if (oPattern.Range.StartDate.Day > 28) {
                            Google.Recurrence.addRule(rrule, "RSCALE", "GREGORIAN");
                            Google.Recurrence.addRule(rrule, "BYMONTHDAY", oPattern.Range.StartDate.Day.ToString());
                            Google.Recurrence.addRule(rrule, "SKIP", "BACKWARD");
                        }
                        break;
                    }

                case RecurrencePatternType.RelativeMonthly: {
                        Google.Recurrence.addRule(rrule, "FREQ", "MONTHLY");
                        Google.Recurrence.setInterval(rrule, oPattern.Pattern.Interval ?? 0);
                        Google.Recurrence.addRule(rrule, "BYDAY", getByDayRelative(rrule, oPattern.Pattern));
                        break;
                    }

                case RecurrencePatternType.AbsoluteYearly: {
                        Google.Recurrence.addRule(rrule, "FREQ", "YEARLY");
                        Google.Recurrence.setInterval(rrule, oPattern.Pattern.Interval ?? 0);
                        break;
                    }

                case RecurrencePatternType.RelativeYearly: {
                        Google.Recurrence.addRule(rrule, "FREQ", "YEARLY");
                        Google.Recurrence.setInterval(rrule, oPattern.Pattern.Interval ?? 0);
                        Google.Recurrence.addRule(rrule, "BYMONTH", oPattern.Pattern.Month.ToString());
                        Google.Recurrence.addRule(rrule, "BYDAY", getByDayRelative(rrule, oPattern.Pattern));
                        break;
                    }
            }
            #endregion

            #region RECURRENCE RANGE
            if (oPattern.Range.Type == RecurrenceRangeType.EndDate && recurrenceEndUtc != null) {
                log.Fine("Checking end date.");
                Google.Recurrence.addRule(rrule, "UNTIL", Google.Recurrence.IANAdate((System.DateTime)recurrenceEndUtc));
            }
            #endregion

            return string.Join(";", rrule.Select(x => x.Key + "=" + x.Value).ToArray());
        }
        /*

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

        */
        private static List<String> getByDay(List<Microsoft.Graph.DayOfWeek> dow) {
            log.Fine("DayOfWeekList = " + string.Join(",", dow));
            List<String> byDay = new List<String>();
            foreach (Microsoft.Graph.DayOfWeek day in dow) {
                byDay.Add(day.ToString().Substring(0, 2).ToUpper());
            }
            return byDay;
        }
        private static String getByDayRelative(Dictionary<String, String> rrule, Microsoft.Graph.RecurrencePattern oPattern) {
            String byDay = "";
            switch (oPattern.Index) {
                case WeekIndex.First: byDay = "1"; break;
                case WeekIndex.Second: byDay = "2"; break;
                case WeekIndex.Third: byDay = "3"; break;
                case WeekIndex.Fourth: byDay = "4"; break;
                case WeekIndex.Last: Google.Recurrence.addRule(rrule, "BYSETPOS", "-1"); break;
            }
            byDay += getByDay(oPattern.DaysOfWeek.ToList()).First();
            return byDay;
        }
        /*
        #region Exceptions
        private static List<Event> googleExceptions;
        public static List<Event> GoogleExceptions {
            get { return googleExceptions; }
        }

        public static Boolean HasExceptions(Event ev, Boolean checkLocalCacheOnly = false) {
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

        /// <summary>
        /// Get occurrence that originally started on a particular date
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
                log.Debug("Google exception cache not being used. Retrieving all recurring instances afresh...");
                //Remove dirty items
                googleExceptions.RemoveAll(ev => ev.RecurringEventId == gRecurringEventID);
            } else {
                foreach (Event gExcp in googleExceptions) {
                    if (gExcp.RecurringEventId == gRecurringEventID) {
                        if (((oIsDeleted == Outlook.Recurrence.DeletionState.NotDeleted || (oIsDeleted == Outlook.Recurrence.DeletionState.Deleted && !oExcp.Deleted)) *//* Weirdness when exception is cancelled by organiser but not yet deleted/accepted by recipient *//*
                            && oExcp.OriginalDate == gExcp.OriginalStartTime.SafeDateTime()
                            ) ||
                            (oIsDeleted == Outlook.Recurrence.DeletionState.Deleted &&
                            oExcp.OriginalDate == gExcp.OriginalStartTime.SafeDateTime().Date
                            )) {
                            return gExcp;
                        }
                    }
                }
                log.Debug("Google exception event is not cached. Retrieving all recurring instances...");
            }
            List<Event> gInstances = Ogcs.Google.Calendar.Instance.GetCalendarEntriesInRecurrence(gRecurringEventID);
            if (gInstances == null) return null;

            //Add any new exceptions to local cache
            googleExceptions = googleExceptions.Union(gInstances.Where(ev => !String.IsNullOrEmpty(ev.RecurringEventId))).ToList();
            foreach (Event gInst in gInstances) {
                if (gInst.RecurringEventId == gRecurringEventID) {
                    if (((oIsDeleted == Outlook.Recurrence.DeletionState.NotDeleted || (oIsDeleted == Outlook.Recurrence.DeletionState.Deleted && !oExcp.Deleted)) *//* Weirdness when exception is cancelled by organiser but not yet deleted/accepted by recipient *//*
                        && oExcp.OriginalDate == gInst.OriginalStartTime.SafeDateTime()
                        ) ||
                        (oIsDeleted == Outlook.Recurrence.DeletionState.Deleted &&
                        oExcp.OriginalDate == gInst.OriginalStartTime.SafeDateTime().Date
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
                    }
                }
            }
            if (!haveMatchingEv) {
                events = Ogcs.Google.Calendar.Instance.GetCalendarEntriesInRange(ai.Start.Date, ai.Start.Date.AddDays(1));
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
                    if (Ogcs.Google.Calendar.SignaturesMatch(Ogcs.Google.Calendar.Signature(ev), Outlook.Calendar.signature(ai))) {
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
            List<Event> gRecurrences = Ogcs.Google.Calendar.Instance.GetCalendarEntriesInRecurrence(recurringEventId);
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

        public static void UpdateGoogleExceptions(AppointmentItem ai, Event ev, Boolean dirtyCache) {
            if (ai.IsRecurring) {
                RecurrencePattern rp = null;
                Exceptions excps = null;
                try {
                    rp = ai.GetRecurrencePattern();
                    excps = rp.Exceptions;
                    if (excps.Count > 0) {
                        log.Debug(Outlook.Calendar.GetEventSummary(ai));
                        log.Debug("This is a recurring appointment with " + excps.Count + " exceptions that will now be iteratively compared.");
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
                                            TimeSpan modifiedDiff = (TimeSpan)(gExcp.Updated - aiExcp.LastModificationTime);
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
                                        } catch (System.Exception ex) {
                                            Forms.Main.Instance.Console.UpdateWithError(Ogcs.Google.Calendar.GetEventSummary("Updated event exception failed to save.", gExcp, out String anonSummary, true), ex, logEntry: anonSummary);
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
        }
        #endregion*/
    }
}
