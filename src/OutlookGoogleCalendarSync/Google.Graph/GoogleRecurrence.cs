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
            if (ai.Recurrence == null || ai.Type != EventType.SeriesMaster) return null;

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
            if (oPattern.Range.Type == RecurrenceRangeType.Numbered) {
                Google.Recurrence.addRule(rrule, "COUNT", oPattern.Range.NumberOfOccurrences.ToString());

            } else if (oPattern.Range.Type == RecurrenceRangeType.EndDate && recurrenceEndUtc != null) {
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
                case WeekIndex.Last: byDay = "-1"; break;
            }
            byDay += getByDay(oPattern.DaysOfWeek.ToList()).First();
            return byDay;
        }

        #region Exceptions
        /*
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
        */

        /// <summary>
        /// Get occurrence that is Outlook exception equivalent
        /// </summary>
        /// <param name="oExcp">Outlook exception to search Google for equivalent</param>
        /// <param name="gRecurringEventID">The ID for the Google series</param
        /// <param name="dirtyCache">Don't used cached items; retrieve them from the cloud</param>
        /// <returns></returns>
        private static GcalData.Event getGoogleInstance(Event oExcp, String gRecurringEventID, Boolean dirtyCache, Boolean isDeleted) {
            log.Debug("Finding Google instance for " + (isDeleted ? "deleted " : "") + "Outlook exception:-");
            if (oExcp.OriginalStart != null) 
                log.Debug("  Original date: " + ((System.DateTimeOffset)oExcp.OriginalStart).UtcDateTime.ToString("dd/MM/yyyy"));
            if (!isDeleted)
                log.Debug("  Current  date: " + oExcp.Start.SafeDateTime().ToString("dd/MM/yyyy"));
            
            if (dirtyCache) {
                log.Debug("Google exception cache not being used. Retrieving all recurring instance exceptions afresh...");
                //Remove dirty items
                Google.Recurrence.GoogleExceptions.RemoveAll(ev => ev.RecurringEventId == gRecurringEventID);
                Google.Calendar.Instance.GetCalendarEntriesInRange(gRecurringEventID);
            }
            foreach (GcalData.Event gExcp in Google.Recurrence.GoogleExceptions.Where(g => g.RecurringEventId == gRecurringEventID).ToList()) {
                if ((isDeleted && oExcp.OriginalStart == gExcp.OriginalStartTime.SafeDateTime().Date)
                    || (!isDeleted && oExcp.OriginalStart == gExcp.OriginalStartTime.SafeDateTime())) {
                    return gExcp;
                }
            }

            log.Debug("Google exception event is not cached. Retrieving all recurring instances...");            
            List<GcalData.Event> gInstances = Ogcs.Google.Calendar.Instance.GetCalendarEntriesInRecurrence(gRecurringEventID);
            if (gInstances == null) return null;

            foreach (GcalData.Event gInst in gInstances) {
                if ((isDeleted && oExcp.OriginalStart == gInst.OriginalStartTime.SafeDateTime().Date)
                    || (!isDeleted && oExcp.OriginalStart == gInst.OriginalStartTime.SafeDateTime())) {
                    return gInst;
                }
            }
            return null;
        }

        /*
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
                    if (Ogcs.Google.Calendar.SignaturesMatch(Ogcs.Google.Calendar.Signature(ev), Outlook.Calendar.signature(ai))) {
                        log.Debug("Master event matched on simple signatures.");
                        return ev;
                    }
                }
            }
            log.Warn("Failed to find master Google event for: " + Outlook.Calendar.GetEventSummary(ai));
            return null;
        }
        */

        public static void CreateGoogleExceptions(List<Event> aiExcps, String recurringEventId) {
            List<System.DateTime> cancelledDates = new();
            Ogcs.Outlook.Graph.Calendar.Instance.CancelledOccurrences.TryGetValue(aiExcps.FirstOrDefault().SeriesMasterId, out cancelledDates);
            if (aiExcps.Count + cancelledDates.Count == 0) return;

            log.Debug("Creating Google recurrence exceptions.");
            List<GcalData.Event> gRecurrences = Ogcs.Google.Calendar.Instance.GetCalendarEntriesInRecurrence(recurringEventId);
            if (gRecurrences == null) return;

            try {
                if (cancelledDates.Count > 0) {
                    log.Debug($"Cancelling {cancelledDates.Count()} occurrences.");
                    foreach (System.DateTime cancelledDate in cancelledDates) {
                        GcalData.Event ev = gRecurrences.Where(ev => ev.OriginalStartTime.SafeDateTime().Date == cancelledDate.Date).First();
                        if (ev == null) {
                            log.Warn($"Could not find a Google occurrence for Outlook's cancellation on {cancelledDate}");
                        } else {
                            gRecurrences.Remove(ev);
                            if (ev.Status == "cancelled") {
                                log.Warn($"Outlook occurrence already deleted on {cancelledDate}");
                            } else {
                                Forms.Main.Instance.Console.Update(Ogcs.Google.Calendar.GetEventSummary("<br/>Occurrence deleted.", ev, out String anonSummary), anonSummary, Console.Markup.calendar, verbose: true);
                                ev.Status = "cancelled";
                                Ogcs.Google.Calendar.Instance.UpdateCalendarEntry_save(ref ev);
                            }
                        }
                    }
                }
            } catch (System.Exception ex) {
                ex.Analyse("Could not process cancelled occurrences.");
            }
            try {
                foreach (Event oExcp in aiExcps) {
                    System.DateTime? oOriginalStart = null;
                    try {
                        oOriginalStart = (oExcp.OriginalStart?.DateTime ?? oExcp.Start.SafeDateTime()).ToLocalTime();
                        GcalData.Event ev = gRecurrences.Where(ev => ev.OriginalStartTime.SafeDateTime() == oOriginalStart).FirstOrDefault();
                        if (ev == null) {
                            log.Warn($"Could not find an occurrence originally starting on {oOriginalStart}");
                        } else {
                            int exceptionItemsModified = 0;
                            GcalData.Event modifiedEv = Calendar.UpdateCalendarEntry(oExcp, ev, ref exceptionItemsModified, forceCompare: true);
                            if (exceptionItemsModified > 0) {
                                Ogcs.Google.Calendar.Instance.UpdateCalendarEntry_save(ref modifiedEv);
                                if (oExcp.OriginalStart?.Date != oExcp.Start.SafeDateTime().Date) {
                                    log.Fine("Double checking there is no other Google item on " + oExcp.Start.SafeDateTime().Date.ToString("dd-MMM-yyyy") + " that " + oExcp.OriginalStart?.Date.ToString("dd-MMM-yyyy") + " was moved to - we don't want a duplicate.");
                                    GcalData.Event duplicate = gRecurrences.FirstOrDefault(g => oExcp.Start.SafeDateTime().Date == g.OriginalStartTime.SafeDateTime().Date);
                                    if (duplicate != null && duplicate.Status != "cancelled") {
                                        log.Warn("Determined a 'duplicate' exists on that date - this will be deleted.");
                                        duplicate.Status = "cancelled";
                                        Ogcs.Google.Calendar.Instance.UpdateCalendarEntry_save(ref duplicate);
                                    }
                                }
                            }
                        }
                    } catch (System.Exception ex) {
                        ex.Analyse($"Failed to process modified occurrence on {oOriginalStart}");
                    }
                }
            } catch (System.Exception ex) {
                ex.Analyse("Could not process modified occurrences.");
            }
        }

        public static void UpdateGoogleExceptions(Event seriesMaster, GcalData.Event ev, Boolean dirtyCache) {
            if (seriesMaster.Type != Microsoft.Graph.EventType.SeriesMaster) return;

            List<Event> aiExcps = Outlook.Graph.Recurrence.GetExceptions(seriesMaster);
            List<System.DateTime> cancelledDates;
            Ogcs.Outlook.Graph.Calendar.Instance.CancelledOccurrences.TryGetValue(seriesMaster.Id, out cancelledDates);
            cancelledDates ??= new();
            if (aiExcps.Count + cancelledDates.Count == 0) return;

            log.Debug(Outlook.Graph.Calendar.GetEventSummary(seriesMaster));
            log.Debug("This is a recurring appointment with " + (aiExcps.Count + cancelledDates.Count) + " exceptions that will now be iteratively compared.");
            
            #region Cancelled occurrences
            try {
                foreach (System.DateTime cancelledDate in cancelledDates) {
                    log.Fine("Cancelled occurrence on " + cancelledDate.ToString("dd/MM/yyyy"));

                    Event aiExcp = new() { OriginalStart = cancelledDate };
                    GcalData.Event gExcp = Google.Recurrence.GetGoogleInstance(ev.RecurringEventId ?? ev.Id, cancelledDate);
                    if (gExcp == null) {
                        log.Fine($"Google has no cached exception yet on {cancelledDate.ToString("dd/MM/yyyy")}");
                        gExcp = getGoogleInstance(aiExcp, ev.RecurringEventId ?? ev.Id, dirtyCache, true);
                    }
                    if (gExcp == null) {
                        log.Warn("No matching Google Event recurrence found, but the Outlook appointment is deleted, so not a problem.");
                    } else {
                        log.Fine("Existing Google exception found");
                        if (gExcp.Status == "cancelled") {
                            log.Fine($"{cancelledDate.ToString("dd/MM/yyyy")} is deleted in Google, which matches Outlook.");
                        } else {
                            System.DateTime movedToStartDate = gExcp.Start.SafeDateTime().Date;
                            log.Fine("Checking if we have another Google instance that /is/ cancelled on " + movedToStartDate.ToString("dd-MMM-yyyy") + " that this one has been moved to.");
                            GcalData.Event duplicate = Google.Recurrence.GetGoogleInstance(gExcp.RecurringEventId, movedToStartDate);
                            DialogResult dr = DialogResult.Yes;
                            String summary = Outlook.Graph.Calendar.GetEventSummary(seriesMaster, out String anonSummary);
                            if (duplicate?.Status == "cancelled") {
                                log.Warn("Another deleted occurrence on the same date " + movedToStartDate.ToString("dd-MMM-yyyy") + " found, so this Google item that has moved to that date cannot be safely deleted automatically.");
                                String msg = summary + "\r\n\r\nAn occurrence on " + movedToStartDate.ToString("dd-MMM-yyyy") + " was previously deleted, before another occurrence on " + ((System.DateTimeOffset)aiExcp.OriginalStart).ToString("dd-MMM-yyyy") +
                                    " was rescheduled to the same date and then deleted again. " +
                                    "Please confirm the Google occurrence, currently on " + movedToStartDate.ToString("dd-MMM-yyyy") + ", should be deleted?";
                                dr = Ogcs.Extensions.MessageBox.Show(msg, "Confirm deletion of recurring series occurrence", MessageBoxButtons.YesNo, MessageBoxIcon.Question, msg.Replace(summary, anonSummary));
                            }
                            if (dr != DialogResult.Yes) continue;

                            Forms.Main.Instance.Console.Update(Ogcs.Google.Calendar.GetEventSummary("<br/>Occurrence deleted.", gExcp, out String anonSummary2), anonSummary2, Console.Markup.calendar, verbose: true);
                            gExcp.Status = "cancelled";
                            log.Debug("Google Exception deleted for " + cancelledDate.ToString("dd/MM/yyyy"));
                            try {
                                Ogcs.Google.Calendar.Instance.UpdateCalendarEntry_save(ref gExcp);
                            } catch (System.Exception ex) {
                                Forms.Main.Instance.Console.UpdateWithError(Ogcs.Google.Calendar.GetEventSummary("Updated event exception failed to save.", gExcp, out String anonSummary3, true), ex, logEntry: anonSummary3);
                                log.Debug(Newtonsoft.Json.JsonConvert.SerializeObject(gExcp));
                                Ogcs.Exception.Analyse(ex, true);
                                if (Ogcs.Extensions.MessageBox.Show("Updated Google event exception failed to save. Continue with synchronisation?", "Sync item failed", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                                    continue;
                                else {
                                    throw new UserCancelledSyncException("User chose not to continue sync.");
                                }
                            }
                        }
                    }
                }
            } catch (System.Exception ex) {
                ex.Analyse("Could not sync deleted occurrences.");
                if (ex is UserCancelledSyncException) throw;
            }
            #endregion

            #region Modified occurrences
            try {
                foreach (Event aiExcp in aiExcps) {
                    try {
                        System.DateTime oExcp_currDate = aiExcp.Start.SafeDateTime();
                        log.Fine("Modified occurrence on " + oExcp_currDate.ToString("dd/MM/yyyy"));

                        GcalData.Event gExcp = getGoogleInstance(aiExcp, ev.RecurringEventId ?? ev.Id, dirtyCache, false);
                        if (gExcp == null || gExcp.Status == "cancelled") {
                            log.Warn(oExcp_currDate.ToString("dd/MM/yyyy") + " is deleted in Google, which does NOT match Outlook!");
                            String syncDirectionTip = (Sync.Engine.Calendar.Instance.Profile.SyncDirection.Id == Sync.Direction.Bidirectional.Id) ? "<br/><i>Ensure you <b>first</b> set OGCS to one-way sync O->G.</i>" : "";
                            Forms.Main.Instance.Console.Update(
                                Outlook.Graph.Calendar.GetEventSummary("<br/>" +
                                    "The occurrence on " + oExcp_currDate.ToShortDateString() + " does not exist in Google, but does in Outlook.<br/>" +
                                    "This can happen if, for example, you declined the occurrence (which is synced to Google) and proposed a new time that is subsequently accepted by the organiser.<br/>" +
                                    "<u>Suggested fix</u>: delete the entire series in Google and let OGCS recreate it." + syncDirectionTip, aiExcp, out String anonSummary)
                                , anonSummary, Console.Markup.warning);
                            continue;
                        } else {
                            log.Fine("Matching Google Event recurrence found.");
                            int excp_itemModified = 0;
                            try {
                                //Force a compare of the exception if both G and O have been modified within 24 hours
                                TimeSpan modifiedDiff = (TimeSpan)(gExcp.Updated - aiExcp.LastModifiedDateTime?.ToLocalTime());
                                log.Fine("Modification time difference (in days) between G and O exception: " + modifiedDiff);
                                Boolean forceCompare = modifiedDiff < TimeSpan.FromDays(1);
                                Calendar.UpdateCalendarEntry(aiExcp, gExcp, ref excp_itemModified, forceCompare);
                                if (forceCompare && excp_itemModified == 0) {
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
                                ex.Analyse("Could not compare modified recurring series exception.", true);
                                throw;
                            }
                            if (excp_itemModified > 0) {
                                try {
                                    Ogcs.Google.Calendar.Instance.UpdateCalendarEntry_save(ref gExcp);
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
                        }
                    } catch (System.Exception ex) {
                        ex.Analyse("Unable to process modified recurring series exception.");
                    }
                }
            } catch (System.Exception ex) {
                ex.Analyse();
            }
            #endregion
        }
        #endregion
    }
}
