using log4net;
using Microsoft.Graph;
using OutlookGoogleCalendarSync.Extensions;
using OutlookGoogleCalendarSync.GraphExtension;
using System;
using System.Collections.Generic;
using System.Linq;
using GcalData = Google.Apis.Calendar.v3.Data;
using Ogcs = OutlookGoogleCalendarSync;

namespace OutlookGoogleCalendarSync.Outlook.Graph {
    public class Recurrence {
        private static readonly ILog log = LogManager.GetLogger(typeof(Recurrence));

        public static PatternedRecurrence BuildOutlookPattern(GcalData.Event ev) {
            if (ev.Recurrence == null) { return null; }

            Dictionary<String, String> ruleBook = Google.Recurrence.ExplodeRrule(ev.Recurrence);
            if (ruleBook == null) {
                throw new ApplicationException("WARNING: The recurrence pattern is not compatible with Outlook. This event cannot be synced.");
            }
            log.Fine("Building Outlook recurrence pattern");
            PatternedRecurrence oPattern = new() { 
                Pattern = new() { Interval = 1 }, 
                Range = new() { Type = RecurrenceRangeType.NoEnd } 
            };

            #region RECURRENCE PATTERN
            //RRULE:FREQ=WEEKLY;UNTIL=20150906T000000Z;BYDAY=SA

            switch (ruleBook["FREQ"]) {
                case "DAILY": {
                        oPattern.Pattern.Type = RecurrencePatternType.Daily;
                        break;
                    }
                case "WEEKLY": {
                        oPattern.Pattern.Type = RecurrencePatternType.Weekly;
                        // Need to work out dayMask from "BY" pattern
                        // Eg "BYDAY=MO,TU,WE,TH,FR"
                        oPattern.Pattern.DaysOfWeek = getDoW(ruleBook["BYDAY"]);
                        break;
                    }
                case "MONTHLY": {
                        if (ruleBook.ContainsKey("BYSETPOS")) {
                            oPattern.Pattern.Type = RecurrencePatternType.RelativeMonthly;
                            oPattern.Pattern.Index = getByDayRelative(ruleBook["BYSETPOS"]);
                        }
                        if (ruleBook.ContainsKey("BYDAY")) {
                            oPattern.Pattern.Type = RecurrencePatternType.RelativeMonthly;
                            oPattern.Pattern.DaysOfWeek = getDoW(ruleBook["BYDAY"].TrimStart("-1".ToCharArray()));
                            if (ruleBook["BYDAY"].StartsWith("-1"))
                                oPattern.Pattern.Index = getByDayRelative("-1");
                            else
                                oPattern.Pattern.Index ??= getByDayRelative(ruleBook["BYDAY"].Substring(0, 1));

                        } else {
                            oPattern.Pattern.Type = RecurrencePatternType.AbsoluteMonthly;
                            if (ruleBook.ContainsKey("BYMONTHDAY"))
                                oPattern.Pattern.DayOfMonth = Convert.ToInt16(ruleBook["BYMONTHDAY"]);
                            else
                                oPattern.Pattern.DayOfMonth = ev.Start.SafeDateTime().Day;
                        }
                        break;
                    }

                case "YEARLY": {
                        if (ruleBook.ContainsKey("BYSETPOS")) {
                            oPattern.Pattern.Type = RecurrencePatternType.RelativeYearly;
                            oPattern.Pattern.Index = getByDayRelative(ruleBook["BYSETPOS"]);
                        } else {
                            oPattern.Pattern.Type = RecurrencePatternType.AbsoluteYearly;
                            if (ruleBook.ContainsKey("BYMONTHDAY"))
                                oPattern.Pattern.DayOfMonth = Convert.ToInt16(ruleBook["BYMONTHDAY"]);
                            else
                                oPattern.Pattern.DayOfMonth = ev.Start.SafeDateTime().Day;
                        }
                        if (ruleBook.ContainsKey("BYMONTH"))
                            oPattern.Pattern.Month = Convert.ToInt16(ruleBook["BYMONTH"]);
                        else
                            oPattern.Pattern.Month ??= ev.Start.SafeDateTime().Month;

                        if (ruleBook.ContainsKey("BYDAY")) {
                            if (ruleBook["BYDAY"].StartsWith("-1"))
                                oPattern.Pattern.Index = getByDayRelative("-1");
                            else
                                oPattern.Pattern.Index ??= getByDayRelative(ruleBook["BYDAY"].Substring(0, 1));
                            if (oPattern.Pattern.Index != null)
                                oPattern.Pattern.Type = RecurrencePatternType.RelativeYearly;
                            oPattern.Pattern.DaysOfWeek = getDoW(ruleBook["BYDAY"]);
                        }                    
                        break;
                    }
            }
            #endregion

            #region RANGE
            oPattern.Range.StartDate = ev.Start.SafeDateTime().ToGraphDate();
            if (ruleBook.ContainsKey("INTERVAL") && Convert.ToInt16(ruleBook["INTERVAL"]) > 1)
                oPattern.Pattern.Interval = Convert.ToInt16(ruleBook["INTERVAL"]);

            if (ruleBook.ContainsKey("COUNT")) {
                oPattern.Range.Type = RecurrenceRangeType.Numbered;
                oPattern.Range.NumberOfOccurrences = Convert.ToInt16(ruleBook["COUNT"]);
            }

            if (ruleBook.ContainsKey("UNTIL")) {
                if (ruleBook["UNTIL"].StartsWith("4500")) {
                    log.Warn("Outlook can't handle end dates this far in the future. Converting to no end date.");
                    oPattern.Range.Type = RecurrenceRangeType.NoEnd;
                    oPattern.Range.EndDate = null;
                } else {
                    System.DateTime endDate;
                    if (ruleBook["UNTIL"].Length == 8 && !ruleBook["UNTIL"].EndsWith("Z"))
                        endDate = System.DateTime.ParseExact(ruleBook["UNTIL"], "yyyyMMdd", System.Globalization.CultureInfo.InvariantCulture).Date;
                    else {
                        endDate = System.DateTime.ParseExact(ruleBook["UNTIL"], "yyyyMMddTHHmmssZ", System.Globalization.CultureInfo.InvariantCulture, System.Globalization.DateTimeStyles.AdjustToUniversal);
                        endDate = endDate.AddHours(TimezoneDB.GetUtcOffset(ev.End.TimeZone)).Date;
                    }
                    System.DateTime patternStart = oPattern.Range.StartDate?.SafeDateTime() ?? ev.Start.SafeDateTime();
                    if (endDate < patternStart) {
                        log.Debug("PatternStartDate: " + patternStart.ToString("yyyyMMddHHmmss"));
                        log.Debug("PatternEndDate:   " + ruleBook["UNTIL"].ToString());
                        String summary = Ogcs.Google.Calendar.GetEventSummary("The recurring Google event has an end date <i>before</i> the start date, which Outlook doesn't allow.<br/>" +
                            "The synced Outlook recurrence has been changed to a single occurrence.", ev, out String anonSummary, onlyIfNotVerbose: true);
                        Forms.Main.Instance.Console.Update(summary, anonSummary, Console.Markup.warning);
                        oPattern.Range.NumberOfOccurrences = 1;
                        oPattern.Range.Type = RecurrenceRangeType.Numbered;
                    } else {
                        oPattern.Range.Type = RecurrenceRangeType.EndDate;
                        oPattern.Range.EndDate = endDate.ToGraphDate();
                    }
                }
            }
            #endregion

            return oPattern;
        }

        public static PatternedRecurrence CompareOutlookPattern(GcalData.Event ev, PatternedRecurrence aiOpattern, Sync.Direction syncDirection, System.Text.StringBuilder sb, ref int itemModified) {
            if (ev.Recurrence == null) return null;
            
            log.Fine("Building a temporary recurrent Appointment generated from Event");
            PatternedRecurrence evOpattern = BuildOutlookPattern(ev);
            
            log.Fine("Comparing Google recurrence to Outlook equivalent");
            #region Recurrence Pattern
            //Set defaults to avoid false changes
            evOpattern.Pattern.FirstDayOfWeek ??= Microsoft.Graph.DayOfWeek.Sunday;
            evOpattern.Pattern.Index ??= WeekIndex.First;

            if (Sync.Engine.CompareAttribute("Recurrence Type", syncDirection,
                evOpattern.Pattern.Type.ToString(), aiOpattern.Pattern.Type.ToString(), sb, ref itemModified)) {
                aiOpattern.Pattern.Type = evOpattern.Pattern.Type;
            }
            if (Sync.Engine.CompareAttribute("Recurrence Interval", syncDirection,
                evOpattern.Pattern.Interval.ToString(), aiOpattern.Pattern.Interval.ToString(), sb, ref itemModified)) {
                aiOpattern.Pattern.Interval = evOpattern.Pattern.Interval;
            }
            if (Sync.Engine.CompareAttribute("Recurrence Index", syncDirection,
                evOpattern.Pattern.Index?.ToString(), aiOpattern.Pattern.Index?.ToString(), sb, ref itemModified)) {
                aiOpattern.Pattern.Index = evOpattern.Pattern.Index;
            }
            if (Sync.Engine.CompareAttribute("Recurrence DoW", syncDirection,
                string.Join(",", evOpattern.Pattern.DaysOfWeek ?? new List<Microsoft.Graph.DayOfWeek>()),
                string.Join(",", aiOpattern.Pattern.DaysOfWeek ?? new List<Microsoft.Graph.DayOfWeek>()), sb, ref itemModified)) {
                aiOpattern.Pattern.DaysOfWeek = evOpattern.Pattern.DaysOfWeek;

            }
            if (Sync.Engine.CompareAttribute("Recurrence MoY", syncDirection,
                convertEquivalenceToNull(evOpattern.Pattern.Month).ToString(), convertEquivalenceToNull(aiOpattern.Pattern.Month).ToString(), sb, ref itemModified)) {
                aiOpattern.Pattern.Month = evOpattern.Pattern.Month ?? 0;
            }
            if (Sync.Engine.CompareAttribute("Recurrence 1stDoW", syncDirection,
                evOpattern.Pattern.FirstDayOfWeek?.ToString(), aiOpattern.Pattern.FirstDayOfWeek?.ToString(), sb, ref itemModified)) {
                aiOpattern.Pattern.FirstDayOfWeek = evOpattern.Pattern.FirstDayOfWeek;
            }
            if (Sync.Engine.CompareAttribute("Recurrence DoM", syncDirection,
                convertEquivalenceToNull(evOpattern.Pattern.DayOfMonth).ToString(), convertEquivalenceToNull(aiOpattern.Pattern.DayOfMonth).ToString(), sb, ref itemModified)) {
                aiOpattern.Pattern.DayOfMonth = evOpattern.Pattern.DayOfMonth ?? 0;
            }
            #endregion
            #region Range
            if (Sync.Engine.CompareAttribute("Recurrence EndDate", syncDirection,
                convertEquivalenceToNull(evOpattern.Range.EndDate, new(1, 1, 1))?.ToString(), convertEquivalenceToNull(aiOpattern.Range.EndDate, new Date(1, 1, 1))?.ToString(), sb, ref itemModified)) {
                aiOpattern.Range.EndDate = evOpattern.Range.EndDate ?? new(1, 1, 1);
                aiOpattern.Range.Type = evOpattern.Range.Type;
            }
            if (Sync.Engine.CompareAttribute("Recurrence Occurences", syncDirection,
                convertEquivalenceToNull(evOpattern.Range.NumberOfOccurrences).ToString(), convertEquivalenceToNull(aiOpattern.Range.NumberOfOccurrences).ToString(), sb, ref itemModified)) {
                aiOpattern.Range.NumberOfOccurrences = evOpattern.Range.NumberOfOccurrences ?? 0;
                aiOpattern.Range.Type = evOpattern.Range.Type;
            }
            #endregion

            return aiOpattern;
        }

        private static List<Microsoft.Graph.DayOfWeek> getDoW(String byDay) {
            List<Microsoft.Graph.DayOfWeek> daysOfWeek = new();
            if (!string.IsNullOrEmpty(byDay)) {
                if (byDay.Contains("MO")) daysOfWeek.Add(Microsoft.Graph.DayOfWeek.Monday);
                if (byDay.Contains("TU")) daysOfWeek.Add(Microsoft.Graph.DayOfWeek.Tuesday);
                if (byDay.Contains("WE")) daysOfWeek.Add(Microsoft.Graph.DayOfWeek.Wednesday);
                if (byDay.Contains("TH")) daysOfWeek.Add(Microsoft.Graph.DayOfWeek.Thursday);
                if (byDay.Contains("FR")) daysOfWeek.Add(Microsoft.Graph.DayOfWeek.Friday);
                if (byDay.Contains("SA")) daysOfWeek.Add(Microsoft.Graph.DayOfWeek.Saturday);
                if (byDay.Contains("SU")) daysOfWeek.Add(Microsoft.Graph.DayOfWeek.Sunday);
            }
            return daysOfWeek;
        }

        private static WeekIndex? getByDayRelative(String byDayRule) {
            switch (byDayRule) {
                case "1": return WeekIndex.First;
                case "2": return WeekIndex.Second; 
                case "3": return WeekIndex.Third; 
                case "4": return WeekIndex.Fourth; 
                case "-1": return WeekIndex.Last; 
            }
            return null;
        }

        #region NULL helper functions
        /// <summary>Return null if the two parameter int values are equivalent</summary>
        private static int? convertEquivalenceToNull(int? value, int nullValue = 0) {
            return ((value ?? nullValue) == nullValue ? null : value);
        }
        /// <summary>Return null if the two parameter Date values are equivalent</summary>
        #pragma warning disable CS8632 // The annotation for nullable reference types should only be used in code within a '#nullable' annotations context.
        private static Date? convertEquivalenceToNull(Date? value, Date nullValue) {
            return ((value ?? nullValue).Compare(nullValue) ? null : value);
        }
        #pragma warning restore CS8632
        #endregion

        #region Exceptions
        private static List<Event> outlookExceptions;
        public static List<Event> OutlookExceptions {
            get { return outlookExceptions; }
        }
        public static List<Event> GetExceptions(Event ai) {
            return outlookExceptions.Where(aiExcp => aiExcp.SeriesMasterId == ai.Id).ToList();
        }
        public static void SeparateOutlookExceptions(List<Event> allAppointments) {
            outlookExceptions = new List<Event>();
            if (allAppointments.Count == 0) return;
            log.Debug("Identifying exceptions in recurring Outlook appointments.");

            for (int o = allAppointments.Count - 1; o >= 0; o--) {
                Event ai = allAppointments[o];
                if (!string.IsNullOrEmpty(ai.SeriesMasterId) && ai.Type == EventType.Exception) {
                    outlookExceptions.Add(ai);
                    allAppointments.Remove(ai);
                }
            }
            log.Debug("Found " + outlookExceptions.Count + " exceptions.");
        }

        /// <summary>
        /// Get missing series master Events for occurrences falling within the sync window.
        /// </summary>
        /// <param name="allAppointments">Single instance, occurences and exceptions.</param>
        public static void GetOutlookMasterEvent(List<Event> allAppointments) {
            List<String> seriesMasterIds = allAppointments.Where(ai => ai.Type == EventType.SeriesMaster).Select(ai => ai.Id).ToList();
            List<String> seriesInstanceIds = allAppointments.Where(ai => ai.SeriesMasterId != null).Select(ai => ai.SeriesMasterId).Distinct().ToList();
            int newMasterEvents = 0;
            if (seriesInstanceIds.Count > 0) {
                log.Info("Retrieving master series appointments for occurrences falling within the sync window.");
                foreach (String masterId in seriesInstanceIds.Except(seriesMasterIds)) {
                    Event ai = Calendar.Instance.GetCalendarEntry(masterId);
                    if (ai != null) {
                        allAppointments.Add(ai);
                        newMasterEvents++;
                    }
                }
                log.Debug(newMasterEvents + " master Graph Events retrieved for occurrences falling within the sync window.");
            }
        }

        public static void CreateOutlookExceptions(GcalData.Event ev, Event createdAi) {
            if (ev.Recurrence == null || ev.RecurringEventId != null) return;

            try {
                List<GcalData.Event> evExceptions = Google.Recurrence.GoogleExceptions.Where(exp => exp.RecurringEventId == ev.Id).ToList();
                if (evExceptions.Count == 0) return;

                Forms.Main.Instance.Console.Update("This is a recurring item with some exceptions:-", verbose: true);

                log.Debug("Creating Google recurrence exceptions.");
                List<Event> oRecurrences = Calendar.Instance.GetCalendarEntriesInRecurrence(createdAi.Id);
                if ((oRecurrences?.Count ?? 0) == 0) return;

                log.Debug($"Modifying {evExceptions.Count} occurrences.");
                foreach (GcalData.Event gExcp in evExceptions) {
                    System.DateTime gExcpOrigDate = gExcp.OriginalStartTime.SafeDateTime();
                    System.DateTime? gExcpCurrDate = gExcp.Start?.SafeDateTime();
                    log.Fine($"Found Google exception with {gExcp.Status} original date " + gExcpOrigDate.ToString() + (gExcpCurrDate != null ? " now on " + gExcpCurrDate?.ToString() : ""));

                    try {
                        Event newAiExcp = oRecurrences.Where(ai => ai.Start.SafeDateTime() == gExcpOrigDate).FirstOrDefault();
                        if (newAiExcp == null) {
                            if (gExcp.Status == "cancelled") {
                                log.Warn($"Could not find Outlook occurrence for Google's cancellation on {gExcpOrigDate.ToString("dd-MM-yyyy")}");
                            } else {
                                log.Error($"Could not find Outlook occurrence for Google's exception on {gExcpCurrDate?.ToString("dd-MM-yyyy")}");
                            }
                            continue;
                        }
                        if (gExcp.Status == "cancelled") {
                            Forms.Main.Instance.Console.Update(Outlook.Graph.Calendar.GetEventSummary("<br/>Occurrence deleted.", newAiExcp, out String anonSummary), anonSummary, Console.Markup.calendar, verbose: true);
                            oRecurrences.Remove(newAiExcp);
                            Calendar.Instance.DeleteCalendarEntry_save(newAiExcp);
                        /*
                        } else if (Sync.Engine.Calendar.Instance.Profile.ExcludeDeclinedInvites && gExcp.Attendees != null && gExcp.Attendees.Count(a => a.Self == true && a.ResponseStatus == "declined") == 1) {
                            Forms.Main.Instance.Console.Update(Outlook.Calendar.GetEventSummary("<br/>Occurrence declined.", newAiExcp, out String anonSummary), anonSummary, Console.Markup.calendar, verbose: true);
                            newAiExcp.Delete();
                        */
                        } else {
                            int itemModified = 0;
                            Event aiPatch = new();
                            if (Outlook.Graph.Calendar.Instance.UpdateCalendarEntry(ref newAiExcp, gExcp, ref itemModified, out aiPatch, true)) {
                                try {
                                    Calendar.Instance.UpdateCalendarEntry_save(ref aiPatch);
                                } catch (Microsoft.Graph.ServiceException ex) {
                                    if (ex.Error.Code == "ErrorOccurrenceCrossingBoundary") {
                                        Forms.Main.Instance.Console.Update(
                                            Outlook.Graph.Calendar.GetEventSummary("Uh oh! Outlook wasn't able to save this recurrence exception! " +
                                                "You may have two occurences on the same day, which it doesn't allow even if the original occurrence has been deleted.",
                                                newAiExcp, out String anonSummary, true)
                                            , anonSummary, Console.Markup.warning);
                                    } else
                                        ex.Analyse();
                                }
                            }
                        }
                    } catch (System.Exception ex) {
                        ex.Analyse($"Failed to process modified Google occurrence on {gExcpOrigDate} for new Outlook series.");
                    }
                }
            } catch (System.Exception ex2) {
                ex2.Analyse("Failed to process modified Google occurrences on new Outlook series.");
            }
            Forms.Main.Instance.Console.Update("Recurring exceptions completed.", verbose: true);
        }

        public static void UpdateOutlookExceptions(GcalData.Event ev, Event ai, Boolean forceCompare) {
            if (ev.Recurrence == null || ev.RecurringEventId != null) return;

            try {
                List<GcalData.Event> evExceptions = Google.Recurrence.GoogleExceptions.Where(exp => exp.RecurringEventId == ev.Id).ToList();
                if (evExceptions.Count == 0) return;

                log.Fine($"{evExceptions.Count} Google recurrence exceptions within sync range to be compared.");
                List<Event> oRecurrences = Calendar.Instance.GetCalendarEntriesInRecurrence(ai.Id);
                if ((oRecurrences?.Count ?? 0) == 0) return;

                foreach (GcalData.Event gExcp in evExceptions) {
                    System.DateTime gExcpOrigDate = gExcp.OriginalStartTime.SafeDateTime();
                    System.DateTime? gExcpCurrDate = gExcp.Start?.SafeDateTime();
                    log.Fine($"Found Google exception with {gExcp.Status} original date " + gExcpOrigDate.ToString() + (gExcpCurrDate != null ? " now on " + gExcpCurrDate?.ToShortDateString() : ""));

                    try {
                        Event newAiExcp = oRecurrences.Where(ai => ai.OriginalStart == gExcpOrigDate).FirstOrDefault();
                        if (newAiExcp == null) {
                            if (gExcp.Status == "cancelled") {
                                if (Calendar.Instance.CancelledOccurrences[ai.Id]?.Contains(gExcpOrigDate.Date) ?? false)
                                    log.Fine($"Outlook occurrence for Google's cancellation on {gExcpOrigDate.ToString("dd-MM-yyyy")} already deleted.");
                                else
                                    log.Warn($"Could not find Outlook occurrence for Google's cancellation on {gExcpOrigDate.ToString("dd-MM-yyyy")}");
                            } else {
                                log.Warn("Unable to find Outlook exception for " + gExcpCurrDate);
                                log.Warn("Google is NOT deleted though - a mismatch has occurred somehow!");
                                String syncDirectionTip = (Sync.Engine.Calendar.Instance.Profile.SyncDirection.Id == Sync.Direction.Bidirectional.Id) ? "<br/><i>Ensure you <b>first</b> set OGCS to one-way sync G->O.</i>" : "";
                                Forms.Main.Instance.Console.Update(Ogcs.Google.Calendar.GetEventSummary(
                                        "<br/>This occurrence cannot be found in Outlook.<br/>" +
                                        "This can happen if, for example, the occurrence has been rearranged to different days more than once.<br/>" +
                                        "<u>Suggested fix</u>: delete the entire series in Outlook and let OGCS recreate it." + syncDirectionTip, gExcp, out String anonSummary),
                                    anonSummary, Console.Markup.warning);
                            }
                            continue;
                        }
                        if (gExcp.Status == "cancelled") {
                            Forms.Main.Instance.Console.Update(Outlook.Graph.Calendar.GetEventSummary("<br/>Occurrence deleted.", newAiExcp, out String anonSummary), anonSummary, Console.Markup.calendar, verbose: true);
                            oRecurrences.Remove(newAiExcp);
                            Calendar.Instance.DeleteCalendarEntry_save(newAiExcp);
                            /*
                            } else if (Sync.Engine.Calendar.Instance.Profile.ExcludeDeclinedInvites && gExcp.Attendees != null && gExcp.Attendees.Count(a => a.Self == true && a.ResponseStatus == "declined") == 1) {
                                Forms.Main.Instance.Console.Update(Outlook.Calendar.GetEventSummary("<br/>Occurrence declined.", newAiExcp, out String anonSummary), anonSummary, Console.Markup.calendar, verbose: true);
                                newAiExcp.Delete();
                            */
                        } else {
                            int itemModified = 0;
                            Event aiPatch = new();
                            Outlook.Graph.Calendar.Instance.UpdateCalendarEntry(ref newAiExcp, gExcp, ref itemModified, out aiPatch,
                                forceCompare || gExcp.Start.SafeDateTime().Date != newAiExcp.Start.SafeDateTime().Date);
                            if (itemModified > 0) {
                                try {
                                    Calendar.Instance.UpdateCalendarEntry_save(ref aiPatch);
                                } catch (Microsoft.Graph.ServiceException ex) {
                                    if (ex.Error.Code == "ErrorOccurrenceCrossingBoundary") {
                                        Forms.Main.Instance.Console.Update(
                                            Outlook.Graph.Calendar.GetEventSummary("Uh oh! Outlook wasn't able to save this recurrence exception! " +
                                                "You may have two occurences on the same day, which it doesn't allow even if the original occurrence has been deleted.",
                                                newAiExcp, out String anonSummary, true)
                                            , anonSummary, Console.Markup.warning);
                                    } else
                                        ex.Analyse();
                                }
                            }
                        }
                    } catch (System.Exception ex) {
                        ex.Analyse($"Failed to process modified Google occurrence on {gExcpOrigDate} for new Outlook series.");
                    }
                }
            } catch (System.Exception ex2) {
                ex2.Analyse("Failed to process modified Google occurrences on new Outlook series.");
            }
        }
        #endregion
    }
}
