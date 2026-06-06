using log4net;
using Microsoft.Kiota.Abstractions;
using OutlookGoogleCalendarSync.Extensions;
using OutlookGoogleCalendarSync.GraphExtension;
using System;
using System.Collections.Generic;
using System.Linq;
using GcalData = Google.Apis.Calendar.v3.Data;
using MsGraph = OutlookGoogleCalendarSync.Outlook.Graph.CustomClient;
using Ogcs = OutlookGoogleCalendarSync;

namespace OutlookGoogleCalendarSync.Outlook.Graph {
    public class Recurrence {
        private static readonly ILog log = LogManager.GetLogger(typeof(Recurrence));

        public static MsGraph.Models.PatternedRecurrence BuildOutlookPattern(GcalData.Event ev) {
            if (ev.Recurrence == null) { return null; }

            Dictionary<String, String> ruleBook = Google.Recurrence.ExplodeRrule(ev.Recurrence);
            if (ruleBook == null) {
                throw new ApplicationException("WARNING: The recurrence pattern is not compatible with Outlook. This event cannot be synced.");
            }
            log.Fine("Building Outlook recurrence pattern");
            MsGraph.Models.PatternedRecurrence oPattern = new() { 
                Pattern = new() { Interval = 1 }, 
                Range = new() { Type = MsGraph.Models.RecurrenceRangeType.NoEnd } 
            };

            #region RECURRENCE PATTERN
            //RRULE:FREQ=WEEKLY;UNTIL=20150906T000000Z;BYDAY=SA

            switch (ruleBook["FREQ"]) {
                case "DAILY": {
                        oPattern.Pattern.Type = MsGraph.Models.RecurrencePatternType.Daily;
                        break;
                    }
                case "WEEKLY": {
                        oPattern.Pattern.Type = MsGraph.Models.RecurrencePatternType.Weekly;
                        // Need to work out dayMask from "BY" pattern
                        // Eg "BYDAY=MO,TU,WE,TH,FR"
                        if (ruleBook.ContainsKey("BYDAY"))
                            oPattern.Pattern.DaysOfWeek = getDoW(ruleBook["BYDAY"]);
                        else
                            oPattern.Pattern.DaysOfWeek = getDoW(ev.Start.SafeDateTimeOffset().DayOfWeek.ToString().ToUpper().Substring(0, 2));
                        break;
                    }
                case "MONTHLY": {
                        if (ruleBook.ContainsKey("BYSETPOS")) {
                            oPattern.Pattern.Type = MsGraph.Models.RecurrencePatternType.RelativeMonthly;
                            oPattern.Pattern.Index = getByDayRelative(ruleBook["BYSETPOS"]);
                        }
                        if (ruleBook.ContainsKey("BYDAY")) {
                            oPattern.Pattern.Type = MsGraph.Models.RecurrencePatternType.RelativeMonthly;
                            oPattern.Pattern.DaysOfWeek = getDoW(ruleBook["BYDAY"].TrimStart("-1".ToCharArray()));
                            if (ruleBook["BYDAY"].StartsWith("-1"))
                                oPattern.Pattern.Index = getByDayRelative("-1");
                            else
                                oPattern.Pattern.Index ??= getByDayRelative(ruleBook["BYDAY"].Substring(0, 1));

                        } else {
                            oPattern.Pattern.Type = MsGraph.Models.RecurrencePatternType.AbsoluteMonthly;
                            if (ruleBook.ContainsKey("BYMONTHDAY"))
                                oPattern.Pattern.DayOfMonth = Convert.ToInt16(ruleBook["BYMONTHDAY"]);
                            else
                                oPattern.Pattern.DayOfMonth = ev.Start.SafeDateTimeOffset().Day;
                        }
                        break;
                    }

                case "YEARLY": {
                        if (ruleBook.ContainsKey("BYSETPOS")) {
                            oPattern.Pattern.Type = MsGraph.Models.RecurrencePatternType.RelativeYearly;
                            oPattern.Pattern.Index = getByDayRelative(ruleBook["BYSETPOS"]);
                        } else {
                            oPattern.Pattern.Type = MsGraph.Models.RecurrencePatternType.AbsoluteYearly;
                            if (ruleBook.ContainsKey("BYMONTHDAY"))
                                oPattern.Pattern.DayOfMonth = Convert.ToInt16(ruleBook["BYMONTHDAY"]);
                            else
                                oPattern.Pattern.DayOfMonth = ev.Start.SafeDateTimeOffset().Day;
                        }
                        if (ruleBook.ContainsKey("BYMONTH"))
                            oPattern.Pattern.Month = Convert.ToInt16(ruleBook["BYMONTH"]);
                        else
                            oPattern.Pattern.Month ??= ev.Start.SafeDateTimeOffset().Month;

                        if (ruleBook.ContainsKey("BYDAY")) {
                            if (ruleBook["BYDAY"].StartsWith("-1"))
                                oPattern.Pattern.Index = getByDayRelative("-1");
                            else
                                oPattern.Pattern.Index ??= getByDayRelative(ruleBook["BYDAY"].Substring(0, 1));
                            if (oPattern.Pattern.Index != null)
                                oPattern.Pattern.Type = MsGraph.Models.RecurrencePatternType.RelativeYearly;
                            oPattern.Pattern.DaysOfWeek = getDoW(ruleBook["BYDAY"]);
                        }                    
                        break;
                    }
            }
            #endregion

            #region RANGE
            oPattern.Range.StartDate = ev.Start.SafeDateTimeOffset().ToGraphDate();
            if (ruleBook.ContainsKey("INTERVAL") && Convert.ToInt16(ruleBook["INTERVAL"]) > 1)
                oPattern.Pattern.Interval = Convert.ToInt16(ruleBook["INTERVAL"]);

            if (ruleBook.ContainsKey("COUNT")) {
                oPattern.Range.Type = MsGraph.Models.RecurrenceRangeType.Numbered;
                oPattern.Range.NumberOfOccurrences = Convert.ToInt16(ruleBook["COUNT"]);
            }

            if (ruleBook.ContainsKey("UNTIL")) {
                if (ruleBook["UNTIL"].StartsWith("4500")) {
                    log.Warn("Outlook can't handle end dates this far in the future. Converting to no end date.");
                    oPattern.Range.Type = MsGraph.Models.RecurrenceRangeType.NoEnd;
                    oPattern.Range.EndDate = null;
                } else {
                    System.DateTimeOffset endDate;
                    if (ruleBook["UNTIL"].Length == 8 && !ruleBook["UNTIL"].EndsWith("Z"))
                        endDate = System.DateTime.ParseExact(ruleBook["UNTIL"], "yyyyMMdd", System.Globalization.CultureInfo.InvariantCulture).Date;
                    else {
                        endDate = System.DateTime.ParseExact(ruleBook["UNTIL"], "yyyyMMddTHHmmssZ", System.Globalization.CultureInfo.InvariantCulture, System.Globalization.DateTimeStyles.AdjustToUniversal);
                        endDate = endDate.AddMinutes(TimezoneDB.GetUtcOffset(ev.End.TimeZone)).Date;
                    }
                    System.DateTimeOffset patternStart = oPattern.Range.StartDate != null ? ((Date)oPattern.Range.StartDate).SafeDateTimeOffset() : ev.Start.SafeDateTimeOffset();
                    if (endDate < patternStart) {
                        log.Debug("PatternStartDate: " + patternStart.ToString("yyyyMMddHHmmss"));
                        log.Debug("PatternEndDate:   " + ruleBook["UNTIL"].ToString());
                        String summary = Ogcs.Google.Calendar.GetEventSummary("The recurring Google event has an end date <i>before</i> the start date, which Outlook doesn't allow.<br/>" +
                            "The synced Outlook recurrence has been changed to a single occurrence.", ev, out String anonSummary, onlyIfNotVerbose: true);
                        Forms.Main.Instance.Console.Update(summary, anonSummary, Console.Markup.warning);
                        oPattern.Range.NumberOfOccurrences = 1;
                        oPattern.Range.Type = MsGraph.Models.RecurrenceRangeType.Numbered;
                    } else {
                        oPattern.Range.Type = MsGraph.Models.RecurrenceRangeType.EndDate;
                        oPattern.Range.EndDate = endDate.ToGraphDate();
                    }
                }
            }
            #endregion

            return oPattern;
        }

        public static MsGraph.Models.PatternedRecurrence CompareOutlookPattern(GcalData.Event ev, MsGraph.Models.PatternedRecurrence aiOpattern, Sync.Direction syncDirection, System.Text.StringBuilder sb, ref int itemModified) {
            if (ev.Recurrence == null) return null;
            
            log.Fine("Building a temporary recurrent Appointment generated from Event");
            MsGraph.Models.PatternedRecurrence evOpattern = BuildOutlookPattern(ev);
            
            log.Fine("Comparing Google recurrence to Outlook equivalent");
            #region Recurrence Pattern
            //Set defaults to avoid false changes
            evOpattern.Pattern.FirstDayOfWeek ??= MsGraph.Models.DayOfWeekObject.Sunday;
            evOpattern.Pattern.Index ??= MsGraph.Models.WeekIndex.First;

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
                string.Join(",", evOpattern.Pattern.DaysOfWeek ?? new List<MsGraph.Models.DayOfWeekObject?>()),
                string.Join(",", aiOpattern.Pattern.DaysOfWeek ?? new List<MsGraph.Models.DayOfWeekObject?>()), sb, ref itemModified)) {
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
                convertEquivalenceToNull(evOpattern.Range.EndDate, new Date(1, 1, 1))?.ToString(), convertEquivalenceToNull(aiOpattern.Range.EndDate, new Date(1, 1, 1))?.ToString(), sb, ref itemModified)) {
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

        private static List<MsGraph.Models.DayOfWeekObject?> getDoW(String byDay) {
            List<MsGraph.Models.DayOfWeekObject?> daysOfWeek = new();
            if (!string.IsNullOrEmpty(byDay)) {
                if (byDay.Contains("MO")) daysOfWeek.Add(MsGraph.Models.DayOfWeekObject.Monday);
                if (byDay.Contains("TU")) daysOfWeek.Add(MsGraph.Models.DayOfWeekObject.Tuesday);
                if (byDay.Contains("WE")) daysOfWeek.Add(MsGraph.Models.DayOfWeekObject.Wednesday);
                if (byDay.Contains("TH")) daysOfWeek.Add(MsGraph.Models.DayOfWeekObject.Thursday);
                if (byDay.Contains("FR")) daysOfWeek.Add(MsGraph.Models.DayOfWeekObject.Friday);
                if (byDay.Contains("SA")) daysOfWeek.Add(MsGraph.Models.DayOfWeekObject.Saturday);
                if (byDay.Contains("SU")) daysOfWeek.Add(MsGraph.Models.DayOfWeekObject.Sunday);
            }
            return daysOfWeek;
        }

        private static MsGraph.Models.WeekIndex? getByDayRelative(String byDayRule) {
            switch (byDayRule) {
                case "1": return MsGraph.Models.WeekIndex.First;
                case "2": return MsGraph.Models.WeekIndex.Second; 
                case "3": return MsGraph.Models.WeekIndex.Third; 
                case "4": return MsGraph.Models.WeekIndex.Fourth; 
                case "-1": return MsGraph.Models.WeekIndex.Last; 
            }
            return null;
        }

        #region NULL helper functions
        /// <summary>Return null if the two parameter int values are equivalent</summary>
        private static int? convertEquivalenceToNull(int? value, int nullValue = 0) {
            return ((value ?? nullValue) == nullValue ? null : value);
        }
        /// <summary>Return null if the two parameter Date values are equivalent</summary>
        private static Date? convertEquivalenceToNull(Date? value, Date nullValue) {
            return ((value ?? nullValue).Compare(nullValue) ? null : value);
        }
        #endregion

        #region Exceptions
        private static List<MsGraph.Models.Event> outlookExceptions;
        public static List<MsGraph.Models.Event> OutlookExceptions {
            get { return outlookExceptions; }
        }
        public static List<MsGraph.Models.Event> GetExceptions(MsGraph.Models.Event ai) {
            return outlookExceptions.Where(aiExcp => aiExcp.SeriesMasterId == ai.Id).ToList();
        }
        public static void SeparateOutlookExceptions(List<MsGraph.Models.Event> allAppointments) {
            outlookExceptions = new List<MsGraph.Models.Event>();
            if (allAppointments.Count == 0) return;
            log.Debug("Identifying exceptions in recurring Outlook appointments.");

            for (int o = allAppointments.Count - 1; o >= 0; o--) {
                MsGraph.Models.Event ai = allAppointments[o];
                if (!string.IsNullOrEmpty(ai.SeriesMasterId) && ai.Type == MsGraph.Models.EventType.Exception) {
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
        public static void GetOutlookMasterEvent(List<MsGraph.Models.Event> allAppointments) {
            List<String> seriesMasterIds = allAppointments.Where(ai => ai.Type == MsGraph.Models.EventType.SeriesMaster).Select(ai => ai.Id).ToList();
            List<String> seriesInstanceIds = allAppointments.Where(ai => ai.SeriesMasterId != null).Select(ai => ai.SeriesMasterId).Distinct().ToList();
            int newMasterEvents = 0;
            if (seriesInstanceIds.Count > 0) {
                log.Info("Retrieving master series appointments for occurrences falling within the sync window.");
                foreach (String masterId in seriesInstanceIds.Except(seriesMasterIds)) {
                    MsGraph.Models.Event ai = Calendar.Instance.GetCalendarEntry(masterId);
                    if (ai != null) {
                        allAppointments.Add(ai);
                        newMasterEvents++;
                    }
                }
                log.Debug(newMasterEvents + " master Graph Events retrieved for occurrences falling within the sync window.");
            }
        }

        public static int CreateOutlookExceptions(GcalData.Event ev, MsGraph.Models.Event createdAi) {
            int updatesMade = 0;
            if (ev.Recurrence == null || ev.RecurringEventId != null) return updatesMade;

            try {
                List<GcalData.Event> evExceptions = Google.Recurrence.GoogleExceptions.Where(exp => exp.RecurringEventId == ev.Id).ToList();
                if (evExceptions.Count == 0) return updatesMade;

                Forms.Main.Instance.Console.Update("This is a recurring item with some exceptions:-", verbose: true);

                log.Debug("Creating Google recurrence exceptions.");
                List<MsGraph.Models.Event> oRecurrences = Calendar.Instance.GetCalendarEntriesInRecurrence(createdAi.Id);
                if ((oRecurrences?.Count ?? 0) == 0) return updatesMade;

                log.Debug($"Modifying {evExceptions.Count} occurrences.");
                foreach (GcalData.Event gExcp in evExceptions) {
                    System.DateTimeOffset gExcpOrigDate = gExcp.OriginalStartTime.SafeDateTimeOffset();
                    System.DateTimeOffset? gExcpCurrDate = gExcp.Start?.SafeDateTimeOffset();
                    log.Fine($"Found Google exception with {gExcp.Status} original date {gExcpOrigDate.DateTime.ToShortDateString()}" + (gExcpCurrDate != null ? " now on " + gExcpCurrDate?.DateTime.ToShortDateString() : ""));

                    try {
                        MsGraph.Models.Event newAiExcp = oRecurrences.Where(ai => ai.Start.SafeDateTimeOffset() == gExcpOrigDate).FirstOrDefault();
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
                            updatesMade++;
                        /*
                        } else if (Sync.Engine.Calendar.Instance.Profile.ExcludeDeclinedInvites && gExcp.Attendees != null && gExcp.Attendees.Count(a => a.Self == true && a.ResponseStatus == "declined") == 1) {
                            Forms.Main.Instance.Console.Update(Outlook.Calendar.GetEventSummary("<br/>Occurrence declined.", newAiExcp, out String anonSummary), anonSummary, Console.Markup.calendar, verbose: true);
                            newAiExcp.Delete();
                            updatesMade++;
                        */
                        } else {
                            int itemModified = 0;
                            MsGraph.Models.Event aiPatch = new();
                            if (Outlook.Graph.Calendar.Instance.UpdateCalendarEntry(ref newAiExcp, gExcp, ref itemModified, out aiPatch, true)) {
                                try {
                                    Calendar.Instance.UpdateCalendarEntry_save(ref aiPatch);
                                    updatesMade++;
                                } catch (Microsoft.Kiota.Abstractions.ApiException ex) {
                                    if (ex.ResponseStatusCode == 400 && O365Errors.GetODataError(ex)?.Error?.Code == "ErrorOccurrenceCrossingBoundary") {
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
            
            return updatesMade;
        }

        public static int UpdateOutlookExceptions(GcalData.Event ev, MsGraph.Models.Event ai, Boolean forceCompare) {
            int updatesMade = 0;
            if (ev.Recurrence == null || ev.RecurringEventId != null) return updatesMade;

            try {
                List<GcalData.Event> evExceptions = Google.Recurrence.GoogleExceptions.Where(exp => exp.RecurringEventId == ev.Id).ToList();
                if (evExceptions.Count == 0) return updatesMade;

                log.Debug($"{evExceptions.Count} Google recurrence exceptions within sync range to be compared.");

                List<GcalData.Event> gCancelledExcps = evExceptions.Where(exp => exp.Status == "cancelled").ToList();
                log.Fine($"{gCancelledExcps.Count} Google cancelled occurrences.");
                log.Fine($"{evExceptions.Count - gCancelledExcps.Count} Google modified exceptions.");

                List<MsGraph.Models.Event> oRecurrences = Calendar.Instance.GetCalendarEntriesInRecurrence(ai.Id);
                if ((oRecurrences?.Count ?? 0) == 0) return updatesMade;

                foreach (GcalData.Event gExcp in evExceptions) {
                    System.DateTimeOffset gExcpOrigDate = gExcp.OriginalStartTime.SafeDateTimeOffset();
                    System.DateTimeOffset gExcpCurrDate = gExcp.Start.SafeDateTimeOffset();
                    log.Fine($"Found Google exception with {gExcp.Status} original date {gExcpOrigDate.DateTime.ToShortDateString()}" + (gExcpCurrDate != null ? " now on " + gExcpCurrDate.DateTime.ToShortDateString() : ""));

                    try {
                        MsGraph.Models.Event newAiExcp = oRecurrences.Where(ai => ai.OriginalStart == gExcpOrigDate).FirstOrDefault();
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
                            updatesMade++;
                            /*
                            } else if (Sync.Engine.Calendar.Instance.Profile.ExcludeDeclinedInvites && gExcp.Attendees != null && gExcp.Attendees.Count(a => a.Self == true && a.ResponseStatus == "declined") == 1) {
                                Forms.Main.Instance.Console.Update(Outlook.Calendar.GetEventSummary("<br/>Occurrence declined.", newAiExcp, out String anonSummary), anonSummary, Console.Markup.calendar, verbose: true);
                                newAiExcp.Delete();
                                updatesMade++;
                            */
                        } else {
                            int itemModified = 0;
                            MsGraph.Models.Event aiPatch = new();
                            Outlook.Graph.Calendar.Instance.UpdateCalendarEntry(ref newAiExcp, gExcp, ref itemModified, out aiPatch,
                                forceCompare || gExcp.Start.SafeDateTimeOffset().Date != newAiExcp.Start.SafeDateTimeOffset().Date);
                            if (itemModified > 0) {
                                try {
                                    Calendar.Instance.UpdateCalendarEntry_save(ref aiPatch);
                                    updatesMade++;
                                } catch (Microsoft.Kiota.Abstractions.ApiException ex) {
                                    if (ex.ResponseStatusCode == 400 && O365Errors.GetODataError(ex)?.Error?.Code == "ErrorOccurrenceCrossingBoundary") {
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

            return updatesMade;
        }
        #endregion
    }
}
