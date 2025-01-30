using Google.Apis.Calendar.v3.Data;
using log4net;
using Microsoft.Office.Interop.Outlook;
using OutlookGoogleCalendarSync.Extensions;
using System;
using System.Collections.Generic;
using System.Linq;
using Ogcs = OutlookGoogleCalendarSync;

namespace OutlookGoogleCalendarSync.Outlook {
    public class Recurrence {
        private static readonly ILog log = LogManager.GetLogger(typeof(Recurrence));

        public static void BuildOutlookPattern(Event ev, AppointmentItem ai) {
            RecurrencePattern ignore;
            buildOutlookPattern(ev, ai, out ignore);
            ignore = (RecurrencePattern)Outlook.Calendar.ReleaseObject(ignore);
        }

        private static void buildOutlookPattern(Event ev, AppointmentItem ai, out RecurrencePattern oPattern) {
            if (ev.Recurrence == null) { oPattern = null; return; }

            Dictionary<String, String> ruleBook = Google.Recurrence.ExplodeRrule(ev.Recurrence);
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
                                ev.Start.SafeDateTime().Day > 28) {
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
            ai = Outlook.Calendar.Instance.IOutlook.WindowsTimeZone_set(ai, ev);
            oPattern.PatternStartDate = ev.Start.SafeDateTime();
            if (ruleBook.ContainsKey("INTERVAL") && Convert.ToInt16(ruleBook["INTERVAL"]) > 1 && ruleBook["FREQ"] != "YEARLY")
                oPattern.Interval = Convert.ToInt16(ruleBook["INTERVAL"]);
            if (ruleBook.ContainsKey("COUNT"))
                oPattern.Occurrences = Convert.ToInt16(ruleBook["COUNT"]);
            if (ruleBook.ContainsKey("UNTIL")) {
                if (ruleBook["UNTIL"].StartsWith("4500")) {
                    log.Warn("Outlook can't handle end dates this far in the future. Converting to no end date.");
                    oPattern.NoEndDate = true;
                } else {
                    System.DateTime endDate;
                    if (ruleBook["UNTIL"].Length == 8 && !ruleBook["UNTIL"].EndsWith("Z"))
                        endDate = System.DateTime.ParseExact(ruleBook["UNTIL"], "yyyyMMdd", System.Globalization.CultureInfo.InvariantCulture).Date;
                    else {
                        endDate = System.DateTime.ParseExact(ruleBook["UNTIL"], "yyyyMMddTHHmmssZ", System.Globalization.CultureInfo.InvariantCulture, System.Globalization.DateTimeStyles.AdjustToUniversal);
                        endDate = endDate.AddHours(TimezoneDB.GetUtcOffset(ev.End.TimeZone)).Date;
                    }
                    if (endDate < oPattern.PatternStartDate) {
                        log.Debug("PatternStartDate: " + oPattern.PatternStartDate.ToString("yyyyMMddHHmmss"));
                        log.Debug("PatternEndDate:   " + ruleBook["UNTIL"].ToString());
                        String summary = Ogcs.Google.Calendar.GetEventSummary("The recurring Google event has an end date <i>before</i> the start date, which Outlook doesn't allow.<br/>" +
                            "The synced Outlook recurrence has been changed to a single occurrence.", ev, out String anonSummary, onlyIfNotVerbose: true);
                        Forms.Main.Instance.Console.Update(summary, anonSummary, Console.Markup.warning);
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

        public static void CompareOutlookPattern(Event ev, ref RecurrencePattern aiOpattern, Sync.Direction syncDirection, System.Text.StringBuilder sb, ref int itemModified) {
            if (ev.Recurrence == null) return;

            log.Fine("Building a temporary recurrent Appointment generated from Event");
            AppointmentItem evAI = Ogcs.Outlook.Calendar.Instance.IOutlook.GetFolderByID(Sync.Engine.Calendar.Instance.Profile.UseOutlookCalendar.Id).Items.Add() as AppointmentItem;
            evAI.Start = ev.Start.SafeDateTime();

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
                evOpattern = (RecurrencePattern)Outlook.Calendar.ReleaseObject(evOpattern);
                evAI.Delete();
                evAI = (AppointmentItem)Outlook.Calendar.ReleaseObject(evAI);
            }
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

        #region Exceptions
        public enum DeletionState {
            Inaccessible,
            Deleted,
            NotDeleted
        }
        public static Boolean HasExceptions(AppointmentItem ai) {
            RecurrencePattern rp = null;
            Exceptions excps = null;
            try {
                rp = ai.GetRecurrencePattern();
                excps = rp.Exceptions;
                return excps.Count != 0;
            } finally {
                excps = (Exceptions)Outlook.Calendar.ReleaseObject(excps);
                rp = (RecurrencePattern)Outlook.Calendar.ReleaseObject(rp);
            }
        }

        public static DeletionState ExceptionIsDeleted(Microsoft.Office.Interop.Outlook.Exception oExcp) {
            if (oExcp.Deleted) return DeletionState.Deleted;
            AppointmentItem ai = null;
            try {
                ai = oExcp.AppointmentItem;
                return DeletionState.NotDeleted;
            } catch (System.Exception ex) {
                Ogcs.Exception.LogAsFail(ref ex);
                String originalDate = oExcp.OriginalDate.ToString("dd/MM/yyyy");
                if (ex.Message == "You changed one of the recurrences of this item, and this instance no longer exists. Close any open items and try again.") {
                    ex.Analyse("This Outlook recurrence instance on " + originalDate + " has become inaccessible, probably due to caching");
                } else {
                    ex.Analyse("Error when determining if Outlook recurrence on " + originalDate + " is deleted or not.");
                }
                return DeletionState.Inaccessible;
            } finally {
                ai = (AppointmentItem)Outlook.Calendar.ReleaseObject(ai);
            }
        }

        public static void CreateOutlookExceptions(Event ev, ref AppointmentItem createdAi) {
            processOutlookExceptions(ev, ref createdAi, forceCompare: true);
        }
        public static void UpdateOutlookExceptions(Event ev, ref AppointmentItem ai, Boolean forceCompare) {
            processOutlookExceptions(ev, ref ai, forceCompare);
        }

        private static void processOutlookExceptions(Event ev, ref AppointmentItem ai, Boolean forceCompare) {
            if (!Google.Recurrence.HasExceptions(ev, checkLocalCacheOnly: true)) return;

            List<Event> gExcps = Google.Recurrence.GoogleExceptions.Where(exp => exp.RecurringEventId == ev.Id).ToList();
            if (gExcps.Count == 0) return;

            log.Debug($"{gExcps.Count} Google recurrence exceptions within sync range to be compared.");

            //Process deleted exceptions first
            List<Event> gCancelledExcps = gExcps.Where(exp => exp.Status == "cancelled").ToList();
            log.Fine($"{gCancelledExcps.Count} Google cancelled occurrences.");
            processOutlookExceptions(gCancelledExcps, ref ai, forceCompare, true);

            //Then process everything else
            gExcps = gExcps.Except(gCancelledExcps).ToList();
            log.Fine($"{gExcps.Count} Google modified exceptions.");
            processOutlookExceptions(gExcps, ref ai, forceCompare, false);
        }

        private static void processOutlookExceptions(List<Event> evExceptions, ref AppointmentItem ai, Boolean forceCompare, Boolean processingDeletions) {
            if (evExceptions.Count == 0) return;

            RecurrencePattern oPattern = null;
            try {
                oPattern = ai.GetRecurrencePattern();

                foreach (Event gExcp in evExceptions) {
                    System.DateTime gExcpOrigDate = gExcp.OriginalStartTime.SafeDateTime();
                    System.DateTime? gExcpCurrDate = gExcp.Start?.SafeDateTime();
                    String gExcpDetails = "Google exception with original date " + gExcpOrigDate.ToString() + (gExcpCurrDate != null ? " now on " + gExcpCurrDate?.ToShortDateString() : "");
                    log.Fine("Found " + gExcpDetails);

                    AppointmentItem newAiExcp = null;
                    try {
                        getOutlookInstance(oPattern, gExcpOrigDate, ref newAiExcp, processingDeletions);
                        if (newAiExcp == null) {
                            if (gExcp.Status != "cancelled") {
                                log.Warn("Unable to find Outlook exception for " + gExcpDetails);
                                log.Warn("Google is NOT deleted though - a mismatch has occurred somehow!");
                                String syncDirectionTip = (Sync.Engine.Calendar.Instance.Profile.SyncDirection.Id == Sync.Direction.Bidirectional.Id) ? "<br/><i>Ensure you <b>first</b> set OGCS to one-way sync G->O.</i>" : "";
                                Forms.Main.Instance.Console.Update(Ogcs.Google.Calendar.GetEventSummary(
                                        "<br/>This occurrence cannot be found in Outlook.<br/>" +
                                        "This can happen if, for example, the occurrence has been rearranged to different days more than once.<br/>" +
                                        "<u>Suggested fix</u>: delete the entire series in Google and let OGCS recreate it." + syncDirectionTip, gExcp, out String anonSummary),
                                    anonSummary, Console.Markup.warning);
                            }
                            continue;
                        }

                        if (gExcp.Status == "cancelled") {
                            Forms.Main.Instance.Console.Update(Outlook.Calendar.GetEventSummary("<br/>Occurrence deleted.", newAiExcp, out String anonSummary), anonSummary, Console.Markup.calendar, verbose: true);
                            newAiExcp.Delete();

                        } else if (Sync.Engine.Calendar.Instance.Profile.ExcludeDeclinedInvites && gExcp.Attendees != null && gExcp.Attendees.Count(a => a.Self == true && a.ResponseStatus == "declined") == 1) {
                            Forms.Main.Instance.Console.Update(Outlook.Calendar.GetEventSummary("<br/>Occurrence declined.", newAiExcp, out String anonSummary), anonSummary, Console.Markup.calendar, verbose: true);
                            newAiExcp.Delete();

                        } else {
                            int itemModified = 0;
                            Outlook.Calendar.Instance.UpdateCalendarEntry(ref newAiExcp, gExcp, ref itemModified,
                                forceCompare || gExcp.Start.SafeDateTime().Date != newAiExcp.Start.Date);
                            if (itemModified > 0) {
                                try {
                                    newAiExcp.Save();
                                } catch (System.Exception ex) {
                                    Ogcs.Exception.Analyse(ex);
                                    if (ex.Message == "Cannot save this item.") {
                                        Forms.Main.Instance.Console.Update(
                                            Outlook.Calendar.GetEventSummary("Uh oh! Outlook wasn't able to save this recurrence exception! " +
                                                "You may have two occurences on the same day, which it doesn't allow.", newAiExcp, out String anonSummary, true)
                                            , anonSummary, Console.Markup.warning);
                                    }
                                }
                            }
                        }
                    } finally {
                        newAiExcp = (AppointmentItem)Outlook.Calendar.ReleaseObject(newAiExcp);
                    }
                }
            } finally {
                oPattern = (RecurrencePattern)Outlook.Calendar.ReleaseObject(oPattern);
            }
        }

        private static void getOutlookInstance(RecurrencePattern oPattern, System.DateTime instanceOrigDate, ref AppointmentItem ai, Boolean processingDeletions) {
            //The Outlook API is rubbish: oPattern.GetOccurrence(instanceDate) returns anything currently on that date NOW, regardless of if it was moved there.
            //Even worse, if 2-Feb was deleted then 1-Feb occurrence is moved to 2-Feb, it will return 2-Feb but there is no OriginalStartDate property to know it was moved.

            //So first we'll check all exceptions by OriginalStartDate, then if not found use oPattern.GetOccurrence(instanceDate)
            Exceptions oExcps = null;
            try {
                oExcps = oPattern.Exceptions;
                for (int e = 1; e <= oExcps.Count; e++) {
                    Microsoft.Office.Interop.Outlook.Exception oExcp = null;
                    try {
                        oExcp = oExcps[e];
                        DeletionState isDeleted = ExceptionIsDeleted(oExcp);

                        if (oExcp.OriginalDate.Date == instanceOrigDate.Date) {
                            log.Debug("Found Outlook exception for original date " + instanceOrigDate);

                            if (isDeleted == DeletionState.Inaccessible) {
                                log.Warn("This exception is inaccessible.");
                                return;
                            } else if (isDeleted == DeletionState.Deleted) {
                                if (processingDeletions) {
                                    log.Debug("This exception is deleted.");
                                    return;
                                }
                            }

                            try {
                                ai = oExcp.AppointmentItem;
                                return;
                            } catch (System.Exception ex) {
                                Forms.Main.Instance.Console.Update(ex.Message + "<br/>If this keeps happening, please restart OGCS.", Console.Markup.error);
                                break;
                            }
                        } else if (processingDeletions && isDeleted != DeletionState.Deleted && oExcp.AppointmentItem.Start.Date == instanceOrigDate.Date) {
                            log.Debug("An Outlook exception has moved to " + instanceOrigDate.Date.ToShortDateString() + " from " + oExcp.OriginalDate.Date.ToShortDateString() + ". This moved exception won't be deleted.");
                            return;
                        }
                    } finally {
                        oExcp = (Microsoft.Office.Interop.Outlook.Exception)Outlook.Calendar.ReleaseObject(oExcp);
                    }
                }
            } finally {
                oExcps = (Exceptions)Outlook.Calendar.ReleaseObject(oExcps);
            }

            //Finally check if the occurrence is not an exception, or an exception has moved to the same date as a deleted exception
            //The two things are stored the same way in Outlook's crazy world
            if (ai == null) {
                try {
                    ai = oPattern.GetOccurrence(instanceOrigDate);
                    return;
                } catch { }
            }
        }
        #endregion
    }
}
