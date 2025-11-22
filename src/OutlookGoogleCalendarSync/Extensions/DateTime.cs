using Google.Apis.Calendar.v3.Data;
using System;
using System.Globalization;

namespace OutlookGoogleCalendarSync.Extensions {
    public class OgcsDateTime {
        private System.DateTime baseDateTime;
        private Boolean dateOnly;
        
        /// <summary>
        /// Extends System.DateTime with date/time precision.
        /// Helps to differentiate a date vs a midnight time.
        /// </summary>
        /// <param name="baseDateTime">The System.DateTime</param>
        /// <param name="dateOnly">Whether the time element should be ignored</param>
        public OgcsDateTime(System.DateTime baseDateTime, Boolean dateOnly = false) {
            this.baseDateTime = baseDateTime;
            this.dateOnly = dateOnly;
        }

        public override string ToString() {
            if (this.dateOnly)
                return this.baseDateTime.ToShortDateString();
            else
                return this.baseDateTime.ToString("g");
        }

        public override bool Equals(Object obj) {
            if (obj is OgcsDateTime)
                return this.baseDateTime == (obj as OgcsDateTime).baseDateTime;
            else
                return false;
        }

        public override int GetHashCode() {
            return this.baseDateTime.GetHashCode() + this.dateOnly.GetHashCode();
        }
    }

    public static class DateTime {
        /// <summary>
        /// Returns the DateTime with UTC time.
        /// This used to be the string format Google held date-times, eg "2012-08-20T00:00:00+02:00"
        /// </summary>
        /// <param name="dt">Date-time value</param>
        /// <returns>Formatted string</returns>
        public static String ToPreciseString(this System.DateTime dt) {
            return dt.ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ssZ", new CultureInfo("en-US"));
        }

        /// <summary>
        /// Returns the DateTime with time and GMT offset.
        /// This used to be the string format Google held date-times, eg "2012-08-20T00:00:00+02:00"
        /// </summary>
        /// <param name="dt">Date-time offset value</param>
        /// <returns>Formatted string</returns>
        public static String ToPreciseString(this System.DateTimeOffset dt) {
            return dt.ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ssZ", new CultureInfo("en-US"));
        }

        /// <summary>
        /// Returns the non-null Date or DateTime properties as a DateTime
        /// </summary>
        /// <returns>DateTime</returns>
        public static System.DateTime SafeDateTime(this EventDateTime evDt) {
            return SafeDateTimeOffset(evDt).DateTime;
        }

        /// <summary>
        /// Returns the non-null Date or DateTime properties as a DateTimeOffset
        /// </summary>
        /// <returns>DateTimeOffset</returns>
        public static System.DateTimeOffset SafeDateTimeOffset(this EventDateTime evDt) {
            return evDt.DateTimeDateTimeOffset?.ToLocalTime() ?? System.DateTimeOffset.ParseExact(evDt.Date, "yyyy-MM-dd", CultureInfo.InvariantCulture, DateTimeStyles.AssumeUniversal);
        }

        /// <summary>
        /// Returns the DateTime for a Graph Date
        /// </summary>
        /// <returns>DateTime</returns>
        public static System.DateTime SafeDateTime(this Microsoft.Graph.Date graphDate) {
            return new System.DateTime(graphDate.Year, graphDate.Month, graphDate.Day);
        }

        /// <summary>
        /// Parses the DateTimeTimeZone string to a local DateTime
        /// </summary>
        /// <returns>Local DateTime</returns>
        public static System.DateTime SafeDateTime(this Microsoft.Graph.DateTimeTimeZone evDt) {
            System.DateTime safeDate;
            if (evDt.TimeZone == "UTC") {
                safeDate = System.DateTime.Parse(evDt.DateTime, null, DateTimeStyles.AssumeUniversal);                
            } else {
                Int16 offset = TimezoneDB.GetUtcOffset(evDt.TimeZone);
                safeDate = System.DateTime.Parse(evDt.DateTime).AddMinutes(-offset);
                safeDate = System.DateTime.SpecifyKind(safeDate, DateTimeKind.Utc);
                safeDate = safeDate.ToLocalTime();
            }
            return safeDate;
        }

        /// <summary>
        /// Converts a System.DateTime to a Graph.Date
        /// </summary>
        /// <returns>Graph.Date</returns>
        public static Microsoft.Graph.Date ToGraphDate(this System.DateTime dt) {
            return new Microsoft.Graph.Date(dt.Year, dt.Month, dt.Day);
        }

        /// <summary>
        /// Whether an Event is all day
        /// </summary>
        /// <param name="ev">The Event to check</param>
        /// <param name="logicallyEquivalent">Midnight to midnight Events treated as all day</param>
        /// <returns></returns>
        public static Boolean AllDayEvent(this Event ev, Boolean logicallyEquivalent = false) {
            if (ev.Start?.Date != null)
                return true;
            if (logicallyEquivalent)
                return (ev.Start?.DateTimeDateTimeOffset?.ToLocalTime().TimeOfDay == new TimeSpan(0, 0, 0) && 
                    ev.Start?.DateTimeDateTimeOffset?.ToLocalTime().TimeOfDay == ev.End?.DateTimeDateTimeOffset?.ToLocalTime().TimeOfDay);
            else
                return false;
        }

        /// <summary>
        /// Whether an Appointment is all day
        /// </summary>
        /// <param name="ai">The Appointment to check</param>
        /// <param name="logicallyEquivalent">Midnight to midnight Appointments treated as all day</param>
        /// <returns></returns>
        public static Boolean AllDayEvent(this Microsoft.Office.Interop.Outlook.AppointmentItem ai, Boolean logicallyEquivalent = false) {
            if (ai.AllDayEvent)
                return true;
            if (logicallyEquivalent)
                return (ai.Start.TimeOfDay == new TimeSpan(0, 0, 0) && ai.Start.TimeOfDay == ai.End.TimeOfDay);
            else
                return false;
        }

        /// <summary>
        /// Whether a Graph Event is all day
        /// </summary>
        /// <param name="ai">The Graph Event to check</param>
        /// <param name="logicallyEquivalent">Midnight to midnight Events treated as all day</param>
        /// <returns></returns>
        public static Boolean AllDayEvent(this Microsoft.Graph.Event ai, Boolean logicallyEquivalent = false) {
            if ((bool)ai.IsAllDay)
                return true;
            if (logicallyEquivalent)
                return (ai.Start.SafeDateTime().TimeOfDay == new TimeSpan(0, 0, 0) && ai.End.SafeDateTime().TimeOfDay == new TimeSpan(0, 0, 0));
            else
                return false;
        }
    }
}
