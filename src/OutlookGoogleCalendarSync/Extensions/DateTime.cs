using Google.Apis.Calendar.v3.Data;
using System;

namespace OutlookGoogleCalendarSync {
    public static class DateTimeExtensions {
        /// <summary>
        /// Returns the DateTime with time and GMT offset.
        /// This used to be the string format Google held date-times, eg "2012-08-20T00:00:00+02:00"
        /// </summary>
        /// <param name="dt">Date-time valule</param>
        /// <returns>Formatted string</returns>
        public static String ToPreciseString(this DateTime dt) {
            return dt.ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ssZ", new System.Globalization.CultureInfo("en-US"));
        }

        /// <summary>
        /// Returns the non-null Date or DateTime properties as a DateTime
        /// </summary>
        /// <returns>DateTime</returns>
        public static DateTime SafeDateTime(this EventDateTime evDt) {
            return evDt.DateTime ?? DateTime.Parse(evDt.Date);
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
                return (ev.Start?.DateTime?.TimeOfDay == new TimeSpan(0, 0, 0) && ev.Start?.DateTime?.TimeOfDay == ev.End?.DateTime?.TimeOfDay);
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
    }
}
