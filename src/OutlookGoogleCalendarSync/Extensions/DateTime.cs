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

        public static Boolean AllDayEvent(this Event ev) {
            return ev.Start.Date != null;
        }
    }
}
