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
    }
}
