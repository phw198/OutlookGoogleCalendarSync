using System;

namespace OutlookGoogleCalendarSync {
    public static class StringExtensions {
        public static String RemoveLineBreaks(this String input) {
            return input?.Replace("\r", "").Replace("\n", "");
        }
    }
}
