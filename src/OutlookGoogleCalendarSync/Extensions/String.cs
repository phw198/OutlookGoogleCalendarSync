using System;

namespace OutlookGoogleCalendarSync.Extensions {
    public static class OgcsString {
        public static String Append(this String input, String append) {
            return (String.IsNullOrEmpty(input) ? input : input + append);
        }
        
        public static String Prepend(this String input, String prepend) {
            return (String.IsNullOrEmpty(input) ? input : prepend + input);
        }

        public static String RemoveLineBreaks(this String input) {
            return input?.Replace("\r", "").Replace("\n", "");
        }
    }
}
