using System;
using System.Globalization;
using System.Text.RegularExpressions;
using log4net;

namespace OutlookGoogleCalendarSync {
    public class EmailAddress {
        private static readonly ILog log = LogManager.GetLogger(typeof(EmailAddress));

        public static String BuildFakeEmailAddress(String recipientName) {
            String buildFakeEmail = Regex.Replace(recipientName, @"[^\w\.-]", "");
            buildFakeEmail += "@unknownemail.com";
            log.Debug("Built a fake email for them: " + buildFakeEmail);
            return buildFakeEmail;
        }

        private static Boolean invalidEmail = false;
        public static Boolean IsValidEmail(string strIn) {
            invalidEmail = false;
            if (String.IsNullOrEmpty(strIn)) {
                MainForm.Instance.Logboxout("ERROR: Recipient has no email address.", notifyBubble: true);
                MainForm.Instance.Logboxout("This must be manually resolved in order to sync attendees for this event.");
                return false;
            }

            // Use IdnMapping class to convert Unicode domain names.
            strIn = Regex.Replace(strIn, @"(@)(.+)$", domainMapper);
            if (!invalidEmail) {
                // Return true if strIn is in valid e-mail format. 
                invalidEmail = !Regex.IsMatch(strIn,
                    @"^(?("")(""[^""]+?""@)|(([0-9a-z]((\.(?!\.))|[-!#\$%&'\*\+/=\?\^`\{\}\|~\w])*)(?<=[0-9a-z])@))" +
                    @"(?(\[)(\[(\d{1,3}\.){3}\d{1,3}\])|(([0-9a-z][-\w]*[0-9a-z]*\.)+[a-z0-9]{2,17}))$",
                    RegexOptions.IgnoreCase);
            }
            if (invalidEmail) {
                MainForm.Instance.Logboxout("ERROR: Recipient with email address \"" + strIn + "\" is invalid.", notifyBubble: true);
                MainForm.Instance.Logboxout("This must be manually resolved in order to sync attendees for this event.");
                return false;
            }
            return true;
        }

        private static string domainMapper(Match match) {
            // IdnMapping class with default property values.
            IdnMapping idn = new IdnMapping();

            string domainName = match.Groups[2].Value;
            try {
                domainName = idn.GetAscii(domainName);
            } catch (ArgumentException) {
                invalidEmail = true;
            }
            return match.Groups[1].Value + domainName;
        }
    }
}
