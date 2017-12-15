using System;
using System.Globalization;
using System.Text.RegularExpressions;
using log4net;

namespace OutlookGoogleCalendarSync {
    public class EmailAddress {
        private static readonly ILog log = LogManager.GetLogger(typeof(EmailAddress));

        public static String BuildFakeEmailAddress(String recipientName, out Boolean builtFakeEmail) {
            String buildFakeEmail = Regex.Replace(recipientName, @"[^\w\.-]", "");
            buildFakeEmail += "@unknownemail.com";
            log.Debug("Built a fake email for them: " + buildFakeEmail);
            builtFakeEmail = true;
            return buildFakeEmail;
        }

        //Sourced from https://msdn.microsoft.com/en-us/library/01escwtf(v=vs.110).aspx
        //Underscores added to regex as allowable
        private static Boolean invalidEmail = false;
        public static Boolean IsValidEmail(string strIn) {
            invalidEmail = false;
            if (String.IsNullOrEmpty(strIn)) {
                return false;
            }
            if (strIn.StartsWith("'") && strIn.EndsWith("'")) {
                strIn = strIn.TrimStart('\'').TrimEnd('\'').Trim();
            }

            // Use IdnMapping class to convert Unicode domain names.
            strIn = Regex.Replace(strIn, @"(@)(.+)$", domainMapper);
            if (!invalidEmail) {
                // Return true if strIn is in valid e-mail format. 
                invalidEmail = !Regex.IsMatch(strIn,
                    @"^(?("")(""[^""]+?""@)|(([0-9a-z_']((\.(?!\.))|[-_!#\$%&'\*\+/=\?\^`\{\}\|~\w])*)(?<=[0-9a-z_'])@))" +
                    @"(?(\[)(\[(\d{1,3}\.){3}\d{1,3}\])|(([0-9a-z][-\w]*[0-9a-z]*\.)+[a-z0-9]{2,17}))$",
                    RegexOptions.IgnoreCase);
            }
            if (invalidEmail) {
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

        public static String MaskAddress(String emailAddress) {
            try {
                int at = emailAddress.IndexOf('@');
                String masked = emailAddress.Substring(0, 2) + "".PadRight(at - 3, '*') + emailAddress.Substring(at - 1);
                return masked;
            } catch (System.Exception) {
                return "*****@masked.com";
            }
        }
    }
}
