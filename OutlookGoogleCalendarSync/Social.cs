using System;
using System.Net;
using log4net;

namespace OutlookGoogleCalendarSync {
    class Social {
        private static readonly ILog log = LogManager.GetLogger(typeof(Social));

        public static void Donate() {
            System.Diagnostics.Process.Start("https://www.paypal.com/cgi-bin/webscr?cmd=_s-xclick&hosted_button_id=RT46CXQDSSYWJ&item_name=Outlook Google Calendar Sync from " + Settings.Instance.GaccountEmail);
        }

        #region Social
        public static void Google_goToCommunity() {
            System.Diagnostics.Process.Start("https://plus.google.com/communities/114412828247015553563");
        }
        public static void Google_share() {
            System.Diagnostics.Process.Start("https://plus.google.com/share?&url=http://bit.ly/OGcalsync");
        }

        public static void Twitter_tweet() {
            string text = "I'm using this Outlook-Google calendar sync tool - completely #free and feature loaded. #recommend";
            System.Diagnostics.Process.Start("http://twitter.com/intent/tweet?&url=http://bit.ly/OGcalsync&text=" + urlEncode(text) + "&via=ogcalsync");
        }
        public static void Twitter_follow() {
            System.Diagnostics.Process.Start("https://twitter.com/OGcalsync");
        }

        public static void Facebook_share() {
            System.Diagnostics.Process.Start("http://www.facebook.com/sharer/sharer.php?u=http://bit.ly/OGcalsync");
        }
        public static void RSS_follow() {
            System.Diagnostics.Process.Start("https://outlookgooglecalendarsync.codeplex.com/project/feeds/rss?ProjectRSSFeed=codeplex%3a%2f%2frelease%2foutlookgooglecalendarsync");
        }
        public static void Linkedin_share() {
            string text = "I'm using this Outlook-Google calendar sync tool - completely #free and feature loaded. #recommend";
            System.Diagnostics.Process.Start("http://www.linkedin.com/shareArticle?mini=true&url=http://bit.ly/OGcalsync&summary=" + urlEncode(text));
        }

        private static String urlEncode(String text) {
            return text.Replace("#", "%23");
        }
        #endregion

        #region Analytics
        public static void TrackVersion() {
            if (System.Diagnostics.Debugger.IsAttached) return;

            string analytics = null;
            switch (OutlookFactory.OutlookVersion) {
                case 11: analytics = "http://goo.gl/LMf6HT"; break; //2003
                case 12: analytics = "http://goo.gl/Xpqzua"; break; //2007
                case 14: analytics = "http://goo.gl/VM9Yaz"; break; //2010
                case 15: analytics = "http://goo.gl/LvIiQd"; break; //2013
                case 16: analytics = "http://goo.gl/Jhyzo5"; break; //2016
            }
            if (analytics != null) {
                log.Debug("Retrieving URL: " + analytics);
                WebClient wc = new WebClient();
                wc.Headers.Add("user-agent", "Mozilla/5.0 (Windows NT 6.1; WOW64; rv:37.0) Gecko/20100101 Firefox/37.0");
                wc.DownloadStringCompleted += new DownloadStringCompletedEventHandler(trackVersion_completed);
                wc.DownloadStringAsync(new Uri(analytics));
            }
        }

        private static void trackVersion_completed(object sender, DownloadStringCompletedEventArgs e) {
            if (e.Error != null)
                log.Error("Failed to access URL: " + e.Error.Message);
        }

        public static void TrackSync() {
            //Use an API that isn't used anywhere else - can use to see how many syncs are happening
            if (System.Diagnostics.Debugger.IsAttached) return;
            GoogleCalendar.Instance.GetSetting("locale");
        }
        #endregion
    }
}
