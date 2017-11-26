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
            System.Diagnostics.Process.Start("https://plus.google.com/share?&url=http://bit.ly/ogcalsync");
        }

        public static void Twitter_tweet() {
            string text = "I'm using this Outlook-Google calendar sync tool - completely #free and feature loaded. #recommend";
            System.Diagnostics.Process.Start("http://twitter.com/intent/tweet?&url=http://bit.ly/ogcalsync&text=" + urlEncode(text) + "&via=ogcalsync");
        }
        public static void Twitter_follow() {
            System.Diagnostics.Process.Start("https://twitter.com/OGcalsync");
        }

        public static void Facebook_share() {
            System.Diagnostics.Process.Start("http://www.facebook.com/sharer/sharer.php?u=http://bit.ly/ogcalsync");
        }
        public static void RSS_follow() {
            System.Diagnostics.Process.Start("https://github.com/phw198/outlookgooglecalendarsync/releases.atom");
        }
        public static void Linkedin_share() {
            string text = "I'm using this Outlook-Google calendar sync tool - completely #free and feature loaded. #recommend";
            System.Diagnostics.Process.Start("http://www.linkedin.com/shareArticle?mini=true&url=http://bit.ly/ogcalsync&summary=" + urlEncode(text));
        }

        private static String urlEncode(String text) {
            return text.Replace("#", "%23");
        }
        #endregion

        #region Analytics
        public static void TrackVersions() {
            if (System.Diagnostics.Debugger.IsAttached) return;

            String cid = GoogleOgcs.Authenticator.HashedGmailAccount ?? "1";
            
            //OUTLOOK CLIENT
            String baseAnalyticsUrl = "https://www.google-analytics.com/collect?v=1&t=event&tid=UA-19426033-4&cid="+ cid +"&ea=version";
            String analyticsUrl = baseAnalyticsUrl + "&ec=outlook&el=";
            try {
                switch (OutlookOgcs.Factory.OutlookVersion) {
                    case 11: analyticsUrl += "2003"; break;
                    case 12: analyticsUrl += "2007"; break;
                    case 14: analyticsUrl += "2010"; break;
                    case 15: analyticsUrl += "2013"; break;
                    case 16: analyticsUrl += "2016"; break;
                    case 17: analyticsUrl += "2019"; break;
                    default: analyticsUrl += "Unknown-" + OutlookOgcs.Factory.OutlookVersion; break;
                }
            } catch (System.Exception ex) {
                log.Error("Failed setting Outlook client analytics URL.");
                OGCSexception.Analyse(ex);
                analyticsUrl = "https://phw198.github.io/OutlookGoogleCalendarSync/track/ogcs?version=Unknown";
            }
            sendVersion(analyticsUrl);

            //OGCS APPLICATION
            analyticsUrl = baseAnalyticsUrl + "&ec=ogcs&el=" + System.Windows.Forms.Application.ProductVersion;
            sendVersion(analyticsUrl);
        }

        private static void sendVersion(String analyticsUrl) {
            if (analyticsUrl != null) {
                log.Debug("Retrieving URL: " + analyticsUrl);
                WebClient wc = new WebClient();
                wc.Headers.Add("user-agent", "Mozilla/5.0 (Windows NT 6.1; WOW64; rv:37.0) Gecko/20100101 Firefox/37.0");
                wc.DownloadStringCompleted += new DownloadStringCompletedEventHandler(trackVersion_completed);
                wc.DownloadStringAsync(new Uri(analyticsUrl), analyticsUrl);
            }
        }

        private static void trackVersion_completed(object sender, DownloadStringCompletedEventArgs e) {
            if (e.Error != null) {
                log.Warn("Failed to access URL " + e.UserState.ToString());
                log.Error(e.Error.Message);
            }
        }

        public static void TrackSync() {
            //Use an API that isn't used anywhere else - can use to see how many syncs are happening
            if (System.Diagnostics.Debugger.IsAttached) return;
            GoogleOgcs.Calendar.Instance.GetSetting("locale");
        }
        #endregion
    }
}
