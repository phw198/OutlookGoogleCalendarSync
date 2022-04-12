using System;
using System.Net;
using log4net;

namespace OutlookGoogleCalendarSync {
    class Telemetry {
        private static readonly ILog log = LogManager.GetLogger(typeof(Telemetry));

        public static void TrackVersions() {
            if (Program.InDeveloperMode) return;

            //OUTLOOK CLIENT
            Send(Analytics.Category.outlook, Analytics.Action.version, OutlookOgcs.Factory.OutlookVersionNameFull.Replace("Outlook", ""));

            //OGCS APPLICATION
            Send(Analytics.Category.ogcs, Analytics.Action.version, System.Windows.Forms.Application.ProductVersion);
        }

        public static void TrackSync() {
            if (Program.InDeveloperMode) return;
            Send(Analytics.Category.ogcs, Analytics.Action.sync, "calendar");
        }

        public static void Send(Analytics.Category category, Analytics.Action action, String label) {
            String cid = GoogleOgcs.Authenticator.HashedGmailAccount ?? "1";
            String baseAnalyticsUrl = "https://www.google-analytics.com/collect?v=1&t=event&tid=UA-19426033-4&aip=1&cid=" + cid;

            if (action == Analytics.Action.debug) {
                label = "v" + System.Windows.Forms.Application.ProductVersion + ";" + label;
            }
            String analyticsUrl = baseAnalyticsUrl + "&ec=" + category.ToString() + "&ea=" + action.ToString() + "&el=" + System.Net.WebUtility.UrlEncode(label);
            log.Debug("Retrieving URL: " + analyticsUrl);
            
            if (Settings.Instance.TelemetryDisabled || Program.InDeveloperMode) {
                log.Debug("Telemetry is disabled.");
                return;
            }

            Extensions.OgcsWebClient wc = new Extensions.OgcsWebClient();
            wc.DownloadStringCompleted += new DownloadStringCompletedEventHandler(sendTelemetry_completed);
            wc.DownloadStringAsync(new Uri(analyticsUrl), analyticsUrl);
        }

        private static void sendTelemetry_completed(object sender, DownloadStringCompletedEventArgs e) {
            if (e.Error != null) {
                log.Warn("Failed to access URL " + e.UserState.ToString());
                log.Fail(e.Error.Message);
                if (e.Error.InnerException != null) log.Fail(e.Error.InnerException.Message);
                if (e.Error is WebException) {
                    WebException we = e.Error as WebException;
                    if (we.Response != null) {
                        log.Debug("Reading response.");
                        System.IO.Stream stream = we.Response.GetResponseStream();
                        System.IO.StreamReader sr = new System.IO.StreamReader(stream);
                        log.Fail(sr.ReadToEnd());
                    }
                }
            }
        }
    }

    public class Analytics {
        private static readonly ILog log = LogManager.GetLogger(typeof(Analytics));

        public enum Category {
            ogcs,
            outlook,
            squirrel
        }
        public enum Action {
            debug,      //ogcs
            donate,     //ogcs
            download,   //squirrel
            install,    //squirrel
            setting,    //ogcs
            sync,       //ogcs
            uninstall,  //squirrel
            upgrade,    //squirrel
            version     //outlook,ogcs
        }        
    }
}
