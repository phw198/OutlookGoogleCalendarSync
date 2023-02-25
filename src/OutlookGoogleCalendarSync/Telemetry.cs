using log4net;
using System;
using System.Collections.Generic;
using System.Net;
//using System.Text.RegularExpressions;

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

        /// <summary>
        /// Deprecated Universal Analytics (dies in Jul 2023)
        /// </summary>
        public static void Send(Analytics.Category category, Analytics.Action action, String label) {
            try {
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
                wc.UploadStringCompleted += new UploadStringCompletedEventHandler(sendTelemetry_completed);
                wc.UploadStringAsync(new Uri(analyticsUrl), "");

            } catch (System.Exception ex) {
                OGCSexception.Analyse(ex);
            }
        }

        public class GA4Event {
            public String client_id { get; }
            public String user_id { get; }
            public Boolean non_personalized_ads { get; }
            public List<Event> events { get; }

            public enum Name {
                application_started
            }
            
            public GA4Event(Name eventName) {
                client_id = System.Windows.Forms.Application.ProductVersion;
                user_id = GoogleOgcs.Authenticator.HashedGmailAccount ?? null;
                non_personalized_ads = true;
                events = new List<Event>();
                events.Add(new Event(eventName));
            }

            public void Send() {
                if (Settings.Instance.TelemetryDisabled || Program.InDeveloperMode) {
                    log.Debug("Telemetry is disabled.");
                    return;
                }

                try {
                    String baseAnalyticsUrl = "https://www.google-analytics.com/mp/collect?api_secret=kWOsAm2tQny1xOjiwMyC5Q&measurement_id=G-S6RMS8GHEE";

                    Extensions.OgcsWebClient wc = new Extensions.OgcsWebClient();
                    wc.Headers[HttpRequestHeader.ContentType] = "application/json";
                    wc.UploadStringCompleted += new UploadStringCompletedEventHandler(sendTelemetry_completed);

                    GA4Event payload = this;
                    String jsonPayload = Newtonsoft.Json.JsonConvert.SerializeObject(payload);

                    log.Debug("GA4: " + jsonPayload);
                    wc.UploadStringAsync(new Uri(baseAnalyticsUrl), "POST", jsonPayload);

                } catch (System.Exception ex) {
                    OGCSexception.Analyse(ex);
                }
            }

            public class Event {
                public String name;

                public Event(Name eventName) {
                    name = eventName.ToString();
                }
            }
        }

        private static void sendTelemetry_completed(object sender, UploadStringCompletedEventArgs e) {
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
