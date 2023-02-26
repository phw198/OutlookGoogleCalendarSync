using log4net;
using System;
using System.Collections.Generic;
using System.Management;
using System.Net;
using System.Text.RegularExpressions;

namespace OutlookGoogleCalendarSync {
    class Telemetry {
        private static readonly ILog log = LogManager.GetLogger(typeof(Telemetry));
        
        private static Telemetry instance;
        public static Telemetry Instance {
            get {
                return instance ??= new Telemetry();
            }
        }

        /// <summary>MD5 hash to identify distinct, anonymous user</summary>
        private String uuId;
        public String AnonymousUniqueUserId {
            get { return uuId; }
        }
        
        /// <summary>
        /// MD5 hash of either Gmail account, or custom thumbprint: ComputerName;Processor;C-driveSerial
        /// </summary>
        /// <returns>An MD5 hash</returns>
        public String UpdateAnonymousUniqueUserId() {
            try {
                if (Settings.AreLoaded && !string.IsNullOrEmpty(Settings.Instance.GaccountEmail)) {
                    log.Debug("Settings have been loaded, which contains Gmail account.");
                    uuId = GoogleOgcs.Authenticator.GetMd5(Settings.Instance.GaccountEmail, true);

                } else {
                    log.Debug("Settings not loaded; checking if the raw settings file has Gmail account set.");
                    String gmailAccount = null;
                    try {
                        gmailAccount = XMLManager.ImportElement("GaccountEmail", Settings.ConfigFile, false);
                    } catch { }

                    if (!string.IsNullOrEmpty(gmailAccount)) {
                        log.Fine("Gmail account found in settings files.");
                        uuId = GoogleOgcs.Authenticator.GetMd5(gmailAccount, true);
                    } else {
                        log.Warn("No Gmail account found, building custom thumbprint instead.");
                        String customThumbprint = "";
                        //Make a "unique" string based on:
                        //ComputerName;Processor;C-driveSerial
                        ManagementClass mc = new ManagementClass("win32_processor");
                        ManagementObjectCollection moc = mc.GetInstances();
                        foreach (ManagementObject mo in moc) {
                            customThumbprint = mo.Properties["SystemName"].Value.ToString();
                            customThumbprint += ";" + mo.Properties["Name"].Value.ToString();
                            break;
                        }
                        String drive = "C";
                        ManagementObject dsk = new ManagementObject(@"win32_logicaldisk.deviceid=""" + drive + @":""");
                        dsk.Get();
                        String volumeSerial = dsk["VolumeSerialNumber"].ToString();
                        customThumbprint += ";" + volumeSerial;

                        uuId = GoogleOgcs.Authenticator.GetMd5(customThumbprint);
                    }
                }

            } catch {
                log.Error("Unable to build accurate anonymous unique ID. Resorting to a random number.");
                Random random = new Random();
                uuId = random.Next().ToString();
            }
            return uuId;
        }

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
                String cid = Telemetry.Instance.AnonymousUniqueUserId;
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
            public Dictionary<String, Dictionary<String,String>> user_properties { get; }
            public List<Event> events { get; }

            public GA4Event(Event.Name eventName, Event throwAway = null) : this(eventName, out throwAway) {
            }

            public GA4Event(Event.Name eventName, out Event theEvent) {
                client_id = Telemetry.Instance.AnonymousUniqueUserId; //Extend this in case more than one instance of OGCS running?
                user_id = Telemetry.Instance.AnonymousUniqueUserId;
                non_personalized_ads = true;
                user_properties = new Dictionary<String, Dictionary<String, String>>();
                user_properties.Add("ogcsVersion", new Dictionary<String, String> { { "value", System.Windows.Forms.Application.ProductVersion } });
                user_properties.Add("isBenefactor", new Dictionary<String, String> { { "value", Settings.Instance.UserIsBenefactor().ToString() } });
                theEvent = new Event(eventName);
                events = new List<Event> { theEvent };
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
                    jsonPayload = jsonPayload.Replace("\"parameters\":", "\"params\":");

                    log.Debug("GA4: " + jsonPayload);
                    wc.UploadStringAsync(new Uri(baseAnalyticsUrl), "POST", jsonPayload);

                } catch (System.Exception ex) {
                    OGCSexception.Analyse(ex);
                }
            }

            public class Event {
                public String name { get; private set; }
                public Dictionary<String, Object> parameters { get; private set; }

                public enum Name {
                    application_started,
                    donate
                }

                public Event(Name eventName) {
                    name = eventName.ToString();
                }

                public void AddParameter(String parameterName, object parameterValue) {
                    if (parameters == null) 
                        parameters = new Dictionary<String, Object>();
                    if (parameterValue is int)
                        parameters.Add(parameterName, (int)parameterValue);
                    else
                        parameters.Add(parameterName, parameterValue.ToString());
                }
            }
        }

        private static void sendTelemetry_completed(object sender, UploadStringCompletedEventArgs e) {
            if (e.Error != null) {
                log.Warn("Failed to access URL " + e.UserState?.ToString());
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
