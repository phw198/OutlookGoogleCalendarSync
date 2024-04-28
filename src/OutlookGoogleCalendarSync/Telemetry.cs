using Ogcs = OutlookGoogleCalendarSync;
using log4net;
using System;
using System.Collections.Generic;
using System.Management;
using System.Net;
using System.Threading.Tasks;

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
                    uuId = Ogcs.Google.Authenticator.GetMd5(Settings.Instance.GaccountEmail, true);

                } else {
                    log.Debug("Settings not loaded; checking if the raw settings file has Gmail account set.");
                    String gmailAccount = null;
                    try {
                        gmailAccount = XMLManager.ImportElement("GaccountEmail", Settings.ConfigFile, false);
                    } catch { }

                    if (!string.IsNullOrEmpty(gmailAccount)) {
                        log.Fine("Gmail account found in settings files.");
                        uuId = Ogcs.Google.Authenticator.GetMd5(gmailAccount, true);
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

                        uuId = Ogcs.Google.Authenticator.GetMd5(customThumbprint);
                    }
                }

            } catch {
                log.Error("Unable to build accurate anonymous unique ID. Resorting to a random number.");
                Random random = new Random();
                uuId = random.Next().ToString();
            }
            return uuId;
        }

        public String OutlookVersion { get; internal set; }
        public String OutlookVersionName { get; internal set; }

        public String Continent { get; private set; }
        public String Country { get; private set; }
        public String CountryCode { get; private set; }
        public String Region { get; private set; }
        public String City { get; private set; }

        public Telemetry() {
            getIpGeoData().ConfigureAwait(false);
        }

        private async Task getIpGeoData() {
            try {
                using (Extensions.OgcsWebClient wc = new()) {
                    //https://api.country.is/
                    String response = await wc.DownloadStringTaskAsync(new Uri("https://api.techniknews.net/ipgeo"));
                    Newtonsoft.Json.Linq.JObject ipGeoInfo = Newtonsoft.Json.Linq.JObject.Parse(response);
                    if (ipGeoInfo.HasValues && ipGeoInfo["status"]?.ToString() == "success") {
                        Continent = ipGeoInfo["continent"]?.ToString();
                        Country = ipGeoInfo["country"]?.ToString();
                        CountryCode = ipGeoInfo["countryCode"]?.ToString();
                        Region = ipGeoInfo["regionName"]?.ToString();
                        City = ipGeoInfo["city"]?.ToString();
                    } else {
                        log.Warn("Could not determine IP geolocation; status=" + ipGeoInfo["status"]);
                    }
                }
            } catch (System.Exception ex) {
                Ogcs.Exception.Analyse("Could not get IP geolocation.", Ogcs.Exception.LogAsFail(ex));
            }
        }

        /// <summary>
        /// This can just be removed once Universal Analytics dies
        /// </summary>
        public static void TrackVersions() {
            if (Program.InDeveloperMode) return;

            //OUTLOOK CLIENT
            Send(Analytics.Category.outlook, Analytics.Action.version, Outlook.Factory.OutlookVersionNameFull.Replace("Outlook", ""));

            //OGCS APPLICATION
            Send(Analytics.Category.ogcs, Analytics.Action.version, System.Windows.Forms.Application.ProductVersion);
        }

        public static void TrackSync() {
            if (Program.InDeveloperMode) return;
            Send(Analytics.Category.ogcs, Analytics.Action.sync, "calendar");
            Telemetry.GA4Event.Event syncGa4Ev = new(Telemetry.GA4Event.Event.Name.sync);
            syncGa4Ev.AddParameter(GA4.General.type, "calendar");
            syncGa4Ev.AddParameter(GA4.General.sync_count, Settings.Instance.CompletedSyncs);
            syncGa4Ev.Send();
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
                Ogcs.Exception.Analyse(ex);
            }
        }

        public class GA4Event {
            public String client_id { get; private set; }
            public String user_id { get; private set; }
            public Boolean non_personalized_ads { get; private set; }
            public Dictionary<String, Dictionary<String, String>> user_properties { get; private set; }
            public List<Event> events { get; private set; }

            /// <summary>
            /// A GA4 measurement protocol containing just the header/envelope propeties
            /// </summary>
            public GA4Event() {
                prepareEnvelope();
            }
            /// <summary>
            /// A GA4 event with no parameters
            /// </summary>
            /// <param name="eventName">The name of the event</param>
            /// <param name="throwAway">Don't pass anything here</param>
            public GA4Event(Event.Name eventName, Event throwAway = null) : this(eventName, out throwAway) { }
            /// <summary>
            /// A GA4 event that will contain event parameters. To support this, returns the nested event named with eventName
            /// </summary>
            /// <param name="eventName">The name of the event</param>
            /// <param name="theEvent">The new event to which parameters will be added</param>
            public GA4Event(Event.Name eventName, out Event _event) {
                prepareEnvelope();
                _event = new Event(eventName);
                events = new List<Event> { _event };
            }

            private void prepareEnvelope() {
                //https://developers.google.com/analytics/devguides/collection/protocol/ga4/sending-events?client_type=gtag#limitations
                //Maximum name length: 24; value length: 36
                client_id = Telemetry.Instance.AnonymousUniqueUserId; //Extend this in case more than one instance of OGCS running?
                user_id = Telemetry.Instance.AnonymousUniqueUserId;
                non_personalized_ads = true;
                user_properties = new Dictionary<String, Dictionary<String, String>>();
                user_properties.Add("ogcs_version", new Dictionary<String, String> { { "value", System.Windows.Forms.Application.ProductVersion } });
                user_properties.Add("benefactor", new Dictionary<String, String> { { "value", Settings.Instance.UserIsBenefactor().ToString() } });
                user_properties.Add("account_present", new Dictionary<String, String> { { "value", (!String.IsNullOrEmpty(Settings.Instance.GaccountEmail)).ToString() } });
                user_properties.Add("profiles", new Dictionary<String, String> { { "value", Settings.Instance.Calendars.Count.ToString() } });
                user_properties.Add("outlook_version", new Dictionary<String, String> { { "value", Telemetry.Instance.OutlookVersion } });
                user_properties.Add("outlook_name", new Dictionary<String, String> { { "value", Telemetry.Instance.OutlookVersionName } });
                user_properties.Add("continent", new Dictionary<String, String> { { "value", Telemetry.Instance.Continent } });
                user_properties.Add("country", new Dictionary<String, String> { { "value", Telemetry.Instance.Country } });
                user_properties.Add("country_code", new Dictionary<String, String> { { "value", Telemetry.Instance.CountryCode } });
                user_properties.Add("region", new Dictionary<String, String> { { "value", Telemetry.Instance.Region } });
                user_properties.Add("city", new Dictionary<String, String> { { "value", Telemetry.Instance.City } });
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
                    Ogcs.Exception.Analyse(ex);
                }
            }

            public class Event {
                public String name { get; private set; }
                public Dictionary<String, Object> parameters { get; private set; }

                public enum Name {
                    application_started,
                    debug,
                    donate,
                    error,
                    setting,
                    squirrel,
                    sync
                }

                public Event(Name eventName) {
                    name = eventName.ToString();
                }

                public void AddParameter(Object parameterName, Object parameterValue) {
                    if (parameters == null)
                        parameters = new Dictionary<String, Object>();

                    String strParamName = parameterName.ToString();
                    if (strParamName.Length > 40)
                        throw new ApplicationException($"The parameter name {strParamName} exceeds maximum length.");

                    if (!parameters.ContainsKey(strParamName))
                        parameters.Add(strParamName, null);

                    if (parameterValue is int)
                        parameters[strParamName] = (int)parameterValue;
                    else {
                        parameterValue ??= "";
                        parameters[strParamName] = parameterValue.ToString().Substring(0, Math.Min(parameterValue.ToString().Length, 100));
                    }
                }

                /// <summary>
                /// When sending an event, the "envelope" is created around it before posting
                /// </summary>
                public void Send() {
                    GA4Event ga4Ev = new GA4Event();
                    ga4Ev.events = new List<Event> { this };
                    ga4Ev.Send();
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

    public static class GA4 {
        public enum Squirrel {
            action_taken,
            error,
            feedback,
            file,
            install,
            result,
            state,
            target_version,
            target_type,
            upgraded_from,
            uninstall
        }
        public enum General {
            github_issue,
            sync_count,
            type
        }
    }
}
