using System.Runtime.Serialization;
using System.Net;
using log4net;

namespace OutlookGoogleCalendarSync {
    [DataContract]
    public class SettingsProxy {
        private static readonly ILog log = LogManager.GetLogger(typeof(SettingsProxy));

        public SettingsProxy() {
            setDefaults();
        }

        //Default values before loading from xml and attribute not yet serialized
        [OnDeserializing]
        void OnDeserializing(StreamingContext context) {
            setDefaults();
        }

        private void setDefaults() {
            //Default values for new class
            this.Type = "IE";
            this.Port = 8888;
        }
        
        [DataMember]
        public string Type { get; set; }

        [DataMember]
        public string ServerName { get; set; }

        [DataMember]
        public int Port { get; set; }

        [DataMember]
        public string UserName { get; set; }

        [DataMember]
        public string Password { get; set; }

        public void Configure() {
            if (Type == "None") {
                log.Info("Removing proxy usage.");
                WebRequest.DefaultWebProxy = null;

            } else if (Type == "Custom") {
                log.Info("Setting custom proxy.");
                WebProxy wp = new WebProxy();
                wp.Address = new System.Uri(string.Format("http://{0}:{1}", ServerName, Port));
                log.Debug("Using " + wp.Address);
                wp.BypassProxyOnLocal = true;
                if (!string.IsNullOrEmpty(UserName) && !string.IsNullOrEmpty(Password)) {
                    if (UserName.Contains("\\")) {
                        try {
                            string[] usernameBits = UserName.Split('\\');
                            wp.Credentials = new NetworkCredential(usernameBits[1], Password, usernameBits[0]);
                        } catch (System.Exception ex) {
                            log.Error("Failed to extract domain from proxy username: " + UserName);
                            log.Error(ex.Message);
                        }
                    } else {
                        wp.Credentials = new NetworkCredential(UserName, Password);
                    }
                }
                WebRequest.DefaultWebProxy = wp;

            } else { //IE
                if (WebRequest.DefaultWebProxy != null) {
                    log.Info("Using default proxy (app.config / IE).");
                } else {
                    log.Info("Setting system-wide proxy.");
                    IWebProxy iwp = WebRequest.GetSystemWebProxy();
                    iwp.Credentials = CredentialCache.DefaultNetworkCredentials;
                    WebRequest.DefaultWebProxy = iwp;
                }
            }

            if (WebRequest.DefaultWebProxy != null) { //Now let's test it
                try {
                    WebRequest wr = WebRequest.CreateDefault(new System.Uri("http://www.google.com"));
                    System.Uri proxyUri = wr.Proxy.GetProxy(new System.Uri("http://www.google.com"));
                    log.Debug("Confirmation of configured proxy: " + proxyUri.OriginalString);
                } catch (System.Exception ex) {
                    log.Error("Failed to confirm proxy settings.");
                    log.Error(ex.Message);
                }
            }
        }
    }
}