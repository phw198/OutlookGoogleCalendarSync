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
            WebProxy wp;
            if (Type == "None") {
                log.Info("Removing proxy usage.");
                WebRequest.DefaultWebProxy = null;

            } else if (Type == "Custom") {
                log.Info("Setting custom proxy.");
                wp = new WebProxy();
                wp.Address = new System.Uri(string.Format("http://{0}:{1}", ServerName, Port));
                log.Debug("Using " + wp.Address);
                wp.BypassProxyOnLocal = true;
                if (!string.IsNullOrEmpty(UserName) && !string.IsNullOrEmpty(Password)) {
                    wp.Credentials = new NetworkCredential(UserName, Password);
                }
                WebRequest.DefaultWebProxy = wp;

            } else { //IE
                log.Info("Setting system proxy.");
                IWebProxy iwp = WebRequest.GetSystemWebProxy();
                iwp.Credentials = CredentialCache.DefaultNetworkCredentials;
                WebRequest.DefaultWebProxy = iwp;
            }
            //IWebProxy iwp = WebRequest.DefaultWebProxy;
        }

    }
}