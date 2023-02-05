using System.Runtime.Serialization;
using System.Net;
using System;
using log4net;

//***
//https://developers.google.com/gdata/articles/proxy_setup
//The Google API needs updating so we don't have to rely on System proxy setting
//service.RequestFactory.Proxy = new WebProxy();

namespace OutlookGoogleCalendarSync.SettingsStore {
    [DataContract(Namespace = "http://schemas.datacontract.org/2004/07/OutlookGoogleCalendarSync")]
    public class Proxy {
        private static readonly ILog log = LogManager.GetLogger(typeof(Proxy));

        public Proxy() {
            setDefaults();
        }

        //Default values before loading from xml and attribute not yet serialized
        [OnDeserializing]
        void OnDeserializing(StreamingContext context) {
            setDefaults();
        }

        public String DefaultBrowserAgent {
            get { return "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/103.0.0.0 Safari/537.36"; }
        }
        private void setDefaults() {
            //Default values for new class
            this.Type = "IE";
            this.Port = 8888;

            //Browser agent can cause "HTTP-403 Forbidden" if target server/URL doesn't like it.
            //"Other" can be used as a fallback
            this.BrowserUserAgent = DefaultBrowserAgent;
        }

        [DataMember]
        public string Type { get; set; }

        [DataMember]
        public string ServerName { get; set; }

        [DataMember]
        public int Port { get; set; }

        [DataMember]
        public Boolean AuthenticationRequired { get; set; }

        [DataMember]
        public string UserName { get; set; }

        [DataMember]
        public string Password { get; set; }

        [DataMember]
        public string BrowserUserAgent { get; set; }

        public void Configure() {
            if (Type == "None") {
                log.Info("Removing proxy usage.");
                WebRequest.DefaultWebProxy = null;

            } else if (Type == "Custom") {
                log.Info("Setting custom proxy.");
                WebProxy wp = new WebProxy { Address = new System.Uri(string.Format("http://{0}:{1}", ServerName, Port)) };
                log.Debug("Using " + wp.Address);
                wp.BypassProxyOnLocal = true;
                if (AuthenticationRequired && !string.IsNullOrEmpty(UserName) && !string.IsNullOrEmpty(Password)) {
                    if (UserName.Contains(@"\")) {
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
                log.Info("Using default proxy (app.config / IE).");
                log.Info("Setting system-wide proxy.");
                IWebProxy iwp = WebRequest.GetSystemWebProxy();
                iwp.Credentials = CredentialCache.DefaultNetworkCredentials;
                WebRequest.DefaultWebProxy = iwp;
            }

            if (WebRequest.DefaultWebProxy != null) {
                try {
                    log.Debug("Testing the system proxy.");
                    String testUrl = "http://www.google.com";
                    WebRequest wr = WebRequest.CreateDefault(new System.Uri(testUrl));
                    System.Uri proxyUri = wr.Proxy.GetProxy(new System.Uri(testUrl));
                    log.Debug("Confirmation of configured proxy: " + proxyUri.OriginalString);
                    if (testUrl != proxyUri.OriginalString) {
                        try {
                            new Extensions.OgcsWebClient().OpenRead(testUrl);
                        } catch (WebException ex) {
                            if (ex.Response != null) {
                                System.IO.Stream stream = null;
                                System.IO.StreamReader sr = null;
                                try {
                                    HttpWebResponse hwr = ex.Response as HttpWebResponse;
                                    log.Debug("Proxy error status code: " + hwr.StatusCode + " = " + hwr.StatusDescription);
                                    stream = hwr.GetResponseStream();
                                    sr = new System.IO.StreamReader(stream);
                                    log.Fail(sr.ReadToEnd());
                                } catch (System.Exception ex2) {
                                    OGCSexception.Analyse("Could not analyse WebException response.", ex2);
                                } finally {
                                    if (sr != null) sr.Close();
                                    if (stream != null) stream.Close();
                                }
                            } else
                                OGCSexception.Analyse("Testing proxy connection failed.", ex);
                        }
                    }
                } catch (System.Exception ex) {
                    OGCSexception.Analyse("Failed to confirm proxy settings.", ex);
                }
            }
        }
    }
}