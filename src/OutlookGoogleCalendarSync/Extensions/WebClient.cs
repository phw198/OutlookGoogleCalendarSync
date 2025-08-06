using Microsoft.Identity.Client;
using System;
using System.Net;
using System.Net.Http;

namespace OutlookGoogleCalendarSync.Extensions {

    public class OgcsWebClient : WebClient {

        protected override WebRequest GetWebRequest(Uri address) {
            HttpWebRequest request = base.GetWebRequest(address) as HttpWebRequest;
            if (Settings.InstanceInitialiased)
                request.UserAgent = Settings.Instance.Proxy.BrowserUserAgent;
            else
                request.UserAgent = SettingsStore.Proxy.DefaultBrowserAgent;
            return request;
        }
    }

    public class MsalHttpClientFactoryAdapter : IMsalHttpClientFactory {
        private readonly IHttpClientFactory _httpClientFactory;

        public MsalHttpClientFactoryAdapter(IHttpClientFactory httpClientFactory) {
            _httpClientFactory = httpClientFactory;
        }

        public HttpClient GetHttpClient() {
            return _httpClientFactory.CreateClient("Msal");
        }
    }
}
