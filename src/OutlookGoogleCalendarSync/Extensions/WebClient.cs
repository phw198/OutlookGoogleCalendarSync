using System;
using System.Net;

namespace OutlookGoogleCalendarSync.Extensions {

    public class OgcsWebClient : WebClient {

        protected override WebRequest GetWebRequest(Uri address) {
            HttpWebRequest request = base.GetWebRequest(address) as HttpWebRequest;
            request.UserAgent = Settings.Instance.Proxy.BrowserUserAgent;
            return request;
        }
    }
}
