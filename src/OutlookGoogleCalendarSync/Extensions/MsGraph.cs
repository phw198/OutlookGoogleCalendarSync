using OutlookGoogleCalendarSync.Extensions;
using System;
using System.Text.RegularExpressions;
using MsGraph = OutlookGoogleCalendarSync.Outlook.Graph.CustomClient;
using Ogcs = OutlookGoogleCalendarSync;

namespace OutlookGoogleCalendarSync.GraphExtension {

    public static class Extensions {

        public static MsGraph.Models.Extension OgcsExtension(this MsGraph.Models.Event ai) {
            if (ai.Extensions == null) return null;

            foreach (MsGraph.Models.Extension ext in ai.Extensions) {
                if (ext.Id == Ogcs.Outlook.Graph.CustomProperty.ExtensionName()
                 || ext.Id == Ogcs.Outlook.Graph.CustomProperty.ExtensionName(prefixWithMsType: true)
                )
                    return ext;
            }
            return null;
        }

        public static MsGraph.Models.Event UpdateOgcsExtension(this MsGraph.Models.Event ai, MsGraph.Models.Extension updatedExt) {
            if (ai.Extensions == null) {
                ai.Extensions = new();
            } else {
                MsGraph.Models.Extension staleExt = ai.OgcsExtension();
                ai.Extensions.Remove(staleExt);
            }
            ai.Extensions.Add(updatedExt);
            return ai;
        }

        public static Boolean Compare(this Microsoft.Kiota.Abstractions.Date date, Microsoft.Kiota.Abstractions.Date otherDate) {
            return (date.Day == otherDate.Day && date.Month == otherDate.Month && date.Year == otherDate.Year);
        }

        /// <summary>Just the HTML within the <body> tags</summary>
        public static String BodyInnerHtml(this MsGraph.Models.ItemBody body) {
            if (body == null || String.IsNullOrEmpty(body.Content)) return "";

            Regex htmlBodyTag = new Regex(@"<body>(?<body>.*?)</body>");
            String bodyInnerHtml = htmlBodyTag.Match(body.Content.RemoveLineBreaks()).Groups["body"]?.Value ?? "";
            if (bodyInnerHtml == "<div></div>") return "";
            else return bodyInnerHtml;
        }

        /// <summary>Get the body content suitable for synchronization, handling both plain text and HTML conversions</summary>
        public static String BodyContentForSync(this MsGraph.Models.ItemBody body) {
            //Although there is a body type "Text" - MS Exchange can automatically convert it to HTML!
            
            if (body == null || String.IsNullOrEmpty(body.Content)) return "";

            //So first check for evidence of this conversion
            Boolean convertedFromText = body.ContentType == MsGraph.Models.BodyType.Text ||
                Regex.IsMatch(body.Content, @"<!--\s*converted from text\s*-->", RegexOptions.IgnoreCase);
            if (!convertedFromText) return body.BodyInnerHtml();

            //Then look for a <div> class "PlainText" which indicates the original plain text content
            Match plainTextDiv = Regex.Match(body.Content,
                @"<div\b[^>]*\bclass\s*=\s*[""']PlainText[""'][^>]*>(?<text>.*?)</div>",
                RegexOptions.IgnoreCase | RegexOptions.Singleline);
            if (!plainTextDiv.Success) return body.ContentType == MsGraph.Models.BodyType.Text ? body.Content : body.BodyInnerHtml();

            String text = Regex.Replace(plainTextDiv.Groups["text"].Value, @"<br\s*/?>", "\r\n", RegexOptions.IgnoreCase);
            text = text.RemoveHtmlTags();
            return System.Net.WebUtility.HtmlDecode(text).RemoveNBSP();
        }

        /// <summary>Add the Authorization header to an HTTP Request Message</summary>
        public static System.Net.Http.HttpRequestMessage AddAuthorisation(this System.Net.Http.HttpRequestMessage a) {
            //This is required due to
            // 1. cancelledOccurrences Graph Event property being on the v1.0 API, but undocumented
            // 2. the Graph SDK only supporting that property on the beta release channel
            // 3. Native GetHttpRequestMessage() to build custom API call doesn't utilise existing Authorization header; https://github.com/microsoftgraph/msgraph-sdk-dotnet/issues/263
            
            a.Headers.Authorization = new("Bearer", Ogcs.Outlook.Graph.Calendar.Instance.Authenticator.AccessToken);
            return a;
        }
    }
}
