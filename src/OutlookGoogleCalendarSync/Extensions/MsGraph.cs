using OutlookGoogleCalendarSync.Extensions;
using System;
using System.Text.RegularExpressions;
using Ogcs = OutlookGoogleCalendarSync;

namespace OutlookGoogleCalendarSync.GraphExtension {

    public static class Extensions {

        public static Microsoft.Graph.Extension OgcsExtension(this Microsoft.Graph.Event ai) {
            if (ai.Extensions == null) return null;

            foreach (Microsoft.Graph.Extension ext in ai.Extensions) {
                if (ext.Id == Ogcs.Outlook.Graph.O365CustomProperty.ExtensionName()
                 || ext.Id == Ogcs.Outlook.Graph.O365CustomProperty.ExtensionName(prefixWithMsType: true)
                )
                    return ext;
            }
            return null;
        }

        /// <summary>Just the HTML within the <body> tags</summary>
        public static String BodyInnerHtml(this Microsoft.Graph.ItemBody body) {
            Regex htmlBodyTag = new Regex(@"<body>(?<body>.*?)</body>");
            return htmlBodyTag.Match(body.Content.RemoveLineBreaks()).Groups["body"]?.Value ?? "";
        }
    }
}