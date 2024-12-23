﻿using OutlookGoogleCalendarSync.Extensions;
using System;
using System.Text.RegularExpressions;
using Ogcs = OutlookGoogleCalendarSync;

namespace OutlookGoogleCalendarSync.GraphExtension {

    public static class Extensions {

        public static Microsoft.Graph.Extension OgcsExtension(this Microsoft.Graph.Event ai) {
            if (ai.Extensions == null) return null;

            foreach (Microsoft.Graph.Extension ext in ai.Extensions) {
                if (ext.Id == Ogcs.Outlook.Graph.CustomProperty.ExtensionName()
                 || ext.Id == Ogcs.Outlook.Graph.CustomProperty.ExtensionName(prefixWithMsType: true)
                )
                    return ext;
            }
            return null;
        }

        public static Microsoft.Graph.Event UpdateOgcsExtension(this Microsoft.Graph.Event ai, Microsoft.Graph.Extension updatedExt) {
            if (ai.Extensions == null) {
                ai.Extensions = new Microsoft.Graph.EventExtensionsCollectionPage();
            } else {
                Microsoft.Graph.Extension staleExt = ai.OgcsExtension();
                ai.Extensions.Remove(staleExt);
            }
            ai.Extensions.Add(updatedExt);
            return ai;
        }

        public static Boolean Compare(this Microsoft.Graph.Date date, Microsoft.Graph.Date otherDate) {
            return (date.Day == otherDate.Day && date.Month == otherDate.Month && date.Year == otherDate.Year);
        }

        /// <summary>Just the HTML within the <body> tags</summary>
        public static String BodyInnerHtml(this Microsoft.Graph.ItemBody body) {
            Regex htmlBodyTag = new Regex(@"<body>(?<body>.*?)</body>");
            String bodyInnerHtml = htmlBodyTag.Match(body.Content.RemoveLineBreaks()).Groups["body"]?.Value ?? "";
            if (bodyInnerHtml == "<div></div>") return "";
            else return bodyInnerHtml;
        }
    }
}