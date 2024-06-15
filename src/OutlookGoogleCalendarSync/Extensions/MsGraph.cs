using Ogcs = OutlookGoogleCalendarSync;

namespace OutlookGoogleCalendarSync.GraphExtension {

    public static class Extensions {

        public static Microsoft.Graph.Extension OgcsExtension(this Microsoft.Graph.Event ai) {
            foreach (Microsoft.Graph.Extension ext in ai.Extensions) {
                if (ext.Id == Ogcs.Outlook.Graph.O365CustomProperty.ExtensionName()
                 || ext.Id == Ogcs.Outlook.Graph.O365CustomProperty.ExtensionName(prefixWithMsType: true)
                )
                    return ext;
            }
            return null;
        }
    }
}