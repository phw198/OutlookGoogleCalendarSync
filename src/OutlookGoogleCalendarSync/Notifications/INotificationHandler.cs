using System.Windows.Forms;

namespace OutlookGoogleCalendarSync {
    internal class NotificationActivation {
        public string Action { get; set; }
        public string BrowserUrl { get; set; }
    }

    internal interface INotificationHandler {
        void ShowNotification(string message, ToolTipIcon iconType = ToolTipIcon.None, string tagValue = "", string browserUrl = "");
    }
}
