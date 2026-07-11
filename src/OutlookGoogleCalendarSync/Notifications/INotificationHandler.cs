using System.Windows.Forms;

namespace OutlookGoogleCalendarSync {
    internal interface INotificationHandler {
        void ShowNotification(string message, ToolTipIcon iconType = ToolTipIcon.None, string tagValue = "");
    }
}
