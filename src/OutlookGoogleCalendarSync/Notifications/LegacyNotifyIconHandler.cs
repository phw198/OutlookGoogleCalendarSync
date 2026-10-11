using System;
using System.Windows.Forms;

namespace OutlookGoogleCalendarSync {
    internal class LegacyNotifyIconHandler : INotificationHandler {
        private readonly NotifyIcon ownerIcon;
        private readonly Action<NotificationActivation> onClick;

        public LegacyNotifyIconHandler(NotifyIcon icon, Action<NotificationActivation> onClick) {
            this.ownerIcon = icon;
            this.onClick = onClick;
        }

        public void ShowNotification(string message, string tagValue, ToolTipIcon iconType = ToolTipIcon.None, string browserUrl = "") {
            try {
                ownerIcon.Icon = Properties.Resources.icon; //Set to standard, non-animated icon
                ownerIcon.Tag = new NotificationActivation { Action = tagValue, BrowserUrl = browserUrl };
                ownerIcon.ShowBalloonTip(
                    500,
                    "Outlook Google Calendar Sync" + (string.IsNullOrEmpty(Program.Title) ? "" : $" - {Program.Title}"),
                    message,
                    iconType
                );
            } catch (System.Exception ex) {
                ex.Analyse("Unable to show legacy balloon notification.");
            }
        }
    }
}
