using System;
using System.Windows.Forms;

namespace OutlookGoogleCalendarSync {
    internal class LegacyNotifyIconHandler : INotificationHandler {
        private readonly NotifyIcon ownerIcon;

        public LegacyNotifyIconHandler(NotifyIcon icon, Action<string> onClick) {
            this.ownerIcon = icon;
        }

        public void ShowNotification(string message, ToolTipIcon iconType = ToolTipIcon.None, string tagValue = "") {
            try {
                ownerIcon.Icon = Properties.Resources.icon; //Set to standard, non-animated icon
                ownerIcon.Tag = tagValue;
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
