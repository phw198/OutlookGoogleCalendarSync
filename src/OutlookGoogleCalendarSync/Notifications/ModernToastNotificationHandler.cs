using log4net;
using Microsoft.Toolkit.Uwp.Notifications;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using Windows.Data.Xml.Dom;
using Windows.UI.Notifications;

namespace OutlookGoogleCalendarSync {
    internal class ModernToastNotificationHandler : INotificationHandler {
        private static readonly ILog log = LogManager.GetLogger(typeof(ModernToastNotificationHandler));
        private readonly Action<string> onClick;
        private readonly NotifyIcon ownerIcon;

        public ModernToastNotificationHandler(NotifyIcon ownerIcon, Action<string> onClick) {
            this.ownerIcon = ownerIcon;
            this.onClick = onClick;
        }

        public void ShowNotification(string message, ToolTipIcon iconType = ToolTipIcon.None, string tagValue = "") {
            String title = "Outlook Google Calendar Sync" + (string.IsNullOrEmpty(Program.Title) ? "" : $" - {Program.Title}");
            try {
                String originalMessage = message;
                List<String> messages = message.Split("\r\n".ToCharArray(), StringSplitOptions.RemoveEmptyEntries).ToList();

                Regex syncSummary = new Regex(@"\d+\screated;\s\d+\sdeleted;\s\d+\supdated");
                if (syncSummary.IsMatch(messages[0])) {
                    if (Sync.Engine.Instance.ActiveProfile is SettingsStore.Calendar profile)
                        title = $"Profile '{profile._ProfileName}' Sync Summary";
                    else
                        title = "Sync Summary";
                } else if (messages.Count > 1) {
                    title = messages[0].TrimEnd('.');
                    messages.RemoveAt(0);
                    message = string.Join("\r\n", messages);
                }

                ToastContent toastContent = new ToastContentBuilder()
                    .AddText(title)
                    .AddText(message)
                    .AddArgument("action", tagValue)
                    .AddArgument("message", originalMessage)
                    .SetToastScenario(ToastScenario.Default)
                    .GetToastContent();

                XmlDocument toastXml = toastContent.GetXml();
                ToastNotifierCompat notifier = ToastNotificationManagerCompat.CreateToastNotifier();
                ToastNotification toast = new Windows.UI.Notifications.ToastNotification(toastXml);
                toast.Activated += (sender, args) => {
                    log.Debug("ModernToastNotificationHandler: Toast activated");
                    onClick?.Invoke(tagValue);
                };
                toast.Dismissed += (sender, args) => {
                    log.Debug("ModernToastNotificationHandler: Toast dismissed");
                };
                notifier.Show(toast);

            } catch (System.Exception ex) {
                ex.Analyse("Unable to show toast");
                ownerIcon?.ShowBalloonTip(500, title, message, iconType);
            }
        }
    }
}
