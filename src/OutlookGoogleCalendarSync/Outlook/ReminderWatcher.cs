using Ogcs = OutlookGoogleCalendarSync;
using log4net;
using Microsoft.Office.Interop.Outlook;

namespace OutlookGoogleCalendarSync.Outlook {
    /// <summary>
    /// Hides classic Outlook's own native reminder popups. Intended for when this Outlook
    /// instance is only being used in the background for OGCS's calendar automation, while
    /// the user's actual day-to-day client (eg "New Outlook") shows its own reminders for the
    /// same mailbox - without this, both clients independently pop a reminder for the same item.
    /// Cancels the popup via Reminders.BeforeReminderShow; never calls Dismiss()/Snooze()/Remove(),
    /// so nothing is written back to the mailbox and no other client's reminders are affected.
    /// </summary>
    class ReminderWatcher {
        private static readonly ILog log = LogManager.GetLogger(typeof(ReminderWatcher));

        private Reminders reminders;

        public ReminderWatcher(Application oApp) {
            Reminders reminders = null;
            try {
                reminders = oApp.Reminders;
                log.Info("Suppressing classic Outlook's own reminder popups, per configuration.");

                this.reminders = reminders;
                this.reminders.BeforeReminderShow += new ReminderCollectionEvents_BeforeReminderShowEventHandler(reminders_BeforeReminderShow);
            } catch (System.Exception ex) {
                ex.Analyse("Failed to set up Outlook reminder suppression.");
            }
        }

        private void reminders_BeforeReminderShow(ref bool Cancel) {
            log.Fine("Suppressing a native Outlook reminder popup.");
            Cancel = true;
        }

        public void Dispose() {
            if (reminders != null) {
                try { reminders.BeforeReminderShow -= reminders_BeforeReminderShow; } catch { }
                reminders = (Reminders)Calendar.ReleaseObject(reminders);
            }
        }
    }
}
