using Ogcs = OutlookGoogleCalendarSync;
using log4net;
using Microsoft.Office.Interop.Outlook;

namespace OutlookGoogleCalendarSync.Outlook {
    /// <summary>
    /// Hides classic Outlook's own native reminder popups when this Outlook instance is only running
    /// in the background for OGCS's calendar automation - ie it has no Explorer windows open. Without this,
    /// a headless classic Outlook pops just its reminders window, and if the user's day-to-day client
    /// (eg "New Outlook") shows its own reminders for the same mailbox, both clients pop a reminder for the same item.
    /// The Explorer count is re-evaluated on every reminder, so as soon as the user opens classic Outlook
    /// themselves, its reminders show as normal again.
    /// Cancels the popup via Reminders.BeforeReminderShow; never calls Dismiss()/Snooze()/Remove(),
    /// so nothing is written back to the mailbox and no other client's reminders are affected.
    /// </summary>
    class ReminderWatcher {
        private static readonly ILog log = LogManager.GetLogger(typeof(ReminderWatcher));

        private Application oApp;
        private Reminders reminders;

        public ReminderWatcher(Application oApp) {
            try {
                this.oApp = oApp;
                this.reminders = oApp.Reminders;
                log.Debug("Setting up Outlook reminder watcher.");
                this.reminders.BeforeReminderShow += new ReminderCollectionEvents_BeforeReminderShowEventHandler(reminders_BeforeReminderShow);
            } catch (System.Exception ex) {
                ex.Analyse("Failed to set up Outlook reminder watcher.");
            }
        }

        private void reminders_BeforeReminderShow(ref bool Cancel) {
            if (!Settings.Instance.SuppressOutlookReminders) return;

            Explorers explorers = null;
            try {
                explorers = oApp.Explorers;
                if (explorers.Count == 0) {
                    log.Fine("Suppressing a native Outlook reminder popup - Outlook is running in the background with no open windows.");
                    Cancel = true;
                } else {
                    log.Fine("Outlook has " + explorers.Count + " open window(s), so not suppressing its reminder popup.");
                }
            } catch (System.Exception ex) {
                ex.Analyse("Could not determine whether Outlook has any open windows; not suppressing reminder.");
            } finally {
                explorers = (Explorers)Calendar.ReleaseObject(explorers);
            }
        }

        public void Dispose() {
            if (reminders != null) {
                try { reminders.BeforeReminderShow -= reminders_BeforeReminderShow; } catch { }
                reminders = (Reminders)Calendar.ReleaseObject(reminders);
            }
            oApp = null;
        }
    }
}
