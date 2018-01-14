using log4net;
using System;
using System.Windows.Forms;

namespace OutlookGoogleCalendarSync {
    public class SyncTimer : Timer {
        private static readonly ILog log = LogManager.GetLogger(typeof(SyncTimer));
        private Timer ogcsTimer;
        public DateTime LastSyncDate { get; set; }
        public DateTime? NextSyncDate { 
            get {
                try {
                    if (Forms.Main.Instance.NextSyncVal == "Inactive" || !ogcsTimer.Enabled || 
                        Forms.Main.Instance.NextSyncVal == "Push Sync Active") {
                        return null;
                    } else {
                        return DateTime.ParseExact(Forms.Main.Instance.NextSyncVal,
                            System.Globalization.CultureInfo.CurrentCulture.DateTimeFormat.LongDatePattern + " - " +
                            System.Globalization.CultureInfo.CurrentCulture.DateTimeFormat.LongTimePattern,
                            System.Globalization.CultureInfo.CurrentCulture);
                    }
                } catch (System.Exception ex) {
                    log.Warn("Failed to determine next sync date from '" + Forms.Main.Instance.NextSyncVal +"'");
                    log.Error(ex.Message);
                    return null;
                }
            }
        }

        public SyncTimer() {
            ogcsTimer = new Timer();
            ogcsTimer.Tag = "AutoSyncTimer";
            ogcsTimer.Tick += new EventHandler(ogcsTimer_Tick);

            //Refresh synchronizations (last and next)
            LastSyncDate = Settings.Instance.LastSyncDate;
            Forms.Main.Instance.LastSyncVal = LastSyncDate.ToLongDateString() + " - " + LastSyncDate.ToLongTimeString();
            SetNextSync(getResyncInterval());
        }

        private void ogcsTimer_Tick(object sender, EventArgs e) {
            log.Debug("Scheduled sync triggered.");

            Forms.Main frm = Forms.Main.Instance;
            frm.NotificationTray.ShowBubbleInfo("Autosyncing calendars: " + Settings.Instance.SyncDirection.Name + "...");
            if (!frm.SyncingNow) {
                frm.Sync_Click(sender, null);
            } else {
                log.Debug("Busy syncing already. Rescheduled for 5 mins time.");
                SetNextSync(5, fromNow: true);
            }
        }

        private int getResyncInterval() {
            int min = Settings.Instance.SyncInterval;
            if (Settings.Instance.SyncIntervalUnit == "Hours") {
                min *= 60;
            }
            return min;
        }

        public void SetNextSync(int? delayMins = null, Boolean fromNow = false) {
            int _delayMins = delayMins ?? getResyncInterval();

            if (Settings.Instance.SyncInterval != 0) {
                DateTime nextSyncDate = this.LastSyncDate.AddMinutes(_delayMins);
                DateTime now = DateTime.Now;
                if (fromNow)
                    nextSyncDate = now.AddMinutes(_delayMins);

                if (ogcsTimer.Interval != (delayMins * 60000)) {
                    ogcsTimer.Stop();
                    TimeSpan diff = nextSyncDate - now;
                    if (diff.TotalMinutes < 1) {
                        nextSyncDate = now.AddMinutes(1);
                        ogcsTimer.Interval = 1 * 60000;
                    } else {
                        ogcsTimer.Interval = (int)(diff.TotalMinutes * 60000);
                    }
                    ogcsTimer.Start();
                }
                Forms.Main.Instance.NextSyncVal = nextSyncDate.ToLongDateString() + " - " + nextSyncDate.ToLongTimeString();
                log.Info("Next sync scheduled for " + Forms.Main.Instance.NextSyncVal);
            } else {
                Forms.Main.Instance.NextSyncVal = "Inactive";
                ogcsTimer.Stop();
                log.Info("Schedule disabled.");
            }
        }

        public void Switch(Boolean enable) {
            if (enable && !ogcsTimer.Enabled) ogcsTimer.Start();
            else if (!enable && ogcsTimer.Enabled) ogcsTimer.Stop();
        }

        public Boolean Running() {
            return ogcsTimer.Enabled;
        }
    }


    public class PushSyncTimer : Timer {
        private static readonly ILog log = LogManager.GetLogger(typeof(PushSyncTimer));
        private Timer ogcsTimer;
        public Int16 ItemsQueued { get; set; }

        public PushSyncTimer() {
            ItemsQueued = 0;
            ogcsTimer = new Timer();
            ogcsTimer.Tag = "PushTimer";
            ogcsTimer.Interval = 2 * 60000;
            ogcsTimer.Tick += new EventHandler(ogcsPushTimer_Tick);
            Forms.Main.Instance.NextSyncVal = "Push Sync Active";
        }

        private void ogcsPushTimer_Tick(object sender, EventArgs e) {
            if (ItemsQueued != 0) {
                log.Debug("Push sync triggered.");
                Forms.Main frm = Forms.Main.Instance;
                frm.NotificationTray.ShowBubbleInfo("Autosyncing calendars: " + Settings.Instance.SyncDirection.Name + "...");
                if (!frm.SyncingNow) {
                    frm.Sync_Click(sender, null);
                } else {
                    log.Debug("Busy syncing already. No need to push.");
                    ItemsQueued = 0;
                }
            } else {
                log.Fine("Push sync triggered, but no items queued.");
            }
        }

        public void Switch(Boolean enable) {
            if (enable && !ogcsTimer.Enabled) {
                ogcsTimer.Start();
                Forms.Main.Instance.NextSyncVal = "Push Sync Active";
            } else if (!enable && ogcsTimer.Enabled) {
                ogcsTimer.Stop();
                Forms.Main.Instance.OgcsTimer.SetNextSync();
            }
        }
        public Boolean Running() {
            return ogcsTimer.Enabled;
        }
    }
}
