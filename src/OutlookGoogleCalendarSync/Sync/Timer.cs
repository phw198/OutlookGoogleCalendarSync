using log4net;
using System;
using System.Windows.Forms;

namespace OutlookGoogleCalendarSync.Sync {
    public class SyncTimer : Timer {
        private static readonly ILog log = LogManager.GetLogger(typeof(SyncTimer));
        private Timer ogcsTimer;
        public DateTime LastSyncDate { get; set; }
        public DateTime? NextSyncDate { 
            get {
                try {
                    if ("Inactive;Push Sync Active;In progress...".Contains(Forms.Main.Instance.NextSyncVal) || !ogcsTimer.Enabled) {
                        return null;
                    } else {
                        return DateTime.ParseExact(Forms.Main.Instance.NextSyncVal.Replace(" + Push",""),
                            System.Globalization.CultureInfo.CurrentCulture.DateTimeFormat.LongDatePattern + " @ " +
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
            Forms.Main.Instance.LastSyncVal = LastSyncDate.ToLongDateString() + " @ " + LastSyncDate.ToLongTimeString();
            SetNextSync(getResyncInterval());
        }

        private void ogcsTimer_Tick(object sender, EventArgs e) {
            log.Debug("Scheduled sync triggered.");

            Forms.Main frm = Forms.Main.Instance;
            frm.NotificationTray.ShowBubbleInfo("Autosyncing calendars: " + Settings.Instance.SyncDirection.Name + "...");
            if (!Sync.Engine.Instance.SyncingNow) {
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
                Forms.Main.Instance.NextSyncVal = nextSyncDate.ToLongDateString() + " @ " + nextSyncDate.ToLongTimeString();
                if (Settings.Instance.OutlookPush) Forms.Main.Instance.NextSyncVal += " + Push";
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
        private DateTime lastRunTime;
        private Int32 lastRunItemCount;
        private Int16 failures = 0;
        private static PushSyncTimer instance;
        public static PushSyncTimer Instance {
            get {
                if (instance == null) {
                    instance = new PushSyncTimer();
                }
                return instance;
            }
        }

        private PushSyncTimer() {
            ResetLastRun();
            ogcsTimer = new Timer();
            ogcsTimer.Tag = "PushTimer";
            ogcsTimer.Interval = 2 * 60000;
            ogcsTimer.Tick += new EventHandler(ogcsPushTimer_Tick);
            Forms.Main.Instance.NextSyncVal = Settings.Instance.SyncInterval == 0 
                ? "Push Sync Active" 
                : Forms.Main.Instance.NextSyncVal = Forms.Main.Instance.NextSyncVal.Replace(" + Push", "") + " + Push";
        }

        public void ResetLastRun() {
            this.lastRunTime = DateTime.Now;
            try {
                log.Fine("Updating calendar item count following Push Sync.");
                this.lastRunItemCount = OutlookOgcs.Calendar.Instance.GetCalendarEntriesInRange(true).Count;
            } catch {
                log.Error("Failed to update item count following a Push Sync.");
            }
        }

        private void ogcsPushTimer_Tick(object sender, EventArgs e) {
            log.Fine("Push sync triggered.");

            try {
                System.Collections.Generic.List<Microsoft.Office.Interop.Outlook.AppointmentItem> items = OutlookOgcs.Calendar.Instance.GetCalendarEntriesInRange(true);

                if (items.Count < this.lastRunItemCount || items.FindAll(x => x.LastModificationTime > this.lastRunTime).Count > 0) {
                    log.Debug("Changes found for Push sync.");
                    Forms.Main.Instance.NotificationTray.ShowBubbleInfo("Autosyncing calendars: " + Settings.Instance.SyncDirection.Name + "...");
                    if (!Sync.Engine.Instance.SyncingNow) {
                        Forms.Main.Instance.Sync_Click(sender, null);
                    } else {
                        log.Debug("Busy syncing already. No need to push.");
                    }
                } else {
                    log.Fine("No changes found.");
                }
                failures = 0;
            } catch (System.Exception ex) {
                failures++;
                log.Warn("Push Sync failed " + failures + " times to check for changed items.");
                OGCSexception.Analyse(ex);
                if (failures == 10)
                    Forms.Main.Instance.Console.UpdateWithError("Push Sync is failing.", ex, notifyBubble: true);
            }
        }

        public void Switch(Boolean enable) {
            if (enable && !ogcsTimer.Enabled) {
                ResetLastRun();
                ogcsTimer.Start();
                if (Settings.Instance.SyncInterval == 0) Forms.Main.Instance.NextSyncVal = "Push Sync Active";
            } else if (!enable && ogcsTimer.Enabled) {
                ogcsTimer.Stop();
                Sync.Engine.Instance.OgcsTimer.SetNextSync();
            }
        }
        public Boolean Running() {
            return ogcsTimer.Enabled;
        }
    }
}
