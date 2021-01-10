using log4net;
using System;
using System.Windows.Forms;

namespace OutlookGoogleCalendarSync.Sync {
    public class SyncTimer : Timer {
        private static readonly ILog log = LogManager.GetLogger(typeof(SyncTimer));
        public DateTime LastSyncDate { private get; set; }
        private DateTime nextSyncDate;
        public DateTime NextSyncDate {
            get { return nextSyncDate; }
            private set {
                nextSyncDate = value;
                NextSyncDateText = nextSyncDate.ToLongDateString() + " @ " + nextSyncDate.ToLongTimeString();
                if (Settings.Instance.OutlookPush) NextSyncDateText += " + Push";
                Forms.Main.Instance.NextSyncVal = NextSyncDateText;
                log.Info("Next sync scheduled for " + NextSyncDateText);
            }
        }
        public String NextSyncDateText { get; private set; }
        
        public void Initialise() {
            this.Tag = "AutoSyncTimer";
            this.Tick += new EventHandler(ogcsTimer_Tick);
            this.Interval = int.MaxValue;

            //Refresh synchronizations (last and next)
            this.LastSyncDate = Settings.Instance.LastSyncDate;
            Forms.Main.Instance.LastSyncVal = LastSyncDate.ToLongDateString() + " @ " + LastSyncDate.ToLongTimeString();
            SetNextSync();
        }

        private void ogcsTimer_Tick(object sender, EventArgs e) {
            if (Forms.ErrorReporting.Instance.Visible) return;
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

        /// <summary>Configure the next sync according to configured schedule in Settings.</summary>
        public void SetNextSync() {
            SetNextSync(getResyncInterval());
        }
        /// <summary>
        /// Configure the next sync that override any configured schedule in Settings.</summary>
        /// </summary>
        /// <param name="delayMins">Number of minutes to next sync</param>
        /// <param name="fromNow">From now or since last successful sync</param>
        /// <param name="calculateInterval">Calculate milliseconds to next sync and activate timer</param>
        public void SetNextSync(int delayMins, Boolean fromNow = false, Boolean calculateInterval = true) {
            if (Settings.Instance.SyncInterval != 0) {
                DateTime now = DateTime.Now;
                this.NextSyncDate = fromNow ? now.AddMinutes(delayMins) : LastSyncDate.AddMinutes(delayMins);
                if (calculateInterval) CalculateInterval();

            } else {
                Forms.Main.Instance.NextSyncVal = "Inactive";
                Activate(false);
                log.Info("Schedule disabled.");
            }
        }
        public void CalculateInterval() {
            DateTime now = DateTime.Now;
            double interval = (this.nextSyncDate - now).TotalMinutes;

            if (this.Interval != (interval * 60000)) {
                Activate(false);
                if (interval < 0) {
                    log.Debug("Moving past sync into imminent future.");
                    this.NextSyncDate = now.AddMinutes(1);
                    this.Interval = 1 * 60000;
                } else {
                    this.Interval = (int)(interval * 60000);
                }
            }
            Activate(true);
        }
        
        public void Activate(Boolean activate) {
            if (Forms.Main.Instance.InvokeRequired) {
                log.Error("Attempted to " + (activate ? "" : "de") + "activate " + this.Tag + " from non-GUI thread will not work.");
                return;
            }

            if (activate && !this.Enabled) this.Start();
            else if (!activate && this.Enabled) this.Stop();
        }

        public Boolean Running() {
            return this.Enabled;
        }
    }


    public class PushSyncTimer : Timer {
        private static readonly ILog log = LogManager.GetLogger(typeof(PushSyncTimer));
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
            this.Tag = "PushTimer";
            this.Interval = 2 * 60000;
            this.Tick += new EventHandler(ogcsPushTimer_Tick);
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
            if (Forms.ErrorReporting.Instance.Visible) return;
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

                String hResult = OGCSexception.GetErrorCode(ex);
                if ((ex is System.InvalidCastException && hResult == "0x80004002" && ex.Message.Contains("0x800706BA")) || //The RPC server is unavailable
                    (ex is System.Runtime.InteropServices.COMException && (
                        ex.Message.Contains("0x80010108(RPC_E_DISCONNECTED)") || //The object invoked has disconnected from its clients
                        hResult == "0x800706BE" || //The remote procedure call failed
                        hResult == "0x800706BA")) //The RPC server is unavailable
                    ) {
                    OGCSexception.Analyse(OGCSexception.LogAsFail(ex));
                    try {
                        OutlookOgcs.Calendar.Instance.Reset();
                    } catch (System.Exception ex2) {
                        OGCSexception.Analyse("Failed resetting Outlook connection.", ex2);
                    }
                } else
                    OGCSexception.Analyse(ex);
                if (failures == 10)
                    Forms.Main.Instance.Console.UpdateWithError("Push Sync is failing.", ex, notifyBubble: true);
            }
        }

        public void Activate(Boolean activate) {
            if (activate && !this.Enabled) {
                ResetLastRun();
                this.Start();
                if (Settings.Instance.SyncInterval == 0) Forms.Main.Instance.NextSyncVal = "Push Sync Active";
            } else if (!activate && this.Enabled) {
                this.Stop();
                Sync.Engine.Instance.OgcsTimer.SetNextSync();
            }
        }
        public Boolean Running() {
            return this.Enabled;
        }
    }
}
