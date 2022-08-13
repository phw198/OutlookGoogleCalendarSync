using log4net;
using System;
using System.Windows.Forms;

namespace OutlookGoogleCalendarSync.Sync {
    public class SyncTimer : Timer {
        private static readonly ILog log = LogManager.GetLogger(typeof(SyncTimer));
        public object owningProfile { get; internal set; }
        
        public DateTime LastSyncDate { internal get; set; }

        private DateTime? nextSyncDate;
        public DateTime? NextSyncDate {
            get { return nextSyncDate; }
            set {
                nextSyncDate = value;
                if (nextSyncDate != null) {
                    DateTime theDate = (DateTime)nextSyncDate;
                    var profile = owningProfile as SettingsStore.Calendar;
                    NextSyncDateText = theDate.ToLongDateString() + " @ " + theDate.ToLongTimeString() + (profile.OutlookPush ? " + Push" : "");
                }
            }
        }

        private String nextSyncDateText;
        public String NextSyncDateText {
            get { return nextSyncDateText; }
            set {
                nextSyncDateText = value;
                var profile = owningProfile as SettingsStore.Calendar;
                if (profile.Equals(Forms.Main.Instance.ActiveCalendarProfile))
                    Forms.Main.Instance.NextSyncVal = value;
            }
        }
        
        public SyncTimer(Object owningProfile) {
            this.owningProfile = owningProfile;
            this.Tag = "AutoSyncTimer";
            this.Tick += new EventHandler(ogcsTimer_Tick);
            this.Interval = int.MaxValue;

            if (owningProfile is SettingsStore.Calendar)
                this.LastSyncDate = (owningProfile as SettingsStore.Calendar).LastSyncDate;

            SetNextSync();
        }

        private void ogcsTimer_Tick(object sender, EventArgs e) {
            if (Forms.ErrorReporting.Instance.Visible) return;

            log.Debug("Scheduled sync triggered for profile: "+ Settings.Profile.Name(this.owningProfile));

            if (!Sync.Engine.Instance.SyncingNow) {
                Forms.Main.Instance.Sync_Click(sender, null);
            } else {
                log.Debug("Busy syncing already. Rescheduled for 5 mins time.");
                SetNextSync(5, fromNow: true);
            }
        }

        private int getResyncInterval() {
            var profile = (this.owningProfile as SettingsStore.Calendar);
            int min = profile.SyncInterval;
            if (profile.SyncIntervalUnit == "Hours") {
                min *= 60;
            }
            return min;
        }

        /// <summary>Configure the next sync according to configured schedule in Settings.</summary>
        public void SetNextSync() {
            SetNextSync(getResyncInterval());
        }

        /// <summary>Configure the next sync that override any configured schedule in Settings.</summary>
        /// <param name="delayMins">Number of minutes to next sync</param>
        /// <param name="fromNow">From now or since last successful sync</param>
        /// <param name="calculateInterval">Calculate milliseconds to next sync and activate timer</param>
        public void SetNextSync(int delayMins, Boolean fromNow = false, Boolean calculateInterval = true) {
            SettingsStore.Calendar profile = null;
            if (owningProfile is SettingsStore.Calendar)
                profile = SettingsStore.Calendar.GetCalendarProfile(owningProfile);
            
            if (profile == null || profile.SyncInterval == 0) {
                this.NextSyncDateText = (profile?.OutlookPush ?? false) ? "Push Sync Active" : "Inactive";
                Activate(false);
                log.Info("Schedule disabled.");
            } else {
                DateTime now = DateTime.Now;
                this.nextSyncDate = fromNow ? now.AddMinutes(delayMins) : this.LastSyncDate.AddMinutes(delayMins);
                if (calculateInterval) CalculateInterval();
                else this.NextSyncDate = this.nextSyncDate;
                log.Info("Next sync scheduled for profile '"+ Settings.Profile.Name(owningProfile) +"' at " + this.NextSyncDateText);
            }
        }

        public void CalculateInterval() {
            if ((owningProfile as SettingsStore.Calendar).SyncInterval == 0) return;

            DateTime now = DateTime.Now;
            double interval = ((DateTime)this.nextSyncDate - now).TotalMinutes;

            if (this.Interval != (interval * 60000)) {
                Activate(false);
                if (interval < 0) {
                    log.Debug("Moving past sync into imminent future.");
                    this.Interval = 1 * 60000;
                } else if (interval == 0)
                    this.Interval = 1000;
                else
                    this.Interval = (int)Math.Min((interval * 60000), int.MaxValue);
                this.NextSyncDate = now.AddMilliseconds(this.Interval);
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

        public String Status() {
            var profile = (owningProfile as SettingsStore.Calendar);
            if (this.Running()) return NextSyncDateText;
            else if (profile.OgcsPushTimer != null && profile.OgcsPushTimer.Running()) return "Push Sync Active";
            else return "Inactive";
        }
    }


    public class PushSyncTimer : Timer {
        private static readonly ILog log = LogManager.GetLogger(typeof(PushSyncTimer));
        public object owningProfile { get; internal set; }
        private DateTime lastRunTime;
        private Int32 lastRunItemCount;
        private Int16 failures = 0;
        public PushSyncTimer(Object owningProfile) {
            this.owningProfile = owningProfile;
            ResetLastRun();
            this.Tag = "PushTimer";
            this.Interval = 2 * 60000;
            if (Program.InDeveloperMode) this.Interval = 30000;
            this.Tick += new EventHandler(ogcsPushTimer_Tick);
        }

        /// <summary>
        /// Recalculate item count as of now.
        /// </summary>
        public void ResetLastRun() {
            this.lastRunTime = DateTime.Now;
            try {
                log.Fine("Updating calendar item count following Push Sync.");
                this.lastRunItemCount = OutlookOgcs.Calendar.Instance.GetCalendarEntriesInRange(this.owningProfile as SettingsStore.Calendar, true).Count;
            } catch (System.Exception ex) {
                OGCSexception.Analyse("Failed to update item count following a Push Sync.", ex);
            }
        }

        private void ogcsPushTimer_Tick(object sender, EventArgs e) {
            if (Forms.ErrorReporting.Instance.Visible) return;
            log.UltraFine("Push sync triggered.");
            
            try {
                SettingsStore.Calendar profile = this.owningProfile as SettingsStore.Calendar;

                //In case the IOutlook.Connect() has to be called which needs an active profile
                if (Sync.Engine.Calendar.Instance.Profile == null)
                    //Force in the push sync profile
                    Sync.Engine.Calendar.Instance.Profile = profile;

                if (OutlookOgcs.Calendar.Instance.IOutlook.NoGUIexists()) return;
                log.Fine("Push sync triggered for profile: " + Settings.Profile.Name(profile));
                System.Collections.Generic.List<Microsoft.Office.Interop.Outlook.AppointmentItem> items = OutlookOgcs.Calendar.Instance.GetCalendarEntriesInRange(profile, true);

                if (items.Count < this.lastRunItemCount || items.FindAll(x => x.LastModificationTime > this.lastRunTime).Count > 0) {
                    log.Debug("Changes found for Push sync.");
                        Forms.Main.Instance.Sync_Click(sender, null);
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
            SettingsStore.Calendar profile = this.owningProfile as SettingsStore.Calendar;
            if (activate && !this.Enabled) {
                ResetLastRun();
                this.Start();
                if (profile.SyncInterval == 0 && profile.Equals(Forms.Main.Instance.ActiveCalendarProfile)) 
                    Forms.Main.Instance.NextSyncVal = "Push Sync Active";
            } else if (!activate && this.Enabled) {
                this.Stop();
                profile.OgcsTimer.SetNextSync();
            }
        }
        public Boolean Running() {
            return this.Enabled;
        }
    }
}
