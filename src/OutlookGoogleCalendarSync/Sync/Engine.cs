using log4net;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace OutlookGoogleCalendarSync.Sync {
    public partial class Engine {
        public class Job {
            public String RequestedBy { get; internal set; }
            public String ProfileName { get; internal set; }
            public Object Profile { get; internal set; }
            public Job(String requestBy, Object profile) {
                this.RequestedBy = requestBy;
                this.ProfileName = Settings.Profile.Name(profile);
                this.Profile = profile;
            }

            public class Queue {
                //Generic Queue object would be nice, but then can't dedupe
                private static readonly ILog log = LogManager.GetLogger(typeof(Queue));

                Timer queueTimer;
                List<Dictionary<String, Job>> queue; //Generic Queue object would be nice, but then can't dedupe

                public Queue() {
                    this.queue = new List<Dictionary<String, Job>>();
                    this.queueTimer = new Timer();
                    this.queueTimer.Interval = 1000;
                    this.queueTimer.Tick += QueueTimer_Tick;
                    this.queueTimer.Start();
                }

                public Boolean Add(Job job) {
                    if (this.queue.Exists(q => q.ContainsKey(job.ProfileName)))
                        return false;
                    else {
                        queue.Add(new Dictionary<string, Job>() { { job.ProfileName, job } });
                        return true;
                    }
                }

                public int Count() {
                    return queue.Count();
                }
                public void Clear() {
                    queue.Clear();
                }

                private void QueueTimer_Tick(object sender, EventArgs e) {
                    log.UltraFine("Sync queue size: " + queue.Count());

                    if (queue.Count() == 0) return;
                    if (Engine.Instance.ActiveProfile != null) return;

                    try {
                        Job job = queue[0].Values.First();
                        queue.RemoveAt(0);
                        log.Info("Scheduled sync started (" + job.RequestedBy + ") for profile: " + job.ProfileName);
                        Engine.Instance.ActiveProfile = job.Profile;
                        Engine.Instance.Start(manualIgnition: false, updateSyncSchedule: (job.RequestedBy == "AutoSyncTimer"));
                    } catch (System.Exception ex) {
                        OGCSexception.Analyse("Scheduled sync encountered a problem.", ex, true);
                    }
                }
            }
        }

        private static readonly ILog log = LogManager.GetLogger(typeof(Engine));

        private static Engine instance;
        public static Engine Instance {
            get {
                if (instance == null) instance = new Engine();
                return instance;
            }
            set {
                instance = value;
            }
        }

        public Engine() {
            this.JobQueue = new Job.Queue();
        }
        public Job.Queue JobQueue { get; protected set; }

        private Object activeProfile;
        /// <summary>
        /// The profile currently set to be synced, either manually from GUI settings or scheduled from a Timer.
        /// </summary>
        public Object ActiveProfile { 
            get { return activeProfile; } 
            set { 
                activeProfile = value;
                log.Debug("ActiveProfile set to: " + Settings.Profile.Name(activeProfile));
            }
        }

        /// <summary>
        /// Get the earliest upcoming sync time
        /// </summary>
        public DateTime? NextSyncDate { get {
                DateTime? retVal = null;
                foreach (SettingsStore.Calendar cal in Settings.Instance.Calendars) {
                    if (cal.OgcsTimer.NextSyncDate != null)
                        retVal = cal.OgcsTimer.NextSyncDate < (retVal ?? DateTime.MaxValue) ? cal.OgcsTimer.NextSyncDate : retVal;
                }
                return retVal;
            }
        }

        /// <summary>The time the current sync started</summary>
        public DateTime SyncStarted { get; protected set; }

        public AbortableBackgroundWorker bwSync { get; private set; }
        public Boolean SyncingNow {
            get {
                if (bwSync == null) return false;
                else return bwSync.IsBusy;
            }
        }
        public Boolean CancellationPending {
            get {
                return (bwSync != null && bwSync.CancellationPending);
            }
        }
        public Boolean ManualForceCompare = false;
        public enum SyncResult {
            OK,
            Fail,
            Abandon,
            AutoRetry,
            ReconnectThenRetry,
            UserCancelled
        }

        public void Sync_Requested(object sender = null, EventArgs e = null) {
            ManualForceCompare = false;
            if (sender != null && sender.GetType().ToString().EndsWith("Timer")) { //Automated sync
                Forms.Main.Instance.NotificationTray.UpdateItem("delayRemove", enabled: false);
                Timer aTimer = sender as Timer;
                Object timerProfile = null;

                if (aTimer.Tag.ToString() == "PushTimer" && aTimer is PushSyncTimer)
                    timerProfile = (aTimer as PushSyncTimer).owningProfile;
                else if (aTimer.Tag.ToString() == "AutoSyncTimer" && aTimer is SyncTimer)
                    timerProfile = (aTimer as SyncTimer).owningProfile;

                if (JobQueue.Add(new Job(aTimer.Tag.ToString(), timerProfile))) {
                    aTimer.Stop();
                } else {
                    log.Warn("Sync of profile '" + Settings.Profile.Name(timerProfile) + "' requested by " + aTimer.Tag.ToString() + " already previously queued.");
                }

            } else { //Manual sync
                if (Forms.Main.Instance.bSyncNow.Text == "Start Sync" || Forms.Main.Instance.bSyncNow.Text == "Start Full Sync") {
                    log.Info("Manual sync requested.");
                    if (SyncingNow) {
                        log.Info("Already busy syncing, cannot accept another sync request.");
                        OgcsMessageBox.Show("A sync is already running. Please wait for it to complete and then try again.", "Sync already running", MessageBoxButtons.OK, MessageBoxIcon.Hand);
                        return;
                    }
                    if (Control.ModifierKeys == Keys.Shift) {
                        if (Forms.Main.Instance.ActiveCalendarProfile.SyncDirection == Direction.Bidirectional) {
                            OgcsMessageBox.Show("Forcing a full sync is not allowed whilst in 2-way sync mode.\r\nPlease temporarily chose a direction to sync in first.",
                                "2-way full sync not allowed", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                            return;
                        }
                        log.Info("Shift-click has forced a compare of all items");
                        ManualForceCompare = true;
                    }
                    this.ActiveProfile = Forms.Main.Instance.ActiveCalendarProfile;
                    Start(manualIgnition: true, updateSyncSchedule: false);

                } else if (Forms.Main.Instance.bSyncNow.Text == "Stop Sync") {
                    GoogleOgcs.Calendar.Instance.Authenticator.CancelTokenSource.Cancel();
                    if (!SyncingNow) return;

                    if (!bwSync.CancellationPending) {
                        Forms.Main.Instance.Console.Update("Sync cancellation requested.", Console.Markup.warning);
                        bwSync.CancelAsync();
                    } else {
                        Forms.Main.Instance.Console.Update("Repeated cancellation requested - forcefully aborting sync!", Console.Markup.warning);
                        AbortSync();
                    }
                    if (this.JobQueue.Count() > 0) {
                        if (OgcsMessageBox.Show("There are " + this.JobQueue.Count() + " sync(s) still queued to run. Would you like to cancel these too?",
                            "Clear queued syncs?", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes) {
                            log.Info("User requested clear down of sync queue.");
                            this.JobQueue.Clear();
                        }
                    }
                }
            }
        }

        public void AbortSync() {
            try {
                bwSync.Abort();
                bwSync.Dispose();
                bwSync = null;
            } catch (System.Exception ex) {
                OGCSexception.Analyse(ex);
            } finally {
                log.Warn("Sync thread forcefully aborted!");
            }
        }

        private void Start(Boolean manualIgnition, Boolean updateSyncSchedule) {
            if (Settings.Profile.GetType(this.ActiveProfile) == Settings.Profile.Type.Calendar) {
                Forms.Main.Instance.NotificationTray.IconAnimator.Start();
                Forms.Main.Instance.NotificationTray.ShowBubbleInfo((manualIgnition ? "S" : "Autos") + "yncing calendars: " + (this.ActiveProfile as SettingsStore.Calendar).SyncDirection.Name + "...");
                Sync.Engine.Calendar.Instance.Profile = this.ActiveProfile as SettingsStore.Calendar;
                Sync.Engine.Calendar.Instance.StartSync(manualIgnition, updateSyncSchedule);
                Forms.Main.Instance.NotificationTray.IconAnimatorStop();
            }
        }

        #region Compare Event Attributes
        public static Boolean CompareAttribute(String attrDesc, Direction fromTo, String googleAttr, String outlookAttr, StringBuilder sb, ref int itemModified) {
            if (googleAttr == null) googleAttr = "";
            if (outlookAttr == null) outlookAttr = "";
            //Truncate long strings
            String googleAttr_stub = ((googleAttr.Length > 50) ? googleAttr.Substring(0, 47) + "..." : googleAttr).Replace("\r\n", " ");
            String outlookAttr_stub = ((outlookAttr.Length > 50) ? outlookAttr.Substring(0, 47) + "..." : outlookAttr).Replace("\r\n", " ");
            log.Fine("Comparing " + attrDesc);
            log.UltraFine("Google  attribute: " + googleAttr);
            log.UltraFine("Outlook attribute: " + outlookAttr);
            if (googleAttr != outlookAttr) {
                if (fromTo == Direction.GoogleToOutlook) {
                    sb.AppendLine(attrDesc + ": " + outlookAttr_stub + " => " + googleAttr_stub);
                } else {
                    sb.AppendLine(attrDesc + ": " + googleAttr_stub + " => " + outlookAttr_stub);
                }
                itemModified++;
                log.Fine("Attributes differ.");
                return true;
            }
            return false;
        }
        public static Boolean CompareAttribute(String attrDesc, Direction fromTo, Boolean googleAttr, Boolean outlookAttr, StringBuilder sb, ref int itemModified) {
            log.Fine("Comparing " + attrDesc);
            log.UltraFine("Google  attribute: " + googleAttr);
            log.UltraFine("Outlook attribute: " + outlookAttr);
            if (googleAttr != outlookAttr) {
                if (fromTo == Direction.GoogleToOutlook) {
                    sb.AppendLine(attrDesc + ": " + outlookAttr + " => " + googleAttr);
                } else {
                    sb.AppendLine(attrDesc + ": " + googleAttr + " => " + outlookAttr);
                }
                itemModified++;
                log.Fine("Attributes differ.");
                return true;
            }
            return false;
        }
        public static Boolean CompareAttribute(String attrDesc, Direction fromTo, DateTime googleAttr, DateTime outlookAttr, StringBuilder sb, ref int itemModified) {
            log.Fine("Comparing " + attrDesc);
            log.UltraFine("Google  attribute: " + googleAttr);
            log.UltraFine("Outlook attribute: " + outlookAttr);
            if (googleAttr != outlookAttr) {
                if (fromTo == Direction.GoogleToOutlook) {
                    sb.AppendLine(attrDesc + ": " + outlookAttr + " => " + googleAttr);
                } else {
                    sb.AppendLine(attrDesc + ": " + googleAttr + " => " + outlookAttr);
                }
                itemModified++;
                log.Fine("Attributes differ.");
                return true;
            }
            return false;
        }
        #endregion
    }
}
