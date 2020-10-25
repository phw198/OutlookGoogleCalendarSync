using Google.Apis.Calendar.v3.Data;
using System;
using System.ComponentModel;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using log4net;
using Microsoft.Office.Interop.Outlook;

namespace OutlookGoogleCalendarSync.Sync {
    public class Engine {
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

        public Engine() { }

        /// <summary>The time the current sync started</summary>
        public DateTime SyncStarted { get; protected set; }
        public SyncTimer OgcsTimer;
        public Sync.PushSyncTimer OgcsPushTimer;
        private AbortableBackgroundWorker bwSync;
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
        private enum SyncResult {
            OK,
            Fail,
            Abandon,
            AutoRetry,
            ReconnectThenRetry,
            UserCancelled
        }
        private int consecutiveSyncFails = 0;

        #region Push Sync
        public void RegisterForPushSync() {
            if (Settings.Instance.SyncDirection != Sync.Direction.GoogleToOutlook) {
                log.Debug("Create the timer for the push synchronisation");
                if (OgcsPushTimer == null)
                    OgcsPushTimer = Sync.PushSyncTimer.Instance;
                if (!OgcsPushTimer.Running())
                    OgcsPushTimer.Switch(true);
            }
        }

        public void DeregisterForPushSync() {
            log.Info("Stop monitoring for Outlook appointment changes...");
            if (OgcsPushTimer != null && OgcsPushTimer.Running())
                OgcsPushTimer.Switch(false);
        }
        #endregion

        public void Sync_Requested(object sender = null, EventArgs e = null) {
            ManualForceCompare = false;
            if (sender != null && sender.GetType().ToString().EndsWith("Timer")) { //Automated sync
                Forms.Main.Instance.NotificationTray.UpdateItem("delayRemove", enabled: false);
                if (Forms.Main.Instance.bSyncNow.Text == "Start Sync") {
                    Timer aTimer = sender as Timer;
                    log.Info("Scheduled sync started (" + aTimer.Tag.ToString() + ").");
                    if (aTimer.Tag.ToString() == "PushTimer") Sync.Engine.Instance.Start(updateSyncSchedule: false);
                    else if (aTimer.Tag.ToString() == "AutoSyncTimer") Sync.Engine.Instance.Start(updateSyncSchedule: true);
                } else if (Forms.Main.Instance.bSyncNow.Text == "Stop Sync") {
                    log.Warn("Automated sync triggered whilst previous sync is still running. Ignoring this new request.");
                    if (bwSync == null)
                        log.Debug("Background worker is null somehow?!");
                    else
                        log.Debug("Background worker is busy? A:" + bwSync.IsBusy.ToString());
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
                        if (Settings.Instance.SyncDirection == Direction.Bidirectional) {
                            OgcsMessageBox.Show("Forcing a full sync is not allowed whilst in 2-way sync mode.\r\nPlease temporarily chose a direction to sync in first.",
                                "2-way full sync not allowed", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                            return;
                        }
                        log.Info("Shift-click has forced a compare of all items");
                        ManualForceCompare = true;
                    }
                    Sync.Engine.Instance.Start(updateSyncSchedule: false);

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

        public void Start(Boolean updateSyncSchedule = true) {
            Forms.Main mainFrm = Forms.Main.Instance;
            try {
                this.SyncStarted = DateTime.Now;
                String cacheNextSync = mainFrm.NextSyncVal;

                mainFrm.Console.Clear();

                if (Settings.Instance.UseGoogleCalendar == null || string.IsNullOrEmpty(Settings.Instance.UseGoogleCalendar.Id)) {
                    OgcsMessageBox.Show("You need to select a Google Calendar first on the 'Settings' tab.", "Configuration Required", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

                if (Settings.Instance.MuteClickSounds) Console.MuteClicks(true);

                //Check network availability
                if (!System.Net.NetworkInformation.NetworkInterface.GetIsNetworkAvailable()) {
                    mainFrm.Console.Update("There does not appear to be any network available! Sync aborted.", Console.Markup.warning, notifyBubble: true);
                    setNextSync(false, updateSyncSchedule, cacheNextSync);
                    return;
                }
                //Check if Outlook is Online
                try {
                    if (OutlookOgcs.Calendar.Instance.IOutlook.Offline() && Settings.Instance.AddAttendees) {
                        mainFrm.Console.Update("<p>You have selected to sync attendees but Outlook is currently offline.</p>" +
                            "<p>Either put Outlook online or do not sync attendees.</p>", Console.Markup.error, notifyBubble: true);
                        setNextSync(false, updateSyncSchedule, cacheNextSync);
                        return;
                    }
                } catch (System.Exception ex) {
                    mainFrm.Console.UpdateWithError(null, ex, notifyBubble: true);
                    OGCSexception.Analyse(ex, true);
                    return;
                }
                GoogleOgcs.Calendar.APIlimitReached_attendee = false;
                mainFrm.SyncNote(Forms.Main.SyncNotes.QuotaExhaustedInfo, null, false);
                mainFrm.bSyncNow.Text = "Stop Sync";
                mainFrm.NotificationTray.UpdateItem("sync", "&Stop Sync");

                mainFrm.NextSyncVal = "In progress...";

                StringBuilder sb = new StringBuilder();
                mainFrm.Console.BuildOutput("Sync version: " + System.Windows.Forms.Application.ProductVersion, ref sb);
                mainFrm.Console.BuildOutput((ManualForceCompare ? "Full s" : "S") + "ync started at " + this.SyncStarted.ToString(), ref sb);
                mainFrm.Console.BuildOutput("Syncing from " + Settings.Instance.SyncStart.ToShortDateString() +
                    " to " + Settings.Instance.SyncEnd.ToShortDateString(), ref sb);
                mainFrm.Console.BuildOutput(Settings.Instance.SyncDirection.Name, ref sb);

                //Make the clock emoji show the right time
                int minsPastHour = DateTime.Now.Minute;
                minsPastHour = (int)minsPastHour - (minsPastHour % 30);
                sb.Insert(0, ":clock" + DateTime.Now.ToString("hh").TrimStart('0') + (minsPastHour == 00 ? "" : "30") + ":");
                mainFrm.Console.Update(sb);

                if (Settings.Instance.OutlookPush) DeregisterForPushSync();

                SyncResult syncResult = SyncResult.Fail;
                int failedAttempts = 0;
                Telemetry.TrackSync();

                while ((syncResult == SyncResult.Fail || syncResult == SyncResult.ReconnectThenRetry) && !Forms.Main.Instance.IsDisposed) {
                    if (failedAttempts > (syncResult == SyncResult.ReconnectThenRetry ? 1 : 0)) {
                        if (OgcsMessageBox.Show("The synchronisation failed - check the Sync tab for further details.\r\nDo you want to try again?", "Sync Failed",
                            MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation) == System.Windows.Forms.DialogResult.No) {
                            log.Info("User opted to abandon further syncs.");
                            syncResult = SyncResult.Abandon;
                            break;
                        } else
                            log.Info("User opted to retry sync straight away.");
                    }

                    //Set up a separate thread for the sync to operate in. Keeps the UI responsive.
                    bwSync = new AbortableBackgroundWorker();
                    //Don't need thread to report back. The logbox is updated from the thread anyway.
                    bwSync.WorkerReportsProgress = false;
                    bwSync.WorkerSupportsCancellation = true;

                    //Kick off the sync in the background thread
                    bwSync.DoWork += new DoWorkEventHandler(
                        delegate (object o, DoWorkEventArgs args) {
                            BackgroundWorker b = o as BackgroundWorker;
                            try {
                                syncResult = synchronize();
                            } catch (System.Exception ex) {
                                String hResult = OGCSexception.GetErrorCode(ex);

                                if (ex.Data.Count > 0 && ex.Data.Contains("OGCS")) {
                                    sb = new StringBuilder();
                                    mainFrm.Console.BuildOutput("The following error was encountered during sync:-", ref sb);
                                    mainFrm.Console.BuildOutput(ex.Data["OGCS"].ToString(), ref sb);
                                    mainFrm.Console.Update(sb, (OGCSexception.LoggingAsFail(ex) ? Console.Markup.fail : Console.Markup.error), notifyBubble: true);
                                    if (ex.Data["OGCS"].ToString().Contains("try again")) {
                                        syncResult = SyncResult.AutoRetry;
                                    }

                                } else if (
                                    (ex is System.InvalidCastException && hResult == "0x80004002" && ex.Message.Contains("0x800706BA")) || //The RPC server is unavailable
                                    (ex is System.Runtime.InteropServices.COMException && (
                                        ex.Message.Contains("0x80010108(RPC_E_DISCONNECTED)") || //The object invoked has disconnected from its clients
                                        hResult == "0x800706BE" || //The remote procedure call failed
                                        hResult == "0x800706BA")) //The RPC server is unavailable
                                    ) {
                                    OGCSexception.Analyse(OGCSexception.LogAsFail(ex));
                                    String message = "It looks like Outlook was closed during the sync.";
                                    if (hResult == "0x800706BE") message = "It looks like Outlook has been restarted and is not yet responsive.";
                                    mainFrm.Console.Update(message + "<br/>Will retry syncing in a few seconds...", Console.Markup.fail, newLine: false);
                                    System.Threading.Thread.Sleep(10 * 1000);
                                    syncResult = SyncResult.ReconnectThenRetry;

                                } else {
                                    OGCSexception.Analyse(ex, true);
                                    mainFrm.Console.UpdateWithError(null, ex, notifyBubble: true);
                                    syncResult = SyncResult.Fail;
                                }
                            }
                        }
                    );

                    bwSync.RunWorkerAsync();
                    while (bwSync != null && (bwSync.IsBusy || bwSync.CancellationPending)) {
                        System.Windows.Forms.Application.DoEvents();
                        System.Threading.Thread.Sleep(100);
                    }
                    try {
                        if (syncResult == SyncResult.ReconnectThenRetry) {
                            mainFrm.Console.Update("Attempting to reconnect to Outlook...");
                            try {
                                OutlookOgcs.Calendar.Instance.Reset();
                            } catch (System.Exception ex) {
                                mainFrm.Console.UpdateWithError("A problem was encountered reconnecting to Outlook.<br/>Further syncs aborted.", ex, notifyBubble: true);
                                syncResult = SyncResult.Abandon;
                            }
                        }
                    } finally {
                        failedAttempts += (syncResult != SyncResult.OK) ? 1 : 0;
                    }
                }

                if (syncResult == SyncResult.OK) {
                    Settings.Instance.CompletedSyncs++;
                    consecutiveSyncFails = 0;
                    mainFrm.Console.Update("Sync finished!", Console.Markup.checkered_flag);
                } else if (syncResult == SyncResult.AutoRetry) {
                    consecutiveSyncFails++;
                    mainFrm.Console.Update("Sync encountered a problem and did not complete successfully.<br/>" + consecutiveSyncFails + " consecutive syncs failed.", Console.Markup.error, notifyBubble: true);
                    //***Simplify this one settings profiles in place
                    if (!("Inactive;Push Sync Active;In progress...".Contains(Forms.Main.Instance.NextSyncVal)) &&
                        DateTime.ParseExact(Forms.Main.Instance.NextSyncVal.Replace(" + Push", ""),
                            System.Globalization.CultureInfo.CurrentCulture.DateTimeFormat.LongDatePattern + " @ " +
                            System.Globalization.CultureInfo.CurrentCulture.DateTimeFormat.LongTimePattern,
                            System.Globalization.CultureInfo.CurrentCulture) > DateTime.Now) {
                        log.Debug("The next sync has already been set (likely through auto retry for new quota at 8AM GMT): " + mainFrm.NextSyncVal);
                        updateSyncSchedule = false;
                        cacheNextSync = mainFrm.NextSyncVal;
                    }
                } else {
                    consecutiveSyncFails += failedAttempts;
                    mainFrm.Console.Update("Sync aborted after " + failedAttempts + " failed attempts!", 
                        new SyncResult[] { SyncResult.UserCancelled, SyncResult.Abandon }.Contains(syncResult) ? Console.Markup.fail : Console.Markup.error);
                }

                setNextSync(syncResult == SyncResult.OK, updateSyncSchedule, cacheNextSync);
                mainFrm.CheckSyncMilestone();

            } finally {
                mainFrm.bSyncNow.Text = "Start Sync";
                mainFrm.NotificationTray.UpdateItem("sync", "&Sync Now");
                if (Settings.Instance.MuteClickSounds) Console.MuteClicks(false);

                if (Settings.Instance.OutlookPush) RegisterForPushSync();

                //Release Outlook reference if GUI not available. 
                //Otherwise, tasktray shows "another program is using outlook" and it doesn't send and receive emails
                OutlookOgcs.Calendar.Disconnect(onlyWhenNoGUI: true);
            }
        }
        
        /// <summary>
        /// Set the next scheduled sync
        /// </summary>
        /// <param name="syncedOk">The result of the current sync</param>
        /// <param name="updateSyncSchedule">Whether to calculate the next sync time or not</param>
        /// <param name="cacheNextSync">The time previously calculated for the next sync when the current one started.
        /// If updateSyncSchedule is false, this value persists.</param>
        private void setNextSync(Boolean syncedOk, Boolean updateSyncSchedule, String cacheNextSync) {
            if (syncedOk) {
                Forms.Main.Instance.LastSyncVal = this.SyncStarted.ToLongDateString() + " @ " + this.SyncStarted.ToLongTimeString();
                Settings.Instance.LastSyncDate = this.SyncStarted;
            }
            if (!updateSyncSchedule) {
                Forms.Main.Instance.NextSyncVal = cacheNextSync;
            } else {
                if (syncedOk) {
                    OgcsTimer.LastSyncDate = this.SyncStarted;
                    OgcsTimer.SetNextSync();
                } else {
                    if (Settings.Instance.SyncInterval != 0) {
                        Forms.Main.Instance.Console.Update("Another sync has been scheduled to automatically run in " + Forms.Main.Instance.MinSyncMinutes + " minutes time.");
                        OgcsTimer.SetNextSync(Forms.Main.Instance.MinSyncMinutes, fromNow: true);
                    }
                }
            }
            Forms.Main.Instance.bSyncNow.Enabled = true;
            if (OgcsPushTimer != null)
                OgcsPushTimer.ResetLastRun(); //Reset Push flag regardless of success (don't want it trying every 2 mins)
        }

        private void skipCorruptedItem(ref List<AppointmentItem> outlookEntries, AppointmentItem cai, String errMsg) {
            try {
                String itemSummary = OutlookOgcs.Calendar.GetEventSummary(cai);
                if (string.IsNullOrEmpty(itemSummary)) {
                    try {
                        itemSummary = cai.Start.Date.ToShortDateString() + " => " + cai.Subject;
                    } catch {
                        itemSummary = cai.Subject;
                    }
                }
                Forms.Main.Instance.Console.Update("<p>" + itemSummary + "</p><p>There is probem with this item - it will not be synced.</p><p>" + errMsg + "</p>",
                    Console.Markup.warning, logit: true);

            } finally {
                log.Debug("Outlook object removed.");
                outlookEntries.Remove(cai);
            }
        }

        private SyncResult synchronize() {
            Console console = Forms.Main.Instance.Console;
            console.Update("Finding Calendar Entries", Console.Markup.mag_right, newLine: false);

            List<AppointmentItem> outlookEntries = null;
            List<Event> googleEntries = null;
            GoogleOgcs.Calendar.Instance.EphemeralProperties.Clear();
            OutlookOgcs.Calendar.Instance.EphemeralProperties.Clear();
            try {
                #region Read Outlook items
                console.Update("Scanning Outlook calendar...");
                outlookEntries = OutlookOgcs.Calendar.Instance.GetCalendarEntriesInRange(false);
                console.Update(outlookEntries.Count + " Outlook calendar entries found.", Console.Markup.sectionEnd, newLine: false);

                if (CancellationPending) return SyncResult.UserCancelled;
                #endregion

                #region Read Google items
                console.Update("Scanning Google calendar...");
                try {
                    GoogleOgcs.Calendar.Instance.GetSettings();
                    googleEntries = GoogleOgcs.Calendar.Instance.GetCalendarEntriesInRange();
                } catch (AggregateException agex) {
                    OGCSexception.AnalyseAggregate(agex);
                } catch (Google.Apis.Auth.OAuth2.Responses.TokenResponseException ex) {
                    OGCSexception.AnalyseTokenResponse(ex, false);
                    return SyncResult.Fail;
                } catch (System.Net.Http.HttpRequestException ex) {
                    if (ex.InnerException != null && ex.InnerException is System.Net.WebException && OGCSexception.GetErrorCode(ex.InnerException) == "0x80131509") {
                        ex = OGCSexception.LogAsFail(ex) as System.Net.Http.HttpRequestException;
                    }
                    OGCSexception.Analyse(ex);
                    ex.Data.Add("OGCS", "ERROR: Unable to connect to the Google calendar. Please try again. " + ((ex.InnerException != null) ? ex.InnerException.Message : ex.Message));
                    throw;
                } catch (System.ApplicationException ex) {
                    if (ex.InnerException != null && ex.InnerException is Google.GoogleApiException &&
                        (ex.Message.Contains("daily Calendar quota has been exhausted") || OGCSexception.GetErrorCode(ex.InnerException) == "0x80131500")) {
                        Forms.Main.Instance.Console.Update(ex.Message, Console.Markup.warning);
                        DateTime newQuota = DateTime.UtcNow.Date.AddHours(8);
                        String tryAfter = "08:00 GMT.";
                        if (newQuota < DateTime.UtcNow) {
                            newQuota = newQuota.AddDays(1);
                            tryAfter = newQuota.ToLocalTime().ToShortTimeString() + " tomorrow.";
                        } else
                            tryAfter = newQuota.ToLocalTime().ToShortTimeString() + ".";

                        //Already rescheduled to run again once new quota available, so just set to retry.
                        ex.Data.Add("OGCS", "ERROR: Unable to connect to the Google calendar" +
                            (Settings.Instance.SyncInterval == 0 ? ". Please try again after " + tryAfter : ", but OGCS is all set to automatically try again after "+ tryAfter));
                        OGCSexception.LogAsFail(ref ex);
                    }
                    throw;
                } catch (System.Exception ex) {
                    OGCSexception.Analyse(ex);
                    ex.Data.Add("OGCS", "ERROR: Unable to connect to the Google calendar.");
                    if (OGCSexception.GetErrorCode(ex) == "0x8013153B") //ex.Message == "A task was canceled." - likely timed out.
                        ex.Data["OGCS"] += " Please try again.";
                    throw;
                }
                Recurrence.Instance.SeparateGoogleExceptions(googleEntries);
                if (Recurrence.Instance.GoogleExceptions != null && Recurrence.Instance.GoogleExceptions.Count > 0) {
                    console.Update(googleEntries.Count + " Google calendar entries found.");
                    console.Update(Recurrence.Instance.GoogleExceptions.Count + " are exceptions to recurring events.", Console.Markup.sectionEnd, newLine: false);
                } else
                    console.Update(googleEntries.Count + " Google calendar entries found.", Console.Markup.sectionEnd, newLine: false);

                if (CancellationPending) return SyncResult.UserCancelled;
                #endregion

                #region Normalise recurring items in sync window
                console.Update("Total inc. recurring items spanning sync date range...");
                //Outlook returns recurring items that span the sync date range, Google doesn't
                //So check for master Outlook items occurring before sync date range, and retrieve Google equivalent
                for (int o = outlookEntries.Count - 1; o >= 0; o--) {
                    log.Fine("Processing " + (o + 1) + "/" + outlookEntries.Count);
                    AppointmentItem ai = null;
                    try {
                        if (outlookEntries[o] is AppointmentItem) ai = outlookEntries[o];
                        else if (outlookEntries[o] is MeetingItem) {
                            log.Info("Calendar object appears to be a MeetingItem, so retrieving associated AppointmentItem.");
                            MeetingItem mi = outlookEntries[o] as MeetingItem;
                            outlookEntries[o] = mi.GetAssociatedAppointment(false);
                            ai = outlookEntries[o];
                        } else {
                            log.Warn("Unknown calendar object type - cannot sync it.");
                            skipCorruptedItem(ref outlookEntries, outlookEntries[o], "Unknown object type.");
                            continue;
                        }
                    } catch (System.Exception ex) {
                        log.Warn("Encountered error casting calendar object to AppointmentItem - cannot sync it.");
                        log.Debug(ex.Message);
                        skipCorruptedItem(ref outlookEntries, outlookEntries[o], ex.Message);
                        ai = (AppointmentItem)OutlookOgcs.Calendar.ReleaseObject(ai);
                        continue;
                    }

                    //Now let's check there's a start/end date - sometimes it can be missing, even though this shouldn't be possible!!
                    String entryID;
                    try {
                        entryID = outlookEntries[o].EntryID;
                        DateTime checkDates = ai.Start;
                        checkDates = ai.End;
                    } catch (System.Exception ex) {
                        log.Warn("Calendar item does not have a proper date range - cannot sync it.");
                        log.Debug(ex.Message);
                        skipCorruptedItem(ref outlookEntries, outlookEntries[o], ex.Message);
                        ai = (AppointmentItem)OutlookOgcs.Calendar.ReleaseObject(ai);
                        continue;
                    }

                    if (ai.IsRecurring && ai.Start.Date < Settings.Instance.SyncStart && ai.End.Date < Settings.Instance.SyncStart) {
                        //We won't bother getting Google master event if appointment is yearly reoccurring in a month outside of sync range
                        //Otherwise, every sync, the master event will have to be retrieved, compared, concluded nothing's changed (probably) = waste of API calls
                        RecurrencePattern oPattern = ai.GetRecurrencePattern();
                        try {
                            if (oPattern.RecurrenceType.ToString().Contains("Year")) {
                                log.Fine("It's an annual event.");
                                Boolean monthInSyncRange = false;
                                DateTime monthMarker = Settings.Instance.SyncStart;
                                while (Convert.ToInt32(monthMarker.ToString("yyyyMM")) <= Convert.ToInt32(Settings.Instance.SyncEnd.ToString("yyyyMM"))
                                    && !monthInSyncRange) {
                                    if (monthMarker.Month == ai.Start.Month) {
                                        monthInSyncRange = true;
                                    }
                                    monthMarker = monthMarker.AddMonths(1);
                                }
                                log.Fine("Found it to be " + (monthInSyncRange ? "inside" : "outside") + " sync range.");
                                if (!monthInSyncRange) { outlookEntries.Remove(ai); log.Fine("Removed."); continue; }
                            }
                            Event masterEv = Recurrence.Instance.GetGoogleMasterEvent(ai);
                            if (masterEv != null && masterEv.Status != "cancelled") {
                                Event cachedEv = googleEntries.Find(x => x.Id == masterEv.Id);
                                if (cachedEv == null) {
                                    googleEntries.Add(masterEv);
                                } else {
                                    if (masterEv.Updated > cachedEv.Updated) {
                                        log.Debug("Refreshing cache for this Event.");
                                        googleEntries.Remove(cachedEv);
                                        googleEntries.Add(masterEv);
                                    }
                                }
                            }
                        } catch (System.Exception ex) {
                            console.Update("Failed to retrieve master for Google recurring event outside of sync range.", OGCSexception.LoggingAsFail(ex) ? Console.Markup.fail : Console.Markup.error);
                            throw;
                        } finally {
                            oPattern = (RecurrencePattern)OutlookOgcs.Calendar.ReleaseObject(oPattern);
                        }
                    }
                    //Completely dereference object and retrieve afresh (due to GetRecurrencePattern earlier) 
                    ai = (AppointmentItem)OutlookOgcs.Calendar.ReleaseObject(ai);
                    OutlookOgcs.Calendar.Instance.IOutlook.GetAppointmentByID(entryID, out ai);
                    outlookEntries[o] = ai;
                }
                console.Update("Outlook " + outlookEntries.Count + ", Google " + googleEntries.Count);

                GoogleOgcs.Calendar.ExportToCSV("Outputting all Events.", "google_events.csv", googleEntries);
                OutlookOgcs.Calendar.ExportToCSV("Outputting all Appointments.", "outlook_appointments.csv", outlookEntries);
                if (CancellationPending) return SyncResult.UserCancelled;
                #endregion

                Boolean success = true;
                String bubbleText = "";
                if (Settings.Instance.ExtirpateOgcsMetadata) {
                    return extirpateCustomProperties(outlookEntries, googleEntries);
                }

                if (Settings.Instance.SyncDirection != Direction.GoogleToOutlook) {
                    success = outlookToGoogle(outlookEntries, googleEntries, ref bubbleText);
                    if (CancellationPending) return SyncResult.UserCancelled;
                }
                if (!success) return SyncResult.Fail;
                if (Settings.Instance.SyncDirection != Direction.OutlookToGoogle) {
                    if (bubbleText != "") bubbleText += "\r\n";
                    success = googleToOutlook(googleEntries, outlookEntries, ref bubbleText);
                    if (CancellationPending) return SyncResult.UserCancelled;
                }
                if (bubbleText != "") Forms.Main.Instance.NotificationTray.ShowBubbleInfo(bubbleText);

                return SyncResult.OK;
            } finally {
                if (outlookEntries != null) {
                    for (int o = outlookEntries.Count() - 1; o >= 0; o--) {
                        outlookEntries[o] = (AppointmentItem)OutlookOgcs.Calendar.ReleaseObject(outlookEntries[o]);
                        outlookEntries.RemoveAt(o);
                    }
                }
            }
        }

        private Boolean outlookToGoogle(List<AppointmentItem> outlookEntries, List<Event> googleEntries, ref String bubbleText) {
            log.Debug("Synchronising from Outlook to Google.");
            if (Settings.Instance.SyncDirection == Sync.Direction.Bidirectional)
                Forms.Main.Instance.Console.Update("Syncing " + Sync.Direction.OutlookToGoogle.Name, Console.Markup.syncDirection, newLine: false);

            //  Make copies of each list of events (Not strictly needed)
            List<AppointmentItem> googleEntriesToBeCreated = new List<AppointmentItem>(outlookEntries);
            List<Event> googleEntriesToBeDeleted = new List<Event>(googleEntries);
            Dictionary<AppointmentItem, Event> entriesToBeCompared = new Dictionary<AppointmentItem, Event>();

            Console console = Forms.Main.Instance.Console;

            try {
                Forms.Main.Instance.Console.Update("Checking for orphaned items", verbose: true);
                GoogleOgcs.Calendar.Instance.ReclaimOrphanCalendarEntries(ref googleEntriesToBeDeleted, ref outlookEntries);
                if (CancellationPending) return false;
            } catch (System.Exception) {
                console.Update("Unable to reclaim orphan calendar entries in Google calendar.", Console.Markup.error);
                throw;
            }

            DateTime timeSection = DateTime.Now;
            try {
                GoogleOgcs.Calendar.Instance.IdentifyEventDifferences(ref googleEntriesToBeCreated, ref googleEntriesToBeDeleted, entriesToBeCompared);
                if (CancellationPending) return false;
            } catch (System.Exception) {
                console.Update("Unable to identify differences in Google calendar.", Console.Markup.error);
                throw;
            }
            TimeSpan sectionDuration = DateTime.Now - timeSection;
            if (sectionDuration.TotalSeconds > 30) {
                log.Warn("That step took a long time! Issue #599");
                Telemetry.Send(Analytics.Category.ogcs, Analytics.Action.debug, "Duration;Google.IdentifyEventDifferences=" + sectionDuration.TotalSeconds);
            }

            StringBuilder sb = new StringBuilder();
            console.BuildOutput(googleEntriesToBeDeleted.Count + " Google calendar entries to be deleted.", ref sb, false);
            console.BuildOutput(googleEntriesToBeCreated.Count + " Google calendar entries to be created.", ref sb, false);
            console.BuildOutput(entriesToBeCompared.Count + " calendar entries to be compared.", ref sb, false);
            console.Update(sb, Console.Markup.info, logit: true);

            //Protect against very first syncs which may trample pre-existing non-Outlook events in Google
            if (!Settings.Instance.DisableDelete && !Settings.Instance.ConfirmOnDelete &&
                googleEntriesToBeDeleted.Count == googleEntries.Count && googleEntries.Count > 0) {
                if (OgcsMessageBox.Show("All Google events are going to be deleted. Do you want to allow this?" +
                    "\r\nNote, " + googleEntriesToBeCreated.Count + " events will then be created.", "Confirm mass deletion",
                    MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation) == DialogResult.No) {
                    googleEntriesToBeDeleted = new List<Event>();
                }
            }

            int entriesUpdated = 0;
            try {
                #region Delete Google Entries
                if (googleEntriesToBeDeleted.Count > 0) {
                    console.Update("Deleting " + googleEntriesToBeDeleted.Count + " Google calendar entries", Console.Markup.h2, newLine: false);
                    try {
                        GoogleOgcs.Calendar.Instance.DeleteCalendarEntries(googleEntriesToBeDeleted);
                    } catch (UserCancelledSyncException ex) {
                        log.Info(ex.Message);
                        return false;
                    } catch (System.Exception ex) {
                        console.UpdateWithError("Unable to delete obsolete entries in Google calendar.", ex);
                        throw;
                    }
                    log.Info("Done.");
                }

                if (CancellationPending) return false;
                #endregion

                #region Create Google Entries
                if (googleEntriesToBeCreated.Count > 0) {
                    console.Update("Creating " + googleEntriesToBeCreated.Count + " Google calendar entries", Console.Markup.h2, newLine: false);
                    try {
                        GoogleOgcs.Calendar.Instance.CreateCalendarEntries(googleEntriesToBeCreated);
                    } catch (UserCancelledSyncException ex) {
                        log.Info(ex.Message);
                        return false;
                    } catch (System.Exception ex) {
                        console.UpdateWithError("Unable to add new entries into the Google Calendar.", ex);
                        throw;
                    }
                    log.Info("Done.");
                }

                if (CancellationPending) return false;
                #endregion

                #region Update Google Entries
                if (entriesToBeCompared.Count > 0) {
                    console.Update("Comparing " + entriesToBeCompared.Count + " existing Google calendar entries", Console.Markup.h2, newLine: false);
                    try {
                        GoogleOgcs.Calendar.Instance.UpdateCalendarEntries(entriesToBeCompared, ref entriesUpdated);
                    } catch (UserCancelledSyncException ex) {
                        log.Info(ex.Message);
                        return false;
                    } catch (System.Exception ex) {
                        console.UpdateWithError("Unable to update existing entries in the Google calendar.", ex);
                        throw;
                    }
                    console.Update(entriesUpdated + " entries updated.");
                }

                if (CancellationPending) return false;
                #endregion

            } finally {
                bubbleText = "Google: " + googleEntriesToBeCreated.Count + " created; " +
                    googleEntriesToBeDeleted.Count + " deleted; " + entriesUpdated + " updated";

                if (Settings.Instance.SyncDirection == Direction.OutlookToGoogle) {
                    while (entriesToBeCompared.Count() > 0) {
                        OutlookOgcs.Calendar.ReleaseObject(entriesToBeCompared.Keys.Last());
                        entriesToBeCompared.Remove(entriesToBeCompared.Keys.Last());
                    }
                }
            }
            return true;
        }

        private Boolean googleToOutlook(List<Event> googleEntries, List<AppointmentItem> outlookEntries, ref String bubbleText) {
            log.Debug("Synchronising from Google to Outlook.");
            if (Settings.Instance.SyncDirection == Sync.Direction.Bidirectional)
                Forms.Main.Instance.Console.Update("Syncing " + Sync.Direction.GoogleToOutlook.Name, Console.Markup.syncDirection, newLine: false);

            List<Event> outlookEntriesToBeCreated = new List<Event>(googleEntries);
            List<AppointmentItem> outlookEntriesToBeDeleted = new List<AppointmentItem>(outlookEntries);
            Dictionary<AppointmentItem, Event> entriesToBeCompared = new Dictionary<AppointmentItem, Event>();

            Console console = Forms.Main.Instance.Console;

            try {
                OutlookOgcs.Calendar.Instance.ReclaimOrphanCalendarEntries(ref outlookEntriesToBeDeleted, ref outlookEntriesToBeCreated);
                if (CancellationPending) return false;
            } catch (System.Exception) {
                console.Update("Unable to reclaim orphan calendar entries in Outlook calendar.", Console.Markup.error);
                throw;
            }
            try {
                OutlookOgcs.Calendar.IdentifyEventDifferences(ref outlookEntriesToBeCreated, ref outlookEntriesToBeDeleted, entriesToBeCompared);
                if (CancellationPending) return false;
            } catch (System.Exception) {
                console.Update("Unable to identify differences in Outlook calendar.", Console.Markup.error);
                throw;
            }

            StringBuilder sb = new StringBuilder();
            console.BuildOutput(outlookEntriesToBeDeleted.Count + " Outlook calendar entries to be deleted.", ref sb, false);
            console.BuildOutput(outlookEntriesToBeCreated.Count + " Outlook calendar entries to be created.", ref sb, false);
            console.BuildOutput(entriesToBeCompared.Count + " calendar entries to be compared.", ref sb, false);
            console.Update(sb, Console.Markup.info, logit: true);

            //Protect against very first syncs which may trample pre-existing non-Google events in Outlook
            if (!Settings.Instance.DisableDelete && !Settings.Instance.ConfirmOnDelete &&
                outlookEntriesToBeDeleted.Count == outlookEntries.Count && outlookEntries.Count > 0) {
                if (OgcsMessageBox.Show("All Outlook events are going to be deleted. Do you want to allow this?" +
                    "\r\nNote, " + outlookEntriesToBeCreated.Count + " events will then be created.", "Confirm mass deletion",
                    MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation) == DialogResult.No) {

                    while (outlookEntriesToBeDeleted.Count() > 0) {
                        OutlookOgcs.Calendar.ReleaseObject(outlookEntriesToBeDeleted.Last());
                        outlookEntriesToBeDeleted.Remove(outlookEntriesToBeDeleted.Last());
                    }
                }
            }

            int entriesUpdated = 0;
            try {
                #region Delete Outlook Entries
                if (outlookEntriesToBeDeleted.Count > 0) {
                    console.Update("Deleting " + outlookEntriesToBeDeleted.Count + " Outlook calendar entries", Console.Markup.h2, newLine: false);
                    try {
                        OutlookOgcs.Calendar.Instance.DeleteCalendarEntries(outlookEntriesToBeDeleted);
                    } catch (UserCancelledSyncException ex) {
                        log.Info(ex.Message);
                        return false;
                    } catch (System.Exception) {
                        console.Update("Unable to delete obsolete entries in Google calendar.", Console.Markup.error);
                        throw;
                    }
                    log.Info("Done.");
                }

                if (CancellationPending) return false;
                #endregion

                #region Create Outlook Entries
                if (outlookEntriesToBeCreated.Count > 0) {
                    console.Update("Creating " + outlookEntriesToBeCreated.Count + " Outlook calendar entries", Console.Markup.h2, newLine: false);
                    try {
                        OutlookOgcs.Calendar.Instance.CreateCalendarEntries(outlookEntriesToBeCreated);
                    } catch (UserCancelledSyncException ex) {
                        log.Info(ex.Message);
                        return false;
                    } catch (System.Exception) {
                        console.Update("Unable to add new entries into the Outlook Calendar.", Console.Markup.error);
                        throw;
                    }
                    log.Info("Done.");
                }

                if (CancellationPending) return false;
                #endregion

                #region Update Outlook Entries
                if (entriesToBeCompared.Count > 0) {
                    console.Update("Comparing " + entriesToBeCompared.Count + " existing Outlook calendar entries", Console.Markup.h2, newLine: false);
                    try {
                        OutlookOgcs.Calendar.Instance.UpdateCalendarEntries(entriesToBeCompared, ref entriesUpdated);
                    } catch (UserCancelledSyncException ex) {
                        log.Info(ex.Message);
                        return false;
                    } catch (System.Exception) {
                        console.Update("Unable to update existing entries in the Outlook calendar.", Console.Markup.error);
                        throw;
                    }
                    console.Update(entriesUpdated + " entries updated.");
                }

                if (CancellationPending) return false;
                #endregion

            } finally {
                bubbleText += "Outlook: " + outlookEntriesToBeCreated.Count + " created; " +
                    outlookEntriesToBeDeleted.Count + " deleted; " + entriesUpdated + " updated";

                while (outlookEntriesToBeCreated.Count() > 0) {
                    OutlookOgcs.Calendar.ReleaseObject(outlookEntriesToBeCreated.Last());
                    outlookEntriesToBeCreated.Remove(outlookEntriesToBeCreated.Last());
                }
                while (outlookEntriesToBeDeleted.Count() > 0) {
                    OutlookOgcs.Calendar.ReleaseObject(outlookEntriesToBeDeleted.Last());
                    outlookEntriesToBeDeleted.Remove(outlookEntriesToBeDeleted.Last());
                }
                while (entriesToBeCompared.Count() > 0) {
                    OutlookOgcs.Calendar.ReleaseObject(entriesToBeCompared.Keys.Last());
                    entriesToBeCompared.Remove(entriesToBeCompared.Keys.Last());
                }
            }
            return true;
        }

        private SyncResult extirpateCustomProperties(List<AppointmentItem> outlookEntries, List<Event> googleEntries) {
            SyncResult returnVal = SyncResult.Fail;
            try {
                Forms.Main.Instance.Console.Update("Cleansing OGCS metadata from Outlook items...", Console.Markup.h2, newLine: false);
                for (int o = 0; o < outlookEntries.Count; o++) {
                    AppointmentItem ai = null;
                    try {
                        ai = outlookEntries[o];
                        OutlookOgcs.CustomProperty.LogProperties(ai, log4net.Core.Level.Debug);
                        if (OutlookOgcs.CustomProperty.Extirpate(ref ai)) {
                            Forms.Main.Instance.Console.Update(OutlookOgcs.Calendar.GetEventSummary(ai), Console.Markup.calendar);
                            ai.Save();
                        }
                    } finally {
                        ai = (AppointmentItem)OutlookOgcs.Calendar.ReleaseObject(ai);
                    }
                    if (CancellationPending) return SyncResult.UserCancelled;
                }

                Forms.Main.Instance.Console.Update("Cleansing OGCS metadata from Google items...", Console.Markup.h2, newLine: false);
                for (int g = 0; g < googleEntries.Count; g++) {
                    Event ev = googleEntries[g];
                    GoogleOgcs.CustomProperty.LogProperties(ev, log4net.Core.Level.Debug);
                    if (GoogleOgcs.CustomProperty.Extirpate(ref ev)) {
                        Forms.Main.Instance.Console.Update(GoogleOgcs.Calendar.GetEventSummary(ev), Console.Markup.calendar);
                        GoogleOgcs.Calendar.Instance.UpdateCalendarEntry_save(ref ev);
                    }
                    if (CancellationPending) return SyncResult.UserCancelled;
                }
                returnVal = SyncResult.OK;
                return returnVal;

            } catch (System.Exception ex) {
                OGCSexception.Analyse("Failed to fully cleanse metadata!", ex);
                Forms.Main.Instance.Console.UpdateWithError(null, ex);
                returnVal = SyncResult.Fail;
                return returnVal;

            } finally {
                if (CancellationPending) {
                    Forms.Main.Instance.Console.Update("Not letting this process run to completion is <b>strongly discouraged</b>.<br>" +
                        "If you are two-way syncing and use OGCS for normal syncing again, unexpected behaviour will ensue.<br>" +
                        "It is recommended to rerun the metadata cleanse to completion.", Console.Markup.warning);
                } else if (returnVal == SyncResult.Fail) {
                    Forms.Main.Instance.Console.Update(
                        "It is recommended to rerun the metadata cleanse to <b>successful completion</b> before using OGCS for normal syncing again.<br>" +
                        "If this is not possible and you wish to continue using OGCS, please " +
                        "<a href='https://github.com/phw198/OutlookGoogleCalendarSync/issues' target='_blank'>raise an issue</a> on the GitHub project.", Console.Markup.warning);
                }
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
