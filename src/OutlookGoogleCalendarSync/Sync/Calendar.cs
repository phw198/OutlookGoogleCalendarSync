using Google.Apis.Calendar.v3.Data;
using log4net;
using Microsoft.Office.Interop.Outlook;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace OutlookGoogleCalendarSync.Sync {
    public partial class Engine {
        protected internal class Calendar {
            private static readonly ILog log = LogManager.GetLogger(typeof(Calendar));

            /// <summary>
            /// The calendar settings profile currently being synced.
            /// </summary>
            public SettingsStore.Calendar Profile { internal set; get; }

            private int consecutiveSyncFails = 0;

            private static Calendar instance;
            public static Calendar Instance {
                get {
                    if (instance == null) instance = new Calendar();
                    return instance;
                }
                set {
                    instance = value;
                }
            }
            public Calendar() {
                Profile = Sync.Engine.Instance.ActiveProfile as SettingsStore.Calendar;
            }

            public void StartSync(Boolean manualIgnition, Boolean updateSyncSchedule = true) {
                Forms.Main mainFrm = Forms.Main.Instance;
                mainFrm.bSyncNow.Text = "Stop Sync";
                mainFrm.NotificationTray.UpdateItem("sync", "&Stop Sync");

                this.Profile.LogSettings();
                try {
                    Sync.Engine.Instance.SyncStarted = DateTime.Now;
                    String cacheNextSync = this.Profile.OgcsTimer.NextSyncDateText;

                    mainFrm.Console.Clear();

                    //Set up a separate thread for the sync to operate in. Keeps the UI responsive.
                    Sync.Engine.Instance.bwSync = new AbortableBackgroundWorker() {
                        //Don't need thread to report back. The logbox is updated from the thread anyway.
                        WorkerReportsProgress = false,
                        WorkerSupportsCancellation = true
                    };

                    if (this.Profile.UseGoogleCalendar == null || string.IsNullOrEmpty(this.Profile.UseGoogleCalendar.Id)) {
                        MessageBox.Show("You need to select a Google Calendar first on the 'Settings' tab.");
                        return;
                    }

                    if (Settings.Instance.MuteClickSounds) Console.MuteClicks(true);

                    //Check network availability
                    if (!System.Net.NetworkInformation.NetworkInterface.GetIsNetworkAvailable()) {
                        mainFrm.Console.Update("There does not appear to be any network available! Sync aborted.", Console.Markup.warning, notifyBubble: true);
                        this.setNextSync(false, updateSyncSchedule);
                        return;
                    }
                    //Check if Outlook is Online
                    try {
                        if (OutlookOgcs.Calendar.Instance.IOutlook.Offline() && this.Profile.AddAttendees) {
                            mainFrm.Console.Update("<p>You have selected to sync attendees but Outlook is currently offline.</p>" +
                                "<p>Either put Outlook online or do not sync attendees.</p>", Console.Markup.error, notifyBubble: true);
                            this.setNextSync(false, updateSyncSchedule);
                            return;
                        }
                    } catch (System.Exception ex) {
                        mainFrm.Console.UpdateWithError(null, ex, notifyBubble: true);
                        OGCSexception.Analyse(ex, true);
                        return;
                    }
                    GoogleOgcs.Calendar.APIlimitReached_attendee = false;

                    this.Profile.OgcsTimer.NextSyncDateText = "In progress...";

                    if (this.Profile.OutlookPush) this.Profile.DeregisterForPushSync();

                    Sync.Engine.SyncResult syncResult = Sync.Engine.SyncResult.Fail;
                    int failedAttempts = 0;
                    Telemetry.TrackSync();

                    while ((syncResult == Sync.Engine.SyncResult.Fail || syncResult == Sync.Engine.SyncResult.ReconnectThenRetry) && !Forms.Main.Instance.IsDisposed) {
                        if (failedAttempts > (syncResult == Sync.Engine.SyncResult.ReconnectThenRetry ? 1 : 0)) {
                            if (OgcsMessageBox.Show("The synchronisation failed - check the Sync tab for further details.\r\nDo you want to try again?", "Sync Failed",
                                MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation) == System.Windows.Forms.DialogResult.No) {
                                log.Info("User opted to abandon further syncs.");
                                syncResult = Sync.Engine.SyncResult.Abandon;
                                break;
                            } else {
                                log.Info("User opted to retry sync straight away.");
                                mainFrm.Console.Clear();
                                Sync.Engine.Instance.bwSync = new AbortableBackgroundWorker() {
                                    //Don't need thread to report back. The logbox is updated from the thread anyway.
                                    WorkerReportsProgress = false,
                                    WorkerSupportsCancellation = true
                                };
                            }
                        }

                        StringBuilder sb = new StringBuilder();
                        mainFrm.Console.BuildOutput("Sync version: " + System.Windows.Forms.Application.ProductVersion, ref sb);
                        mainFrm.Console.BuildOutput("Profile: " + this.Profile._ProfileName, ref sb);
                        mainFrm.Console.BuildOutput((Sync.Engine.Instance.ManualForceCompare ? "Full s" : "S") + "ync started at " + DateTime.Now.ToString(), ref sb);
                        mainFrm.Console.BuildOutput("Syncing from " + this.Profile.SyncStart.ToShortDateString() +
                            " to " + this.Profile.SyncEnd.ToShortDateString(), ref sb);
                        mainFrm.Console.BuildOutput(this.Profile.SyncDirection.Name, ref sb);

                        //Make the clock emoji show the right time
                        int minsPastHour = DateTime.Now.Minute;
                        minsPastHour = (int)minsPastHour - (minsPastHour % 30);
                        sb.Insert(0, ":clock" + DateTime.Now.ToString("hh").TrimStart('0') + (minsPastHour == 00 ? "" : "30") + ":");
                        mainFrm.Console.Update(sb);

                        //Kick off the sync in the background thread
                        Sync.Engine.Instance.bwSync.DoWork += new DoWorkEventHandler(
                            delegate (object o, DoWorkEventArgs args) {
                                BackgroundWorker b = o as BackgroundWorker;
                                try {
                                    syncResult = manualIgnition ? manualSynchronize() : synchronize();
                                } catch (System.Exception ex) {
                                    String hResult = OGCSexception.GetErrorCode(ex);

                                    if (ex.Data.Count > 0 && ex.Data.Contains("OGCS")) {
                                        sb = new StringBuilder();
                                        mainFrm.Console.BuildOutput("The following error was encountered during sync:-", ref sb);
                                        mainFrm.Console.BuildOutput(ex.Data["OGCS"].ToString(), ref sb);
                                        mainFrm.Console.Update(sb, (OGCSexception.LoggingAsFail(ex) ? Console.Markup.fail : Console.Markup.error), notifyBubble: true);
                                        if (ex.Data["OGCS"].ToString().Contains("try again")) {
                                            syncResult = Sync.Engine.SyncResult.AutoRetry;
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
                                        syncResult = SyncResult.ReconnectThenRetry;

                                    } else if (ex is System.Runtime.InteropServices.COMException && 
                                        OGCSexception.GetErrorCode(ex) == "0x80040201") { //The operation failed.  The messaging interfaces have returned an unknown error. If the problem persists, restart Outlook.
                                        mainFrm.Console.Update(ex.Message, Console.Markup.fail, newLine: false);
                                        syncResult = SyncResult.ReconnectThenRetry;

                                    } else {
                                        OGCSexception.Analyse(ex, true);
                                        mainFrm.Console.UpdateWithError(null, ex, notifyBubble: true);
                                        syncResult = SyncResult.Fail;
                                    }
                                }
                            }
                        );

                        Sync.Engine.Instance.bwSync.RunWorkerAsync();
                        while (Sync.Engine.Instance.bwSync != null && (Sync.Engine.Instance.bwSync.IsBusy || Sync.Engine.Instance.bwSync.CancellationPending)) {
                            System.Windows.Forms.Application.DoEvents();
                            System.Threading.Thread.Sleep(100);
                        }
                        try {
                            if (syncResult == SyncResult.ReconnectThenRetry) {
                                mainFrm.Console.Update("Will retry syncing in a few seconds...");
                                DateTime waitUntil = DateTime.Now.AddSeconds(10);
                                while (DateTime.Now < waitUntil) {
                                    System.Windows.Forms.Application.DoEvents();
                                    System.Threading.Thread.Sleep(100);
                                }                                
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
                        this.consecutiveSyncFails = 0;
                        mainFrm.Console.Update("Sync finished!", Console.Markup.checkered_flag);
                        mainFrm.SyncNote(Forms.Main.SyncNotes.QuotaExhaustedInfo, null, false);
                    } else if (syncResult == SyncResult.AutoRetry) {
                        this.consecutiveSyncFails++;
                        mainFrm.Console.Update("Sync encountered a problem and did not complete successfully.<br/>" + this.consecutiveSyncFails + " consecutive syncs failed.", Console.Markup.error, notifyBubble: true);
                        if (this.Profile.OgcsTimer.NextSyncDate != null && this.Profile.OgcsTimer.NextSyncDate > DateTime.Now) {
                            log.Debug("The next sync has already been set (likely through auto retry for new quota at 8AM GMT): " + this.Profile.OgcsTimer.NextSyncDateText);
                            updateSyncSchedule = false;
                        }
                    } else {
                        this.consecutiveSyncFails += failedAttempts;
                        mainFrm.Console.Update("Sync aborted after " + failedAttempts + " failed attempts!", syncResult == Sync.Engine.SyncResult.UserCancelled ? Console.Markup.fail : Console.Markup.error);
                    }

                    this.setNextSync(syncResult == Sync.Engine.SyncResult.OK, updateSyncSchedule);
                    this.Profile.OgcsTimer.CalculateInterval();
                    mainFrm.CheckSyncMilestone();

                } finally {
                    Sync.Engine.Instance.bwSync?.Dispose();
                    Sync.Engine.Instance.bwSync = null;
                    Sync.Engine.Instance.ActiveProfile = null;
                    mainFrm.bSyncNow.Text = "Start Sync";
                    mainFrm.NotificationTray.UpdateItem("sync", "&Sync Now");
                    if (Settings.Instance.MuteClickSounds) Console.MuteClicks(false);

                    this.Profile.OgcsTimer.Start();
                    this.Profile.RegisterForPushSync();

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
            /// If updateSyncSchedule is false, this value persists.</param>
            private void setNextSync(Boolean syncedOk, Boolean updateSyncSchedule) {
                if (syncedOk) {
                    this.Profile.LastSyncDate = Sync.Engine.Instance.SyncStarted;
                }
                if (!updateSyncSchedule) {
                    if (this.Profile.SyncInterval != 0) {
                        this.Profile.OgcsTimer.NextSyncDate = this.Profile.OgcsTimer.NextSyncDate; //Force update of MainForm, if profile displaying
                        this.Profile.OgcsTimer.Activate(true);
                    } else
                        this.Profile.OgcsTimer.NextSyncDateText = "Inactive";
                } else {
                    if (syncedOk) {
                        this.Profile.OgcsTimer.LastSyncDate = Sync.Engine.Instance.SyncStarted;
                        this.Profile.OgcsTimer.SetNextSync();
                    } else {
                        if (this.Profile.SyncInterval != 0) {
                            Forms.Main.Instance.Console.Update("Another sync has been scheduled to automatically run in " + Forms.Main.Instance.MinSyncMinutes + " minutes time.");
                            this.Profile.OgcsTimer.SetNextSync(Forms.Main.Instance.MinSyncMinutes, fromNow: true);
                        }
                    }
                }
                Forms.Main.Instance.bSyncNow.Enabled = true;
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
                    Forms.Main.Instance.Console.Update("<p>" + itemSummary + "</p><p>There is problem with this item - it will not be synced.</p><p>" + errMsg + "</p>",
                        Console.Markup.warning, logit: true);

                } finally {
                    log.Debug("Outlook object removed.");
                    outlookEntries.Remove(cai);
                }
            }

            private SyncResult manualSynchronize() {
                //This function is just a shim to determine how the sync was triggered when looking at the call stack
                return synchronize();
            }

            private SyncResult synchronize() {
                Console console = Forms.Main.Instance.Console;
                console.Update("Finding Calendar Entries", Console.Markup.mag_right, newLine: false);

                List<AppointmentItem> outlookEntries = null;
                List<Event> googleEntries = null;
                if (!GoogleOgcs.Calendar.IsInstanceNull)
                    GoogleOgcs.Calendar.Instance.EphemeralProperties.Clear();
                OutlookOgcs.Calendar.Instance.EphemeralProperties.Clear();

                try {
                    #region Read Outlook items
                    console.Update("Scanning Outlook calendar...");
                    outlookEntries = OutlookOgcs.Calendar.Instance.GetCalendarEntriesInRange(null, false);
                    console.Update(outlookEntries.Count + " Outlook calendar entries found.", Console.Markup.sectionEnd, newLine: false);

                    if (Sync.Engine.Instance.CancellationPending) return SyncResult.UserCancelled;
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
                                (this.Profile.SyncInterval == 0 ? ". Please try again after " + tryAfter : ", but OGCS is all set to automatically try again after " + tryAfter));
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

                    if (Sync.Engine.Instance.CancellationPending) return SyncResult.UserCancelled;
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
                            OGCSexception.Analyse("Encountered error casting calendar object to AppointmentItem - cannot sync it. ExchangeMode=" +
                                OutlookOgcs.Calendar.Instance.IOutlook.ExchangeConnectionMode().ToString(), ex);
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
                        } catch (System.Runtime.InteropServices.COMException ex) {
                            if (OGCSexception.GetErrorCode(ex, 0x0000FFFF) == "0x00004005") { //You must specify a time/hour
                                skipCorruptedItem(ref outlookEntries, outlookEntries[o], ex.Message);
                            } else {
                                //"Your server administrator has limited the number of items you can open simultaneously."
                                //Once we have the error code for above message, need to abort sync - and suggest using cached Exchange mode
                                OGCSexception.Analyse("Calendar item does not have a proper date range - cannot sync it. ExchangeMode=" +
                                    OutlookOgcs.Calendar.Instance.IOutlook.ExchangeConnectionMode().ToString(), ex);
                                skipCorruptedItem(ref outlookEntries, outlookEntries[o], ex.Message);
                            }
                            ai = (AppointmentItem)OutlookOgcs.Calendar.ReleaseObject(ai);
                            continue;
                        }

                        if (ai.IsRecurring && ai.Start.Date < this.Profile.SyncStart && ai.End.Date < this.Profile.SyncStart) {
                            //We won't bother getting Google master event if appointment is yearly reoccurring in a month outside of sync range
                            //Otherwise, every sync, the master event will have to be retrieved, compared, concluded nothing's changed (probably) = waste of API calls
                            RecurrencePattern oPattern = ai.GetRecurrencePattern();
                            try {
                                if (oPattern.RecurrenceType.ToString().Contains("Year")) {
                                    log.Fine("It's an annual event.");
                                    Boolean monthInSyncRange = false;
                                    DateTime monthMarker = this.Profile.SyncStart;
                                    while (Convert.ToInt32(monthMarker.ToString("yyyyMM")) <= Convert.ToInt32(this.Profile.SyncEnd.ToString("yyyyMM"))
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
                    if (Sync.Engine.Instance.CancellationPending) return SyncResult.UserCancelled;
                    #endregion

                    Boolean success = true;
                    String bubbleText = "";
                    if (this.Profile.ExtirpateOgcsMetadata) {
                        return extirpateCustomProperties(outlookEntries, googleEntries);
                    }

                    //Reclaim orphans
                    GoogleOgcs.Calendar.Instance.ReclaimOrphanCalendarEntries(ref googleEntries, ref outlookEntries);
                    if (Sync.Engine.Instance.CancellationPending) return SyncResult.UserCancelled;

                    OutlookOgcs.Calendar.Instance.ReclaimOrphanCalendarEntries(ref outlookEntries, ref googleEntries);
                    if (Sync.Engine.Instance.CancellationPending) return SyncResult.UserCancelled;

                    //Sync
                    if (this.Profile.SyncDirection.Id != Direction.GoogleToOutlook.Id) {
                        success = outlookToGoogle(outlookEntries, googleEntries, ref bubbleText);
                        if (Sync.Engine.Instance.CancellationPending) return SyncResult.UserCancelled;
                    }
                    if (!success) return SyncResult.Fail;
                    if (this.Profile.SyncDirection.Id != Sync.Direction.OutlookToGoogle.Id) {
                        if (bubbleText != "") bubbleText += "\r\n";
                        success = googleToOutlook(googleEntries, outlookEntries, ref bubbleText);
                        if (Sync.Engine.Instance.CancellationPending) return SyncResult.UserCancelled;
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
                if (this.Profile.SyncDirection.Id == Sync.Direction.Bidirectional.Id)
                    Forms.Main.Instance.Console.Update("Syncing " + Sync.Direction.OutlookToGoogle.Name, Console.Markup.syncDirection, newLine: false);

                //  Make copies of each list of events (Not strictly needed)
                List<AppointmentItem> googleEntriesToBeCreated = new List<AppointmentItem>(outlookEntries);
                List<Event> googleEntriesToBeDeleted = new List<Event>(googleEntries);
                Dictionary<AppointmentItem, Event> entriesToBeCompared = new Dictionary<AppointmentItem, Event>();

                Console console = Forms.Main.Instance.Console;

                DateTime timeSection = DateTime.Now;
                try {
                    GoogleOgcs.Calendar.Instance.IdentifyEventDifferences(ref googleEntriesToBeCreated, ref googleEntriesToBeDeleted, entriesToBeCompared);
                    if (Sync.Engine.Instance.CancellationPending) return false;
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
                if (!this.Profile.DisableDelete && !this.Profile.ConfirmOnDelete &&
                    googleEntriesToBeDeleted.Count == googleEntries.Count && googleEntries.Count > 1) {
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

                    if (Sync.Engine.Instance.CancellationPending) return false;
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

                    if (Sync.Engine.Instance.CancellationPending) return false;
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

                    if (Sync.Engine.Instance.CancellationPending) return false;
                    #endregion

                } finally {
                    bubbleText = "Google: " + googleEntriesToBeCreated.Count + " created; " +
                        googleEntriesToBeDeleted.Count + " deleted; " + entriesUpdated + " updated";

                    if (this.Profile.SyncDirection.Id == Direction.OutlookToGoogle.Id) {
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
                if (this.Profile.SyncDirection.Id == Sync.Direction.Bidirectional.Id)
                    Forms.Main.Instance.Console.Update("Syncing " + Sync.Direction.GoogleToOutlook.Name, Console.Markup.syncDirection, newLine: false);

                List<Event> outlookEntriesToBeCreated = new List<Event>(googleEntries);
                List<AppointmentItem> outlookEntriesToBeDeleted = new List<AppointmentItem>(outlookEntries);
                Dictionary<AppointmentItem, Event> entriesToBeCompared = new Dictionary<AppointmentItem, Event>();

                Console console = Forms.Main.Instance.Console;

                try {
                    OutlookOgcs.Calendar.IdentifyEventDifferences(ref outlookEntriesToBeCreated, ref outlookEntriesToBeDeleted, entriesToBeCompared);
                    if (Sync.Engine.Instance.CancellationPending) return false;
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
                if (!this.Profile.DisableDelete && !this.Profile.ConfirmOnDelete &&
                    outlookEntriesToBeDeleted.Count == outlookEntries.Count && outlookEntries.Count > 1) {
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

                    if (Sync.Engine.Instance.CancellationPending) return false;
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

                    if (Sync.Engine.Instance.CancellationPending) return false;
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

                    if (Sync.Engine.Instance.CancellationPending) return false;
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
                Console console = Forms.Main.Instance.Console;
                try {
                    console.Update("Cleansing OGCS metadata from Outlook items...", Console.Markup.h2, newLine: false);
                    for (int o = 0; o < outlookEntries.Count; o++) {
                        AppointmentItem ai = null;
                        try {
                            ai = outlookEntries[o];
                            OutlookOgcs.CustomProperty.LogProperties(ai, log4net.Core.Level.Debug);
                            if (OutlookOgcs.CustomProperty.Extirpate(ref ai)) {
                                console.Update(OutlookOgcs.Calendar.GetEventSummary(ai), Console.Markup.calendar);
                                ai.Save();
                            }
                        } finally {
                            ai = (AppointmentItem)OutlookOgcs.Calendar.ReleaseObject(ai);
                        }
                        if (Sync.Engine.Instance.CancellationPending) return SyncResult.UserCancelled;
                    }

                    console.Update("Cleansing OGCS metadata from Google items...", Console.Markup.h2, newLine: false);
                    for (int g = 0; g < googleEntries.Count; g++) {
                        Event ev = googleEntries[g];
                        GoogleOgcs.CustomProperty.LogProperties(ev, log4net.Core.Level.Debug);
                        if (GoogleOgcs.CustomProperty.Extirpate(ref ev)) {
                            console.Update(GoogleOgcs.Calendar.GetEventSummary(ev), Console.Markup.calendar);
                            GoogleOgcs.Calendar.Instance.UpdateCalendarEntry_save(ref ev);
                        }
                        if (Sync.Engine.Instance.CancellationPending) return SyncResult.UserCancelled;
                    }
                    returnVal = SyncResult.OK;
                    return returnVal;

                } catch (System.Exception ex) {
                    OGCSexception.Analyse("Failed to fully cleanse metadata!", ex);
                    console.UpdateWithError(null, ex);
                    returnVal = SyncResult.Fail;
                    return returnVal;

                } finally {
                    if (Sync.Engine.Instance.CancellationPending) {
                        console.Update("Not letting this process run to completion is <b>strongly discouraged</b>.<br>" +
                            "If you are two-way syncing and use OGCS for normal syncing again, unexpected behaviour will ensue.<br>" +
                            "It is recommended to rerun the metadata cleanse to completion.", Console.Markup.warning);
                    } else if (returnVal == SyncResult.Fail) {
                        console.Update(
                            "It is recommended to rerun the metadata cleanse to <b>successful completion</b> before using OGCS for normal syncing again.<br>" +
                            "If this is not possible and you wish to continue using OGCS, please " +
                            "<a href='https://github.com/phw198/OutlookGoogleCalendarSync/issues' target='_blank'>raise an issue</a> on the GitHub project.", Console.Markup.warning);
                    }
                }
            }
        }
    }
}
