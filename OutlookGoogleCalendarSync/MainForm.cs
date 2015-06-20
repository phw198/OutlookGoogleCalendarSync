using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.IO;
using System.Net;
using System.Windows.Forms;
using Google.Apis.Calendar.v3.Data;
using log4net;
using Microsoft.Office.Interop.Outlook;

namespace OutlookGoogleCalendarSync {
    /// <summary>
    /// Description of MainForm.
    /// </summary>
    public partial class MainForm : Form {
        public static MainForm Instance;
        public ToolTip ToolTips;

        public Timer OgcsPushTimer;
        private Timer ogcsTimer;
        private DateTime lastSyncDate;

        private BackgroundWorker bwSync;
        public Boolean SyncingNow {
            get {
                if (bwSync == null) return false;
                else return bwSync.IsBusy; 
            }
        }
        private static readonly ILog log = LogManager.GetLogger(typeof(MainForm));
        private Rectangle tabAppSettings_background = new Rectangle();

        public MainForm(string startingTab = null) {
            log.Debug("Initialiasing MainForm.");
            InitializeComponent();
            if (startingTab!=null && startingTab=="Help") this.tabApp.SelectedTab = this.tabPage_Help;
            lVersion.Text = lVersion.Text.Replace("{version}", System.Windows.Forms.Application.ProductVersion);

            Instance = this;

            Social.TrackVersion();
            updateGUIsettings();
            Settings.Instance.LogSettings();
            
            log.Debug("Create the timer for the auto synchronisation");
            ogcsTimer = new Timer();
            ogcsTimer.Tag = "AutoSyncTimer";
            ogcsTimer.Tick += new EventHandler(ogcsTimer_Tick);

            #region Tooltips
            //set up tooltips for some controls
            ToolTips = new ToolTip();
            ToolTips.AutoPopDelay = 10000;
            ToolTips.InitialDelay = 500;
            ToolTips.ReshowDelay = 200;
            ToolTips.ShowAlways = true;
            //Outlook
            ToolTips.SetToolTip(cbOutlookCalendars,
                "The Outlook calendar to synchonize with.");
            //Google
            ToolTips.SetToolTip(cbGoogleCalendars,
                "The Google calendar to synchonize with.");
            ToolTips.SetToolTip(btResetGCal,
                "Reset the Google account being used to synchonize with.");
            //Settings
            ToolTips.SetToolTip(tbInterval,
                "Set to zero to disable");
            ToolTips.SetToolTip(rbOutlookAltMB,
                "Only choose this if you need to use an Outlook Calendar that is not in the default mailbox");
            ToolTips.SetToolTip(cbMergeItems,
                "If the destination calendar has pre-existing items, don't delete them");
            ToolTips.SetToolTip(cbOutlookPush,
                "Synchronise changes in Outlook to Google within a few minutes.");
            ToolTips.SetToolTip(cbOfuscate,
                "Mask specified words in calendar item subject.\nTakes effect for new or updated calendar items.");
            ToolTips.SetToolTip(dgObfuscateRegex,
                "All rules are applied using AND logic");
            //Application behaviour
            ToolTips.SetToolTip(cbPortable,
                "For ZIP deployments, store configuration files in the application folder (useful if running from a USB thumb drive.\n" +
                "Default is in your User roaming profile.");
            ToolTips.SetToolTip(cbCreateFiles,
                "If checked, all entries found in Outlook/Google and identified for creation/deletion will be exported \n" +
                "to CSV files in the application's directory (named \"*.csv\"). \n" +
                "Only for debug/diagnostic purposes.");
            ToolTips.SetToolTip(rbProxyIE,
                "If IE settings have been changed, a restart of the Sync application may be required");
            #endregion

            //Refresh synchronizations (last and next)
            lLastSyncVal.Text = lastSyncDate.ToLongDateString() + " - " + lastSyncDate.ToLongTimeString();
            setNextSync(getResyncInterval());

            //Set up listener for Outlook calendar changes
            if (Settings.Instance.OutlookPush) OutlookCalendar.Instance.RegisterForAutoSync();

            //Start in tray?
            if (cbStartInTray.Checked) {
                this.WindowState = FormWindowState.Minimized;
                notifyIcon1.Visible = true;
                this.Hide();
                this.ShowInTaskbar = false;
            }
        }

        private void updateGUIsettings() {
            this.SuspendLayout();
            lastSyncDate = Settings.Instance.LastSyncDate;
            cbVerboseOutput.Checked = Settings.Instance.VerboseOutput;
            #region Outlook box
            gbEWS.Enabled = false;
            if (Settings.Instance.OutlookService == OutlookCalendar.Service.AlternativeMailbox) {
                rbOutlookAltMB.Checked = true;
            } else if (Settings.Instance.OutlookService == OutlookCalendar.Service.EWS) {
                rbOutlookEWS.Checked = true;
                gbEWS.Enabled = true;
            } else {
                rbOutlookDefaultMB.Checked = true;
                ddMailboxName.Enabled = false;
            }
            txtEWSPass.Text = Settings.Instance.EWSpassword;
            txtEWSUser.Text = Settings.Instance.EWSuser;
            txtEWSServerURL.Text = Settings.Instance.EWSserver;

            //Mailboxes the user has access to
            log.Debug("Find Accounts");
            if (OutlookCalendar.Instance.Accounts.Count == 1) {
                rbOutlookAltMB.Enabled = false;
                rbOutlookAltMB.Checked = false;
                ddMailboxName.Enabled = false;
            }
            for (int acc = 1; acc <= OutlookCalendar.Instance.Accounts.Count-1; acc++) {
                String mailbox = OutlookCalendar.Instance.Accounts[acc];
                ddMailboxName.Items.Add(mailbox);
                if (Settings.Instance.MailboxName == mailbox) { ddMailboxName.SelectedIndex = acc - 1; }
            }
            if (ddMailboxName.SelectedIndex == -1 && ddMailboxName.Items.Count > 0) { ddMailboxName.SelectedIndex = 0; }

            log.Debug("List Calendar folders");
            cbOutlookCalendars.SelectedIndexChanged -= cbOutlookCalendar_SelectedIndexChanged;
            cbOutlookCalendars.DataSource = new BindingSource(OutlookCalendar.Instance.CalendarFolders, null);
            cbOutlookCalendars.DisplayMember = "Key";
            cbOutlookCalendars.ValueMember = "Value";
            cbOutlookCalendars.SelectedIndex = -1; //Reset to nothing selected
            cbOutlookCalendars.SelectedIndexChanged += cbOutlookCalendar_SelectedIndexChanged;
            //Select the right calendar
            int c = 0;
            foreach (KeyValuePair<String, MAPIFolder> calendarFolder in OutlookCalendar.Instance.CalendarFolders) {
                if (calendarFolder.Value.EntryID == Settings.Instance.UseOutlookCalendar.Id) {
                    cbOutlookCalendars.SelectedIndex = c;
                }
                c++;
            }
            if (cbOutlookCalendars.SelectedIndex == -1) cbOutlookCalendars.SelectedIndex = 0;
            #endregion
            #region Google box
            if (Settings.Instance.UseGoogleCalendar != null && Settings.Instance.UseGoogleCalendar.Id != null) {
                cbGoogleCalendars.Items.Add(Settings.Instance.UseGoogleCalendar);
                cbGoogleCalendars.SelectedIndex = 0;
            }
            #endregion
            #region Sync Options box
            #region How
            this.gbSyncOptions_How.Height = 109;
            syncDirection.Items.Add(SyncDirection.OutlookToGoogle);
            syncDirection.Items.Add(SyncDirection.GoogleToOutlook);
            syncDirection.Items.Add(SyncDirection.Bidirectional);
            cbObfuscateDirection.Items.Add(SyncDirection.OutlookToGoogle);
            cbObfuscateDirection.Items.Add(SyncDirection.GoogleToOutlook);
            //Sync Direction dropdown
            for (int i = 0; i < syncDirection.Items.Count; i++) {
                SyncDirection sd = (syncDirection.Items[i] as SyncDirection);
                if (sd.Id == Settings.Instance.SyncDirection.Id) {
                    syncDirection.SelectedIndex = i;
                }
            }
            if (syncDirection.SelectedIndex == -1) syncDirection.SelectedIndex = 0;
            this.gbSyncOptions_How.SuspendLayout();
            cbMergeItems.Checked = Settings.Instance.MergeItems;
            cbDisableDeletion.Checked = Settings.Instance.DisableDelete;
            cbConfirmOnDelete.Enabled = !Settings.Instance.DisableDelete;
            cbConfirmOnDelete.Checked = Settings.Instance.ConfirmOnDelete;
            //Obfuscate Direction dropdown
            for (int i = 0; i < cbObfuscateDirection.Items.Count; i++) {
                SyncDirection sd = (cbObfuscateDirection.Items[i] as SyncDirection);
                if (sd.Id == Settings.Instance.Obfuscation.Direction.Id) {
                    cbObfuscateDirection.SelectedIndex = i;
                }
            }
            if (cbObfuscateDirection.SelectedIndex == -1) cbObfuscateDirection.SelectedIndex = 0;
            cbObfuscateDirection.Enabled = Settings.Instance.SyncDirection == SyncDirection.Bidirectional;
            Settings.Instance.Obfuscation.LoadRegex(dgObfuscateRegex);
            this.gbSyncOptions_How.ResumeLayout();
            #endregion
            #region When
            this.gbSyncOptions_When.SuspendLayout();
            tbDaysInThePast.Text = Settings.Instance.DaysInThePast.ToString();
            tbDaysInTheFuture.Text = Settings.Instance.DaysInTheFuture.ToString();
            tbInterval.ValueChanged -= new System.EventHandler(this.tbMinuteOffsets_ValueChanged);
            tbInterval.Value = Settings.Instance.SyncInterval;
            tbInterval.ValueChanged += new System.EventHandler(this.tbMinuteOffsets_ValueChanged);
            cbIntervalUnit.SelectedIndexChanged -= new System.EventHandler(this.cbIntervalUnit_SelectedIndexChanged);
            cbIntervalUnit.Text = Settings.Instance.SyncIntervalUnit;
            cbIntervalUnit.SelectedIndexChanged += new System.EventHandler(this.cbIntervalUnit_SelectedIndexChanged); 
            cbOutlookPush.Checked = Settings.Instance.OutlookPush;
            this.gbSyncOptions_When.ResumeLayout();
            #endregion
            #region What
            this.gbSyncOptions_What.SuspendLayout();
            cbAddDescription.Checked = Settings.Instance.AddDescription;
            cbAddAttendees.Checked = Settings.Instance.AddAttendees;
            cbAddReminders.Checked = Settings.Instance.AddReminders;
            this.gbSyncOptions_What.ResumeLayout();
            #endregion
            #endregion
            #region Application behaviour
            cbShowBubbleTooltips.Checked = Settings.Instance.ShowBubbleTooltipWhenSyncing;
            cbStartOnStartup.Checked = Settings.Instance.StartOnStartup;
            cbStartInTray.Checked = Settings.Instance.StartInTray;
            cbMinimizeToTray.Checked = Settings.Instance.MinimizeToTray;
            cbPortable.Checked = Settings.Instance.Portable;
            cbPortable.Enabled = !Program.isClickOnceInstall();
            cbCreateFiles.Checked = Settings.Instance.CreateCSVFiles;
            for (int i = 0; i < cbLoggingLevel.Items.Count; i++) {
                if (cbLoggingLevel.Items[i].ToString().ToLower() == Settings.Instance.LoggingLevel.ToLower()) {
                    cbLoggingLevel.SelectedIndex = i;
                    break;
                }
            }
            updateGUIsettings_Proxy();
            #endregion
            cbAlphaReleases.Checked = Settings.Instance.AlphaReleases;
            cbAlphaReleases.Visible = !Program.isClickOnceInstall();
            this.ResumeLayout();
        }

        private void updateGUIsettings_Proxy() {
            rbProxyIE.Checked = true;
            rbProxyNone.Checked = (Settings.Instance.Proxy.Type == "None");
            rbProxyCustom.Checked = (Settings.Instance.Proxy.Type == "Custom");
            cbProxyAuthRequired.Enabled = (Settings.Instance.Proxy.Type == "Custom");
            txtProxyServer.Text = Settings.Instance.Proxy.ServerName;
            txtProxyPort.Text = Settings.Instance.Proxy.Port.ToString();
            txtProxyServer.Enabled = rbProxyCustom.Checked;
            txtProxyPort.Enabled = rbProxyCustom.Checked;

            if (!string.IsNullOrEmpty(Settings.Instance.Proxy.UserName) &&
                !string.IsNullOrEmpty(Settings.Instance.Proxy.Password)) {
                cbProxyAuthRequired.Checked = true;
            } else {
                cbProxyAuthRequired.Checked = false;
            }
            txtProxyUser.Text = Settings.Instance.Proxy.UserName;
            txtProxyPassword.Text = Settings.Instance.Proxy.Password;
            txtProxyUser.Enabled = cbProxyAuthRequired.Checked;
            txtProxyPassword.Enabled = cbProxyAuthRequired.Checked;
        }
        
        private void applyProxy() {
            if (rbProxyNone.Checked) Settings.Instance.Proxy.Type = rbProxyNone.Tag.ToString();
            else if (rbProxyCustom.Checked) Settings.Instance.Proxy.Type = rbProxyCustom.Tag.ToString();
            else Settings.Instance.Proxy.Type = rbProxyIE.Tag.ToString();
            
            if (rbProxyCustom.Checked) {
                if (String.IsNullOrEmpty(txtProxyServer.Text) || String.IsNullOrEmpty(txtProxyPort.Text)) {
                    MessageBox.Show("A proxy server name and port must be provided.", "Proxy Authentication Enabled", 
                        MessageBoxButtons.OK, MessageBoxIcon.Error);
                    txtProxyServer.Focus();
                    return;
                }
                int nPort;
                if (!int.TryParse(txtProxyPort.Text, out nPort)) {
                    MessageBox.Show("Proxy server port must be a number.", "Invalid Proxy Port",
                        MessageBoxButtons.OK, MessageBoxIcon.Error);
                    txtProxyPort.Focus();
                    return;
                }

                string userName = null;
                string password = null;
                if (cbProxyAuthRequired.Checked) {
                    userName = txtProxyUser.Text;
                    password = txtProxyPassword.Text;
                } else {
                    userName = string.Empty;
                    password = string.Empty;
                }

                Settings.Instance.Proxy.ServerName = txtProxyServer.Text;
                Settings.Instance.Proxy.Port = nPort;
                Settings.Instance.Proxy.UserName = userName;
                Settings.Instance.Proxy.Password = password;
            }
            Settings.Instance.Proxy.Configure();
        }

        #region Autosync functions
        int getResyncInterval() {
            int min = Settings.Instance.SyncInterval;
            if (Settings.Instance.SyncIntervalUnit == "Hours") {
                min *= 60;
            }
            return min;
        }

        private void ogcsTimer_Tick(object sender, EventArgs e) {
            log.Debug("Scheduled sync triggered.");

            showBubbleInfo("Autosyncing calendars: " + Settings.Instance.SyncDirection.Name + "...");
            if (!this.SyncingNow) {
                sync_Click(sender, null);
            } else {
                log.Debug("Busy syncing already. Rescheduled for 5 mins time.");
                setNextSync(5, fromNow:true);
            }
        }

        public void OgcsPushTimer_Tick(object sender, EventArgs e) {
            if (Convert.ToInt16(bSyncNow.Tag) != 0) {
                log.Debug("Push sync triggered.");
                showBubbleInfo("Autosyncing calendars: " + Settings.Instance.SyncDirection.Name + "...");
                if (!this.SyncingNow) {
                    sync_Click(sender, null);
                } else {
                    log.Debug("Busy syncing already. No need to push.");
                    bSyncNow.Tag = 0;
                }
            }
        }
        
        void setNextSync(int delayMins, Boolean fromNow=false) {
            if (tbInterval.Value != 0) {
                DateTime nextSyncDate = lastSyncDate.AddMinutes(delayMins);
                DateTime now = DateTime.Now;
                if (fromNow)
                    nextSyncDate = now.AddMinutes(delayMins);

                if (ogcsTimer.Interval != (delayMins * 60000)) {
                    ogcsTimer.Stop();
                    TimeSpan diff = nextSyncDate - now;
                    if (diff.Minutes < 1) {
                        nextSyncDate = now.AddMinutes(1);
                        ogcsTimer.Interval = 1 * 60000;
                    } else {
                        ogcsTimer.Interval = diff.Minutes * 60000;
                    }
                    ogcsTimer.Start();
                }
                lNextSyncVal.Text = nextSyncDate.ToLongDateString() + " - " + nextSyncDate.ToLongTimeString();
                log.Info("Next sync scheduled for " + lNextSyncVal.Text);
            } else {
                lNextSyncVal.Text = "Inactive";
                ogcsTimer.Stop();
                log.Info("Schedule disabled.");
            }
        }
        #endregion

        private void sync_Click(object sender, EventArgs e) {
            if (bSyncNow.Text == "Start Sync") {
                if (sender.GetType().ToString().EndsWith("Timer")) {
                    log.Debug("Scheduled sync started.");
                    Timer aTimer = sender as Timer;
                    if (aTimer.Tag.ToString() == "PushTimer") sync_Start(updateSyncSchedule: false);
                    else if (aTimer.Tag.ToString() == "AutoSyncTimer") sync_Start(updateSyncSchedule: true);
                } else {
                    log.Debug("Manual sync started.");
                    sync_Start(updateSyncSchedule: false);
                }

            } else if (bSyncNow.Text == "Stop Sync") {
                if (bwSync != null && !bwSync.CancellationPending) {
                    log.Warn("Sync cancellation requested.");
                    bwSync.CancelAsync();
                } else {
                    log.Warn("Repeated cancellation requested - forcefully aborting thread!");
                    bwSync = null;
                }
            }
        } 

        private void sync_Start(Boolean updateSyncSchedule=true) {
            LogBox.Clear();

            if (Settings.Instance.UseGoogleCalendar == null || 
                Settings.Instance.UseGoogleCalendar.Id == null ||
                Settings.Instance.UseGoogleCalendar.Id == "") {
                MessageBox.Show("You need to select a Google Calendar first on the 'Settings' tab.");
                return;
            }
            //Check network availability
            if (!System.Net.NetworkInformation.NetworkInterface.GetIsNetworkAvailable()) {
                Logboxout("There does not appear to be any network available! Sync aborted.", notifyBubble:true);
                return;
            }
            //Check if Outlook is Online
            try {
                if (OutlookCalendar.Instance.IOutlook.Offline() && Settings.Instance.AddAttendees) {
                    Logboxout("You have selected to sync attendees but Outlook is currently offline.");
                    Logboxout("Either put Outlook online or do not sync attendees.", notifyBubble: true);
                    return;
                }
            } catch (System.Exception ex) {
                Logboxout(ex.Message, notifyBubble: true);
                log.Error(ex.StackTrace);
                return;
            }
            GoogleCalendar.APIlimitReached_attendee = false;
            bSyncNow.Text = "Stop Sync";
            String cacheNextSync = lNextSyncVal.Text;
            lNextSyncVal.Text = "In progress...";

            DateTime SyncStarted = DateTime.Now;
            log.Info("Sync version: " + System.Windows.Forms.Application.ProductVersion);
            Logboxout("Sync started at " + SyncStarted.ToString());
            Logboxout("Syncing from "+ DateTime.Today.AddDays(-Settings.Instance.DaysInThePast).ToShortDateString() +
                " to "+ DateTime.Today.AddDays(+Settings.Instance.DaysInTheFuture+1).ToShortDateString());
            Logboxout(Settings.Instance.SyncDirection.Name);
            Logboxout("--------------------------------------------------");

            if (Settings.Instance.OutlookPush) OutlookCalendar.Instance.DeregisterForAutoSync();

            Boolean syncOk = false;
            int failedAttempts = 0;
            Social.TrackSync();
            while (!syncOk) {
                if (failedAttempts > 0 &&
                    MessageBox.Show("The synchronisation failed. Do you want to try again?", "Sync Failed",
                    MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation) == System.Windows.Forms.DialogResult.No) 
                {
                    bSyncNow.Text = "Start Sync";
                    break;
                }

                //Set up a separate thread for the sync to operate in. Keeps the UI responsive.
                bwSync = new BackgroundWorker();
                //Don't need thread to report back. The logbox is updated from the thread anyway.
                bwSync.WorkerReportsProgress = false;
                bwSync.WorkerSupportsCancellation = true;

                //Kick off the sync in the background thread
                bwSync.DoWork += new DoWorkEventHandler(
                delegate(object o, DoWorkEventArgs args) {
                    BackgroundWorker b = o as BackgroundWorker;
                    syncOk = synchronize();
                });

                bwSync.RunWorkerAsync();
                while (bwSync != null && (bwSync.IsBusy || bwSync.CancellationPending)) {
                    System.Windows.Forms.Application.DoEvents();
                    System.Threading.Thread.Sleep(100);
                }
                failedAttempts += !syncOk ? 1 : 0;
            }
            Settings.Instance.CompletedSyncs += syncOk ? 1 : 0;
            bSyncNow.Text = "Start Sync";

            Logboxout(syncOk ? "Sync finished with success!" : "Operation aborted after "+ failedAttempts +" failed attempts!");

            if (Settings.Instance.OutlookPush) OutlookCalendar.Instance.RegisterForAutoSync();

            lLastSyncVal.Text = SyncStarted.ToLongDateString() + " - " + SyncStarted.ToLongTimeString();
            Settings.Instance.LastSyncDate = SyncStarted;
            if (!updateSyncSchedule) {
                lNextSyncVal.Text = cacheNextSync;
            } else {
                if (syncOk) {
                    lastSyncDate = SyncStarted;
                    setNextSync(getResyncInterval());
                } else {
                    if (Settings.Instance.SyncInterval != 0) {
                        Logboxout("Another sync has been scheduled to automatically run in 5 minutes time.");
                        setNextSync(5, fromNow: true);
                    }
                }
            }
            bSyncNow.Enabled = true;
            bSyncNow.Tag = 0; //Reset Push flag regardless of success (don't want it trying every 2 mins)

            checkSyncMilestone();
        }

        private Boolean synchronize() {
            #region Read Outlook items
            Logboxout("Reading Outlook Calendar Entries...");
            List<AppointmentItem> OutlookEntries = null;
            try {
                OutlookEntries = OutlookCalendar.Instance.getCalendarEntriesInRange();
            } catch (System.Exception ex) {
                Logboxout("Unable to access the Outlook calendar. The following error occurred:");
                Logboxout(ex.Message + "\r\n => Retry later.");
                log.Error(ex.StackTrace);
                try { OutlookCalendar.Instance.Reset(); } catch { }
                return false;
            }
            Logboxout(OutlookEntries.Count + " Outlook calendar entries found.");
            Logboxout("--------------------------------------------------");
            #endregion

            #region Read Google items
            Logboxout("Reading Google Calendar Entries...");
            List<Event> googleEntries = null;
            try {
                googleEntries = GoogleCalendar.Instance.GetCalendarEntriesInRange();
            } catch (System.Exception ex) {
                Logboxout("Unable to connect to the Google calendar. The following error occurred:");
                Logboxout(ex.Message);
                log.Error(ex.StackTrace);
                if (Settings.Instance.Proxy.Type == "IE") {
                    if (MessageBox.Show("Please ensure you can access the internet with Internet Explorer.\r\n"+
                        "Test it now? If successful, please retry synchronising your calendar.", 
                        "Test IE Internet Access",
                        MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes) {
                        System.Diagnostics.Process.Start("iexplore.exe", "http://www.google.com");
                    }
                }
                return false;
            }
            Logboxout(googleEntries.Count + " Google calendar entries found.");
            Logboxout("--------------------------------------------------");
            #endregion

            Boolean success = true;
            String bubbleText = "";
            if (Settings.Instance.SyncDirection != SyncDirection.GoogleToOutlook) {
                success = sync_outlookToGoogle(OutlookEntries, googleEntries, ref bubbleText);
            }
            if (!success) return false;
            if (Settings.Instance.SyncDirection != SyncDirection.OutlookToGoogle) {
                if (bubbleText != "") bubbleText += "\r\n";
                success = sync_googleToOutlook(googleEntries, OutlookEntries, ref bubbleText);
            }
            if (bubbleText != "") showBubbleInfo(bubbleText);
            return success;
        }

        private Boolean sync_outlookToGoogle(List<AppointmentItem> outlookEntries, List<Event> googleEntries, ref String bubbleText) {
            log.Debug("Synchronising from Outlook to Google.");
            
            //  Make copies of each list of events (Not strictly needed)
            List<AppointmentItem> googleEntriesToBeCreated = new List<AppointmentItem>(outlookEntries);
            List<Event> googleEntriesToBeDeleted = new List<Event>(googleEntries);
            Dictionary<AppointmentItem, Event> entriesToBeCompared = new Dictionary<AppointmentItem, Event>();

            try {
                GoogleCalendar.Instance.ReclaimOrphanCalendarEntries(ref googleEntriesToBeDeleted, ref outlookEntries);
            } catch (System.Exception ex) {
                MainForm.Instance.Logboxout("Unable to reclaim orphan calendar entries in Google calendar. The following error occurred:");
                MainForm.Instance.Logboxout(ex.Message, notifyBubble:true);
                log.Error(ex.StackTrace);
                return false;
            }
            try {
                GoogleCalendar.IdentifyEventDifferences(ref googleEntriesToBeCreated, ref googleEntriesToBeDeleted, entriesToBeCompared);
            } catch (System.Exception ex) {
                MainForm.Instance.Logboxout("Unable to identify differences in Google calendar. The following error occurred:");
                MainForm.Instance.Logboxout(ex.Message, notifyBubble: true);
                log.Error(ex.StackTrace);
                return false;
            }
            
            Logboxout(googleEntriesToBeDeleted.Count + " Google calendar entries to be deleted.");
            Logboxout(googleEntriesToBeCreated.Count + " Google calendar entries to be created.");

            //Protect against very first syncs which may trample pre-existing non-Outlook events in Google
            if (!Settings.Instance.DisableDelete && !Settings.Instance.ConfirmOnDelete &&
                googleEntriesToBeDeleted.Count == googleEntries.Count && googleEntries.Count > 0) {
                if (MessageBox.Show("All Google events are going to be deleted. Do you want to allow this?" +
                    "\r\nNote, " + googleEntriesToBeCreated.Count + " events will then be created.", "Confirm mass deletion",
                    MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation) == DialogResult.No) 
                {
                    googleEntriesToBeDeleted = new List<Event>();
                }
            }

            #region Delete Google Entries
            if (googleEntriesToBeDeleted.Count > 0) {
                Logboxout("--------------------------------------------------");
                Logboxout("Deleting " + googleEntriesToBeDeleted.Count + " Google calendar entries...");
                try {
                    GoogleCalendar.Instance.DeleteCalendarEntries(googleEntriesToBeDeleted);
                } catch (System.Exception ex) {
                    MainForm.Instance.Logboxout("Unable to delete obsolete entries in Google calendar. The following error occurred:");
                    MainForm.Instance.Logboxout(ex.Message, notifyBubble: true);
                    log.Error(ex.StackTrace);
                    return false;
                }
                Logboxout("Done.");
            }
            #endregion

            #region Create Google Entries
            if (googleEntriesToBeCreated.Count > 0) {
                Logboxout("--------------------------------------------------");
                Logboxout("Creating " + googleEntriesToBeCreated.Count + " Google calendar entries...");
                try {
                    GoogleCalendar.Instance.CreateCalendarEntries(googleEntriesToBeCreated);
                } catch (System.Exception ex) {
                    Logboxout("Unable to add new entries into the Google Calendar. The following error occurred:");
                    Logboxout(ex.Message, notifyBubble: true);
                    log.Error(ex.StackTrace);
                    return false;
                }
                Logboxout("Done.");
            }
            #endregion

            #region Update Google Entries
            int entriesUpdated = 0;
            if (entriesToBeCompared.Count > 0) {
                Logboxout("--------------------------------------------------");
                Logboxout("Comparing " + entriesToBeCompared.Count + " existing Google calendar entries...");
                try {
                    GoogleCalendar.Instance.UpdateCalendarEntries(entriesToBeCompared, ref entriesUpdated);
                } catch (System.Exception ex) {
                    Logboxout("Unable to update existing entries in the Google calendar. The following error occurred:");
                    Logboxout(ex.Message, notifyBubble: true);
                    log.Error(ex.StackTrace);
                    return false;
                }
                Logboxout(entriesUpdated + " entries updated.");
            }
            #endregion
            Logboxout("--------------------------------------------------");

            bubbleText = "Google: " + googleEntriesToBeCreated.Count + " created; "+ 
                googleEntriesToBeDeleted.Count +" deleted; "+ entriesUpdated + " updated";

            return true;
        }

        private Boolean sync_googleToOutlook(List<Event> googleEntries, List<AppointmentItem> outlookEntries, ref String bubbleText) {
            log.Debug("Synchronising from Google to Outlook.");
            
            //  Make copies of each list of events (Not strictly needed)
            List<Event> outlookEntriesToBeCreated = new List<Event>(googleEntries);
            List<AppointmentItem> outlookEntriesToBeDeleted = new List<AppointmentItem>(outlookEntries);
            Dictionary<AppointmentItem, Event> entriesToBeCompared = new Dictionary<AppointmentItem, Event>();
            
            try {
                OutlookCalendar.Instance.ReclaimOrphanCalendarEntries(ref outlookEntriesToBeDeleted, ref googleEntries);
            } catch (System.Exception ex) {
                MainForm.Instance.Logboxout("Unable to reclaim orphan calendar entries in Outlook calendar. The following error occurred:");
                MainForm.Instance.Logboxout(ex.Message, notifyBubble:true);
                log.Error(ex.StackTrace);
                return false;
            }
            try {
                OutlookCalendar.IdentifyEventDifferences(ref outlookEntriesToBeCreated, ref outlookEntriesToBeDeleted, entriesToBeCompared);
            } catch (System.Exception ex) {
                MainForm.Instance.Logboxout("Unable to identify differences in Outlook calendar. The following error occurred:");
                MainForm.Instance.Logboxout(ex.Message, notifyBubble: true);
                log.Error(ex.StackTrace);
                return false;
            }
            
            Logboxout(outlookEntriesToBeDeleted.Count + " Outlook calendar entries to be deleted.");
            Logboxout(outlookEntriesToBeCreated.Count + " Outlook calendar entries to be created.");

            //Protect against very first syncs which may trample pre-existing non-Google events in Outlook
            if (!Settings.Instance.DisableDelete && !Settings.Instance.ConfirmOnDelete &&
                outlookEntriesToBeDeleted.Count == outlookEntries.Count && outlookEntries.Count > 0) {
                if (MessageBox.Show("All Outlook events are going to be deleted. Do you want to allow this?" +
                    "\r\nNote, " + outlookEntriesToBeCreated.Count + " events will then be created.", "Confirm mass deletion",
                    MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation) == DialogResult.No) {
                    outlookEntriesToBeDeleted = new List<AppointmentItem>();
                }
            }
            
            #region Delete Outlook Entries
            if (outlookEntriesToBeDeleted.Count > 0) {
                Logboxout("--------------------------------------------------");
                Logboxout("Deleting " + outlookEntriesToBeDeleted.Count + " Outlook calendar entries...");
                try {
                    OutlookCalendar.Instance.DeleteCalendarEntries(outlookEntriesToBeDeleted);
                } catch (System.Exception ex) {
                    MainForm.Instance.Logboxout("Unable to delete obsolete entries in Google calendar. The following error occurred:");
                    MainForm.Instance.Logboxout(ex.Message, notifyBubble:true);
                    log.Error(ex.StackTrace);
                    return false;
                }
                Logboxout("Done.");
            }
            #endregion
            
            #region Create Outlook Entries
            if (outlookEntriesToBeCreated.Count > 0) {
                Logboxout("--------------------------------------------------");
                Logboxout("Creating " + outlookEntriesToBeCreated.Count + " Outlook calendar entries...");
                try {
                    OutlookCalendar.Instance.CreateCalendarEntries(outlookEntriesToBeCreated);
                } catch (System.Exception ex) {
                    Logboxout("Unable to add new entries into the Outlook Calendar. The following error occurred:");
                    Logboxout(ex.Message, notifyBubble:true);
                    log.Error(ex.StackTrace);
                    return false;
                }
                Logboxout("Done.");
            }
            #endregion
            
            #region Update Outlook Entries
            int entriesUpdated = 0;
            if (entriesToBeCompared.Count > 0) {
                Logboxout("--------------------------------------------------");
                Logboxout("Comparing " + entriesToBeCompared.Count + " existing Outlook calendar entries...");
                try {
                    OutlookCalendar.Instance.UpdateCalendarEntries(entriesToBeCompared, ref entriesUpdated);
                } catch (System.Exception ex) {
                    Logboxout("Unable to update new entries into the Outlook calendar. The following error occurred:");
                    Logboxout(ex.Message, notifyBubble:true);
                    log.Error(ex.StackTrace);
                    return false;
                }
                Logboxout(entriesUpdated + " entries updated.");
            }
            #endregion
            Logboxout("--------------------------------------------------");

            bubbleText += "Outlook: " + outlookEntriesToBeCreated.Count + " created; " +
                outlookEntriesToBeDeleted.Count + " deleted; " + entriesUpdated + " updated";
            
            return true;
        }

        #region Compare Event Attributes 
        public static Boolean CompareAttribute(String attrDesc, SyncDirection fromTo, String googleAttr, String outlookAttr, System.Text.StringBuilder sb, ref int itemModified) {
            if (googleAttr == null) googleAttr = "";
            if (outlookAttr == null) outlookAttr = "";
            //Truncate long strings
            String googleAttr_stub = (googleAttr.Length > 50) ? googleAttr.Substring(0, 47) + "..." : googleAttr;
            String outlookAttr_stub = (outlookAttr.Length > 50) ? outlookAttr.Substring(0, 47) + "..." : outlookAttr;
            log.Fine("Comparing " + attrDesc);
            if (googleAttr != outlookAttr) {
                if (fromTo == SyncDirection.GoogleToOutlook) {
                    sb.AppendLine(attrDesc + ": " + outlookAttr_stub + " => " + googleAttr_stub);
                } else {
                    sb.AppendLine(attrDesc + ": " + googleAttr_stub + " => " + outlookAttr_stub);
                }
                itemModified++;
                return true;
            } 
            return false;
        }
        public static Boolean CompareAttribute(String attrDesc, SyncDirection fromTo, Boolean googleAttr, Boolean outlookAttr, System.Text.StringBuilder sb, ref int itemModified) {
            log.Fine("Comparing " + attrDesc);
            if (googleAttr != outlookAttr) {
                if (fromTo == SyncDirection.GoogleToOutlook) {
                    sb.AppendLine(attrDesc + ": " + outlookAttr + " => " + googleAttr);
                } else {
                    sb.AppendLine(attrDesc + ": " + googleAttr + " => " + outlookAttr);
                }
                itemModified++;
                return true;
            } else {
                return false;
            }
        }
        #endregion

        private void showBubbleInfo(string message, ToolTipIcon iconType = ToolTipIcon.Info) {
            notifyIcon1.Visible = true; //How best to do this?
            if (Settings.Instance.ShowBubbleTooltipWhenSyncing) {
                notifyIcon1.ShowBalloonTip(
                    500,
                    "Outlook Google Calendar Sync",
                    message,
                    iconType
                );
            }
        }

        public void Logboxout(string s, bool newLine=true, bool verbose=false, bool notifyBubble=false) {
            if ((verbose && Settings.Instance.VerboseOutput) || !verbose) {
                String existingText = GetControlPropertyThreadSafe(LogBox, "Text") as String;
                SetControlPropertyThreadSafe(LogBox, "Text", existingText + s + (newLine ? Environment.NewLine : ""));
            }
            if (notifyBubble & Settings.Instance.ShowBubbleTooltipWhenSyncing) {
                showBubbleInfo("Issue encountered.\n" +
                    "Please review output on the main 'Sync' tab", ToolTipIcon.Warning);
            }
            if (verbose) log.Debug(s);
            else log.Info(s);
        }

        public void HandleException(System.Exception ex) {
            MessageBox.Show(ex.ToString(), "Exception!", MessageBoxButtons.OK, MessageBoxIcon.Error);
            TextWriter tw = new StreamWriter("exception.txt");
            tw.WriteLine(ex.ToString());
            tw.Close();

            this.Close();
            System.Environment.Exit(-1);
            System.Windows.Forms.Application.Exit();
        }
                
        
        #region EVENTS
        #region Form actions
        void Save_Click(object sender, EventArgs e) {
            applyProxy();
            Settings.Instance.Save();
            
            Program.CreateStartupShortcut();

            //Push Sync
            if (Settings.Instance.OutlookPush) OutlookCalendar.Instance.RegisterForAutoSync();
            else OutlookCalendar.Instance.DeregisterForAutoSync();

            Settings.Instance.LogSettings();
        }

        private void NotifyIcon1_Click(object sender, EventArgs e) {
            this.WindowState = FormWindowState.Normal;
            this.Show();
        }

        private void mainFormResize(object sender, EventArgs e) {
            if (!cbMinimizeToTray.Checked) return;
            if (this.WindowState == FormWindowState.Minimized) {
                notifyIcon1.Visible = true;
                this.Hide();
                this.ShowInTaskbar = false;
            } else if (this.WindowState == FormWindowState.Normal) {
                notifyIcon1.Visible = false;
                this.ShowInTaskbar = true;
            }
        }

        void lAboutURL_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e) {
            System.Diagnostics.Process.Start(lAboutURL.Text);
        }
        #endregion
        private void tabAppSettings_DrawItem(object sender, DrawItemEventArgs e) {
            //Want to have horizontal sub-tabs on the left of the Settings tab.
            //Need to handle this manually

            Graphics g = e.Graphics;

            //Tab is rotated, so width is height and vica-versa :-|
            if (tabAppSettings.ItemSize.Width != 35 || tabAppSettings.ItemSize.Height != 75) {
                tabAppSettings.ItemSize = new Size(35, 75);
            }
            // Get the item from the collection.
            TabPage tabPage = tabAppSettings.TabPages[e.Index];

            // Get the real bounds for the tab rectangle.
            Rectangle tabBounds = tabAppSettings.GetTabRect(e.Index);
            Font tabFont = new Font("Microsoft Sans Serif", (float)11, FontStyle.Regular, GraphicsUnit.Pixel);
            Brush textBrush = new SolidBrush(Color.Black);

            if (e.State == DrawItemState.Selected) {
                tabFont = new Font("Microsoft Sans Serif", (float)11, FontStyle.Bold, GraphicsUnit.Pixel);
                Rectangle tabColour = e.Bounds;
                //Blue highlight
                int highlightWidth = 5;
                tabColour.Width = highlightWidth;
                tabColour.X = 0;
                g.FillRectangle(Brushes.Blue, tabColour);
                //Tab main background
                tabColour = e.Bounds;
                tabColour.Width -= highlightWidth;
                tabColour.X += highlightWidth;
                g.FillRectangle(Brushes.White, tabColour);
            } else {
                // Draw a different background color, and don't paint a focus rectangle.
                g.FillRectangle(SystemBrushes.ButtonFace, e.Bounds);
            }

            //Draw white rectangle below the tabs (this would be nice and easy in .Net4)
            Rectangle lastTabRect = tabAppSettings.GetTabRect(tabAppSettings.TabPages.Count - 1);
            tabAppSettings_background.Location = new Point(0, ((lastTabRect.Height + 1) * tabAppSettings.TabPages.Count));
            tabAppSettings_background.Size = new Size(lastTabRect.Width, tabAppSettings.Height - (lastTabRect.Height * tabAppSettings.TabPages.Count));
            e.Graphics.FillRectangle(Brushes.White, tabAppSettings_background);

            // Draw string and align the text.
            StringFormat stringFlags = new StringFormat();
            stringFlags.Alignment = StringAlignment.Far;
            stringFlags.LineAlignment = StringAlignment.Center;
            g.DrawString(tabPage.Text, tabFont, textBrush, tabBounds, new StringFormat(stringFlags));
        }
        #region Outlook settings
        public void rbOutlookDefaultMB_CheckedChanged(object sender, EventArgs e) {
            if (rbOutlookDefaultMB.Checked) {
                Settings.Instance.OutlookService = OutlookCalendar.Service.DefaultMailbox;
                OutlookCalendar.Instance.Reset();
                gbEWS.Enabled = false;
                //Update available calendars
                cbOutlookCalendars.DataSource = new BindingSource(OutlookCalendar.Instance.CalendarFolders, null);
            }
        }

        private void rbOutlookAltMB_CheckedChanged(object sender, EventArgs e) {
            if (rbOutlookAltMB.Checked) {
                Settings.Instance.OutlookService = OutlookCalendar.Service.AlternativeMailbox;
                Settings.Instance.MailboxName = ddMailboxName.Text;
                OutlookCalendar.Instance.Reset();
                gbEWS.Enabled = false;
                //Update available calendars
                cbOutlookCalendars.DataSource = new BindingSource(OutlookCalendar.Instance.CalendarFolders, null);
            }
            Settings.Instance.MailboxName = (rbOutlookAltMB.Checked ? ddMailboxName.Text : "");
            ddMailboxName.Enabled = rbOutlookAltMB.Checked;
        }

        private void rbOutlookEWS_CheckedChanged(object sender, EventArgs e) {
            if (rbOutlookEWS.Checked) {
                Settings.Instance.OutlookService = OutlookCalendar.Service.EWS;
                OutlookCalendar.Instance.Reset();
                gbEWS.Enabled = true;
                //Update available calendars
                cbOutlookCalendars.DataSource = new BindingSource(OutlookCalendar.Instance.CalendarFolders, null);
            }
        }

        private void ddMailboxName_SelectedIndexChanged(object sender, EventArgs e) {
            if (this.Visible) {
                Settings.Instance.MailboxName = ddMailboxName.Text;
                OutlookCalendar.Instance.Reset();
            }
        }

        private void txtEWSUser_TextChanged(object sender, EventArgs e) {
            Settings.Instance.EWSuser = txtEWSUser.Text;
        }

        private void txtEWSPass_TextChanged(object sender, EventArgs e) {
            Settings.Instance.EWSpassword = txtEWSPass.Text;
        }

        private void txtEWSServerURL_TextChanged(object sender, EventArgs e) {
            Settings.Instance.EWSserver = txtEWSServerURL.Text;
        }

        private void cbOutlookCalendar_SelectedIndexChanged(object sender, EventArgs e) {
            KeyValuePair<String, MAPIFolder> calendar = (KeyValuePair<String, MAPIFolder>)cbOutlookCalendars.SelectedItem;
            OutlookCalendar.Instance.UseOutlookCalendar = calendar.Value;
        }
        #endregion
        #region Google settings
        private void GetMyGoogleCalendars_Click(object sender, EventArgs e) {
            this.bGetGoogleCalendars.Text = "Retrieving Calendars...";
            bGetGoogleCalendars.Enabled = false;
            cbGoogleCalendars.Enabled = false;
            List<MyGoogleCalendarListEntry> calendars = null;
            try {
                calendars = GoogleCalendar.Instance.GetCalendars();
            } catch (ApplicationException) {
            } catch (System.Exception ex) {
                MessageBox.Show("Failed to retrieve Google calendars. \r\n" +
                    "Please check the output on the Sync tab for more details.", "Google calendar retrieval failed",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
                Logboxout("Unable to get the list of Google calendars. The following error occurred:");
                Logboxout(ex.Message);
                if (ex.InnerException!=null) Logboxout(ex.InnerException.Message);
                if (Settings.Instance.Proxy.Type == "IE") {
                    if (MessageBox.Show("Please ensure you can access the internet with Internet Explorer.\r\n" +
                        "Test it now? If successful, please retry retrieving your Google calendars.",
                        "Test IE Internet Access",
                        MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes) {
                        System.Diagnostics.Process.Start("iexplore.exe", "http://www.google.com");
                    }
                }
            } 
            if (calendars != null) {
                cbGoogleCalendars.Items.Clear();
                foreach (MyGoogleCalendarListEntry mcle in calendars) {
                    cbGoogleCalendars.Items.Add(mcle);
                }
                cbGoogleCalendars.SelectedIndex = 0;
            }

            bGetGoogleCalendars.Enabled = true;
            cbGoogleCalendars.Enabled = true;
            this.bGetGoogleCalendars.Text = "Retrieve Calendars";
        }

        private void cbGoogleCalendars_SelectedIndexChanged(object sender, EventArgs e) {
            Settings.Instance.UseGoogleCalendar = (MyGoogleCalendarListEntry)cbGoogleCalendars.SelectedItem;
        }

        private void btResetGCal_Click(object sender, EventArgs e) {
            if (MessageBox.Show("This will reset the Google account you are using to synchronise with.\r\n" +
                "Useful if you want to start syncing to a different account.",
                "Reset Google account?", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == System.Windows.Forms.DialogResult.Yes) {
                    Settings.Instance.UseGoogleCalendar.Id = null;
                    Settings.Instance.UseGoogleCalendar.Name = null;
                    this.cbGoogleCalendars.Items.Clear();
                    GoogleCalendar.Instance.Reset();
            }
        }
        #endregion
        #region Sync options
        #region How
        private void syncDirection_SelectedIndexChanged(object sender, EventArgs e) {
            Settings.Instance.SyncDirection = (SyncDirection)syncDirection.SelectedItem;
            if (Settings.Instance.SyncDirection == SyncDirection.Bidirectional) {
                cbObfuscateDirection.Enabled = true;
                cbObfuscateDirection.SelectedIndex = SyncDirection.OutlookToGoogle.Id-1;
            } else {
                cbObfuscateDirection.Enabled = false;
                cbObfuscateDirection.SelectedIndex = Settings.Instance.SyncDirection.Id-1;
            }
            showWhatPostit();
        }

        private void cbMergeItems_CheckedChanged(object sender, EventArgs e) {
            Settings.Instance.MergeItems = cbMergeItems.Checked;
        }

        private void cbConfirmOnDelete_CheckedChanged(object sender, System.EventArgs e) {
            Settings.Instance.ConfirmOnDelete = cbConfirmOnDelete.Checked;
        }

        private void cbDisableDeletion_CheckedChanged(object sender, System.EventArgs e) {
            Settings.Instance.DisableDelete = cbDisableDeletion.Checked;
            cbConfirmOnDelete.Enabled = !cbDisableDeletion.Checked;
        }

        private void cbOfuscate_CheckedChanged(object sender, EventArgs e) {
            Settings.Instance.Obfuscation.Enabled = cbOfuscate.Checked;
        }
        
        private void btObfuscateRules_CheckedChanged(object sender, EventArgs e) {
            const int minPanelHeight = 109;
            const int maxPanelHeight = 251;
            this.gbSyncOptions_How.BringToFront();
            if ((sender as CheckBox).Checked) {
                while (this.gbSyncOptions_How.Height < maxPanelHeight) {
                    this.gbSyncOptions_How.Height += 2;
                    System.Windows.Forms.Application.DoEvents();
                    System.Threading.Thread.Sleep(1);
                }
                this.gbSyncOptions_How.Height = maxPanelHeight;
            } else {
                while (this.gbSyncOptions_How.Height > minPanelHeight && this.Visible) {
                    this.gbSyncOptions_How.Height -= 2;
                    System.Windows.Forms.Application.DoEvents();
                    System.Threading.Thread.Sleep(1);
                }
                this.gbSyncOptions_How.Height = minPanelHeight;
            }
        }

        private void cbObfuscateDirection_SelectedIndexChanged(object sender, EventArgs e) {
            Settings.Instance.Obfuscation.Direction = (SyncDirection)cbObfuscateDirection.SelectedItem;
        }

        private void dgObfuscateRegex_Leave(object sender, EventArgs e) {
            Settings.Instance.Obfuscation.SaveRegex(dgObfuscateRegex);
        }
        #endregion
        #region When
        private void tbDaysInThePast_ValueChanged(object sender, EventArgs e) {
            Settings.Instance.DaysInThePast = (int)tbDaysInThePast.Value;
        }

        private void tbDaysInTheFuture_ValueChanged(object sender, EventArgs e) {
            Settings.Instance.DaysInTheFuture = (int)tbDaysInTheFuture.Value;
        }

        private void tbMinuteOffsets_ValueChanged(object sender, EventArgs e) {
            Settings.Instance.SyncInterval = (int)tbInterval.Value;
            setNextSync(getResyncInterval());
        }

        private void cbIntervalUnit_SelectedIndexChanged(object sender, EventArgs e) {
            Settings.Instance.SyncIntervalUnit = cbIntervalUnit.Text;
            setNextSync(getResyncInterval());
        }

        private void cbOutlookPush_CheckedChanged(object sender, EventArgs e) {
            Settings.Instance.OutlookPush = cbOutlookPush.Checked;
        }
        #endregion
        #region What
        private void showWhatPostit() {
            Boolean visible = (Settings.Instance.AddDescription &&
                Settings.Instance.SyncDirection == SyncDirection.Bidirectional);
            WhatPostit.Visible = visible;
            cbAddDescription_OnlyToGoogle.Visible = visible;
        }

        private void CbAddDescriptionCheckedChanged(object sender, EventArgs e) {
            Settings.Instance.AddDescription = cbAddDescription.Checked;
            cbAddDescription_OnlyToGoogle.Enabled = cbAddDescription.Checked;
            showWhatPostit();
        }
        private void cbAddDescription_OnlyToGoogle_CheckedChanged(object sender, EventArgs e) {
            Settings.Instance.AddDescription_OnlyToGoogle = cbAddDescription_OnlyToGoogle.Checked;
        }
        
        private void CbAddRemindersCheckedChanged(object sender, EventArgs e) {
            Settings.Instance.AddReminders = cbAddReminders.Checked;
        }

        private void cbAddAttendees_CheckedChanged(object sender, EventArgs e) {
            Settings.Instance.AddAttendees = cbAddAttendees.Checked;
        }
        #endregion
        #endregion
        #region Application settings
        private void cbStartOnStartup_CheckedChanged(object sender, EventArgs e) {
            Settings.Instance.StartOnStartup = cbStartOnStartup.Checked;
        }

        private void cbShowBubbleTooltipsCheckedChanged(object sender, System.EventArgs e) {
            Settings.Instance.ShowBubbleTooltipWhenSyncing = cbShowBubbleTooltips.Checked;
        }

        private void cbStartInTrayCheckedChanged(object sender, System.EventArgs e) {
            Settings.Instance.StartInTray = cbStartInTray.Checked;
        }

        private void cbMinimizeToTrayCheckedChanged(object sender, System.EventArgs e) {
            Settings.Instance.MinimizeToTray = cbMinimizeToTray.Checked;
        }

        private void cbPortable_CheckedChanged(object sender, EventArgs e) {
            if (this.Visible) {
                Settings.Instance.Portable = cbPortable.Checked;
                Program.MakePortable(cbPortable.Checked);
            }
        }

        private void cbCreateFiles_CheckedChanged(object sender, EventArgs e) {
            Settings.Instance.CreateCSVFiles = cbCreateFiles.Checked;
        }
        
        private void cbLoggingLevel_SelectedIndexChanged(object sender, EventArgs e) {
            Settings.configureLoggingLevel(MainForm.Instance.cbLoggingLevel.Text);
            Settings.Instance.LoggingLevel = MainForm.Instance.cbLoggingLevel.Text.ToUpper();
        }

        private void btLogLocation_Click(object sender, EventArgs e) {
            log4net.Appender.IAppender[] appenders = log.Logger.Repository.GetAppenders();
            String logFileLocation = (((log4net.Appender.FileAppender)appenders[0]).File);
            logFileLocation = logFileLocation.Substring(0, logFileLocation.LastIndexOf("\\"));
            System.Diagnostics.Process.Start(@logFileLocation);
        }

        #region Proxy
        private void rbProxyCustom_CheckedChanged(object sender, EventArgs e) {
            bool result = rbProxyCustom.Checked;
            txtProxyServer.Enabled = result;
            txtProxyPort.Enabled = result;
            cbProxyAuthRequired.Enabled = result;
            if (result) {
                result = !string.IsNullOrEmpty(txtProxyUser.Text) && !string.IsNullOrEmpty(txtProxyPassword.Text);
                cbProxyAuthRequired.Checked = result;
                txtProxyUser.Enabled = result;
                txtProxyPassword.Enabled = result;
            }
        }

        private void cbProxyAuthRequired_CheckedChanged(object sender, EventArgs e) {
            bool result = cbProxyAuthRequired.Checked;
            this.txtProxyPassword.Enabled = result;
            this.txtProxyUser.Enabled = result;
        }

        private void gbProxy_Leave(object sender, EventArgs e) {
            applyProxy();
        }
        #endregion
        #endregion

        private void cbVerboseOutput_CheckedChanged(object sender, EventArgs e) {
            Settings.Instance.VerboseOutput = cbVerboseOutput.Checked;
        }

        private void pbDonate_Click(object sender, EventArgs e) {
            Social.Donate();
        }

        private void btCheckForUpdate_Click(object sender, EventArgs e) {
            Program.checkForUpdate(true);
        }
        private void cbAlphaReleases_CheckedChanged(object sender, EventArgs e) {
            if (this.Visible)
                Settings.Instance.AlphaReleases = cbAlphaReleases.Checked;
        }
        #endregion

        #region Thread safe access to form components
        //private delegate Control getControlThreadSafeDelegate(Control control);
        //Used to update the logbox from the Sync() thread
        public delegate void SetControlPropertyThreadSafeDelegate(Control control, string propertyName, object propertyValue);
        public delegate object GetControlPropertyThreadSafeDelegate(Control control, string propertyName);
        
        //private static Control getControlThreadSafe(Control control) {
        //    if (control.InvokeRequired) {
        //        return (Control)control.Invoke(new getControlThreadSafeDelegate(getControlThreadSafe), new object[] { control });
        //    } else {
        //        return control;
        //    }
        //}
        public object GetControlPropertyThreadSafe(Control control, string propertyName) {
            if (control.InvokeRequired) {
                return control.Invoke(new GetControlPropertyThreadSafeDelegate(GetControlPropertyThreadSafe), new object[] { control, propertyName });
            } else {
                return control.GetType().InvokeMember(propertyName, System.Reflection.BindingFlags.GetProperty, null, control, null);
            }
        }
        public void SetControlPropertyThreadSafe(Control control, string propertyName, object propertyValue) {
            if (control.InvokeRequired) {
                control.Invoke(new SetControlPropertyThreadSafeDelegate(SetControlPropertyThreadSafe), new object[] { control, propertyName, propertyValue });
            } else {
                var theObject = control.GetType().InvokeMember(propertyName, System.Reflection.BindingFlags.SetProperty, null, control, new object[] { propertyValue });
                if (control.GetType().Name == "TextBox") {
                    TextBox tb = control as TextBox;
                    tb.SelectionStart = tb.Text.Length;
                    tb.ScrollToCaret();
                }
            }
        }
        #endregion

        #region Social Media & Analytics
        private void checkSyncMilestone() {
            Boolean isMilestone = false;
            Int32 syncs = Settings.Instance.CompletedSyncs;
            String blurb = "You've completed "+ String.Format("{0:n0}",syncs) +" syncs! Why not let people know how useful this tool is...";
            
            lMilestone.Text = String.Format("{0:n0}",syncs) +" Syncs!";
            lMilestoneBlurb.Text = blurb;

            switch (syncs) {
                case 10: isMilestone = true; break;
                case 100: isMilestone = true; break;
                case 250: isMilestone = true; break;
                case 1000: isMilestone = true; break;
            }
            if (isMilestone) {
                if (MessageBox.Show(blurb, "Spread the Word", MessageBoxButtons.OKCancel, MessageBoxIcon.Exclamation) == System.Windows.Forms.DialogResult.OK)
                    tabApp.SelectedTab = tabPage_Social;
            }
        }

        private void btSocialTweet_Click(object sender, EventArgs e) {
            Social.Twitter_tweet();
        }
        private void pbSocialTwitterFollow_Click(object sender, EventArgs e) {
            Social.Twitter_follow();
        }

        private void btSocialGplus_Click(object sender, EventArgs e) {
            Social.Google_share();
        }
        private void pbSocialGplusCommunity_Click(object sender, EventArgs e) {
            Social.Google_goToCommunity();
        }

        private void btSocialFB_Click(object sender, EventArgs e) {
            Social.Facebook_share();
        }

        private void btSocialRSSfeed_Click(object sender, EventArgs e) {
            Social.RSS_follow();
        }

        private void btSocialLinkedin_Click(object sender, EventArgs e) {
            Social.Linkedin_share();
        }

        #endregion

    }
}
