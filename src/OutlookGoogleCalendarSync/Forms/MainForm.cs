using Google.Apis.Calendar.v3.Data;
using log4net;
using Microsoft.Office.Interop.Outlook;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace OutlookGoogleCalendarSync.Forms {
    /// <summary>
    /// Description of MainForm.
    /// </summary>
    public partial class Main : Form {
        public static Main Instance;
        public NotificationTray NotificationTray { get; set; }
        public ToolTip ToolTips;
        private Console console;
        public Console Console {
            get { return console; }
        }

        private static readonly ILog log = LogManager.GetLogger(typeof(Main));
        private Rectangle tabAppSettings_background = new Rectangle();
        private float magnification = Graphics.FromHwnd(IntPtr.Zero).DpiY / 96; //Windows Display Magnifier (96DPI = 100%)

        public Main(string startingTab = null) {
            log.Debug("Initialiasing MainForm.");
            InitializeComponent();

            if (startingTab != null && startingTab == "Help") this.tabApp.SelectedTab = this.tabPage_Help;

            Instance = this;

            console = new Console(consoleWebBrowser);
            Telemetry.TrackVersions();
            updateGUIsettings();
            Settings.Instance.LogSettings();
            NotificationTray = new NotificationTray(this.trayIcon);

            log.Debug("Create the timer for the auto synchronisation");
            Sync.Engine.Instance.OgcsTimer = new Sync.SyncTimer();

            //Set up listener for Outlook calendar changes
            if (Settings.Instance.OutlookPush) Sync.Engine.Instance.RegisterForPushSync();

            if (Settings.Instance.StartInTray) {
                this.CreateHandle();
                this.WindowState = FormWindowState.Minimized;
            }
            if (((Sync.Engine.Instance.OgcsTimer.NextSyncDate ?? DateTime.Now.AddMinutes(10)) - DateTime.Now).TotalMinutes > 5) {
                OutlookOgcs.Calendar.Instance.Disconnect(onlyWhenNoGUI: true);
            }
        }

        private void updateGUIsettings() {
            this.SuspendLayout();
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
            ToolTips.SetToolTip(btTestOutlookFilter,
                "Check how many appointments are returned for the date range being synced.");

            //Google
            ToolTips.SetToolTip(cbGoogleCalendars,
                "The Google calendar to synchonize with.");
            ToolTips.SetToolTip(btResetGCal,
                "Reset the Google account being used to synchonize with.");

            //Settings
            ToolTips.SetToolTip(tbInterval,
                "Set to zero to disable automated syncs");
            ToolTips.SetToolTip(rbOutlookAltMB,
                "Only choose this if you need to use an Outlook Calendar that is not in the default mailbox");
            ToolTips.SetToolTip(cbMergeItems,
                "If the destination calendar has pre-existing items, don't delete them");
            ToolTips.SetToolTip(cbOutlookPush,
                "Synchronise changes in Outlook to Google within a few minutes.");
            ToolTips.SetToolTip(btCloseRegexRules,
                "Close obfuscation rules.");
            ToolTips.SetToolTip(cbOfuscate,
                "Mask specified words in calendar item subject.\nTakes effect for new or updated calendar items.");
            ToolTips.SetToolTip(dgObfuscateRegex,
                "All rules are applied in order provided using AND logic.\nSupports use of regular expressions.");
            ToolTips.SetToolTip(cbUseGoogleDefaultReminder,
                "If the calendar settings in Google have a default reminder configured, use this when Outlook has no reminder.");
            ToolTips.SetToolTip(cbUseOutlookDefaultReminder,
                "If the calendar settings in Outlook have a default reminder configured, use this when Google has no reminder.");
            ToolTips.SetToolTip(cbAddAttendees,
                "BE AWARE: Deleting Google event through mobile calendar app will notify all attendees.");
            ToolTips.SetToolTip(cbCloakEmail,
                "Google has been known to send meeting updates to attendees without your consent.\n" +
                "This option safeguards against that by appending '"+ GoogleOgcs.EventAttendee.EmailCloak +"' to their email address.");
            ToolTips.SetToolTip(cbReminderDND,
                "Do Not Disturb: Don't sync reminders to Google if they will trigger between these times.");

            //Application behaviour
            if (Settings.Instance.StartOnStartup)
                ToolTips.SetToolTip(tbStartupDelay, "Try setting a delay if COM errors occur on startup.");
            if (!Settings.Instance.UserIsBenefactor()) {
                ToolTips.SetToolTip(cbHideSplash, "Donate £10 or more to enable this feature.");
                ToolTips.SetToolTip(cbSuppressSocialPopup, "Donate £10 or more to enable this feature.");
            }
            ToolTips.SetToolTip(cbPortable,
                "For ZIP deployments, store configuration files in the application folder (useful if running from a USB thumb drive).\n" +
                "Default is in your User roaming profile.");
            ToolTips.SetToolTip(cbCreateFiles,
                "If checked, all entries found in Outlook/Google and identified for creation/deletion will be exported \n" +
                "to CSV files in the application's directory (named \"*.csv\"). \n" +
                "Only for debug/diagnostic purposes.");
            ToolTips.SetToolTip(rbProxyIE,
                "If IE settings have been changed, a restart of the Sync application may be required");
            ToolTips.SetToolTip(cbMuteClicks, "Mute any sounds when sync summary updates.");
            #endregion

            cbVerboseOutput.Checked = Settings.Instance.VerboseOutput;
            cbMuteClicks.Checked = Settings.Instance.MuteClickSounds;
            #region Outlook box
            #region Mailbox
            if (OutlookOgcs.Factory.is2003()) {
                rbOutlookDefaultMB.Checked = true;
                rbOutlookAltMB.Enabled = false;
                rbOutlookSharedCal.Enabled = false;
            } else {
                if (Settings.Instance.OutlookService == OutlookOgcs.Calendar.Service.AlternativeMailbox) {
                    rbOutlookAltMB.Checked = true;
                } else if (Settings.Instance.OutlookService == OutlookOgcs.Calendar.Service.SharedCalendar) {
                    rbOutlookSharedCal.Checked = true;
                } else {
                    rbOutlookDefaultMB.Checked = true;
                }
            }

            //Mailboxes the user has access to
            log.Debug("Find Folders");
            if (OutlookOgcs.Calendar.Instance.Folders.Count == 1) {
                rbOutlookAltMB.Enabled = false;
                rbOutlookAltMB.Checked = false;
            }
            Folders theFolders = OutlookOgcs.Calendar.Instance.Folders;
            Dictionary<String, List<String>> folderIDs = new Dictionary<String, List<String>>();
            for (int fld = 1; fld <= theFolders.Count; fld++) {
                MAPIFolder theFolder = theFolders[fld];
                try {
                    if (theFolder.Name != OutlookOgcs.Calendar.Instance.IOutlook.CurrentUserSMTP()) { //Not the default Exchange folder (assuming the default mailbox folder name hasn't been changed
                        //Create a dictionary of folder names and a list of their ID(s)
                        if (!folderIDs.ContainsKey(theFolder.Name)) {
                            folderIDs.Add(theFolder.Name, new List<String>(new String[] { theFolder.EntryID }));
                        } else if (!folderIDs[theFolder.Name].Contains(theFolder.EntryID)) {
                            folderIDs[theFolder.Name].Add(theFolder.EntryID);
                        }
                    }
                } catch (System.Exception ex) {
                    log.Debug("Failed to get EntryID for folder: " + theFolder.Name);
                    log.Debug(ex.Message);
                } finally {
                    theFolder = (MAPIFolder)OutlookOgcs.Calendar.ReleaseObject(theFolder);
                }
            }
            theFolders = (Folders)OutlookOgcs.Calendar.ReleaseObject(theFolders);
            ddMailboxName.Items.AddRange(folderIDs.Keys.ToArray());
            ddMailboxName.SelectedItem = Settings.Instance.MailboxName;

            if (ddMailboxName.SelectedIndex == -1 && ddMailboxName.Items.Count > 0) { ddMailboxName.SelectedIndex = 0; }

            log.Debug("List Calendar folders");
            cbOutlookCalendars.SelectedIndexChanged -= cbOutlookCalendar_SelectedIndexChanged;
            cbOutlookCalendars.DataSource = new BindingSource(OutlookOgcs.Calendar.Instance.CalendarFolders, null);
            cbOutlookCalendars.DisplayMember = "Key";
            cbOutlookCalendars.ValueMember = "Value";
            cbOutlookCalendars.SelectedIndex = -1; //Reset to nothing selected
            cbOutlookCalendars.SelectedIndexChanged += cbOutlookCalendar_SelectedIndexChanged;
            //Select the right calendar
            int c = 0;
            foreach (KeyValuePair<String, MAPIFolder> calendarFolder in OutlookOgcs.Calendar.Instance.CalendarFolders) {
                if (calendarFolder.Value.EntryID == Settings.Instance.UseOutlookCalendar.Id) {
                    cbOutlookCalendars.SelectedIndex = c;
                }
                c++;
            }
            if (cbOutlookCalendars.SelectedIndex == -1) cbOutlookCalendars.SelectedIndex = 0;
            #endregion
            #region Categories
            cbCategoryFilter.SelectedItem = Settings.Instance.CategoriesRestrictBy == Settings.RestrictBy.Include ?
                "Include" : "Exclude";
            if (OutlookOgcs.Factory.OutlookVersion < 12) {
                clbCategories.Items.Clear();
                cbCategoryFilter.Enabled = false;
                clbCategories.Enabled = false;
                lFilterCategories.Enabled = false;
            } else {
                OutlookOgcs.Calendar.Categories.BuildPicker(ref clbCategories);
                enableOutlookSettingsUI(true);
            }
            #endregion
            cbOnlyRespondedInvites.Checked = Settings.Instance.OnlyRespondedInvites;
            #region DateTime Format / Locale
            Dictionary<string, string> customDates = new Dictionary<string, string>();
            customDates.Add("Default", "g");
            String shortDate = System.Globalization.CultureInfo.CurrentCulture.DateTimeFormat.ShortDatePattern;
            //Outlook can't handle dates or times formatted with a . delimeter!
            switch (shortDate) {
                case "yyyy.MMdd": shortDate = "yyyy-MM-dd"; break;
                default: break;
            }
            String shortTime = System.Globalization.CultureInfo.CurrentCulture.DateTimeFormat.ShortTimePattern.Replace(".", ":");
            customDates.Add("Short Date & Time", shortDate + " " + shortTime);
            customDates.Add("Full (Short Time)", "f");
            customDates.Add("Full Month", "MMMM dd, yyyy hh:mm tt");
            customDates.Add("Generic", "yyyy-MM-dd hh:mm tt");
            customDates.Add("Custom", "yyyy-MM-dd hh:mm tt");
            cbOutlookDateFormat.DataSource = new BindingSource(customDates, null);
            cbOutlookDateFormat.DisplayMember = "Key";
            cbOutlookDateFormat.ValueMember = "Value";
            for (int i = 0; i < cbOutlookDateFormat.Items.Count; i++) {
                KeyValuePair<string, string> aFormat = (KeyValuePair<string, string>)cbOutlookDateFormat.Items[i];
                if (aFormat.Value == Settings.Instance.OutlookDateFormat) {
                    cbOutlookDateFormat.SelectedIndex = i;
                    break;
                } else if (i == cbOutlookDateFormat.Items.Count - 1 && cbOutlookDateFormat.SelectedIndex == 0) {
                    cbOutlookDateFormat.SelectedIndex = i;
                    tbOutlookDateFormat.Text = Settings.Instance.OutlookDateFormat;
                    tbOutlookDateFormat.ReadOnly = false;
                }
            }
            #endregion
            #endregion
            #region Google box
            tbConnectedAcc.Text = string.IsNullOrEmpty(Settings.Instance.GaccountEmail) ? "Not connected" : Settings.Instance.GaccountEmail;
            if (Settings.Instance.UseGoogleCalendar != null && Settings.Instance.UseGoogleCalendar.Id != null) {
                cbGoogleCalendars.Items.Add(Settings.Instance.UseGoogleCalendar);
                cbGoogleCalendars.SelectedIndex = 0;
                tbClientID.ReadOnly = true;
                tbClientSecret.ReadOnly = true;
            } else {
                tbClientID.ReadOnly = false;
                tbClientSecret.ReadOnly = false;
            }

            if (Settings.Instance.UsingPersonalAPIkeys()) {
                cbShowDeveloperOptions.Checked = true;
                tbClientID.Text = Settings.Instance.PersonalClientIdentifier;
                tbClientSecret.Text = Settings.Instance.PersonalClientSecret;
            }
            #endregion
            #region Sync Options box
            syncOptionSizing(gbSyncOptions_How, pbExpandHow, true);
            syncOptionSizing(gbSyncOptions_When, pbExpandWhen, false);
            syncOptionSizing(gbSyncOptions_What, pbExpandWhat, false);
            #region How
            syncDirection.Items.Add(Sync.Direction.OutlookToGoogle);
            syncDirection.Items.Add(Sync.Direction.GoogleToOutlook);
            syncDirection.Items.Add(Sync.Direction.Bidirectional);
            cbObfuscateDirection.Items.Add(Sync.Direction.OutlookToGoogle);
            cbObfuscateDirection.Items.Add(Sync.Direction.GoogleToOutlook);
            //Sync Direction dropdown
            for (int i = 0; i < syncDirection.Items.Count; i++) {
                Sync.Direction sd = (syncDirection.Items[i] as Sync.Direction);
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
            cbOfuscate.Checked = Settings.Instance.Obfuscation.Enabled;
            //More Options
            howObfuscatePanel.Visible = false;
            if (Settings.Instance.SyncDirection == Sync.Direction.Bidirectional) {
                tbCreatedItemsOnly.SelectedIndex = Settings.Instance.CreatedItemsOnly ? 1 : 0;
                if (Settings.Instance.TargetCalendar.Id == Sync.Direction.OutlookToGoogle.Id) tbTargetCalendar.SelectedIndex = 0;
                if (Settings.Instance.TargetCalendar.Id == Sync.Direction.GoogleToOutlook.Id) tbTargetCalendar.SelectedIndex = 1;
            } else {
                tbCreatedItemsOnly.SelectedIndex = 0;
                tbTargetCalendar.SelectedIndex = 2;
            }
            tbCreatedItemsOnly_SelectedItemChanged(null, null);
            tbTargetCalendar_SelectedItemChanged(null, null);
            cbPrivate.Checked = Settings.Instance.SetEntriesPrivate;
            cbAvailable.Checked = Settings.Instance.SetEntriesAvailable;
            cbColour.Checked = Settings.Instance.SetEntriesColour;
            foreach (Extensions.ColourPicker.ColourInfo cInfo in ddCategoryColour.Items) {
                if (cInfo.OutlookCategory.ToString() == Settings.Instance.SetEntriesColourValue &&
                    cInfo.Text == Settings.Instance.SetEntriesColourName) {
                    ddCategoryColour.SelectedItem = cInfo;
                }
            }
            ddCategoryColour.Enabled = cbColour.Checked;
            //Obfuscate Direction dropdown
            for (int i = 0; i < cbObfuscateDirection.Items.Count; i++) {
                Sync.Direction sd = (cbObfuscateDirection.Items[i] as Sync.Direction);
                if (sd.Id == Settings.Instance.Obfuscation.Direction.Id) {
                    cbObfuscateDirection.SelectedIndex = i;
                }
            }
            if (cbObfuscateDirection.SelectedIndex == -1) cbObfuscateDirection.SelectedIndex = 0;
            cbObfuscateDirection.Enabled = Settings.Instance.SyncDirection == Sync.Direction.Bidirectional;
            Settings.Instance.Obfuscation.LoadRegex(dgObfuscateRegex);
            this.gbSyncOptions_How.ResumeLayout();
            #endregion
            #region When
            this.gbSyncOptions_When.SuspendLayout();
            tbDaysInThePast.Text = Settings.Instance.DaysInThePast.ToString();
            tbDaysInTheFuture.Text = Settings.Instance.DaysInTheFuture.ToString();
            if (Settings.Instance.UsingPersonalAPIkeys()) {
                tbDaysInTheFuture.Maximum = 365*10;
                tbDaysInThePast.Maximum = 365*10;
            }
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
            cbLocation.Checked = Settings.Instance.AddLocation;
            cbAddDescription.Checked = Settings.Instance.AddDescription;
            cbAddDescription_OnlyToGoogle.Checked = Settings.Instance.AddDescription_OnlyToGoogle;
            cbAddAttendees.Checked = Settings.Instance.AddAttendees;
            cbCloakEmail.Checked = Settings.Instance.CloakEmail;
            cbCloakEmail.Visible = cbAddAttendees.Checked && Settings.Instance.SyncDirection != Sync.Direction.GoogleToOutlook;
            cbAddReminders.Checked = Settings.Instance.AddReminders;
            cbUseGoogleDefaultReminder.Checked = Settings.Instance.UseGoogleDefaultReminder;
            cbUseOutlookDefaultReminder.Checked = Settings.Instance.UseOutlookDefaultReminder;
            cbReminderDND.Enabled = Settings.Instance.AddReminders;
            cbReminderDND.Checked = Settings.Instance.ReminderDND;
            dtDNDstart.Enabled = Settings.Instance.AddReminders;
            dtDNDend.Enabled = Settings.Instance.AddReminders;
            dtDNDstart.Value = Settings.Instance.ReminderDNDstart;
            dtDNDend.Value = Settings.Instance.ReminderDNDend;
            cbAddColours.Checked = Settings.Instance.AddColours;
            this.gbSyncOptions_What.ResumeLayout();
            #endregion
            #endregion
            #region Application behaviour
            syncOptionSizing(gbAppBehaviour_Logging, pbExpandLogging, true);
            syncOptionSizing(gbAppBehaviour_Proxy, pbExpandProxy, false);
            cbShowBubbleTooltips.Checked = Settings.Instance.ShowBubbleTooltipWhenSyncing;
            cbStartOnStartup.Checked = Settings.Instance.StartOnStartup;
            tbStartupDelay.Value = Settings.Instance.StartupDelay;
            tbStartupDelay.Enabled = cbStartOnStartup.Checked;
            cbHideSplash.Checked = Settings.Instance.HideSplashScreen;
            cbSuppressSocialPopup.Checked = Settings.Instance.SuppressSocialPopup;
            cbStartInTray.Checked = Settings.Instance.StartInTray;
            cbMinimiseToTray.Checked = Settings.Instance.MinimiseToTray;
            cbMinimiseNotClose.Checked = Settings.Instance.MinimiseNotClose;
            cbPortable.Checked = Settings.Instance.Portable;
            cbPortable.Enabled = !Program.IsInstalled;
            #region Logging
            for (int i = 0; i < cbLoggingLevel.Items.Count; i++) {
                if (cbLoggingLevel.Items[i].ToString().ToLower() == Settings.Instance.LoggingLevel.ToLower()) {
                    cbLoggingLevel.SelectedIndex = i;
                    break;
                }
            }
            cbCloudLogging.CheckState = Settings.Instance.CloudLogging == null ? CheckState.Indeterminate : (CheckState)(Convert.ToInt16((bool)Settings.Instance.CloudLogging));
            cbCreateFiles.Checked = Settings.Instance.CreateCSVFiles;
            #endregion
            updateGUIsettings_Proxy();
            #endregion
            linkTShoot_logfile.Text = log4net.GlobalContext.Properties["LogFilename"] + " file";
            #region About
            int r = 0;
            dgAbout.Rows.Add();
            dgAbout.Rows[r].Cells[0].Value = "Version";
            dgAbout.Rows[r].Cells[1].Value = System.Windows.Forms.Application.ProductVersion;
            dgAbout.Rows.Add(); r++;
            dgAbout.Rows[r].Cells[0].Value = "Running From";
            dgAbout.Rows[r].Cells[1].Value = System.Windows.Forms.Application.ExecutablePath;
            dgAbout.Rows.Add(); r++;
            dgAbout.Rows[r].Cells[0].Value = "Config In";
            dgAbout.Rows[r].Cells[1].Value = Settings.ConfigFile;
            dgAbout.Rows.Add(); r++;
            dgAbout.Rows[r].Cells[0].Value = "Subscription";
            dgAbout.Rows[r].Cells[1].Value = (Settings.Instance.Subscribed == DateTime.Parse("01-Jan-2000")) ? "N/A" : Settings.Instance.Subscribed.ToShortDateString();
            dgAbout.Rows.Add(); r++;
            dgAbout.Rows[r].Cells[0].Value = "Timezone DB";
            dgAbout.Rows[r].Cells[1].Value = TimezoneDB.Instance.Version;
            dgAbout.Height = (dgAbout.Rows[r].Height * (r + 1)) + 2;

            this.lAboutMain.Text = this.lAboutMain.Text.Replace("20xx",
                (new DateTime(2000, 1, 1).Add(new TimeSpan(TimeSpan.TicksPerDay * System.Reflection.Assembly.GetEntryAssembly().GetName().Version.Build))).Year.ToString());

            cbAlphaReleases.Checked = Settings.Instance.AlphaReleases;
            #endregion
            FeaturesBlockedByCorpPolicy(Settings.Instance.OutlookGalBlocked);
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
            tbBrowserAgent.Text = Settings.Instance.Proxy.BrowserUserAgent;
            tbBrowserAgent.Enabled = rbProxyCustom.Checked;
            btCheckBrowserAgent.Enabled = rbProxyCustom.Checked;

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

        public void FeaturesBlockedByCorpPolicy(Boolean isTrue) {
            String tooltip = "Your corporate policy is blocking the ability to use this feature.";
            try {
                ToolTips.SetToolTip(cbAddAttendees, isTrue ? tooltip : "BE AWARE: Deleting Google event through mobile calendar app will notify all attendees.");
                ToolTips.SetToolTip(cbAddDescription, isTrue ? tooltip : "");
                ToolTips.SetToolTip(rbOutlookSharedCal, isTrue ? tooltip : "");
            } catch (System.InvalidOperationException ex) {
                if (OGCSexception.GetErrorCode(ex) == "0x80131509") { //Cross-thread operation
                    log.Warn("Can't set form tooltips from sync thread.");
                    //Won't worry too much - will work fine on OGCS startup, and will only arrive here if GAL has been blocked *after* startup. Should be very unlikely.
                }
            }
            if (isTrue) {
                SetControlPropertyThreadSafe(cbAddDescription, "Checked", false);
                SetControlPropertyThreadSafe(cbAddAttendees, "Checked", false);
                SetControlPropertyThreadSafe(rbOutlookSharedCal, "Checked", false);
                //Mimic appearance of disabled control - but can't disable else tooltip doesn't work
                cbAddAttendees.ForeColor = SystemColors.GrayText;
                cbAddDescription.ForeColor = SystemColors.GrayText;
                rbOutlookSharedCal.ForeColor = SystemColors.GrayText;
            } else {
                cbAddAttendees.ForeColor = SystemColors.ControlText;
                cbAddDescription.ForeColor = SystemColors.ControlText;
                rbOutlookSharedCal.ForeColor = SystemColors.ControlText;
            }
        }

        private void applyProxy() {
            if (rbProxyNone.Checked) Settings.Instance.Proxy.Type = rbProxyNone.Tag.ToString();
            else if (rbProxyCustom.Checked) Settings.Instance.Proxy.Type = rbProxyCustom.Tag.ToString();
            else Settings.Instance.Proxy.Type = rbProxyIE.Tag.ToString();

            if (rbProxyCustom.Checked) {
                if (String.IsNullOrEmpty(txtProxyServer.Text) || String.IsNullOrEmpty(txtProxyPort.Text)) {
                    MessageBox.Show("A proxy server name and port must be provided.", "Proxy Authentication Enabled",
                        MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
                int nPort;
                if (!int.TryParse(txtProxyPort.Text, out nPort)) {
                    MessageBox.Show("Proxy server port must be a number.", "Invalid Proxy Port",
                        MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                Settings.Instance.Proxy.BrowserUserAgent = tbBrowserAgent.Text;

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

        public void Sync_Click(object sender, EventArgs e) {
            try {
                Sync.Engine.Instance.Sync_Requested(sender, e);
            } catch (System.AggregateException ex) {
                OGCSexception.AnalyseAggregate(ex, false);
            } catch (System.ApplicationException ex) {
                if (ex.Message.ToLower().Contains("try again") && sender != null) {
                    Sync_Click(null, null);
                }
            } catch (System.Exception ex) {
                console.UpdateWithError("Problem encountered during synchronisation.", ex);
                OGCSexception.Analyse(ex, true);
            } finally {
                if (!Sync.Engine.Instance.SyncingNow) {
                    bSyncNow.Text = "Start Sync";
                    NotificationTray.UpdateItem("sync", "&Sync Now");
                }
            }
        }

        public enum SyncNotes {
            QuotaExhaustedInfo,
            RecentSubscription,
            SubscriptionPendingExpire,
            SubscriptionExpired,
            NotLogFile
        }
        public void SyncNote(SyncNotes syncNote, Object extraData, Boolean show = true) {
            if (!this.tbSyncNote.Visible && !show) return; //Already hidden

            String note = "";
            String url = "";
            String urlStub = "https://www.paypal.com/cgi-bin/webscr?cmd=_s-xclick&hosted_button_id=E595EQ7SNDBHA&item_name=";
            String cr = "\r\n";
            switch (syncNote) {
                case SyncNotes.QuotaExhaustedInfo:
                    note =  "  Google's daily free calendar quota is exhausted!" + cr +
                            "     Either wait for new quota at 08:00GMT or     " + cr +
                            "  get yourself guaranteed quota for just £1/month.";
                    url = urlStub + "OGCS Premium for " + Settings.Instance.GaccountEmail;
                    break;
                case SyncNotes.RecentSubscription:
                    note =  "                                                  " + cr +
                            "   Thank you for your subscription and support!   " + cr +
                            "                                                  ";
                    break;
                case SyncNotes.SubscriptionPendingExpire:
                    DateTime expiration = (DateTime)extraData;
                    note =  "  Your annual subscription for guaranteed quota   " + cr +
                            "  for Google calendar usage is expiring on " + expiration.ToString("dd-MMM") + "." + cr +
                            "         Click to renew for just £1/month.        ";
                    url = urlStub + "OGCS Premium renewal from " + expiration.ToString("dd-MMM-yy") + " for " + Settings.Instance.GaccountEmail;
                    break;
                case SyncNotes.SubscriptionExpired:
                    expiration = (DateTime)extraData;
                    note =  "  Your annual subscription for guaranteed quota   " + cr +
                            "    for Google calendar usage expired on " + expiration.ToString("dd-MMM") + "." + cr +
                            "         Click to renew for just £1/month.        ";
                    url = urlStub + "OGCS Premium renewal for " + Settings.Instance.GaccountEmail;
                    break;
                case SyncNotes.NotLogFile:
                    note =  "                       This is not the log file. " + cr +
                            "                                     --------- " + cr +
                            "  Click here to open the folder with OGcalsync.log ";
                    url = "file://" + Program.UserFilePath;
                    break;
            }
            String existingNote = GetControlPropertyThreadSafe(tbSyncNote, "Text") as String;
            if (note != existingNote.Replace("\n", "\r\n") && !show) return; //Trying to hide a note that isn't currently displaying
            SetControlPropertyThreadSafe(tbSyncNote, "Text", note);
            SetControlPropertyThreadSafe(tbSyncNote, "Tag", url);
            SetControlPropertyThreadSafe(tbSyncNote, "Visible", show);
            SetControlPropertyThreadSafe(panelSyncNote, "Visible", show);
        }

        #region Accessors
        public String NextSyncVal {
            get { return GetControlPropertyThreadSafe(lNextSyncVal, "Text").ToString(); }
            set { SetControlPropertyThreadSafe(lNextSyncVal, "Text", value); }
        }
        public String LastSyncVal {
            get { return lLastSyncVal.Text; }
            set { lLastSyncVal.Text = value; }
        }
        public void StrikeOutNextSyncVal(Boolean strikeout) {
            lNextSyncVal.Font = new Font(lNextSyncVal.Font, strikeout ? FontStyle.Strikeout : FontStyle.Regular);
        }
        #endregion

        #region EVENTS
        #region Form actions
        /// <summary>
        /// Navigates up the parents of a control to the first TabControl control
        /// </summary>
        private static Control findFocusedTab(Control control) {
            Control parentControl = control.Parent as Control;
            while (parentControl != null && !(control is TabControl)) {
                control = control.Parent;
                parentControl = control.Parent;
            }
            return control;
        }

        /// <summary>
        /// Detect when F1 is pressed for help
        /// </summary>
        protected override bool ProcessCmdKey(ref Message msg, Keys keyData) {
            try {
                if (keyData == Keys.F1) {
                    try {
                        log.Fine("Active control: " + this.ActiveControl.ToString());

                        Control focusedTab = null;
                        Control focusedPage = null;

                        focusedTab = findFocusedTab(this.ActiveControl);

                        if (focusedTab is TabControl)
                            focusedPage = (focusedTab as TabControl).SelectedTab;

                        if (focusedPage == null) {
                            System.Diagnostics.Process.Start("https://phw198.github.io/OutlookGoogleCalendarSync/guide");
                            return true;
                        }

                        if (focusedPage.Name == "tabPage_Sync")
                            System.Diagnostics.Process.Start("https://phw198.github.io/OutlookGoogleCalendarSync/guide/sync");

                        else if (focusedPage.Name == "tabPage_Settings") {
                            if (this.tabAppSettings.SelectedTab.Name == "tabOutlook")
                                System.Diagnostics.Process.Start("https://phw198.github.io/OutlookGoogleCalendarSync/guide/outlook");
                            else if (this.tabAppSettings.SelectedTab.Name == "tabGoogle")
                                System.Diagnostics.Process.Start("https://phw198.github.io/OutlookGoogleCalendarSync/guide/google");
                            else if (this.tabAppSettings.SelectedTab.Name == "tabSyncOptions")
                                System.Diagnostics.Process.Start("https://phw198.github.io/OutlookGoogleCalendarSync/guide/syncoptions");
                            else if (this.tabAppSettings.SelectedTab.Name == "tabAppBehaviour")
                                System.Diagnostics.Process.Start("https://phw198.github.io/OutlookGoogleCalendarSync/guide/appbehaviour");
                            else
                                System.Diagnostics.Process.Start("https://phw198.github.io/OutlookGoogleCalendarSync/guide/settings");

                        } else if (focusedPage.Name == "tabPage_Help")
                            System.Diagnostics.Process.Start("https://phw198.github.io/OutlookGoogleCalendarSync/guide/help");

                        else if (focusedPage.Name == "tabPage_About")
                            System.Diagnostics.Process.Start("https://phw198.github.io/OutlookGoogleCalendarSync/guide/about");

                        else
                            System.Diagnostics.Process.Start("https://phw198.github.io/OutlookGoogleCalendarSync/guide");

                        return true; //This keystroke was handled, don't pass to the control with the focus

                    } catch (System.Exception ex) {
                        log.Warn("Failed to process captured F1 key.");
                        OGCSexception.Analyse(ex);
                        System.Diagnostics.Process.Start("https://phw198.github.io/OutlookGoogleCalendarSync/guide");
                        return true;
                    }
                }

            } catch (System.Exception ex) {
                log.Warn("Failed to process captured command key.");
                OGCSexception.Analyse(ex);
            }

            return base.ProcessCmdKey(ref msg, keyData);
        }

        Boolean shiftKeyPressed = false;
        private void tabApp_KeyDown(object sender, KeyEventArgs e) {
            if (e.Shift && bSyncNow.Text == "Start Sync") {
                bSyncNow.Text = "Start Full Sync";
                shiftKeyPressed = true;
            }
        }

        private void tabApp_KeyUp(object sender, KeyEventArgs e) {
            if (shiftKeyPressed && bSyncNow.Text == "Start Full Sync") {
                bSyncNow.Text = "Start Sync";
                shiftKeyPressed = false;
            }
        }

        void Save_Click(object sender, EventArgs e) {
            if (tbStartupDelay.Value != Settings.Instance.StartupDelay) {
                Settings.Instance.StartupDelay = Convert.ToInt32(tbStartupDelay.Value);
                if (cbStartOnStartup.Checked) Program.ManageStartupRegKey(true);
            }
            bSave.Enabled = false;
            bSave.Text = "Saving...";
            try {
                Settings.Instance.Save();
                Settings.Instance.LogSettings();
                bSave.Enabled = true;
                bSave.Text = "Saved";
                DateTime saved = DateTime.Now;
                while (saved.AddSeconds(2) > DateTime.Now) {
                    System.Threading.Thread.Sleep(250);
                    System.Windows.Forms.Application.DoEvents();
                }
            } finally {
                bSave.Enabled = true;
                bSave.Text = "Save";
            }
        }

        public void MainFormShow() {
            this.Show(); //Show minimised back in taskbar
            this.ShowInTaskbar = true;
            this.WindowState = FormWindowState.Normal;
            this.Show(); //Now restore
        }

        private void mainFormResize(object sender, EventArgs e) {
            if (Settings.Instance.MinimiseToTray && this.WindowState == FormWindowState.Minimized) {
                this.ShowInTaskbar = false;
                this.Hide();
                if (Settings.Instance.ShowBubbleWhenMinimising) {
                    NotificationTray.ShowBubbleInfo("OGCS is still running.\r\nClick here to disable this notification.", tagValue: "ShowBubbleWhenMinimising");
                } else {
                    trayIcon.Tag = "";
                }
            }
        }

        #region Anti "Log" File
        //Try and stop people pasting the sync summary text as their log file!!!
        private void Console_KeyDown(object sender, PreviewKeyDownEventArgs e) {
            if (e.KeyData == (Keys.Control | Keys.C) || e.KeyData == (Keys.Control | Keys.A)) {
                if (e.KeyData == (Keys.Control | Keys.A))
                    consoleWebBrowser.Document.ExecCommand("SelectAll", false, null);
                if (e.KeyData == (Keys.Control | Keys.C) && consoleWebBrowser.Document.Body.InnerText != null)
                    Clipboard.SetText(consoleWebBrowser.Document.Body.InnerText);
                notLogFile();
            }
        }

        private void notLogFile() {
            SyncNote(SyncNotes.NotLogFile, null);
            System.Windows.Forms.Application.DoEvents();
            for (int i = 0; i <= 50; i++) {
                System.Threading.Thread.Sleep(100);
                System.Windows.Forms.Application.DoEvents();
            }
            SyncNote(SyncNotes.NotLogFile, null, false);
        }
        #endregion

        private void cbVerboseOutput_CheckedChanged(object sender, EventArgs e) {
            Settings.Instance.VerboseOutput = cbVerboseOutput.Checked;
        }

        private void cbMuteClicks_CheckedChanged(object sender, EventArgs e) {
            Settings.Instance.MuteClickSounds = cbMuteClicks.Checked;

            if (Sync.Engine.Instance.SyncingNow)
                Console.MuteClicks(cbMuteClicks.Checked);
        }

        private void tbSyncNote_Click(object sender, EventArgs e) {
            if (!String.IsNullOrEmpty(tbSyncNote.Tag.ToString())) {
                if (tbSyncNote.Tag.ToString().EndsWith("for ")) {
                    log.Info("User wanted to subscribe, but Google account username is not known :(");
                    DialogResult authorise = MessageBox.Show("Thank you for your interest in subscribing. " +
                       "To kick things off, you'll need to re-authorise OGCS to manage your Google calendar. " +
                       "Would you like to do that now?", "Proceed with authorisation?", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                    if (authorise == DialogResult.Yes) {
                        log.Debug("Resetting Google account access.");
                        GoogleOgcs.Calendar.Instance.Authenticator.Reset();
                        GoogleOgcs.Calendar.Instance.Authenticator.UserSubscriptionCheck();
                    }
                } else {
                    System.Diagnostics.Process.Start(tbSyncNote.Tag.ToString());
                }
            }
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
        private void enableOutlookSettingsUI(Boolean enable) {
            this.clbCategories.Enabled = enable;
            this.cbOutlookCalendars.Enabled = enable;
            this.ddMailboxName.Enabled = enable;
        }

        public void rbOutlookDefaultMB_CheckedChanged(object sender, EventArgs e) {
            if (!this.Visible) return;

            if (rbOutlookDefaultMB.Checked) {
                enableOutlookSettingsUI(false);
                Settings.Instance.OutlookService = OutlookOgcs.Calendar.Service.DefaultMailbox;
                OutlookOgcs.Calendar.Instance.Reset();
                //Update available calendars
                cbOutlookCalendars.DataSource = new BindingSource(OutlookOgcs.Calendar.Instance.CalendarFolders, null);
                refreshCategories();
            }
        }

        private void rbOutlookAltMB_CheckedChanged(object sender, EventArgs e) {
            if (!this.Visible) return;

            if (rbOutlookAltMB.Checked) {
                enableOutlookSettingsUI(false);
                Settings.Instance.OutlookService = OutlookOgcs.Calendar.Service.AlternativeMailbox;
                Settings.Instance.MailboxName = ddMailboxName.Text;
                OutlookOgcs.Calendar.Instance.Reset();
                //Update available calendars
                cbOutlookCalendars.DataSource = new BindingSource(OutlookOgcs.Calendar.Instance.CalendarFolders, null);
                refreshCategories();
            }
            Settings.Instance.MailboxName = (rbOutlookAltMB.Checked ? ddMailboxName.Text : "");
        }

        private void rbOutlookSharedCal_CheckedChanged(object sender, EventArgs e) {
            if (!this.Visible) return;

            if (rbOutlookSharedCal.Checked && Settings.Instance.OutlookGalBlocked) {
                rbOutlookSharedCal.Checked = false;
                return;
            }
            if (rbOutlookSharedCal.Checked) {
                enableOutlookSettingsUI(false);
                Settings.Instance.OutlookService = OutlookOgcs.Calendar.Service.SharedCalendar;
                OutlookOgcs.Calendar.Instance.Reset();
                //Update available calendars
                cbOutlookCalendars.DataSource = new BindingSource(OutlookOgcs.Calendar.Instance.CalendarFolders, null);
                refreshCategories();
            }
        }

        private void ddMailboxName_SelectedIndexChanged(object sender, EventArgs e) {
            if (this.Visible && Settings.Instance.MailboxName != ddMailboxName.Text) {
                rbOutlookAltMB.Checked = true;
                Settings.Instance.MailboxName = ddMailboxName.Text;
                enableOutlookSettingsUI(false);
                OutlookOgcs.Calendar.Instance.Reset();
                refreshCategories();
            }
        }

        public void cbOutlookCalendar_SelectedIndexChanged(object sender, EventArgs e) {
            KeyValuePair<String, MAPIFolder> calendar = (KeyValuePair<String, MAPIFolder>)cbOutlookCalendars.SelectedItem;
            OutlookOgcs.Calendar.Instance.UseOutlookCalendar = calendar.Value;
        }

        #region Categories
        private void cbCategoryFilter_SelectedIndexChanged(object sender, EventArgs e) {
            if (!this.Visible) return;
            Settings.Instance.CategoriesRestrictBy = (cbCategoryFilter.SelectedItem.ToString() == "Include") ?
                Settings.RestrictBy.Include : Settings.RestrictBy.Exclude;
            //Invert selection
            for (int i = 0; i < clbCategories.Items.Count; i++) {
                clbCategories.SetItemChecked(i, !clbCategories.CheckedIndices.Contains(i));
            }
            clbCategories_SelectedIndexChanged(null, null);
        }

        private void clbCategories_SelectedIndexChanged(object sender, EventArgs e) {
            if (!this.Visible) return;

            Settings.Instance.Categories.Clear();
            foreach (object item in clbCategories.CheckedItems) {
                Settings.Instance.Categories.Add(item.ToString());
            }
        }

        private void refreshCategories() {
            OutlookOgcs.Calendar.Instance.IOutlook.RefreshCategories();
            OutlookOgcs.Calendar.Categories.BuildPicker(ref clbCategories);
            enableOutlookSettingsUI(true);
        }

        private void miCatRefresh_Click(object sender, EventArgs e) {
            refreshCategories();
        }
        private void miCatSelectNone_Click(object sender, EventArgs e) {
            for (int i = 0; i < clbCategories.Items.Count; i++) {
                clbCategories.SetItemCheckState(i, CheckState.Unchecked);
            }
            clbCategories_SelectedIndexChanged(null, null);
        }
        private void miCatSelectAll_Click(object sender, EventArgs e) {
            for (int i = 0; i < clbCategories.Items.Count; i++) {
                clbCategories.SetItemCheckState(i, CheckState.Checked);
            }
            clbCategories_SelectedIndexChanged(null, null);
        }
        #endregion

        private void cbOnlyRespondedInvites_CheckedChanged(object sender, EventArgs e) {
            Settings.Instance.OnlyRespondedInvites = cbOnlyRespondedInvites.Checked;
        }

        #region Datetime Format
        private void cbOutlookDateFormat_SelectedIndexChanged(object sender, EventArgs e) {
            KeyValuePair<string, string> selectedFormat = (KeyValuePair<string, string>)cbOutlookDateFormat.SelectedItem;
            if (selectedFormat.Key != "Custom") {
                tbOutlookDateFormat.Text = selectedFormat.Value;
                if (this.Visible) Settings.Instance.OutlookDateFormat = tbOutlookDateFormat.Text;
            }
            tbOutlookDateFormat.ReadOnly = (selectedFormat.Key != "Custom");
        }

        private void tbOutlookDateFormat_TextChanged(object sender, EventArgs e) {
            try {
                tbOutlookDateFormatResult.Text = DateTime.Now.ToString(tbOutlookDateFormat.Text);
            } catch (System.FormatException) {
                tbOutlookDateFormatResult.Text = "Not a valid date format";
            }
        }

        private void tbOutlookDateFormat_Leave(object sender, EventArgs e) {
            if (String.IsNullOrEmpty(tbOutlookDateFormat.Text) || tbOutlookDateFormatResult.Text == "Not a valid date format") {
                cbOutlookDateFormat.SelectedIndex = 0;
            }
            Settings.Instance.OutlookDateFormat = tbOutlookDateFormat.Text;
        }

        private void btTestOutlookFilter_Click(object sender, EventArgs e) {
            log.Debug("Testing the Outlook filter string.");
            int filterCount = OutlookOgcs.Calendar.Instance.FilterCalendarEntries(OutlookOgcs.Calendar.Instance.UseOutlookCalendar.Items, false).Count();
            String msg = "The format '" + tbOutlookDateFormat.Text + "' returns " + filterCount + " calendar items within the date range ";
            msg += Settings.Instance.SyncStart.ToString(System.Globalization.CultureInfo.CurrentCulture.DateTimeFormat.ShortDatePattern);
            msg += " and " + Settings.Instance.SyncEnd.ToString(System.Globalization.CultureInfo.CurrentCulture.DateTimeFormat.ShortDatePattern);

            log.Info(msg);
            MessageBox.Show(msg, "Date-Time Format Results", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void urlDateFormats_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e) {
            System.Diagnostics.Process.Start("https://msdn.microsoft.com/en-us/library/az4se3k1%28v=vs.90%29.aspx");
        }
        #endregion
        #endregion
        #region Google settings
        private void GetMyGoogleCalendars_Click(object sender, EventArgs e) {
            if (bGetGoogleCalendars.Text == "Cancel retrieval") {
                log.Warn("User cancelled retrieval of Google calendars.");
                GoogleOgcs.Calendar.Instance.Authenticator.CancelTokenSource.Cancel();
                return;
            }

            this.bGetGoogleCalendars.Text = "Cancel retrieval";
            cbGoogleCalendars.Enabled = false;
            List<GoogleCalendarListEntry> calendars = null;
            try {
                calendars = GoogleOgcs.Calendar.Instance.GetCalendars();
            } catch (AggregateException agex) {
                OGCSexception.AnalyseAggregate(agex, false);
            } catch (Google.Apis.Auth.OAuth2.Responses.TokenResponseException ex) {
                OGCSexception.AnalyseTokenResponse(ex, false);
            } catch (System.Exception ex) {
                OGCSexception.Analyse(ex);
                MessageBox.Show("Failed to retrieve Google calendars.\r\n" +
                    "Please check the output on the Sync tab for more details.", "Google calendar retrieval failed",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
                StringBuilder sb = new StringBuilder();
                console.BuildOutput("Unable to get the list of Google calendars. The following error occurred:", ref sb, false);
                if (ex is ApplicationException && ex.InnerException != null && ex.InnerException is Google.GoogleApiException) {
                    console.BuildOutput(ex.Message, ref sb, false);
                    console.Update(sb, Console.Markup.fail, logit: true);
                } else {
                    console.BuildOutput(OGCSexception.FriendlyMessage(ex), ref sb, false);
                    console.Update(sb, Console.Markup.error, logit: true);
                    if (Settings.Instance.Proxy.Type == "IE") {
                        if (MessageBox.Show("Please ensure you can access the internet with Internet Explorer.\r\n" +
                            "Test it now? If successful, please retry retrieving your Google calendars.",
                            "Test IE Internet Access",
                            MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes) {
                            System.Diagnostics.Process.Start("iexplore.exe", "http://www.google.com");
                        }
                    }
                }
            }
            if (calendars != null) {
                cbGoogleCalendars.Items.Clear();
                foreach (GoogleCalendarListEntry mcle in calendars) {
                    cbGoogleCalendars.Items.Add(mcle);
                }
                cbGoogleCalendars.SelectedIndex = 0;
                tbClientID.ReadOnly = true;
                tbClientSecret.ReadOnly = true;
            }

            bGetGoogleCalendars.Enabled = true;
            cbGoogleCalendars.Enabled = true;
            bGetGoogleCalendars.Text = "Retrieve Calendars";
        }

        private void cbGoogleCalendars_SelectedIndexChanged(object sender, EventArgs e) {
            Settings.Instance.UseGoogleCalendar = (GoogleCalendarListEntry)cbGoogleCalendars.SelectedItem;
        }

        private void btResetGCal_Click(object sender, EventArgs e) {
            if (MessageBox.Show("This will reset the Google account you are using to synchronise with.\r\n" +
                "Useful if you want to start syncing to a different account.",
                "Reset Google account?", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == System.Windows.Forms.DialogResult.Yes) {
                log.Info("User requested reset of Google authentication details.");
                Settings.Instance.UseGoogleCalendar.Id = null;
                Settings.Instance.UseGoogleCalendar.Name = null;
                this.cbGoogleCalendars.Items.Clear();
                this.tbClientID.ReadOnly = false;
                this.tbClientSecret.ReadOnly = false;
                if (!GoogleOgcs.Calendar.IsInstanceNull && GoogleOgcs.Calendar.Instance.Authenticator != null)
                    GoogleOgcs.Calendar.Instance.Authenticator.Reset(reauthorise: false);
                else {
                    Settings.Instance.AssignedClientIdentifier = "";
                    Settings.Instance.GaccountEmail = "";
                    tbConnectedAcc.Text = "Not connected";
                    System.IO.File.Delete(System.IO.Path.Combine(Program.UserFilePath, GoogleOgcs.Authenticator.TokenFile));
                }
            }
        }

        #region Developer Options
        private void cbShowDeveloperOptions_CheckedChanged(object sender, EventArgs e) {
            //Toggle visibility
            gbDeveloperOptions.Visible =
            lGoogleAPIInstructions.Visible =
            llAPIConsole.Visible =
            lClientID.Visible =
            lSecret.Visible =
            tbClientID.Visible =
            tbClientSecret.Visible =
            cbShowClientSecret.Visible =
                cbShowDeveloperOptions.Checked;
        }

        private void llAPIConsole_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e) {
            System.Diagnostics.Process.Start(llAPIConsole.Text);
        }

        private void tbClientID_TextChanged(object sender, EventArgs e) {
            Settings.Instance.PersonalClientIdentifier = tbClientID.Text;
        }

        private void tbClientSecret_TextChanged(object sender, EventArgs e) {
            Settings.Instance.PersonalClientSecret = tbClientSecret.Text;
            cbShowClientSecret.Enabled = (tbClientSecret.Text != "");
        }
        private void cbShowClientSecret_CheckedChanged(object sender, EventArgs e) {
            tbClientSecret.UseSystemPasswordChar = !cbShowClientSecret.Checked;
        }
        #endregion
        #endregion
        #region Sync options
        private void syncOptionSizing(GroupBox section, PictureBox sectionImage, Boolean? expand = null) {
            int minSectionHeight = Convert.ToInt16(22 * magnification);
            Boolean expandSection = expand ?? false || section.Height - minSectionHeight <= 5;
            if (expandSection) {
                if (!(expand ?? false)) sectionImage.Image.RotateFlip(RotateFlipType.Rotate90FlipNone);
                switch (section.Name.ToString().Split('_').LastOrDefault()) {
                    case "How": section.Height = btCloseRegexRules.Visible ? 251 : 188; break;
                    case "When": section.Height = 119; break;
                    case "What": section.Height = 155; break;
                    case "Logging": section.Height = 93; break;
                    case "Proxy": section.Height = 197; break;
                }
                section.Height = Convert.ToInt16(section.Height * magnification);
            } else {
                sectionImage.Image.RotateFlip(RotateFlipType.Rotate270FlipNone);
                section.Height = minSectionHeight;
            }
            sectionImage.Refresh();

            if ("pbExpandHow|pbExpandWhen|pbExpandWhat".Contains(sectionImage.Name)) {
                gbSyncOptions_When.Top = gbSyncOptions_How.Location.Y + gbSyncOptions_How.Height + Convert.ToInt16(10 * magnification);
                pbExpandWhen.Top = gbSyncOptions_When.Top - Convert.ToInt16(2 * magnification);
                gbSyncOptions_What.Top = gbSyncOptions_When.Location.Y + gbSyncOptions_When.Height + Convert.ToInt16(10 * magnification);
                pbExpandWhat.Top = gbSyncOptions_What.Top - Convert.ToInt16(2 * magnification);

            } else if ("pbExpandLogging|pbExpandProxy".Contains(sectionImage.Name)) {
                gbAppBehaviour_Proxy.Top = gbAppBehaviour_Logging.Location.Y + gbAppBehaviour_Logging.Height + Convert.ToInt16(10 * magnification);
                pbExpandProxy.Top = gbAppBehaviour_Proxy.Top - Convert.ToInt16(2 * magnification);
            }
        }

        private void pbExpandHow_Click(object sender, EventArgs e) {
            syncOptionSizing(gbSyncOptions_How, pbExpandHow);
        }
        private void pbExpandWhen_Click(object sender, EventArgs e) {
            syncOptionSizing(gbSyncOptions_When, pbExpandWhen);
        }
        private void pbExpandWhat_Click(object sender, EventArgs e) {
            syncOptionSizing(gbSyncOptions_What, pbExpandWhat);
        }

        #region How
        private void syncDirection_SelectedIndexChanged(object sender, EventArgs e) {
            Settings.Instance.SyncDirection = (Sync.Direction)syncDirection.SelectedItem;
            if (Settings.Instance.SyncDirection == Sync.Direction.Bidirectional) {
                cbObfuscateDirection.Enabled = true;
                cbObfuscateDirection.SelectedIndex = Sync.Direction.OutlookToGoogle.Id - 1;

                tbCreatedItemsOnly.Enabled = true;

                if (tbTargetCalendar.Items.Contains("target calendar"))
                    tbTargetCalendar.Items.Remove("target calendar");
                tbTargetCalendar.SelectedIndex = 0;
                tbTargetCalendar.Enabled = cbPrivate.Checked || cbAvailable.Checked;
            } else {
                cbObfuscateDirection.Enabled = false;
                cbObfuscateDirection.SelectedIndex = Settings.Instance.SyncDirection.Id - 1;

                tbCreatedItemsOnly.Enabled = false;
                tbCreatedItemsOnly.SelectedIndex = 0;

                if (!tbTargetCalendar.Items.Contains("target calendar"))
                    tbTargetCalendar.Items.Add("target calendar");
                if (tbTargetCalendar.SelectedIndex == 2) tbTargetCalendar_SelectedItemChanged(null, null);
                tbTargetCalendar.SelectedIndex = 2;
                tbTargetCalendar.Enabled = false;
            }
            if (Settings.Instance.SyncDirection == Sync.Direction.GoogleToOutlook) {
                Sync.Engine.Instance.DeregisterForPushSync();
                this.cbOutlookPush.Checked = false;
                this.cbOutlookPush.Enabled = false;
                this.cbReminderDND.Visible = false;
                this.dtDNDstart.Visible = false;
                this.dtDNDend.Visible = false;
                this.lDNDand.Visible = false;
            }
            if (Settings.Instance.SyncDirection == Sync.Direction.OutlookToGoogle) {
                this.cbOutlookPush.Enabled = true;
                this.cbReminderDND.Visible = true;
                this.dtDNDstart.Visible = true;
                this.dtDNDend.Visible = true;
                this.lDNDand.Visible = true;
            }
            cbAddAttendees_CheckedChanged(null, null);
            cbAddReminders_CheckedChanged(null, null);
            showWhatPostit("Description");
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

        private void btObfuscateRules_Click(object sender, EventArgs e) {
            this.howObfuscatePanel.Visible = true;
            this.howMorePanel.Visible = false;
            this.btCloseRegexRules.Visible = true;
            syncOptionSizing(gbSyncOptions_How, pbExpandHow, true);
        }
        private void btCloseRegexRules_Click(object sender, EventArgs e) {
            this.btCloseRegexRules.Visible = false;
            this.howMorePanel.Visible = true;
            this.howObfuscatePanel.Visible = false;
            syncOptionSizing(gbSyncOptions_How, pbExpandHow, true);
        }
        private void gbSyncOptions_HowExpand(Boolean show, Int16 newHeight) {
            int minPanelHeight = Convert.ToInt16(50 * magnification);
            int maxPanelHeight = Convert.ToInt16(newHeight * magnification);
            this.gbSyncOptions_How.BringToFront();
            if (show) {
                while (this.gbSyncOptions_How.Height < maxPanelHeight) {
                    this.gbSyncOptions_How.Height += 2;
                    System.Windows.Forms.Application.DoEvents();
                    System.Threading.Thread.Sleep(1);
                }
                this.gbSyncOptions_How.Height = maxPanelHeight;
                this.gbSyncOptions_What.Height = 20;
            } else {
                while (this.gbSyncOptions_How.Height > minPanelHeight && this.Visible) {
                    this.gbSyncOptions_How.Height -= 2;
                    System.Windows.Forms.Application.DoEvents();
                    System.Threading.Thread.Sleep(1);
                }
                this.gbSyncOptions_How.Height = minPanelHeight;
                this.gbSyncOptions_What.Height = 112;
            }
        }

        #region More Options Panel
        private void tbCreatedItemsOnly_SelectedItemChanged(object sender, EventArgs e) {
            Settings.Instance.CreatedItemsOnly = tbCreatedItemsOnly.SelectedIndex == 1;
            if (tbCreatedItemsOnly.SelectedIndex == 0)
                lTargetSyncCondition.Text = "synced to";
            else
                lTargetSyncCondition.Text = "by sync in";
        }

        private void tbTargetCalendar_SelectedItemChanged(object sender, EventArgs e) {
            if (!this.Visible) return;

            switch (tbTargetCalendar.Text) {
                case "Google calendar": Settings.Instance.TargetCalendar = Sync.Direction.OutlookToGoogle; break;
                case "Outlook calendar": Settings.Instance.TargetCalendar = Sync.Direction.GoogleToOutlook; break;
                case "target calendar": Settings.Instance.TargetCalendar = Settings.Instance.SyncDirection; break;
            }
        }

        private void cbPrivate_CheckedChanged(object sender, EventArgs e) {
            Settings.Instance.SetEntriesPrivate = cbPrivate.Checked;
            tbTargetCalendar.Enabled = cbPrivate.Checked && Settings.Instance.SyncDirection == Sync.Direction.Bidirectional;
        }

        private void cbAvailable_CheckedChanged(object sender, EventArgs e) {
            Settings.Instance.SetEntriesAvailable = cbAvailable.Checked;
            tbTargetCalendar.Enabled = cbAvailable.Checked && Settings.Instance.SyncDirection == Sync.Direction.Bidirectional;
        }

        private void cbColour_CheckedChanged(object sender, EventArgs e) {
            Settings.Instance.SetEntriesColour = cbColour.Checked;
            ddCategoryColour.Enabled = cbColour.Checked;
            tbTargetCalendar.Enabled = cbColour.Checked && Settings.Instance.SyncDirection == Sync.Direction.Bidirectional;
        }

        private void ddCategoryColour_SelectedIndexChanged(object sender, EventArgs e) {
            if (!this.Visible) return;

            Settings.Instance.SetEntriesColourValue = ddCategoryColour.SelectedItem.OutlookCategory.ToString();
            Settings.Instance.SetEntriesColourName = ddCategoryColour.SelectedItem.Text;
        }
        #endregion

        #region Obfuscation Panel
        private void cbObfuscateDirection_SelectedIndexChanged(object sender, EventArgs e) {
            if (this.Visible)
                Settings.Instance.Obfuscation.Direction = (Sync.Direction)cbObfuscateDirection.SelectedItem;
        }

        private void dgObfuscateRegex_Leave(object sender, EventArgs e) {
            Settings.Instance.Obfuscation.SaveRegex(dgObfuscateRegex);
        }
        #endregion
        #endregion

        #region When
        public int MinSyncMinutes {
            get {
                if (System.Diagnostics.Debugger.IsAttached) return 1;
                else {
                    if (Settings.Instance.OutlookPush && Settings.Instance.SyncDirection != Sync.Direction.GoogleToOutlook)
                        return 120;
                    else
                        return 15;
                }
            }
        }

        private void tbDaysInThePast_ValueChanged(object sender, EventArgs e) {
            Settings.Instance.DaysInThePast = (int)tbDaysInThePast.Value;
            if (this.Visible && !Settings.Instance.UsingPersonalAPIkeys() && tbDaysInThePast.Value == tbDaysInThePast.Maximum) {
                this.ToolTips.Show("Limited to 1 year unless personal API keys are used. See 'Developer Options' on Google tab.", tbDaysInThePast);
            }
        }

        private void tbDaysInTheFuture_ValueChanged(object sender, EventArgs e) {
            Settings.Instance.DaysInTheFuture = (int)tbDaysInTheFuture.Value;
            if (this.Visible && !Settings.Instance.UsingPersonalAPIkeys() && tbDaysInTheFuture.Value == tbDaysInTheFuture.Maximum) {
                this.ToolTips.Show("Limited to 1 year unless personal API keys are used. See 'Developer Options' on Google tab.", tbDaysInTheFuture);
            }
        }

        private void tbMinuteOffsets_ValueChanged(object sender, EventArgs e) {
            if (!Settings.Instance.UsingPersonalAPIkeys()) {
                //Fair usage - most frequent sync interval is 2 hours when Push enabled
                tbInterval.ValueChanged -= new System.EventHandler(this.tbMinuteOffsets_ValueChanged);
                if (cbIntervalUnit.SelectedItem.ToString() == "Minutes") {
                    if ((int)tbInterval.Value < MinSyncMinutes)
                        tbInterval.Value = (tbInterval.Value < Convert.ToInt16(tbInterval.Text)) ? 0 : MinSyncMinutes;
                    else if ((int)tbInterval.Value > 120) {
                        tbInterval.Value = 3;
                        cbIntervalUnit.Text = "Hours";
                    }

                } else if (cbIntervalUnit.SelectedItem.ToString() == "Hours") {
                    if (((int)tbInterval.Value * 60) < MinSyncMinutes)
                        tbInterval.Value = (tbInterval.Value < Convert.ToInt16(tbInterval.Text)) ? 0 : (MinSyncMinutes / 60);
                }
                tbInterval.ValueChanged += new System.EventHandler(this.tbMinuteOffsets_ValueChanged);
            }

            Settings.Instance.SyncInterval = (int)tbInterval.Value;
            Sync.Engine.Instance.OgcsTimer.SetNextSync();
            NotificationTray.UpdateAutoSyncItems();
        }

        private void cbIntervalUnit_SelectedIndexChanged(object sender, EventArgs e) {
            if (cbIntervalUnit.Text == "Minutes" && (int)tbInterval.Value > 0 && (int)tbInterval.Value < MinSyncMinutes) {
                tbInterval.Value = MinSyncMinutes;
            }
            Settings.Instance.SyncIntervalUnit = cbIntervalUnit.Text;
            Sync.Engine.Instance.OgcsTimer.SetNextSync();
        }

        private void cbOutlookPush_CheckedChanged(object sender, EventArgs e) {
            Settings.Instance.OutlookPush = cbOutlookPush.Checked;
            if (this.Visible) {
                if (tbInterval.Value != 0) tbMinuteOffsets_ValueChanged(null, null);
                if (cbOutlookPush.Checked) Sync.Engine.Instance.RegisterForPushSync();
                else Sync.Engine.Instance.DeregisterForPushSync();
                NotificationTray.UpdateAutoSyncItems();
            }
        }
        #endregion

        #region What
        private void lWhatInfo_MouseHover(object sender, EventArgs e) {
            showWhatPostit("AffectedItems");
        }
        private void lWhatInfo_MouseLeave(object sender, EventArgs e) {
            showWhatPostit("Description");
        }
        private void showWhatPostit(String info) {
            switch (info) {
                case "Description": {
                        tbWhatHelp.Text = "Google event descriptions don't support rich text (RTF) and truncate at 8Kb. So make sure you REALLY want to 2-way sync descriptions!";
                        Boolean visible = (Settings.Instance.AddDescription &&
                            Settings.Instance.SyncDirection == Sync.Direction.Bidirectional);
                        WhatPostit.Visible = visible && !Settings.Instance.AddDescription_OnlyToGoogle;
                        cbAddDescription_OnlyToGoogle.Visible = visible;
                        break;
                    }
                case "AffectedItems": {
                        tbWhatHelp.Text = "Changes will only affect items synced hereon in.\r" +
                            "To update ALL items, click the Sync button whilst pressing the shift key.";
                        WhatPostit.Visible = true;
                        break;
                    }
            }
            tbWhatHelp.SelectAll();
            tbWhatHelp.SelectionAlignment = HorizontalAlignment.Center;
            tbWhatHelp.DeselectAll();
        }

        private void cbLocation_CheckedChanged(object sender, EventArgs e) {
            Settings.Instance.AddLocation = cbLocation.Checked;
        }

        private void cbAddDescription_CheckedChanged(object sender, EventArgs e) {
            if (cbAddDescription.Checked && Settings.Instance.OutlookGalBlocked) {
                cbAddDescription.Checked = false;
                return;
            }
            Settings.Instance.AddDescription = cbAddDescription.Checked;
            cbAddDescription_OnlyToGoogle.Enabled = cbAddDescription.Checked;
            showWhatPostit("Description");
        }
        private void cbAddDescription_OnlyToGoogle_CheckedChanged(object sender, EventArgs e) {
            Settings.Instance.AddDescription_OnlyToGoogle = cbAddDescription_OnlyToGoogle.Checked;
            showWhatPostit("Description");
        }

        private void cbAddReminders_CheckedChanged(object sender, EventArgs e) {
            if (this.Visible) Settings.Instance.AddReminders = cbAddReminders.Checked;
            cbUseGoogleDefaultReminder.Enabled = Settings.Instance.SyncDirection != Sync.Direction.GoogleToOutlook;
            cbUseOutlookDefaultReminder.Enabled = Settings.Instance.SyncDirection != Sync.Direction.OutlookToGoogle;
            cbReminderDND.Enabled = cbAddReminders.Checked;
            dtDNDstart.Enabled = cbAddReminders.Checked;
            dtDNDend.Enabled = cbAddReminders.Checked;
            lDNDand.Enabled = cbAddReminders.Checked;
        }
        private void cbUseGoogleDefaultReminder_CheckedChanged(object sender, EventArgs e) {
            Settings.Instance.UseGoogleDefaultReminder = cbUseGoogleDefaultReminder.Checked;
        }
        private void cbUseOutlookDefaultReminder_CheckedChanged(object sender, EventArgs e) {
            Settings.Instance.UseOutlookDefaultReminder = cbUseOutlookDefaultReminder.Checked;
        }
        private void cbReminderDND_CheckedChanged(object sender, EventArgs e) {
            Settings.Instance.ReminderDND = cbReminderDND.Checked;
        }
        private void dtDNDstart_ValueChanged(object sender, EventArgs e) {
            Settings.Instance.ReminderDNDstart = dtDNDstart.Value;
        }
        private void dtDNDend_ValueChanged(object sender, EventArgs e) {
            Settings.Instance.ReminderDNDend = dtDNDend.Value;
        }

        private void cbAddAttendees_CheckedChanged(object sender, EventArgs e) {
            if (cbAddAttendees.Checked && Settings.Instance.OutlookGalBlocked) {
                cbAddAttendees.Checked = false;
                cbCloakEmail.Enabled = false;
                return;
            }
            if (this.Visible) Settings.Instance.AddAttendees = cbAddAttendees.Checked;
            cbCloakEmail.Visible = Settings.Instance.SyncDirection != Sync.Direction.GoogleToOutlook;
            cbCloakEmail.Enabled = cbAddAttendees.Checked;
            if (cbAddAttendees.Checked && string.IsNullOrEmpty(OutlookOgcs.Calendar.Instance.IOutlook.CurrentUserSMTP())) {
                OutlookOgcs.Calendar.Instance.IOutlook.GetCurrentUser(null);
            }
        }
        private void cbCloakEmail_CheckedChanged(object sender, EventArgs e) {
            Settings.Instance.CloakEmail = cbCloakEmail.Checked;
        }
        private void cbAddColours_CheckedChanged(object sender, EventArgs e) {
            Settings.Instance.AddColours = cbAddColours.Checked;
        }
        #endregion
        #endregion
        #region Application settings
        private void cbStartOnStartup_CheckedChanged(object sender, EventArgs e) {
            Settings.Instance.StartOnStartup = cbStartOnStartup.Checked;
            tbStartupDelay.Enabled = cbStartOnStartup.Checked;
            Program.ManageStartupRegKey();
        }

        private void cbHideSplash_CheckedChanged(object sender, EventArgs e) {
            if (!Settings.Instance.UserIsBenefactor()) {
                cbHideSplash.CheckedChanged -= cbHideSplash_CheckedChanged;
                cbHideSplash.Checked = false;
                cbHideSplash.CheckedChanged += cbHideSplash_CheckedChanged;
                ToolTips.SetToolTip(cbHideSplash, "Donate £10 or more to enable this feature.");
                ToolTips.Show(ToolTips.GetToolTip(cbHideSplash), cbHideSplash, 5000);
            }
            Settings.Instance.HideSplashScreen = cbHideSplash.Checked;
        }

        private void cbSuppressSocialPopup_CheckedChanged(object sender, EventArgs e) {
            if (!Settings.Instance.UserIsBenefactor()) {
                cbSuppressSocialPopup.CheckedChanged -= cbSuppressSocialPopup_CheckedChanged;
                cbSuppressSocialPopup.Checked = false;
                cbSuppressSocialPopup.CheckedChanged += cbSuppressSocialPopup_CheckedChanged;
                ToolTips.SetToolTip(cbSuppressSocialPopup, "Donate £10 or more to enable this feature.");
            }
            Settings.Instance.SuppressSocialPopup = cbSuppressSocialPopup.Checked;
        }

        private void cbShowBubbleTooltipsCheckedChanged(object sender, System.EventArgs e) {
            Settings.Instance.ShowBubbleTooltipWhenSyncing = cbShowBubbleTooltips.Checked;
        }

        private void cbStartInTrayCheckedChanged(object sender, System.EventArgs e) {
            Settings.Instance.StartInTray = cbStartInTray.Checked;
        }

        private void cbMinimiseToTrayCheckedChanged(object sender, System.EventArgs e) {
            Settings.Instance.MinimiseToTray = cbMinimiseToTray.Checked;
        }

        private void cbMinimiseNotCloseCheckedChanged(object sender, System.EventArgs e) {
            Settings.Instance.MinimiseNotClose = cbMinimiseNotClose.Checked;
        }

        private void cbPortable_CheckedChanged(object sender, EventArgs e) {
            if (this.Visible) {
                if (Program.StartedWithFileArgs)
                    MessageBox.Show("It is not possible to change portability of OGCS when it is started with command line parameters.",
                        "Cannot change portability", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                else {
                    Settings.Instance.Portable = cbPortable.Checked;
                    Program.MakePortable(cbPortable.Checked);
                }
            }
        }

        private void pbExpandLogging_Click(object sender, EventArgs e) {
            syncOptionSizing(gbAppBehaviour_Logging, pbExpandLogging);
        }

        private void pbExpandProxy_Click(object sender, EventArgs e) {
            syncOptionSizing(gbAppBehaviour_Proxy, pbExpandProxy);
        }

        private void cbCreateFiles_CheckedChanged(object sender, EventArgs e) {
            Settings.Instance.CreateCSVFiles = cbCreateFiles.Checked;
        }

        private void cbLoggingLevel_SelectedIndexChanged(object sender, EventArgs e) {
            Settings.configureLoggingLevel(this.cbLoggingLevel.Text);
            Settings.Instance.LoggingLevel = this.cbLoggingLevel.Text.ToUpper();
        }

        private void btLogLocation_Click(object sender, EventArgs e) {
            try {
                log4net.Appender.IAppender[] appenders = log.Logger.Repository.GetAppenders();
                String logFileLocation = (((log4net.Appender.FileAppender)appenders[0]).File);
                logFileLocation = logFileLocation.Substring(0, logFileLocation.LastIndexOf("\\"));
                System.Diagnostics.Process.Start(@logFileLocation);
            } catch {
                System.Diagnostics.Process.Start(@Program.UserFilePath);
            }
        }

        private void cbCloudLogging_CheckStateChanged(object sender, EventArgs e) {
            if (cbCloudLogging.CheckState == CheckState.Indeterminate)
                Settings.Instance.CloudLogging = null;
            else
                Settings.Instance.CloudLogging = cbCloudLogging.Checked;
        }

        #region Proxy
        private void rbProxyCustom_CheckedChanged(object sender, EventArgs e) {
            bool result = rbProxyCustom.Checked;
            txtProxyServer.Enabled = result;
            txtProxyPort.Enabled = result;
            tbBrowserAgent.Enabled = result;
            btCheckBrowserAgent.Enabled = result;
            cbProxyAuthRequired.Enabled = result;
            if (result) {
                result = !string.IsNullOrEmpty(txtProxyUser.Text) && !string.IsNullOrEmpty(txtProxyPassword.Text);
                cbProxyAuthRequired.Checked = result;
                txtProxyUser.Enabled = result;
                txtProxyPassword.Enabled = result;
            }
        }

        private void btCheckBrowserAgent_Click(object sender, EventArgs e) {
            try {
                System.Diagnostics.Process.Start("https://phw198.github.io/OutlookGoogleCalendarSync/browseruseragent");
            } catch (System.Exception ex) {
                OGCSexception.Analyse("Failed to check browser's user agent.", ex);
            }
        }

        private void cbProxyAuthRequired_CheckedChanged(object sender, EventArgs e) {
            bool result = cbProxyAuthRequired.Checked;
            Settings.Instance.Proxy.AuthenticationRequired = result;
            this.txtProxyPassword.Enabled = result;
            this.txtProxyUser.Enabled = result;
        }

        private void gbProxy_Leave(object sender, EventArgs e) {
            applyProxy();
        }
        #endregion
        #endregion

        #region Help
        private void linkTShoot_loglevel_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e) {
            this.tabApp.SelectedTab = this.tabPage_Settings;
            this.tabAppSettings.SelectedTab = this.tabAppSettings.TabPages["tabAppBehaviour"];
        }

        private void linkTShoot_issue_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e) {
            System.Diagnostics.Process.Start("https://github.com/phw198/OutlookGoogleCalendarSync/issues");
        }

        private void linkTShoot_logfile_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e) {
            this.btLogLocation_Click(null, null);
        }
        #endregion

        #region About
        private void dgAbout_CellContentDoubleClick(object sender, DataGridViewCellEventArgs e) {
            try {
                if (dgAbout[1, 1] == dgAbout.CurrentCell || dgAbout[1, 2] == dgAbout.CurrentCell) {
                    String path = dgAbout.CurrentCell.Value.ToString();
                    path = path.Substring(0, path.LastIndexOf("\\"));
                    System.Diagnostics.Process.Start(path);
                }
            } catch { }
        }

        private void lAboutURL_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e) {
            System.Diagnostics.Process.Start(lAboutURL.Text);
        }

        private void pbDonate_Click(object sender, EventArgs e) {
            Program.Donate();
        }

        private void btCheckForUpdate_Click(object sender, EventArgs e) {
            Program.Updater.CheckForUpdate(btCheckForUpdate);
        }
        private void cbAlphaReleases_CheckedChanged(object sender, EventArgs e) {
            if (this.Visible)
                Settings.Instance.AlphaReleases = cbAlphaReleases.Checked;
        }
        #endregion
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

        #region Social Media
        public void CheckSyncMilestone() {
            try {
                if (Settings.Instance.SuppressSocialPopup && Settings.Instance.UserIsBenefactor()) return;

                Boolean isMilestone = false;
                Int32 syncs = Settings.Instance.CompletedSyncs;
                String blurb = "You've completed " + String.Format("{0:n0}", syncs) + " syncs! Why not let people know how useful this tool is...";

                lMilestone.Text = String.Format("{0:n0}", syncs) + " Syncs!";
                lMilestoneBlurb.Text = blurb;

                switch (syncs) {
                    case 10: isMilestone = true; break;
                    case 100: isMilestone = true; break;
                    case 250: isMilestone = true; break;
                    case 500: isMilestone = true; break;
                    case 1000: isMilestone = true; break;
                    case 5000: isMilestone = true; break;
                    case 10000: isMilestone = true; break;
                }
                if (isMilestone) {
                    new Forms.Social().Show();
                }
            } catch (System.Exception ex) {
                log.Warn("Failed checking sync milestone.");
                OGCSexception.Analyse(ex);
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
