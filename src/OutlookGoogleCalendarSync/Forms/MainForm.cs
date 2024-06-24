﻿using Ogcs = OutlookGoogleCalendarSync;
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
        public Boolean LoadingProfileConfig { get; private set; }

        public Main(string startingTab = null) {
            log.Debug("Initialiasing MainForm.");
            InitializeComponent();
            //MinimumSize is set in Designer to stop it keep messing around with the width
            //Then unsetting here, so the scrollbars can reduce width if necessary
            gbGoogle_GAccount.MinimumSize =
            gbGoogle_GConfig.MinimumSize =
            gbGoogle_GOAuth.MinimumSize =
            gbSyncOptions_How.MinimumSize =
            gbSyncOptions_When.MinimumSize =
            gbSyncOptions_What.MinimumSize =
            gbAppBehaviour_Proxy.MinimumSize =
            gbAppBehaviour_Logging.MinimumSize = new System.Drawing.Size(0, 0);

            Program.Updater.CheckForUpdate();

            if (startingTab != null && startingTab == "Help") this.tabApp.SelectedTab = this.tabPage_Help;

            Instance = this;

            console = new Console(consoleWebBrowser);
            Telemetry.TrackVersions();
            updateGUIsettings();
            Settings.Instance.LogSettings();
            NotificationTray = new NotificationTray(this.trayIcon);

            log.Debug("Initialise the timer(s) for the auto synchronisation");
            Settings.Instance.Calendars.ForEach(cal => { cal.InitialiseTimer(); cal.RegisterForPushSync(); });

            if (Settings.Instance.StartInTray) {
                if (!this.IsHandleCreated) this.CreateHandle();
                this.WindowState = FormWindowState.Minimized;
            }
            if (((Sync.Engine.Instance.NextSyncDate ?? DateTime.Now.AddMinutes(10)) - DateTime.Now).TotalMinutes > 5) {
                Outlook.Calendar.Disconnect(onlyWhenNoGUI: true);
            }
            while (!Forms.Splash.BeenAndGone) {
                System.Threading.Thread.Sleep(100);
            }
        }

        private void updateGUIsettings() {
            log.Debug("Configuring main form components.");
            this.Text += (string.IsNullOrEmpty(Program.Title) ? "" : " - " + Program.Title);

            this.SuspendLayout();
            #region Tooltips
            //set up tooltips for some controls
            ToolTips = new ToolTip {
                AutoPopDelay = 10000,
                InitialDelay = 500,
                ReshowDelay = 200,
                ShowAlways = true
            };

            //Outlook
            ToolTips.SetToolTip(cbOutlookCalendars,
                "The Outlook calendar to synchonize with.");
            ToolTips.SetToolTip(cbDeleteWhenCatExcl,
                "If items are already synced in Google and subsequently excluded by a category filter.");
            ToolTips.SetToolTip(btTestOutlookFilter,
                "Check how many appointments are returned for the date range being synced.");

            //Google
            ToolTips.SetToolTip(cbGoogleCalendars,
                "The Google calendar to synchonize with.");
            ToolTips.SetToolTip(btResetGCal,
                "Disconnect the Google account being used to synchonize with.");
            ToolTips.SetToolTip(cbListHiddenGcals,
                "Include hidden calendars in the above drop down.");
            ToolTips.SetToolTip(cbColourFilter,
                "Individual and entire series. Not for occurrences of series.");
            ToolTips.SetToolTip(cbDeleteWhenColourExcl,
                "If items are already synced in Outlook and subsequently excluded by a colour filter.");
            ToolTips.SetToolTip(cbAddGMeet,
                "Sync conference details embedded in Outlook appointment body.");

            //Settings
            ToolTips.SetToolTip(tbInterval,
                "Set to zero to disable automated syncs");
            ToolTips.SetToolTip(rbOutlookAltMB,
                "Only choose this if you need to use an Outlook Calendar that is not in the default mailbox");
            ToolTips.SetToolTip(cbMergeItems,
                "If the destination calendar has pre-existing items, don't delete them");
            ToolTips.SetToolTip(cbIgnoreBusyEntries,
                "Ignore Calendar Entries Named \"Busy\"");
            ToolTips.SetToolTip(cbOutlookPush,
                "Synchronise changes in Outlook to Google within a few minutes.");
            ToolTips.SetToolTip(btCloseRegexRules,
                "Close obfuscation rules.");
            ToolTips.SetToolTip(cbOfuscate,
                "Mask specified words in calendar item subject.\nTakes effect for new or updated calendar items.");
            ToolTips.SetToolTip(cbUseGoogleDefaultReminder,
                "If the calendar settings in Google have a default reminder configured, use this when Outlook has no reminder.");
            ToolTips.SetToolTip(cbUseOutlookDefaultReminder,
                "If the calendar settings in Outlook have a default reminder configured, use this when Google has no reminder.");
            ToolTips.SetToolTip(cbAddDescription_OnlyToGoogle,
                "Helps avoid data loss due to Google's 8Kb limit.");
            ToolTips.SetToolTip(cbAddAttendees,
                "BE AWARE: Deleting Google event through mobile/web calendar app will notify all attendees.");
            ToolTips.SetToolTip(tbMaxAttendees,
                "Only sync attendees if total fewer than this number. Google allows up to 200 attendees.");
            ToolTips.SetToolTip(cbCloakEmail,
                "Google has been known to send meeting updates to attendees without your consent.\n" +
                "This option safeguards against that by appending '" + Ogcs.Google.EventAttendee.EmailCloak + "' to their email address.");
            ToolTips.SetToolTip(cbSingleCategoryOnly,
                "Only allow a single Outlook category - ie 1:1 sync with Google.\n" +
                "Otherwise, for multiple categories and only one synced with OGCS, manually prefix the category name(s) with \"OGCS \".");
            ToolTips.SetToolTip(cbReminderDND,
                "Do Not Disturb: Don't sync reminders to Google if they will trigger between these times.");
            ToolTips.SetToolTip(cbExcludeSubject,
                "Supports use of regular expressions.");

            //Application behaviour
            ToolTips.SetToolTip(cbStartOnStartup, "Start OGCS when current Windows user logs in.");
            ToolTips.SetToolTip(tbStartupDelay, "Try setting a delay if COM errors occur on startup.");
            ToolTips.SetToolTip(cbStartOnStartupAllUsers, "Also try this if 'current user' isn't effective.");
            if (!Settings.Instance.UserIsBenefactor()) {
                ToolTips.SetToolTip(cbHideSplash, "Donate £10 or more to enable this feature.");
                ToolTips.SetToolTip(cbSuppressSocialPopup, "Donate £10 or more to enable this feature.");
            }
            ToolTips.SetToolTip(cbPortable,
                "For ZIP deployments, store configuration files in the application folder (useful if running from a USB thumb drive).\n" +
                "Default is in your User roaming profile.");
            ToolTips.SetToolTip(cbTelemetryDisabled, "Prevent OGCS sending anonymous usage statistics.");
            ToolTips.SetToolTip(cbCreateFiles,
                "If checked, all entries found in Outlook/Google and identified for creation/deletion will be exported \n" +
                "to CSV files in the application's directory (named \"*.csv\"). \n" +
                "Only for debug/diagnostic purposes.");
            ToolTips.SetToolTip(rbProxyIE,
                "If IE settings have been changed, a restart of the Sync application may be required");
            ToolTips.SetToolTip(cbMuteClicks, "Mute any sounds when sync summary updates.");
            #endregion

            #region Profile
            log.Debug("Loading profiles.");
            foreach (SettingsStore.Calendar calendar in Settings.Instance.Calendars) {
                ddProfile.Items.Add(calendar._ProfileName);
            }
            ddProfile.SelectedIndex = 0;
            #endregion

            #region Sync
            if (ActiveCalendarProfile.ExtirpateOgcsMetadata) {
                bSyncNow.FlatStyle = FlatStyle.Flat;
                bSyncNow.BackColor = System.Drawing.Color.PaleVioletRed;
                console.Update("<b>An advanced setting has been enabled.</b><br>If you perform a sync, it will remove all OGCS metadata from your calendar items within the synced date range, " +
                    "but it will <i>not</i> remove the actual calendar items themselves.<br>This can be useful if you wish to 'reset' your calendars to a state similar to before you ever used OGCS.",
                    Console.Markup.warning);
            }
            cbVerboseOutput.Checked = Settings.Instance.VerboseOutput;
            cbMuteClicks.Checked = Settings.Instance.MuteClickSounds;
            #endregion
            UpdateGUIsettings_Profile();
            #region Application behaviour
            groupboxSizing(gbAppBehaviour_Logging, pbExpandLogging, true);
            groupboxSizing(gbAppBehaviour_Proxy, pbExpandProxy, false);
            cbShowSystemNotifications.Checked = Settings.Instance.ShowSystemNotifications;
            cbShowSystemNotificationsIfChange.Enabled = Settings.Instance.ShowSystemNotifications;
            cbShowSystemNotificationsIfChange.Checked = Settings.Instance.ShowSystemNotificationsIfChange;
            cbStartOnStartup.Checked = Settings.Instance.StartOnStartup;
            cbStartOnStartupAllUsers.Enabled = Settings.Instance.StartOnStartup;
            cbStartOnStartupAllUsers.Checked = Settings.Instance.StartOnStartupAllUsers;
            tbStartupDelay.Value = Settings.Instance.StartupDelay;
            tbStartupDelay.Enabled = cbStartOnStartup.Checked;
            cbHideSplash.Checked = Settings.Instance.HideSplashScreen ?? false;
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
            cbAnonymiseLogs.Checked = Settings.Instance.AnonymiseLogs;
            cbTelemetryDisabled.Checked = Settings.Instance.TelemetryDisabled;
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
            dgAbout.Rows[r].Cells[1].Value = (Settings.Instance.Subscribed <= Ogcs.Google.Authenticator.SubscribedBefore) ? "N/A" : Settings.Instance.Subscribed.ToShortDateString();
            dgAbout.Rows.Add(); r++;
            dgAbout.Rows[r].Cells[0].Value = "Timezone DB";
            dgAbout.Rows[r].Cells[1].Value = TimezoneDB.Instance.Version;
            dgAbout.Height = (dgAbout.Rows[r].Height * (r + 1)) + 2;

            this.lAboutMain.Text = this.lAboutMain.Text.Replace("20xx",
                (new DateTime(2000, 1, 1).Add(new TimeSpan(TimeSpan.TicksPerDay * System.Reflection.Assembly.GetEntryAssembly().GetName().Version.Build))).Year.ToString());

            cbAlphaReleases.Checked = Settings.Instance.AlphaReleases;
            #endregion
            FeaturesBlockedByCorpPolicy(ActiveCalendarProfile.OutlookGalBlocked);
            this.ResumeLayout();
            Settings.AreApplied = true;
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

        public void UpdateGUIsettings_Profile() {
            if (ActiveCalendarProfile == null) {
                log.Warn("No Profile active yet!");
                return;

            } else {
                SettingsStore.Calendar profile = ActiveCalendarProfile;
                this.LoadingProfileConfig = true;
                try {
                    #region Profile
                    ProfileVal = profile._ProfileName;
                    LastSyncVal = profile.LastSyncDateText;
                    NextSyncVal = profile.OgcsTimer?.NextSyncDateText;
                    #endregion
                    #region Outlook page
                    #region Mailbox
                    if (Outlook.Factory.OutlookVersionName == Outlook.Factory.OutlookVersionNames.Outlook2003) {
                        rbOutlookDefaultMB.Checked = true;
                        rbOutlookAltMB.Enabled = false;
                        rbOutlookSharedCal.Enabled = false;
                    } else {
                        if (profile.OutlookService == Outlook.Calendar.Service.AlternativeMailbox) {
                            if (rbOutlookAltMB.Checked) {
                                //Toggle check to force refresh of calendar dropdowns
                                rbOutlookAltMB.CheckedChanged -= new System.EventHandler(this.rbOutlookAltMB_CheckedChanged);
                                rbOutlookAltMB.Checked = false;
                                rbOutlookAltMB.CheckedChanged += new System.EventHandler(this.rbOutlookAltMB_CheckedChanged);
                            }
                            rbOutlookAltMB.Checked = true;
                        } else if (profile.OutlookService == Outlook.Calendar.Service.SharedCalendar) {
                            if (rbOutlookSharedCal.Checked) {
                                //Toggle check to force refresh of calendar dropdowns
                                rbOutlookSharedCal.CheckedChanged -= new System.EventHandler(this.rbOutlookSharedCal_CheckedChanged);
                                rbOutlookSharedCal.Checked = false;
                                rbOutlookSharedCal.CheckedChanged += new System.EventHandler(this.rbOutlookSharedCal_CheckedChanged);
                            }
                            rbOutlookSharedCal.Checked = true;
                        } else {
                            rbOutlookDefaultMB.Checked = true;
                        }
                    }

                    //Mailboxes the user has access to
                    log.Debug("Find calendar folders");
                    if (Outlook.Calendar.Instance.Folders.Count == 1) {
                        rbOutlookAltMB.Enabled = false;
                        rbOutlookAltMB.Checked = false;
                    }
                    Folders theFolders = Outlook.Calendar.Instance.Folders;
                    Dictionary<String, List<String>> folderIDs = new Dictionary<String, List<String>>();
                    for (int fld = 1; fld <= theFolders.Count; fld++) {
                        MAPIFolder theFolder = theFolders[fld];
                        try {
                            //Create a dictionary of folder names and a list of their ID(s)
                            if (!folderIDs.ContainsKey(theFolder.Name)) {
                                folderIDs.Add(theFolder.Name, new List<String>(new String[] { theFolder.EntryID }));
                            } else if (!folderIDs[theFolder.Name].Contains(theFolder.EntryID)) {
                                folderIDs[theFolder.Name].Add(theFolder.EntryID);
                            }
                        } catch (System.Exception ex) {
                            ex.LogAsFail().Analyse("Failed to get EntryID for folder: " + theFolder.Name);
                        } finally {
                            theFolder = (MAPIFolder)Outlook.Calendar.ReleaseObject(theFolder);
                        }
                    }
                    ddMailboxName.Items.Clear();
                    ddMailboxName.Items.AddRange(folderIDs.Keys.ToArray());
                    ddMailboxName.SelectedItem = profile.MailboxName;

                    if (ddMailboxName.SelectedIndex == -1 && ddMailboxName.Items.Count > 0) {
                        if (profile.OutlookService == Outlook.Calendar.Service.AlternativeMailbox && string.IsNullOrEmpty(profile.MailboxName))
                            log.Warn("Could not find mailbox '" + profile.MailboxName + "' in Alternate Mailbox dropdown. Defaulting to the first in the list.");

                        ddMailboxName.SelectedIndexChanged -= new System.EventHandler(this.ddMailboxName_SelectedIndexChanged);
                        ddMailboxName.SelectedIndex = 0;
                        ddMailboxName.SelectedIndexChanged += new System.EventHandler(this.ddMailboxName_SelectedIndexChanged);
                    }

                    log.Debug("List Calendar folders");
                    cbOutlookCalendars.SelectedIndexChanged -= cbOutlookCalendar_SelectedIndexChanged;
                    cbOutlookCalendars.DataSource = new BindingSource(Outlook.Calendar.Instance.CalendarFolders, null);
                    cbOutlookCalendars.DisplayMember = "Key";
                    cbOutlookCalendars.ValueMember = "Value";
                    cbOutlookCalendars.SelectedIndex = -1; //Reset to nothing selected
                    cbOutlookCalendars.SelectedIndexChanged += cbOutlookCalendar_SelectedIndexChanged;
                    //Select the right calendar
                    int c = 0;
                    foreach (KeyValuePair<String, MAPIFolder> calendarFolder in Outlook.Calendar.Instance.CalendarFolders) {
                        if (calendarFolder.Value.EntryID == profile.UseOutlookCalendar.Id) {
                            cbOutlookCalendars.SelectedIndex = c;
                            break;
                        }
                        c++;
                    }
                    if (cbOutlookCalendars.SelectedIndex == -1) {
                        if (!string.IsNullOrEmpty(profile.UseOutlookCalendar.Id)) {
                            log.Warn("Outlook calendar '" + profile.UseOutlookCalendar.Name + "' could no longer be found. Selected calendar '" + Outlook.Calendar.Instance.CalendarFolders.First().Key + "' instead.");
                            Ogcs.Extensions.MessageBox.Show("The Outlook calendar '" + profile.UseOutlookCalendar.Name + "' previously configured for syncing is no longer available.\r\n\r\n" +
                                "'" + Outlook.Calendar.Instance.CalendarFolders.First().Key + "' calendar has been selected instead and any automated syncs have been temporarily disabled.",
                                "Outlook Calendar Unavailable", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                            profile.SyncInterval = 0;
                            profile.OutlookPush = false;
                            Forms.Main.Instance.tabApp.SelectTab("tabPage_Settings");
                        }
                        cbOutlookCalendars.SelectedIndex = 0;
                    }
                    #endregion
                    #region Categories
                    cbCategoryFilter.SelectedItem = profile.CategoriesRestrictBy == SettingsStore.Calendar.RestrictBy.Include ?
                    "Include" : "Exclude";
                    if (Outlook.Factory.OutlookVersionName == Outlook.Factory.OutlookVersionNames.Outlook2003) {
                        clbCategories.Items.Clear();
                        clbCategories.Items.Add("Outlook 2003 has no categories");
                        cbCategoryFilter.Enabled = false;
                        clbCategories.Enabled = false;
                        lFilterCategories.Enabled = false;
                        btColourMap.Visible = false;
                        profile.AddColours = false;
                        cbAddColours.Enabled = false;
                    } else {
                        Outlook.Calendar.Categories.BuildPicker(ref clbCategories);
                        enableOutlookSettingsUI(true);
                    }
                    cbDeleteWhenCatExcl.Checked = profile.DeleteWhenCategoryExcluded;
                    #endregion
                    cbOnlyRespondedInvites.Checked = profile.OnlyRespondedInvites;
                    btCustomTzMap.Visible = Settings.Instance.TimezoneMaps.Count != 0;
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
                        if (aFormat.Value == profile.OutlookDateFormat) {
                            cbOutlookDateFormat.SelectedIndex = i;
                            break;
                        } else if (i == cbOutlookDateFormat.Items.Count - 1 && cbOutlookDateFormat.SelectedIndex == 0) {
                            cbOutlookDateFormat.SelectedIndex = i;
                            tbOutlookDateFormat.Text = profile.OutlookDateFormat;
                            tbOutlookDateFormat.ReadOnly = false;
                        }
                    }
                    #endregion
                    #endregion
                    #region Google page
                    groupboxSizing(gbGoogle_GAccount, pbExpandGoogleAccount, true);
                    groupboxSizing(gbGoogle_GConfig, pbExpandGoogleConfig, true);
                    groupboxSizing(gbGoogle_GOAuth, pbExpandGoogleOauth, false);

                    tbConnectedAcc.Text = string.IsNullOrEmpty(Settings.Instance.GaccountEmail) ? "Not connected" : Settings.Instance.GaccountEmail;
                    if (profile.UseGoogleCalendar?.Id != null) {
                        foreach (GoogleCalendarListEntry cle in this.cbGoogleCalendars.Items) {
                            if (cle.Id == profile.UseGoogleCalendar.Id) {
                                this.cbGoogleCalendars.SelectedItem = cle;
                                break;
                            }
                        }
                        if (cbGoogleCalendars.SelectedIndex == -1 || (cbGoogleCalendars.SelectedItem as GoogleCalendarListEntry).Id != profile.UseGoogleCalendar.Id) {
                            cbGoogleCalendars.Items.Add(profile.UseGoogleCalendar);
                            cbGoogleCalendars.SelectedIndex = cbGoogleCalendars.Items.Count - 1;
                        }
                        tbClientID.ReadOnly = true;
                        tbClientSecret.ReadOnly = true;
                    } else {
                        tbClientID.ReadOnly = false;
                        tbClientSecret.ReadOnly = false;
                    }

                    cbColourFilter.SelectedItem = profile.ColoursRestrictBy == SettingsStore.Calendar.RestrictBy.Include ? "Include" : "Exclude";
                    Ogcs.Google.Calendar.BuildOfflineColourPicker(clbColours);
                    cbDeleteWhenColourExcl.Checked = profile.DeleteWhenColourExcluded;
                    cbExcludeDeclinedInvites.Checked = profile.ExcludeDeclinedInvites;
                    cbExcludeGoals.Checked = profile.ExcludeGoals;
                    cbExcludeGoals.Enabled = Ogcs.Google.Calendar.IsDefaultCalendar() ?? true;
                    cbAddGMeet.Checked = profile.AddGMeet;
                    
                    if (Settings.Instance.UsingPersonalAPIkeys()) {
                        cbShowDeveloperOptions.Checked = true;
                        tbClientID.Text = Settings.Instance.PersonalClientIdentifier;
                        tbClientSecret.Text = Settings.Instance.PersonalClientSecret;
                    }
                    #endregion
                    #region Sync Options page
                    groupboxSizing(gbSyncOptions_How, pbExpandHow, true);
                    groupboxSizing(gbSyncOptions_When, pbExpandWhen, false);
                    groupboxSizing(gbSyncOptions_What, pbExpandWhat, false);
                    #region How
                    if (syncDirection.Items.Count == 0) {
                        syncDirection.Items.Add(Sync.Direction.OutlookToGoogle);
                        syncDirection.Items.Add(Sync.Direction.GoogleToOutlook);
                        syncDirection.Items.Add(Sync.Direction.Bidirectional);
                        cbObfuscateDirection.Items.Add(Sync.Direction.OutlookToGoogle);
                        cbObfuscateDirection.Items.Add(Sync.Direction.GoogleToOutlook);
                    }
                    //Sync Direction dropdown
                    for (int i = 0; i < syncDirection.Items.Count; i++) {
                        Sync.Direction sd = (syncDirection.Items[i] as Sync.Direction);
                        if (sd.Id == profile.SyncDirection.Id) {
                            syncDirection.SelectedIndex = i;
                            break;
                        }
                    }
                    if (syncDirection.SelectedIndex == -1) syncDirection.SelectedIndex = 0;
                    this.gbSyncOptions_How.SuspendLayout();
                    cbMergeItems.Checked = profile.MergeItems;
                    cbIgnoreBusyEntries.Checked = profile.IgnoreBusy;
                    cbDisableDeletion.Checked = profile.DisableDelete;
                    cbConfirmOnDelete.Enabled = !profile.DisableDelete;
                    cbConfirmOnDelete.Checked = profile.ConfirmOnDelete;
                    cbOfuscate.Checked = profile.Obfuscation.Enabled;
                    howObfuscatePanel.Visible = false;
                    if (profile.SyncDirection.Id == Sync.Direction.Bidirectional.Id) {
                        tbCreatedItemsOnly.SelectedIndex = profile.CreatedItemsOnly ? 1 : 0;
                        if (profile.TargetCalendar.Id == Sync.Direction.OutlookToGoogle.Id) tbTargetCalendar.SelectedIndex = 0;
                        if (profile.TargetCalendar.Id == Sync.Direction.GoogleToOutlook.Id) tbTargetCalendar.SelectedIndex = 1;
                    } else {
                        tbCreatedItemsOnly.SelectedIndex = 0;
                        tbTargetCalendar.SelectedIndex = 2;
                    }
                    tbCreatedItemsOnly_SelectedItemChanged(null, null);
                    tbTargetCalendar_SelectedItemChanged(null, null);

                    cbPrivate.Checked = profile.SetEntriesPrivate;
                    ddPrivacy.Enabled = profile.SetEntriesPrivate;
                    ddPrivacy.DataSource = null;
                    ddPrivacy.DisplayMember = "Value";
                    ddPrivacy.ValueMember = "Key";
                    ddPrivacy.Items.Clear();
                    Dictionary<OlSensitivity, String> privacy = new Dictionary<OlSensitivity, String>();
                    privacy.Add(OlSensitivity.olPrivate, "Private");
                    privacy.Add(OlSensitivity.olNormal, "Public");
                    ddPrivacy.DataSource = new BindingSource(privacy, null);
                    ddPrivacy.SelectedValue = Enum.Parse(typeof(OlSensitivity), profile.PrivacyLevel);

                    cbAvailable.Checked = profile.SetEntriesAvailable;
                    buildAvailabilityDropdown();
                    cbColour.Checked = profile.SetEntriesColour;
                    ddOutlookColour.AddColourItems();

                    ddOutlookColour.SelectedIndexChanged -= ddOutlookColour_SelectedIndexChanged;
                    foreach (Outlook.Categories.ColourInfo cInfo in ddOutlookColour.Items) {
                        if (cInfo.OutlookCategory.ToString() == profile.SetEntriesColourValue &&
                            cInfo.Text == profile.SetEntriesColourName) {
                            ddOutlookColour.SelectedItem = cInfo;
                            break;
                        }
                    }
                    if (ddOutlookColour.SelectedIndex == -1 && ddOutlookColour.Items.Count > 0)
                        ddOutlookColour.SelectedIndex = 0;

                    ddOutlookColour.SelectedIndexChanged += ddOutlookColour_SelectedIndexChanged;
                    ddOutlookColour.Enabled = cbColour.Checked;

                    ddGoogleColour.SelectedIndexChanged -= ddGoogleColour_SelectedIndexChanged;
                    offlineAddGoogleColour();
                    ddGoogleColour.SelectedIndexChanged += ddGoogleColour_SelectedIndexChanged;
                    ddGoogleColour.Enabled = cbColour.Checked;

                    //Obfuscate Direction dropdown
                    for (int i = 0; i < cbObfuscateDirection.Items.Count; i++) {
                        Sync.Direction sd = (cbObfuscateDirection.Items[i] as Sync.Direction);
                        if (sd.Id == profile.Obfuscation.Direction.Id) {
                            cbObfuscateDirection.SelectedIndex = i;
                            break;
                        }
                    }
                    if (cbObfuscateDirection.SelectedIndex == -1) cbObfuscateDirection.SelectedIndex = 0;
                    cbObfuscateDirection.Enabled = profile.SyncDirection.Id == Sync.Direction.Bidirectional.Id;
                    profile.Obfuscation.LoadRegex(dgObfuscateRegex);
                    this.gbSyncOptions_How.ResumeLayout();
                    #endregion
                    #region When
                    this.gbSyncOptions_When.SuspendLayout();
                    tbDaysInThePast.Text = profile.DaysInThePast.ToString();
                    tbDaysInTheFuture.Text = profile.DaysInTheFuture.ToString();
                    setMaxSyncRange();
                    tbInterval.ValueChanged -= new System.EventHandler(this.tbMinuteOffsets_ValueChanged);
                    tbInterval.Value = profile.SyncInterval;
                    tbInterval.ValueChanged += new System.EventHandler(this.tbMinuteOffsets_ValueChanged);
                    cbIntervalUnit.SelectedIndexChanged -= new System.EventHandler(this.cbIntervalUnit_SelectedIndexChanged);
                    cbIntervalUnit.Text = profile.SyncIntervalUnit;
                    cbIntervalUnit.SelectedIndexChanged += new System.EventHandler(this.cbIntervalUnit_SelectedIndexChanged);
                    cbOutlookPush.Checked = profile.OutlookPush;
                    this.gbSyncOptions_When.ResumeLayout();
                    #endregion
                    #region What
                    this.gbSyncOptions_What.SuspendLayout();
                    cbLocation.Checked = profile.AddLocation;
                    cbAddDescription.Checked = profile.AddDescription;
                    cbAddDescription_OnlyToGoogle.Checked = profile.AddDescription_OnlyToGoogle;
                    cbAddAttendees.Checked = profile.AddAttendees;
                    tbMaxAttendees.Value = profile.MaxAttendees;
                    cbCloakEmail.Checked = profile.CloakEmail;
                    cbCloakEmail.Visible = cbAddAttendees.Checked && profile.SyncDirection.Id != Sync.Direction.GoogleToOutlook.Id;
                    cbAddReminders.Checked = profile.AddReminders;
                    cbUseGoogleDefaultReminder.Checked = profile.UseGoogleDefaultReminder;
                    cbUseOutlookDefaultReminder.Checked = profile.UseOutlookDefaultReminder;
                    cbReminderDND.Enabled = profile.AddReminders;
                    cbReminderDND.Checked = profile.ReminderDND;
                    dtDNDstart.Enabled = profile.AddReminders;
                    dtDNDend.Enabled = profile.AddReminders;
                    dtDNDstart.Value = profile.ReminderDNDstart;
                    dtDNDend.Value = profile.ReminderDNDend;
                    cbAddColours.Checked = profile.AddColours;
                    btColourMap.Enabled = profile.AddColours;
                    cbSingleCategoryOnly.Checked = profile.SingleCategoryOnly;
                    cbSingleCategoryOnly.Enabled = profile.AddColours && profile.SyncDirection.Id != Sync.Direction.OutlookToGoogle.Id;
                    cbExcludeAllDays.Checked = profile.ExcludeAllDays;
                    cbExcludeFreeAllDays.Checked = profile.ExcludeFreeAllDays;
                    cbExcludeFreeAllDays.Enabled = cbExcludeAllDays.Checked;
                    cbExcludeFree.Checked = profile.ExcludeFree;
                    cbExcludeTentative.Checked = profile.ExcludeTentative;
                    cbExcludePrivate.Checked = profile.ExcludePrivate;
                    cbExcludeSubject.Checked = profile.ExcludeSubject;
                    tbExcludeSubjectText.Text = profile.ExcludeSubjectText;
                    this.gbSyncOptions_What.ResumeLayout();
                    #endregion
                    #endregion
                } catch (System.Exception ex) {
                    ex.Analyse("Unable to set GUI profile.");
                    throw;
                } finally {
                    this.LoadingProfileConfig = false;
                }
            }
        }

        public void FeaturesBlockedByCorpPolicy(Boolean isTrue) {
            String tooltip = "Your corporate policy is blocking the ability to use this feature.";
            try {
                ToolTips.SetToolTip(cbAddAttendees, isTrue ? tooltip : "BE AWARE: Deleting Google event through mobile calendar app will notify all attendees.");
                ToolTips.SetToolTip(cbAddDescription, isTrue ? tooltip : "");
                ToolTips.SetToolTip(rbOutlookSharedCal, isTrue ? tooltip : "");
            } catch (System.InvalidOperationException ex) {
                if (ex.GetErrorCode() == "0x80131509") { //Cross-thread operation
                    log.Warn("Can't set form tooltips from sync thread.");
                    //Won't worry too much - will work fine on OGCS startup, and will only arrive here if GAL has been blocked *after* startup. Should be very unlikely.
                }
            }
            if (isTrue) {
                //Mimic appearance of disabled control - but can't disable else tooltip doesn't work
                checkboxSoftRestrict(cbAddAttendees, true);
                checkboxSoftRestrict(cbAddDescription, true);
                checkboxSoftRestrict(rbOutlookSharedCal, true);
                checkboxSoftRestrict(cbAddGMeet, true);
                //If a sync is running, disable relevant config in that profile
                SettingsStore.Calendar activeProfile = Settings.Profile.InPlay();
                if (activeProfile != null) {
                    activeProfile.AddAttendees = false;
                    activeProfile.AddDescription = false;
                    activeProfile.AddGMeet = false;
                }
            } else {
                checkboxSoftRestrict(cbAddAttendees, false);
                checkboxSoftRestrict(cbAddDescription, false);
                checkboxSoftRestrict(rbOutlookSharedCal, false);
                checkboxSoftRestrict(cbAddGMeet, false);
            }
        }

        /// <summary>
        /// Make a checkbox look disabled, but still able to show a tooltip
        /// </summary>
        /// <param name="cb">The form control</param>
        /// <param name="disable">Disable or enable</param>
        private void checkboxSoftRestrict(Control cb, Boolean disable) {
            if (disable) {
                cb.ForeColor = SystemColors.GrayText;
                SetControlPropertyThreadSafe(cb, "Checked", false);
            } else {
                cb.ForeColor = SystemColors.ControlText;
            }
        }

        private void applyProxy() {
            if (rbProxyNone.Checked) Settings.Instance.Proxy.Type = rbProxyNone.Tag.ToString();
            else if (rbProxyCustom.Checked) Settings.Instance.Proxy.Type = rbProxyCustom.Tag.ToString();
            else Settings.Instance.Proxy.Type = rbProxyIE.Tag.ToString();

            if (rbProxyCustom.Checked) {
                if (String.IsNullOrEmpty(txtProxyServer.Text) || String.IsNullOrEmpty(txtProxyPort.Text)) {
                    Ogcs.Extensions.MessageBox.Show("A proxy server name and port must be provided.", "Proxy Authentication Enabled",
                        MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
                int nPort;
                if (!int.TryParse(txtProxyPort.Text, out nPort)) {
                    Ogcs.Extensions.MessageBox.Show("Proxy server port must be a number.", "Invalid Proxy Port",
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

        private void buildAvailabilityDropdown() {
            SettingsStore.Calendar profile = Forms.Main.Instance.ActiveCalendarProfile;
            try {
                this.ddAvailabilty.SelectedIndexChanged -= new System.EventHandler(this.ddAvailabilty_SelectedIndexChanged);
                ddAvailabilty.DataSource = null;
                ddAvailabilty.DisplayMember = "Value";
                ddAvailabilty.ValueMember = "Key";
                ddAvailabilty.Items.Clear();
                Dictionary<OlBusyStatus, String> availability = new Dictionary<OlBusyStatus, String>();
                availability.Add(OlBusyStatus.olFree, "Free");
                availability.Add(OlBusyStatus.olBusy, "Busy");
                if (profile.SyncDirection.Id != Sync.Direction.OutlookToGoogle.Id && tbTargetCalendar.Text != "Google calendar") {
                    availability.Add(OlBusyStatus.olTentative, "Tentative");
                    availability.Add(OlBusyStatus.olOutOfOffice, "Out of Office");
                }
                ddAvailabilty.DataSource = new BindingSource(availability, null);
                ddAvailabilty.Enabled = profile.SetEntriesAvailable;
            } catch (System.Exception ex) {
                ex.Analyse("Failed building availability dropdown values.");
                return;
            }
            try {
                ddAvailabilty.SelectedValue = Enum.Parse(typeof(OlBusyStatus), profile.AvailabilityStatus);
            } catch (System.Exception ex) {
                ex.Analyse("Failed selecting availability dropdown value from Settings.");
            } finally {
                if (ddAvailabilty.SelectedIndex == -1 && ddAvailabilty.Items.Count > 0)
                    ddAvailabilty.SelectedIndex = 0;
                this.ddAvailabilty.SelectedIndexChanged += new System.EventHandler(this.ddAvailabilty_SelectedIndexChanged);
            }
        }

        public void Sync_Click(object sender, EventArgs e) {
            try {
                Sync.Engine.Instance.Sync_Requested(sender, e);
            } catch (System.AggregateException ex) {
                ex.AnalyseAggregate(false);
            } catch (System.ApplicationException ex) {
                if (ex.Message.ToLower().Contains("try again") && sender != null) {
                    Sync_Click(null, null);
                }
            } catch (System.Exception ex) {
                console.UpdateWithError("Problem encountered during synchronisation.", ex);
                Ogcs.Exception.Analyse(ex, true);
            } finally {
                if (!Sync.Engine.Instance.SyncingNow) {
                    bSyncNow.Text = "Start Sync";
                    NotificationTray.UpdateItem("sync", "&Sync Now");
                }
            }
        }

        private void miSyncDelta_Click(object sender, EventArgs e) {
            this.bSyncNow.Text = "Start Sync";
        }
        private void miSyncFull_Click(object sender, EventArgs e) {
            this.bSyncNow.Text = "Start Full Sync";
        }

        public enum SyncNotes {
            DailyQuotaExhaustedInfo,
            DailyQuotaExhaustedPreviously,
            QuotaExceededInfo,
            QuotaExceededPreviously,
            RecentSubscription,
            SubscriptionPendingExpire,
            SubscriptionExpired,
            NotLogFile
        }
        public void SyncNote(SyncNotes syncNote, Object extraData, Boolean show = true) {
            if (this.Visible && !this.tbSyncNote.Visible && !show) return; //Already hidden

            String note = "";
            String url = "";
            String urlStub = "https://www.paypal.com/cgi-bin/webscr?cmd=_s-xclick&hosted_button_id=E595EQ7SNDBHA&item_name=";
            String cr = "\r\n";

            if (syncNote == SyncNotes.DailyQuotaExhaustedInfo && !show && this.tbSyncNote.Text.Contains("quota is exhausted")) {
                syncNote = SyncNotes.DailyQuotaExhaustedPreviously;
                show = true;
            } else if (syncNote == SyncNotes.QuotaExceededInfo && !show && this.tbSyncNote.Text.Contains("quota is being exceeded")) {
                syncNote = SyncNotes.QuotaExceededPreviously;
                show = true;
            }
            String existingNote = GetControlPropertyThreadSafe(tbSyncNote, "Text") as String;

            switch (syncNote) {
                case SyncNotes.DailyQuotaExhaustedInfo:
                    note = "  Google's daily free calendar quota is exhausted!" + cr +
                            "     Either wait for new quota at 08:00GMT or     " + cr +
                            "  get yourself guaranteed quota for just £1/month.";
                    url = urlStub + "OGCS Premium for " + Settings.Instance.GaccountEmail;
                    break;

                case SyncNotes.DailyQuotaExhaustedPreviously:
                    DateTime utcNow = DateTime.UtcNow;
                    DateTime quotaReset = utcNow.Date.AddHours(8).AddMinutes(utcNow.Minute);
                    if ((utcNow - quotaReset).Ticks < -TimeSpan.TicksPerMinute) {
                        //Successful sync before new quota at 8GMT
                        SetControlPropertyThreadSafe(tbSyncNote, "Visible", false);
                        SetControlPropertyThreadSafe(panelSyncNote, "Visible", false);
                        show = false;
                        break;
                    }
                    int delayHours = (int)(DateTime.Now - ActiveCalendarProfile.LastSyncDate).TotalHours;
                    String delay = delayHours + " hours";
                    if (delayHours == 0) {
                        delay = (int)(DateTime.Now - ActiveCalendarProfile.LastSyncDate).TotalMinutes + " mins";
                    }
                    note = "Google's daily free calendar quota was exhausted!" + cr +
                            "    Previous successful sync was " + delay + " ago." + cr +
                            " Get yourself guaranteed quota for just £1/month.";
                    url = urlStub + "OGCS Premium for " + Settings.Instance.GaccountEmail;

                    if (!show && existingNote.Contains("free calendar quota was exhausted")) {
                        log.Debug("Removing quota exhausted advisory notice.");
                        SetControlPropertyThreadSafe(tbSyncNote, "Visible", show);
                        SetControlPropertyThreadSafe(panelSyncNote, "Visible", show);
                    } else {
                        //Display the note for 3 hours after the quota has been renewed
                        System.ComponentModel.BackgroundWorker bwHideNote = new System.ComponentModel.BackgroundWorker {
                            WorkerReportsProgress = false,
                            WorkerSupportsCancellation = true
                        };
                        bwHideNote.DoWork += new System.ComponentModel.DoWorkEventHandler(
                            delegate (object o, System.ComponentModel.DoWorkEventArgs args) {
                                try {
                                    DateTime showUntil = DateTime.Now.AddHours(3);
                                    log.Debug("Showing quota exhausted advisory until " + showUntil.ToString());
                                    while (DateTime.Now < showUntil) {
                                        System.Threading.Thread.Sleep(60 * 1000);
                                    }
                                    log.Debug("Quota exhausted advisory notice period ending.");
                                    SyncNote(SyncNotes.DailyQuotaExhaustedPreviously, null, false);
                                } catch { }
                            });
                        bwHideNote.RunWorkerAsync();
                    }
                    break;

                case SyncNotes.QuotaExceededInfo:
                    note = "  Google's free calendar quota is being exceeded! " + cr +
                            "     Either wait for new quota or get yourself    " + cr +
                            "         guaranteed quota for just £1/month.      ";
                    url = urlStub + "OGCS Premium for " + Settings.Instance.GaccountEmail;
                    break;

                case SyncNotes.QuotaExceededPreviously:
                    delayHours = (int)(DateTime.Now - ActiveCalendarProfile.LastSyncDate).TotalHours;
                    delay = delayHours + " hours";
                    if (delayHours == 0) {
                        delay = (int)(DateTime.Now - ActiveCalendarProfile.LastSyncDate).TotalMinutes + " mins";
                    }
                    note = "    Google's free calendar quota was exceeded!    " + cr +
                            "      Previous successful sync was " + delay + " ago." + cr +
                            " Get yourself guaranteed quota for just £1/month.";
                    url = urlStub + "OGCS Premium for " + Settings.Instance.GaccountEmail;

                    if (!show && existingNote.Contains("free calendar quota was exceeded")) {
                        log.Debug("Removing quota exceeded advisory notice.");
                        SetControlPropertyThreadSafe(tbSyncNote, "Visible", show);
                        SetControlPropertyThreadSafe(panelSyncNote, "Visible", show);
                    } else {
                        //Display the note for 3 hours after the quota has been renewed
                        System.ComponentModel.BackgroundWorker bwHideNote = new System.ComponentModel.BackgroundWorker {
                            WorkerReportsProgress = false,
                            WorkerSupportsCancellation = true
                        };
                        bwHideNote.DoWork += new System.ComponentModel.DoWorkEventHandler(
                            delegate (object o, System.ComponentModel.DoWorkEventArgs args) {
                                try {
                                    DateTime showUntil = DateTime.Now.AddHours(3);
                                    log.Debug("Showing quota exceeded advisory until " + showUntil.ToString());
                                    while (DateTime.Now < showUntil) {
                                        System.Threading.Thread.Sleep(60 * 1000);
                                    }
                                    log.Debug("Quota exceeded advisory notice period ending.");
                                    SyncNote(SyncNotes.QuotaExceededPreviously, null, false);
                                } catch { }
                            });
                        bwHideNote.RunWorkerAsync();
                    }
                    break;

                case SyncNotes.RecentSubscription:
                    note = "                                                  " + cr +
                            "   Thank you for your subscription and support!   " + cr +
                            "                                                  ";
                    break;

                case SyncNotes.SubscriptionPendingExpire:
                    DateTime expiration = (DateTime)extraData;
                    note = "  Your annual subscription for guaranteed quota   " + cr +
                            "  for Google calendar usage is expiring on " + expiration.ToString("dd-MMM") + "." + cr +
                            "         Click to renew for just £1/month.        ";
                    url = urlStub + "OGCS Premium renewal from " + expiration.ToString("dd-MMM-yy", new System.Globalization.CultureInfo("en-US")) +
                        " for " + Settings.Instance.GaccountEmail;
                    break;

                case SyncNotes.SubscriptionExpired:
                    expiration = (DateTime)extraData;
                    note = "  Your annual subscription for guaranteed quota   " + cr +
                            "    for Google calendar usage expired on " + expiration.ToString("dd-MMM") + "." + cr +
                            "         Click to renew for just £1/month.        ";
                    url = urlStub + "OGCS Premium renewal for " + Settings.Instance.GaccountEmail;
                    break;

                case SyncNotes.NotLogFile:
                    note = "                       This is not the log file. " + cr +
                            "                                     --------- " + cr +
                            "  Click here to open the folder with OGcalsync.log ";
                    url = "file://" + Program.UserFilePath;
                    break;
            }
            if (note != existingNote.Replace("\n", "\r\n") && !show) return; //Trying to hide a note that isn't currently displaying
            SetControlPropertyThreadSafe(tbSyncNote, "Text", note);
            SetControlPropertyThreadSafe(tbSyncNote, "Tag", url);
            SetControlPropertyThreadSafe(tbSyncNote, "Visible", show);
            SetControlPropertyThreadSafe(panelSyncNote, "Visible", show);
        }

        #region Accessors
        public String ProfileVal {
            get { return lProfileVal.Text; }
            set { SetControlPropertyThreadSafe(lProfileVal, "Text", value); }
        }
        public String NextSyncVal {
            set { SetControlPropertyThreadSafe(lNextSyncVal, "Text", value); }
        }
        public String LastSyncVal {
            set { SetControlPropertyThreadSafe(lLastSyncVal, "Text", value); }
        }
        public void StrikeOutNextSyncVal(Boolean strikeout) {
            lNextSyncVal.Font = new Font(lNextSyncVal.Font, strikeout ? FontStyle.Strikeout : FontStyle.Regular);
        }
        #endregion

        #region EVENTS
        #region Form actions
        /// <summary>
        /// Detect when F1 is pressed for help
        /// </summary>
        protected override bool ProcessCmdKey(ref Message msg, Keys keyData) {
            try {
                if (keyData == Keys.F1) {
                    try {
                        log.Fine("Active control: " + this.ActiveControl.ToString());

                        Control focusedPage = null;
                        focusedPage = Forms.Main.Instance.tabApp.SelectedTab;

                        if (focusedPage == null) {
                            Helper.OpenBrowser(Program.OgcsWebsite + "/guide");
                            return true;
                        }

                        if (focusedPage.Name == "tabPage_Sync")
                            Helper.OpenBrowser(Program.OgcsWebsite + "/guide/sync");

                        else if (focusedPage.Name == "tabPage_Settings") {
                            if (this.tabAppSettings.SelectedTab.Name == "tabOutlook")
                                Helper.OpenBrowser(Program.OgcsWebsite + "/guide/outlook");
                            else if (this.tabAppSettings.SelectedTab.Name == "tabGoogle")
                                Helper.OpenBrowser(Program.OgcsWebsite + "/guide/google");
                            else if (this.tabAppSettings.SelectedTab.Name == "tabSyncOptions")
                                Helper.OpenBrowser(Program.OgcsWebsite + "/guide/syncoptions");
                            else if (this.tabAppSettings.SelectedTab.Name == "tabAppBehaviour")
                                Helper.OpenBrowser(Program.OgcsWebsite + "/guide/appbehaviour");
                            else
                                Helper.OpenBrowser(Program.OgcsWebsite + "/guide/settings");

                        } else if (focusedPage.Name == "tabPage_Help")
                            Helper.OpenBrowser(Program.OgcsWebsite + "/guide/help");

                        else if (focusedPage.Name == "tabPage_About")
                            Helper.OpenBrowser(Program.OgcsWebsite + "/guide/about");

                        else
                            Helper.OpenBrowser(Program.OgcsWebsite + "/guide");

                        return true; //This keystroke was handled, don't pass to the control with the focus

                    } catch (System.Exception ex) {
                        log.Warn("Failed to process captured F1 key.");
                        Ogcs.Exception.Analyse(ex);
                        System.Diagnostics.Process.Start(Program.OgcsWebsite + "/guide");
                        return true;
                    }
                }

            } catch (System.Exception ex) {
                log.Warn("Failed to process captured command key.");
                Ogcs.Exception.Analyse(ex);
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

        private void Main_Load(object sender, EventArgs e) {
            this.Activate();
        }

        public void MainFormShow(Boolean forceToTop = false) {
            if (this.WindowState == FormWindowState.Minimized || !this.Visible || !this.TopMost || !this.ShowInTaskbar) {
                this.tbSyncNote.ScrollBars = RichTextBoxScrollBars.None; //Reset scrollbar
                this.Show(); //Show minimised back in taskbar
                this.WindowState = FormWindowState.Normal;
                this.ShowInTaskbar = true;
                if (forceToTop) this.TopMost = true;
                this.tbSyncNote.ScrollBars = RichTextBoxScrollBars.Vertical; //Show scrollbar if necessary
                this.Show(); //Now restore
                if (this.Location.X < 0 || this.Location.Y < 0) this.CenterToScreen();
                this.TopMost = false;
                this.Refresh();
                System.Windows.Forms.Application.DoEvents();
                log.Info("Application window restored.");
            }
        }

        private void mainFormResize(object sender, EventArgs e) {
            if (Settings.Instance.MinimiseToTray && this.WindowState == FormWindowState.Minimized) {
                log.Info("Minimising application to task tray.");
                this.ShowInTaskbar = false;
                this.Hide();
                if (Settings.Instance.ShowSystemNotificationWhenMinimising) {
                    NotificationTray.ShowBubbleInfo("OGCS is still running.\r\nClick here to disable this notification.", ToolTipIcon.Info, "ShowSystemNotificationWhenMinimising");
                } else {
                    trayIcon.Tag = "";
                }
            }
        }
        #endregion

        #region Sync
        #region Anti "Log" File
        //Try and stop people pasting the sync summary text as their log file!!!
        private void Console_KeyDown(object sender, PreviewKeyDownEventArgs e) {
            try {
                if (e.KeyData == (Keys.Control | Keys.C) || e.KeyData == (Keys.Control | Keys.A)) {
                    if (e.KeyData == (Keys.Control | Keys.A))
                        consoleWebBrowser.Document.ExecCommand("SelectAll", false, null);
                    if (e.KeyData == (Keys.Control | Keys.C) && consoleWebBrowser.Document.Body.InnerText != null)
                        Clipboard.SetText(consoleWebBrowser.Document.Body.InnerText);
                    notLogFile();
                }
            } catch (System.Exception ex) {
                ex.LogAsFail().Analyse("Console_KeyDown detected.");
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
                    DialogResult authorise = Ogcs.Extensions.MessageBox.Show("Thank you for your interest in subscribing. " +
                       "To kick things off, you'll need to re-authorise OGCS to manage your Google calendar. " +
                       "Would you like to do that now?", "Proceed with authorisation?", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                    if (authorise == DialogResult.Yes) {
                        log.Debug("Resetting Google account access.");
                        Ogcs.Google.Calendar.Instance.Authenticator.Reset();
                        Ogcs.Google.Calendar.Instance.Authenticator.UserSubscriptionCheck();
                    }
                } else {
                    if (tbSyncNote.Tag.ToString().Contains("OGCS Premium renewal")) {
                        Ogcs.Extensions.MessageBox.Show("Before renewing, please ensure you don't already have an active recurring annual payment set up in PayPal :-)",
                            "Recurring payment already configured?", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    }
                    Helper.OpenBrowser(tbSyncNote.Tag.ToString());
                }
            }
        }
        #endregion

        #region Save settings
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
            StringFormat stringFlags = new StringFormat {
                Alignment = StringAlignment.Far,
                LineAlignment = StringAlignment.Center
            };
            g.DrawString(tabPage.Text, tabFont, textBrush, tabBounds, new StringFormat(stringFlags));
        }

        private void Save_Click(object sender, EventArgs e) {
            if (tbStartupDelay.Value != Settings.Instance.StartupDelay) {
                Settings.Instance.StartupDelay = Convert.ToInt32(tbStartupDelay.Value);
                if (cbStartOnStartup.Checked) Program.ManageStartupRegKey();
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
        private void miExportSettings_Click(object sender, EventArgs e) {
            SaveFileDialog exportFile = new SaveFileDialog {
                Title = "Backup OGCS Settings to File",
                FileName = "OGCS_v" + Settings.Instance.Version + ".xml",
                Filter = "XML File|*.xml|All Files|*",
                DefaultExt = "xml",
                AddExtension = true,
                OverwritePrompt = true
            };
            if (exportFile.ShowDialog() == DialogResult.OK) {
                log.Info("Exporting settings to " + exportFile.FileName);
                Settings.Instance.Save(exportFile.FileName);
            }
        }
        private void miImportSettings_Click(object sender, EventArgs e) {
            OpenFileDialog importFile = new OpenFileDialog {
                Title = "Import OGCS Settings from File",
                Filter = "XML File|*.xml|All Files|*",
                DefaultExt = "xml",
                CheckFileExists = true,
                Multiselect = false
            };
            if (importFile.ShowDialog() == DialogResult.OK) {
                log.Info("Importing settings from " + importFile.FileName);
                Settings.Load(importFile.FileName);
                updateGUIsettings();
            }
        }
        #endregion
        #region Profile
        /// <summary>
        /// The calendar settings profile currently displayed in the GUI.
        /// </summary>
        public SettingsStore.Calendar ActiveCalendarProfile { get; internal set; }

        private void ddProfile_SelectedIndexChanged(object sender, EventArgs e) {
            foreach (SettingsStore.Calendar cal in Settings.Instance.Calendars) {
                if (cal._ProfileName == ddProfile.Text) {
                    try {
                        try {
                            if (this.tabAppSettings.SelectedTab != this.tabOutlook) {
                                this.tabAppSettings.SelectedTab.Controls.Add(this.panelObscure);
                                this.panelObscure.BringToFront();
                                this.panelObscure.Dock = DockStyle.Fill;
                                this.panelObscure.Visible = true;
                            }
                        } catch (System.Exception ex) {
                            Ogcs.Exception.Analyse(ex);
                        }
                        cal.SetActive();
                        break;
                    } finally {
                        this.panelObscure.Visible = false;
                        this.tabAppSettings.Enabled = true;
                    }
                }
            }
        }

        private void btProfileAction_Click(object sender, EventArgs e) {
            if (btProfileAction.Text.StartsWith("Add"))
                miAddProfile_Click(null, null);
            else if (btProfileAction.Text.StartsWith("Delete"))
                miDeleteProfile_Click(null, null);
            else if (btProfileAction.Text.StartsWith("Rename"))
                miRenameProfile_Click(null, null);
        }
        private void miAddProfile_Click(object sender, EventArgs e) {
            btProfileAction.Text = miAddProfile.Text;
            new Forms.ProfileManage("Add", ddProfile).ShowDialog();
        }
        private void miDeleteProfile_Click(object sender, EventArgs e) {
            btProfileAction.Text = miDeleteProfile.Text;
            if (ddProfile.Items.Count == 1) {
                MessageBox.Show("At least one profile must always exist.\nIf you don't want it to automatically sync, set the schedule value to zero.",
                    "Profile deletion", MessageBoxButtons.OK, MessageBoxIcon.Hand);
                return;
            }

            String profileName = ddProfile.Text;
            if (MessageBox.Show("Are you sure you want to remove the calendar settings for profile '" + profileName + "'?",
                "Confirm profile deletion", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No) return;

            try {
                Settings.Instance.Calendars.Remove(ActiveCalendarProfile);
                log.Info("Deleted calendar settings '" + profileName + "'.");

                ActiveCalendarProfile.DeregisterForPushSync();
                ActiveCalendarProfile.OgcsTimer.Enabled = false;
                ActiveCalendarProfile.OgcsTimer.Dispose();

                NotificationTray.RemoveProfileItem(profileName);
                ddProfile.Items.Remove(ddProfile.SelectedItem);
                ddProfile.SelectedIndex = 0;

                Settings.Instance.Save();
            } catch (System.Exception ex) {
                ex.Analyse("Failed to delete profile '" + profileName + "'.");
                throw;
            }
        }
        private void miRenameProfile_Click(object sender, EventArgs e) {
            btProfileAction.Text = miRenameProfile.Text;
            new Forms.ProfileManage("Rename", ddProfile).ShowDialog();
        }
        #endregion

        private void groupboxSizing(GroupBox section, PictureBox sectionImage, Boolean? expand = null) {
            int minSectionHeight = Convert.ToInt16(22 * magnification);
            Boolean expandSection = expand ?? section.Height - minSectionHeight <= 5;
            if (expandSection) {
                if (!(expand ?? false)) sectionImage.Image.RotateFlip(RotateFlipType.Rotate90FlipNone);
                switch (section.Name.ToString().Split('_').LastOrDefault()) {
                    //Google
                    case "GAccount": section.Height = 242; break;
                    case "GConfig": section.Height = 130; break;
                    case "GOAuth": section.Height = 174; break;
                    //Settings
                    case "How": section.Height = btCloseRegexRules.Visible ? 251 : 198; break;
                    case "When": section.Height = 119; break;
                    case "What": section.Height = 265; break;
                    //Application Behaviour
                    case "Logging": section.Height = 125; break;
                    case "Proxy": section.Height = 197; break;
                }
                section.Height = Convert.ToInt16(section.Height * magnification);
            } else {
                if (section.Height > minSectionHeight)
                    sectionImage.Image.RotateFlip(RotateFlipType.Rotate270FlipNone);
                section.Height = minSectionHeight;
            }
            sectionImage.Refresh();

            if ("pbExpandGoogleAccount|pbExpandGoogleConfig|pbExpandGoogleOauth".Contains(sectionImage.Name)) {
                gbGoogle_GConfig.Top = gbGoogle_GAccount.Location.Y + gbGoogle_GAccount.Height + Convert.ToInt16(10 * magnification);
                pbExpandGoogleConfig.Top = gbGoogle_GConfig.Top - Convert.ToInt16(2 * magnification);
                cbShowDeveloperOptions.Top = gbGoogle_GConfig.Location.Y + gbGoogle_GConfig.Height + Convert.ToInt16(5 * magnification);
                gbGoogle_GOAuth.Top = cbShowDeveloperOptions.Location.Y + cbShowDeveloperOptions.Height + Convert.ToInt16(5 * magnification);
                pbExpandGoogleOauth.Top = gbGoogle_GOAuth.Top - Convert.ToInt16(2 * magnification);

            } else if ("pbExpandHow|pbExpandWhen|pbExpandWhat".Contains(sectionImage.Name)) {
                gbSyncOptions_When.Top = gbSyncOptions_How.Location.Y + gbSyncOptions_How.Height + Convert.ToInt16(10 * magnification);
                pbExpandWhen.Top = gbSyncOptions_When.Top - Convert.ToInt16(2 * magnification);
                gbSyncOptions_What.Top = gbSyncOptions_When.Location.Y + gbSyncOptions_When.Height + Convert.ToInt16(10 * magnification);
                pbExpandWhat.Top = gbSyncOptions_What.Top - Convert.ToInt16(2 * magnification);

            } else if ("pbExpandLogging|pbExpandProxy".Contains(sectionImage.Name)) {
                gbAppBehaviour_Proxy.Top = gbAppBehaviour_Logging.Location.Y + gbAppBehaviour_Logging.Height + Convert.ToInt16(10 * magnification);
                pbExpandProxy.Top = gbAppBehaviour_Proxy.Top - Convert.ToInt16(2 * magnification);
            }
        }

        #region Outlook settings
        private void enableOutlookSettingsUI(Boolean enable) {
            this.clbCategories.Enabled = enable;
            this.cbOutlookCalendars.Enabled = enable;
            this.ddMailboxName.Enabled = rbOutlookAltMB.Checked ? enable : false;
        }

        public void rbOutlookDefaultMB_CheckedChanged(object sender, EventArgs e) {
            if (!Settings.AreApplied) return;

            if (rbOutlookDefaultMB.Checked) {
                enableOutlookSettingsUI(false);
                ActiveCalendarProfile.OutlookService = Outlook.Calendar.Service.DefaultMailbox;
                Outlook.Calendar.Instance.Reset();
                //Update available calendars
                if (LoadingProfileConfig)
                    cbOutlookCalendars.SelectedIndexChanged -= cbOutlookCalendar_SelectedIndexChanged;
                cbOutlookCalendars.DataSource = new BindingSource(Outlook.Calendar.Instance.CalendarFolders, null);
                if (LoadingProfileConfig)
                    cbOutlookCalendars.SelectedIndexChanged += cbOutlookCalendar_SelectedIndexChanged;
                refreshCategories();
            }
        }

        private void rbOutlookAltMB_CheckedChanged(object sender, EventArgs e) {
            if (!Settings.AreApplied) return;

            if (rbOutlookAltMB.Checked) {
                enableOutlookSettingsUI(false);
                ActiveCalendarProfile.OutlookService = Outlook.Calendar.Service.AlternativeMailbox;
                if (!LoadingProfileConfig)
                    ActiveCalendarProfile.MailboxName = ddMailboxName.Text;
                Outlook.Calendar.Instance.Reset();
                //Update available calendars
                if (LoadingProfileConfig)
                    cbOutlookCalendars.SelectedIndexChanged -= cbOutlookCalendar_SelectedIndexChanged;
                cbOutlookCalendars.DataSource = new BindingSource(Outlook.Calendar.Instance.CalendarFolders, null);
                if (LoadingProfileConfig)
                    cbOutlookCalendars.SelectedIndexChanged += cbOutlookCalendar_SelectedIndexChanged;
                refreshCategories();
            }
            if (!LoadingProfileConfig)
                ActiveCalendarProfile.MailboxName = (rbOutlookAltMB.Checked ? ddMailboxName.Text : "");
        }

        private void rbOutlookSharedCal_CheckedChanged(object sender, EventArgs e) {
            if (!Settings.AreApplied) return;

            if (rbOutlookSharedCal.Checked && ActiveCalendarProfile.OutlookGalBlocked) {
                rbOutlookSharedCal.Checked = false;
                return;
            }
            if (rbOutlookSharedCal.Checked) {
                enableOutlookSettingsUI(false);
                ActiveCalendarProfile.OutlookService = Outlook.Calendar.Service.SharedCalendar;
                Outlook.Calendar.Instance.Reset();
                //Update available calendars
                if (LoadingProfileConfig)
                    cbOutlookCalendars.SelectedIndexChanged -= cbOutlookCalendar_SelectedIndexChanged;
                cbOutlookCalendars.DataSource = new BindingSource(Outlook.Calendar.Instance.CalendarFolders, null);
                if (LoadingProfileConfig)
                    cbOutlookCalendars.SelectedIndexChanged += cbOutlookCalendar_SelectedIndexChanged;
                refreshCategories();
            }
        }

        private void ddMailboxName_SelectedIndexChanged(object sender, EventArgs e) {
            if (Settings.AreApplied && ActiveCalendarProfile.MailboxName != ddMailboxName.Text) {
                rbOutlookAltMB.Checked = true;
                ActiveCalendarProfile.MailboxName = ddMailboxName.Text;
                enableOutlookSettingsUI(false);
                Outlook.Calendar.Instance.Reset();
                refreshCategories();
            }
        }

        public void cbOutlookCalendar_SelectedIndexChanged(object sender, EventArgs e) {
            KeyValuePair<String, MAPIFolder> calendar = (KeyValuePair<String, MAPIFolder>)cbOutlookCalendars.SelectedItem;
            ActiveCalendarProfile.UseOutlookCalendar = new OutlookCalendarListEntry(calendar.Value);

            log.Warn("Outlook calendar selection changed to: " + ActiveCalendarProfile.UseOutlookCalendar.ToString());
        }

        #region Categories
        private void cbCategoryFilter_SelectedIndexChanged(object sender, EventArgs e) {
            if (this.LoadingProfileConfig) return;

            ActiveCalendarProfile.CategoriesRestrictBy = (cbCategoryFilter.SelectedItem.ToString() == "Include") ?
                SettingsStore.Calendar.RestrictBy.Include : SettingsStore.Calendar.RestrictBy.Exclude;
            //Invert selection
            miCatSelectInvert_Click(null, null);
        }

        private void cbDeleteWhenCatExcl_CheckedChanged(object sender, EventArgs e) {
            ActiveCalendarProfile.DeleteWhenCategoryExcluded = cbDeleteWhenCatExcl.Checked;
        }

        private void clbCategories_SelectedIndexChanged(object sender, EventArgs e) {
            if (this.LoadingProfileConfig) return;

            ActiveCalendarProfile.Categories.Clear();
            foreach (object item in clbCategories.CheckedItems) {
                ActiveCalendarProfile.Categories.Add(item.ToString());
            }
        }

        private void refreshCategories() {
            Outlook.Calendar.Instance.IOutlook.RefreshCategories();
            Outlook.Calendar.Categories.BuildPicker(ref clbCategories);
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
        private void miCatSelectInvert_Click(object sender, EventArgs e) {
            for (int i = 0; i < clbCategories.Items.Count; i++) {
                clbCategories.SetItemChecked(i, !clbCategories.CheckedIndices.Contains(i));
            }
            clbCategories_SelectedIndexChanged(null, null);
        }
        #endregion

        private void cbOnlyRespondedInvites_CheckedChanged(object sender, EventArgs e) {
            ActiveCalendarProfile.OnlyRespondedInvites = cbOnlyRespondedInvites.Checked;
        }

        private void btCustomTzMap_Click(object sender, EventArgs e) {
            new Forms.TimezoneMap().ShowDialog(this);
        }

        #region Datetime Format
        private void cbOutlookDateFormat_SelectedIndexChanged(object sender, EventArgs e) {
            KeyValuePair<string, string> selectedFormat = (KeyValuePair<string, string>)cbOutlookDateFormat.SelectedItem;
            if (selectedFormat.Key != "Custom") {
                tbOutlookDateFormat.Text = selectedFormat.Value;
                if (!this.LoadingProfileConfig) ActiveCalendarProfile.OutlookDateFormat = tbOutlookDateFormat.Text;
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
            ActiveCalendarProfile.OutlookDateFormat = tbOutlookDateFormat.Text;
        }

        private void btTestOutlookFilter_Click(object sender, EventArgs e) {
            log.Debug("Testing the Outlook filter string.");
            try {
                MAPIFolder calendar = Outlook.Calendar.Instance.IOutlook.GetFolderByID(this.ActiveCalendarProfile.UseOutlookCalendar.Id);
                int filterCount = Outlook.Calendar.Instance.FilterCalendarEntries(this.ActiveCalendarProfile, false).Count();
                Outlook.Calendar.Disconnect(true);
                String msg = "The format '" + tbOutlookDateFormat.Text + "' returns " + filterCount + " calendar items within the date range ";
                msg += ActiveCalendarProfile.SyncStart.ToString(System.Globalization.CultureInfo.CurrentCulture.DateTimeFormat.ShortDatePattern);
                msg += " and " + ActiveCalendarProfile.SyncEnd.ToString(System.Globalization.CultureInfo.CurrentCulture.DateTimeFormat.ShortDatePattern);

                log.Info(msg);
                Ogcs.Extensions.MessageBox.Show(msg, "Date-Time Format Results", MessageBoxButtons.OK, MessageBoxIcon.Information);
            } catch (System.Exception ex) {
                ex.Analyse("Profile '" + Settings.Profile.Name(ActiveCalendarProfile) + "', calendar ID " + ActiveCalendarProfile.UseOutlookCalendar.Id);
                Ogcs.Extensions.MessageBox.Show(ex.Message, "Unable to perform test", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void urlDateFormats_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e) {
            Helper.OpenBrowser("https://msdn.microsoft.com/en-us/library/az4se3k1%28v=vs.90%29.aspx");
        }
        #endregion
        #endregion
        #region Google settings
        private void pbExpandGoogleAccount_Click(object sender, EventArgs e) {
            groupboxSizing(gbGoogle_GAccount, pbExpandGoogleAccount);
        }
        private void pbExpandGoogleConfig_Click(object sender, EventArgs e) {
            groupboxSizing(gbGoogle_GConfig, pbExpandGoogleConfig);
        }
        private void pbExpandGoogleOauth_Click(object sender, EventArgs e) {
            groupboxSizing(gbGoogle_GOAuth, pbExpandGoogleOauth);
        }

        #region Google Account
        private void llMultipleOGCS_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e) {
            Helper.OpenBrowser("https://github.com/phw198/OutlookGoogleCalendarSync/wiki/Running-Multiple-Instances-of-OGCS");
        }

        private void GetMyGoogleCalendars_Click(object sender, EventArgs e) {
            if (bGetGoogleCalendars.Text == "Cancel retrieval") {
                log.Warn("User cancelled retrieval of Google calendars.");
                Ogcs.Google.Calendar.Instance.Authenticator.CancelTokenSource.Cancel();
                return;
            }

            log.Debug("Retrieving Google calendar list.");
            this.bGetGoogleCalendars.Text = "Cancel retrieval";
            try {
                Ogcs.Google.Calendar.Instance.GetCalendars();
            } catch (AggregateException agex) {
                agex.AnalyseAggregate(false);
            } catch (global::Google.Apis.Auth.OAuth2.Responses.TokenResponseException ex) {
                ex.AnalyseTokenResponse(false);
            } catch (OperationCanceledException) {
            } catch (System.Exception ex) {
                Ogcs.Exception.Analyse(ex);
                Ogcs.Extensions.MessageBox.Show("Failed to retrieve Google calendars.\r\n" +
                    "Please check the output on the Sync tab for more details.", "Google calendar retrieval failed",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
                StringBuilder sb = new StringBuilder();
                console.BuildOutput("Unable to get the list of Google calendars. The following error occurred:", ref sb, false);
                if (ex is ApplicationException && ex.InnerException != null && ex.InnerException is global::Google.GoogleApiException) {
                    console.BuildOutput(ex.Message, ref sb, false);
                    console.Update(sb, Console.Markup.fail, logit: true);
                } else {
                    console.BuildOutput(ex.FriendlyMessage(), ref sb, false);
                    console.Update(sb, Console.Markup.error, logit: true);
                    if (Settings.Instance.Proxy.Type == "IE") {
                        if (Ogcs.Extensions.MessageBox.Show("Please ensure you can access the internet with Internet Explorer.\r\n" +
                            "Test it now? If successful, please retry retrieving your Google calendars.",
                            "Test IE Internet Access",
                            MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes) {
                            System.Diagnostics.Process.Start("iexplore.exe", "http://www.google.com");
                        }
                    }
                }
            }

            cbGoogleCalendars_BuildList();

            bGetGoogleCalendars.Enabled = true;
            cbGoogleCalendars.Enabled = true;
            bGetGoogleCalendars.Text = "Retrieve Calendars";
        }

        private void cbGoogleCalendars_BuildList() {
            if (Ogcs.Google.Calendar.Instance.CalendarList.Count > 0) {
                cbGoogleCalendars.Items.Clear();
                Ogcs.Google.Calendar.Instance.CalendarList.Sort((x, y) => (x.Sorted()).CompareTo(y.Sorted()));
                foreach (GoogleCalendarListEntry mcle in Ogcs.Google.Calendar.Instance.CalendarList) {
                    if (!cbListHiddenGcals.Checked && mcle.Hidden) continue;
                    cbGoogleCalendars.Items.Add(mcle);
                    if (cbGoogleCalendars.SelectedIndex == -1 && mcle.Id == ActiveCalendarProfile.UseGoogleCalendar?.Id)
                        cbGoogleCalendars.SelectedItem = mcle;
                }
                if (cbGoogleCalendars.SelectedIndex == -1) {
                    cbGoogleCalendars.SelectedIndex = 0;
                }
                tbClientID.ReadOnly = true;
                tbClientSecret.ReadOnly = true;
            }
        }

        private void cbGoogleCalendars_SelectedIndexChanged(object sender, EventArgs e) {
            if (this.LoadingProfileConfig) return;

            ActiveCalendarProfile.UseGoogleCalendar = (GoogleCalendarListEntry)cbGoogleCalendars.SelectedItem;
            if (cbGoogleCalendars.Text.StartsWith("[Read Only]") && ActiveCalendarProfile.SyncDirection.Id != Sync.Direction.GoogleToOutlook.Id) {
                Ogcs.Extensions.MessageBox.Show("You cannot " + (ActiveCalendarProfile.SyncDirection.Id == Sync.Direction.Bidirectional.Id ? "two-way " : "") + "sync with a read-only Google calendar.\n" +
                    "Please review your calendar selection.", "Read-only Sync", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                this.tabAppSettings.SelectedTab = this.tabAppSettings.TabPages["tabGoogle"];
            }
            cbExcludeGoals.Enabled = Ogcs.Google.Calendar.IsDefaultCalendar() ?? true;
            if (sender != null)
                log.Warn("Google calendar selection changed to: " + ActiveCalendarProfile.UseGoogleCalendar.ToString(true));
        }

        private void btResetGCal_Click(object sender, EventArgs e) {
            if (Ogcs.Extensions.MessageBox.Show("This will disconnect the Google account you are using to synchronise with.\r\n" +
                "Useful if you want to start syncing to a different account.",
                "Disconnect Google account?", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == System.Windows.Forms.DialogResult.Yes) {
                log.Info("User requested reset of Google authentication details.");
                ActiveCalendarProfile.UseGoogleCalendar = new GoogleCalendarListEntry();
                this.cbGoogleCalendars.Items.Clear();
                this.tbClientID.ReadOnly = false;
                this.tbClientSecret.ReadOnly = false;
                if (!Ogcs.Google.Calendar.IsInstanceNull && Ogcs.Google.Calendar.Instance.Authenticator != null)
                    Ogcs.Google.Calendar.Instance.Authenticator.Reset(reauthorise: false);
                else {
                    Settings.Instance.AssignedClientIdentifier = "";
                    Settings.Instance.GaccountEmail = "";
                    tbConnectedAcc.Text = "Not connected";
                    System.IO.File.Delete(System.IO.Path.Combine(Program.UserFilePath, Ogcs.Google.Authenticator.TokenFile));
                }
            }
        }

        private void cbListHiddenGcals_CheckedChanged(object sender, EventArgs e) {
            cbGoogleCalendars_BuildList();
        }
        #endregion

        #region GoogleConfig
        #region Colours
        private void cbColourFilter_SelectedIndexChanged(object sender, EventArgs e) {
            if (this.LoadingProfileConfig) return;

            ActiveCalendarProfile.ColoursRestrictBy = (cbColourFilter.SelectedItem.ToString() == "Include") ?
                SettingsStore.Calendar.RestrictBy.Include : SettingsStore.Calendar.RestrictBy.Exclude;
            //Invert selection
            miColourSelectInvert_Click(null, null);
        }

        private void cbDeleteWhenColourExcl_CheckedChanged(object sender, EventArgs e) {
            ActiveCalendarProfile.DeleteWhenColourExcluded = cbDeleteWhenColourExcl.Checked;
        }

        private void clbColours_SelectedIndexChanged(object sender, EventArgs e) {
            if (this.LoadingProfileConfig) return;

            ActiveCalendarProfile.Colours.Clear();
            foreach (object item in clbColours.CheckedItems) {
                ActiveCalendarProfile.Colours.Add(item.ToString());
            }
        }

        private void refreshColours() {
            Ogcs.Google.Calendar.Instance.ColourPalette.Get();
            Ogcs.Google.Calendar.Instance.ColourPalette.BuildPicker(clbColours);
        }
        private void miColourRefresh_Click(object sender, EventArgs e) {
            refreshColours();
        }
        /// <summary>Shim function to work around x-thread call of BuildPicker()</summary>
        public void miColourBuildPicker_Click(object sender, EventArgs e) {
            CheckedListBox clb = GetControlThreadSafe(clbColours) as CheckedListBox;
            Ogcs.Google.Calendar.Instance.ColourPalette.BuildPicker(clb);
            SetControlPropertyThreadSafe(clbColours, "Items", clb.Items);
        }
        private void miColourSelectNone_Click(object sender, EventArgs e) {
            for (int i = 0; i < clbColours.Items.Count; i++) {
                clbColours.SetItemCheckState(i, CheckState.Unchecked);
            }
            clbColours_SelectedIndexChanged(null, null);
            this.clbColours.SelectedIndexChanged += new System.EventHandler(this.clbColours_SelectedIndexChanged);
        }
        private void miColourSelectAll_Click(object sender, EventArgs e) {
            for (int i = 0; i < clbColours.Items.Count; i++) {
                clbColours.SetItemCheckState(i, CheckState.Checked);
            }
            clbColours_SelectedIndexChanged(null, null);
        }
        private void miColourSelectInvert_Click(object sender, EventArgs e) {
            for (int i = 0; i < clbColours.Items.Count; i++) {
                clbColours.SetItemChecked(i, !clbColours.CheckedIndices.Contains(i));
            }
            clbColours_SelectedIndexChanged(null, null);
        }
        #endregion

        private void cbExcludeDeclinedInvites_CheckedChanged(object sender, EventArgs e) {
            ActiveCalendarProfile.ExcludeDeclinedInvites = cbExcludeDeclinedInvites.Checked;
        }
        private void cbExcludeGoals_CheckedChanged(object sender, EventArgs e) {
            ActiveCalendarProfile.ExcludeGoals = cbExcludeGoals.Checked;
        }
        private void cbGMeet_CheckedChanged(object sender, EventArgs e) {
            if (!this.LoadingProfileConfig && !cbAddDescription.Checked) {
                cbAddGMeet.Checked = false;
            }
            ActiveCalendarProfile.AddGMeet = cbAddGMeet.Checked;
        }

        #endregion

        #region Developer Options
        private void cbShowDeveloperOptions_CheckedChanged(object sender, EventArgs e) {
            //Toggle visibility
            pbExpandGoogleOauth.Visible =
            gbGoogle_GOAuth.Visible =
                cbShowDeveloperOptions.Checked;
        }

        private void llAPIConsole_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e) {
            Helper.OpenBrowser(llAPIConsole.Text);
        }

        private void tbClientID_TextChanged(object sender, EventArgs e) {
            Settings.Instance.PersonalClientIdentifier = tbClientID.Text;
        }
        private void tbClientSecret_TextChanged(object sender, EventArgs e) {
            Settings.Instance.PersonalClientSecret = tbClientSecret.Text;
            cbShowClientSecret.Enabled = (tbClientSecret.Text != "");
        }
        private void personalApiKey_Leave(object sender, EventArgs e) {
            setMaxSyncRange();
        }

        private void cbShowClientSecret_CheckedChanged(object sender, EventArgs e) {
            tbClientSecret.UseSystemPasswordChar = !cbShowClientSecret.Checked;
        }

        private void setMaxSyncRange() {
            if (Settings.Instance.UsingPersonalAPIkeys()) {
                tbDaysInTheFuture.Maximum = Int32.MaxValue;
                tbDaysInThePast.Maximum = Int32.MaxValue;
            } else {
                tbDaysInTheFuture.Maximum = 365;
                tbDaysInThePast.Maximum = 365;
            }
        }
        #endregion
        #endregion
        #region Sync options
        private void pbExpandHow_Click(object sender, EventArgs e) {
            groupboxSizing(gbSyncOptions_How, pbExpandHow);
        }
        private void pbExpandWhen_Click(object sender, EventArgs e) {
            groupboxSizing(gbSyncOptions_When, pbExpandWhen);
        }
        private void pbExpandWhat_Click(object sender, EventArgs e) {
            groupboxSizing(gbSyncOptions_What, pbExpandWhat);
        }

        #region How
        private void syncDirection_SelectedIndexChanged(object sender, EventArgs e) {
            ActiveCalendarProfile.SyncDirection = (Sync.Direction)syncDirection.SelectedItem;
            if (ActiveCalendarProfile.SyncDirection.Id == Sync.Direction.Bidirectional.Id) {
                ActiveCalendarProfile.RegisterForPushSync();
                cbObfuscateDirection.Enabled = true;
                cbObfuscateDirection.SelectedIndex = Sync.Direction.OutlookToGoogle.Id - 1;

                tbCreatedItemsOnly.Enabled = true;

                if (tbTargetCalendar.Items.Contains("target calendar"))
                    tbTargetCalendar.Items.Remove("target calendar");
                tbTargetCalendar.SelectedIndex = 0;
                tbTargetCalendar.Enabled = true;
                cbOutlookPush.Enabled = true;
                cbReminderDND.Visible = true;
                dtDNDstart.Visible = true;
                dtDNDend.Visible = true;
                lDNDand.Visible = true;
                cbSingleCategoryOnly.Visible = true;
                lExcludeItems.Text = "Exclude items. Affects newly synced items:-";
                lWhatExcludeInfo.Left = 207;
                cbExcludeTentative.Visible = true;
            } else {
                cbObfuscateDirection.Enabled = false;
                cbObfuscateDirection.SelectedIndex = ActiveCalendarProfile.SyncDirection.Id - 1;

                tbCreatedItemsOnly.Enabled = false;
                tbCreatedItemsOnly.SelectedIndex = 0;

                if (!tbTargetCalendar.Items.Contains("target calendar"))
                    tbTargetCalendar.Items.Add("target calendar");
                if (tbTargetCalendar.SelectedIndex == 2) tbTargetCalendar_SelectedItemChanged(null, null);
                tbTargetCalendar.SelectedIndex = 2;
                tbTargetCalendar.Enabled = false;
                lExcludeItems.Text = "Exclude items. Affects those previously synced:-";
                lWhatExcludeInfo.Left = 228;
            }
            if (ActiveCalendarProfile.SyncDirection.Id == Sync.Direction.GoogleToOutlook.Id) {
                ActiveCalendarProfile.DeregisterForPushSync();
                cbOutlookPush.Checked = false;
                cbOutlookPush.Enabled = false;
                cbReminderDND.Visible = false;
                dtDNDstart.Visible = false;
                dtDNDend.Visible = false;
                lDNDand.Visible = false;
                ddGoogleColour.Visible = false;
                ddOutlookColour.Visible = true;
                cbSingleCategoryOnly.Visible = true;
                cbExcludeTentative.Visible = false;
            }
            if (ActiveCalendarProfile.SyncDirection.Id == Sync.Direction.OutlookToGoogle.Id) {
                ActiveCalendarProfile.RegisterForPushSync();
                cbOutlookPush.Enabled = true;
                cbReminderDND.Visible = true;
                dtDNDstart.Visible = true;
                dtDNDend.Visible = true;
                lDNDand.Visible = true;
                ddGoogleColour.Visible = true;
                ddOutlookColour.Visible = false;
                cbSingleCategoryOnly.Visible = false;
                cbExcludeTentative.Visible = true;
            }
            cbAddAttendees_CheckedChanged(null, null);
            cbAddReminders_CheckedChanged(null, null);
            cbGoogleCalendars_SelectedIndexChanged(null, null);
            buildAvailabilityDropdown();
            showWhatPostit("Description");
        }

        private void cbMergeItems_CheckedChanged(object sender, EventArgs e) {
            ActiveCalendarProfile.MergeItems = cbMergeItems.Checked;
        }

        private void cbIgnoreBusyEntries_CheckedChanged(object sender, EventArgs e)
        {
            ActiveCalendarProfile.IgnoreBusy = cbIgnoreBusyEntries.Checked;
        }

        private void cbConfirmOnDelete_CheckedChanged(object sender, System.EventArgs e) {
            ActiveCalendarProfile.ConfirmOnDelete = cbConfirmOnDelete.Checked;
        }

        private void cbDisableDeletion_CheckedChanged(object sender, System.EventArgs e) {
            ActiveCalendarProfile.DisableDelete = cbDisableDeletion.Checked;
            cbConfirmOnDelete.Enabled = !cbDisableDeletion.Checked;
        }

        private void cbOfuscate_CheckedChanged(object sender, EventArgs e) {
            ActiveCalendarProfile.Obfuscation.Enabled = cbOfuscate.Checked;
        }

        private void btObfuscateRules_Click(object sender, EventArgs e) {
            this.howObfuscatePanel.Visible = true;
            this.howMorePanel.Visible = false;
            this.btCloseRegexRules.Visible = true;
            groupboxSizing(gbSyncOptions_How, pbExpandHow, true);
        }
        private void btCloseRegexRules_Click(object sender, EventArgs e) {
            this.btCloseRegexRules.Visible = false;
            this.howMorePanel.Visible = true;
            this.howObfuscatePanel.Visible = false;
            groupboxSizing(gbSyncOptions_How, pbExpandHow, true);
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
            ActiveCalendarProfile.CreatedItemsOnly = tbCreatedItemsOnly.SelectedIndex == 1;
            if (tbCreatedItemsOnly.SelectedIndex == 0)
                lTargetSyncCondition.Text = "synced to";
            else
                lTargetSyncCondition.Text = "by sync in";
        }

        private void tbTargetCalendar_SelectedItemChanged(object sender, EventArgs e) {
            if (this.LoadingProfileConfig) return;

            switch (tbTargetCalendar.Text) {
                case "Google calendar": {
                        ActiveCalendarProfile.TargetCalendar = Sync.Direction.OutlookToGoogle;
                        this.ddGoogleColour.Visible = true;
                        this.ddOutlookColour.Visible = false;
                        break;
                    }
                case "Outlook calendar": {
                        ActiveCalendarProfile.TargetCalendar = Sync.Direction.GoogleToOutlook;
                        this.ddGoogleColour.Visible = false;
                        this.ddOutlookColour.Visible = true;
                        if (Outlook.Factory.OutlookVersionName == Outlook.Factory.OutlookVersionNames.Outlook2003)
                            this.cbColour.Checked = false;
                        break;
                    }
                case "target calendar": {
                        ActiveCalendarProfile.TargetCalendar = ActiveCalendarProfile.SyncDirection;
                        if (Outlook.Factory.OutlookVersionName == Outlook.Factory.OutlookVersionNames.Outlook2003
                            && ActiveCalendarProfile.SyncDirection.Id == Sync.Direction.GoogleToOutlook.Id)
                            this.cbColour.Checked = false;
                        break;
                    }
            }
            buildAvailabilityDropdown();
        }

        private void cbPrivate_CheckedChanged(object sender, EventArgs e) {
            ddPrivacy.Enabled = cbPrivate.Checked;
            if (this.LoadingProfileConfig) return;

            ActiveCalendarProfile.SetEntriesPrivate = cbPrivate.Checked;
        }
        private void ddPrivacy_SelectedIndexChanged(object sender, EventArgs e) {
            if (this.LoadingProfileConfig) return;

            ActiveCalendarProfile.PrivacyLevel = ddPrivacy.SelectedValue.ToString();
        }

        private void cbAvailable_CheckedChanged(object sender, EventArgs e) {
            ddAvailabilty.Enabled = cbAvailable.Checked;
            if (this.LoadingProfileConfig) return;

            ActiveCalendarProfile.SetEntriesAvailable = cbAvailable.Checked;
        }
        private void ddAvailabilty_SelectedIndexChanged(object sender, EventArgs e) {
            if (this.LoadingProfileConfig) return;

            ActiveCalendarProfile.AvailabilityStatus = ddAvailabilty.SelectedValue.ToString();
        }

        private void cbColour_CheckedChanged(object sender, EventArgs e) {
            ddOutlookColour.Enabled = cbColour.Checked;
            ddGoogleColour.Enabled = cbColour.Checked;
            if (this.LoadingProfileConfig) return;

            ActiveCalendarProfile.SetEntriesColour = cbColour.Checked;
        }

        private void ddOutlookColour_SelectedIndexChanged(object sender, EventArgs e) {
            if (this.LoadingProfileConfig) return;

            ActiveCalendarProfile.SetEntriesColourValue = ddOutlookColour.SelectedItem.OutlookCategory.ToString();
            ActiveCalendarProfile.SetEntriesColourName = ddOutlookColour.SelectedItem.Text;

            if (sender == null) return;
            try {
                ddGoogleColour.SelectedIndexChanged -= ddGoogleColour_SelectedIndexChanged;

                Ogcs.Google.EventColour.Palette palette = Ogcs.Google.EventColour.Palette.NullPalette;
                if (Ogcs.Google.Calendar.IsColourPaletteNull || !Ogcs.Google.Calendar.Instance.ColourPalette.IsCached())
                    offlineAddGoogleColour();
                else {
                    if (ddGoogleColour.Items.Count != Ogcs.Google.Calendar.Instance.ColourPalette.ActivePalette.Count)
                        ddGoogleColour.AddPaletteColours();
                    palette = Ogcs.Google.Calendar.Instance.GetColour(ddOutlookColour.SelectedItem.OutlookCategory);
                    ddGoogleColour.SelectedIndex = Convert.ToInt16(palette.Id);
                }

                if (ddGoogleColour.SelectedIndex == -1)
                    log.Warn("Could not find the Google colour for: " + palette.ToString());
                else
                    ddGoogleColour_SelectedIndexChanged(null, null);

            } catch (System.Exception ex) {
                ex.Analyse("ddOutlookColour_SelectedIndexChanged(): Could not update ddGoogleColour.");
            } finally {
                ddGoogleColour.SelectedIndexChanged += ddGoogleColour_SelectedIndexChanged;
            }
        }

        private void ddGoogleColour_SelectedIndexChanged(object sender, EventArgs e) {
            if (this.LoadingProfileConfig) return;

            ActiveCalendarProfile.SetEntriesColourGoogleId = ddGoogleColour.SelectedItem.Id;

            if (sender == null) return;
            try {
                ddOutlookColour.SelectedIndexChanged -= ddOutlookColour_SelectedIndexChanged;

                String oCatName = null;
                if (Ogcs.Google.Calendar.IsColourPaletteNull || !Ogcs.Google.Calendar.Instance.ColourPalette.IsCached())
                    oCatName = ActiveCalendarProfile.SetEntriesColourName;
                else
                    oCatName = Outlook.Calendar.Instance.GetCategoryColour(ddGoogleColour.SelectedItem.Id);

                foreach (Outlook.Categories.ColourInfo cInfo in ddOutlookColour.Items) {
                    if (cInfo.Text == oCatName) {
                        ddOutlookColour.SelectedItem = cInfo;
                        break;
                    }
                }

                if (ddOutlookColour.SelectedIndex == -1)
                    log.Warn("Could not find the Outlook category for '" + oCatName + "'");
                else
                    ddOutlookColour_SelectedIndexChanged(null, null);

            } catch (System.Exception ex) {
                ex.Analyse("ddGoogleColour_SelectedIndexChanged(): Could not update ddOutlookColour.");
            } finally {
                ddOutlookColour.SelectedIndexChanged += ddOutlookColour_SelectedIndexChanged;
            }
        }

        /// <summary>
        /// Avoid connecting to Google simply to add correct profile colour to dropdown
        /// </summary>
        private void offlineAddGoogleColour() {
            Ogcs.Google.EventColour.Palette localPalette = new Ogcs.Google.EventColour.Palette(
                    Ogcs.Google.EventColour.Palette.Type.Event, ActiveCalendarProfile.SetEntriesColourGoogleId, null, Color.Transparent);
            if (!ddGoogleColour.Items.Cast<Ogcs.Google.EventColour.Palette>().Any(cbi => cbi.Id == localPalette.Id)) {
                ddGoogleColour.Items.Add(localPalette);
                ddGoogleColour.SelectedItem = localPalette;
                return;
            }
            foreach (Ogcs.Google.EventColour.Palette item in ddGoogleColour.Items) {
                if (item.Id == localPalette.Id) {
                    ddGoogleColour.SelectedItem = item;
                    break;
                }
            }
        }
        #endregion

        #region Obfuscation Panel
        private void cbObfuscateDirection_SelectedIndexChanged(object sender, EventArgs e) {
            if (!this.LoadingProfileConfig)
                ActiveCalendarProfile.Obfuscation.Direction = (Sync.Direction)cbObfuscateDirection.SelectedItem;
        }

        private void dgObfuscateRegex_Leave(object sender, EventArgs e) {
            ActiveCalendarProfile.Obfuscation.SaveRegex(dgObfuscateRegex);
        }

        private void dgObfuscateRegex_CellValidating(object sender, DataGridViewCellValidatingEventArgs e) {
            if (dgObfuscateRegex.Rows[e.RowIndex].IsNewRow) return;
            if (e.ColumnIndex != ((int)Obfuscate.Columns.target)) return;

            String strVal = e.FormattedValue.ToString().ToUpper().Replace(" ", "");
            if (String.IsNullOrEmpty(strVal)) strVal = "S";

            System.Text.RegularExpressions.Regex rgx = new System.Text.RegularExpressions.Regex("[^SLD]", System.Text.RegularExpressions.RegexOptions.IgnoreCase);

            if (rgx.IsMatch(strVal)) {
                e.Cancel = true;
                Ogcs.Extensions.MessageBox.Show("Cell must only include the characters:-\r   S = Subject\r   L = Location\r   D = Description",
                    "Invalid target values", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }
            dgObfuscateRegex.Rows[e.RowIndex].Cells[((int)Obfuscate.Columns.target)].Value = strVal;
            dgObfuscateRegex.RefreshEdit();
        }
        #endregion
        #endregion

        #region When
        public int MinSyncMinutes {
            get {
                if (Program.InDeveloperMode) return 1;
                else {
                    if (ActiveCalendarProfile.OutlookPush && ActiveCalendarProfile.SyncDirection.Id != Sync.Direction.GoogleToOutlook.Id)
                        return 120;
                    else
                        return 15;
                }
            }
        }

        private void tbDaysInThePast_ValueChanged(object sender, EventArgs e) {
            ActiveCalendarProfile.DaysInThePast = (int)tbDaysInThePast.Value;
        }

        private void tbDaysInTheFuture_ValueChanged(object sender, EventArgs e) {
            ActiveCalendarProfile.DaysInTheFuture = (int)tbDaysInTheFuture.Value;
        }

        private void tbMinuteOffsets_ValueChanged(object sender, EventArgs e) {
            String tooltip = "Set to zero to disable automated syncs";
            if (!Settings.Instance.UsingPersonalAPIkeys()) {
                String fup = "Fair usage policy: Minimum sync interval of " + MinSyncMinutes + "mins" + (ActiveCalendarProfile.OutlookPush ? " with Push Sync enabled" : "") + ".";

                tbInterval.ValueChanged -= new System.EventHandler(this.tbMinuteOffsets_ValueChanged);
                cbIntervalUnit.SelectedIndexChanged -= new System.EventHandler(this.cbIntervalUnit_SelectedIndexChanged);
                try {
                    if (cbIntervalUnit.SelectedItem.ToString() == "Minutes") {
                        if ((int)tbInterval.Value <= MinSyncMinutes)
                            tooltip = fup;
                        if ((int)tbInterval.Value < MinSyncMinutes) {
                            tbInterval.Value = (tbInterval.Value < Convert.ToInt16(tbInterval.Text)) ? 0 : MinSyncMinutes;
                        } else if ((int)tbInterval.Value > (tbInterval.Maximum - 1)) {
                            tbInterval.Value = ((int)tbInterval.Value / 60) + 1;
                            cbIntervalUnit.Text = "Hours";
                        }

                    } else if (cbIntervalUnit.SelectedItem.ToString() == "Hours") {
                        if (((int)tbInterval.Value * 60) <= MinSyncMinutes)
                            tooltip = fup;
                        if (((int)tbInterval.Value * 60) < MinSyncMinutes)
                            tbInterval.Value = (tbInterval.Value < Convert.ToInt16(tbInterval.Text)) ? 0 : (MinSyncMinutes / 60);
                    }
                } finally {
                    tbInterval.ValueChanged += new System.EventHandler(this.tbMinuteOffsets_ValueChanged);
                    cbIntervalUnit.SelectedIndexChanged += new System.EventHandler(this.cbIntervalUnit_SelectedIndexChanged);
                }
            }
            if (tbInterval.Value == 0)
                tooltip = "Set to non-zero to enable automated syncs";
            ToolTips.SetToolTip(tbInterval, tooltip);

            ActiveCalendarProfile.SyncInterval = (int)tbInterval.Value;
            ActiveCalendarProfile.OgcsTimer.SetNextSync();
            NotificationTray.UpdateAutoSyncItems();
        }

        private void cbIntervalUnit_SelectedIndexChanged(object sender, EventArgs e) {
            if (cbIntervalUnit.Text == "Minutes" && (int)tbInterval.Value > 0 && (int)tbInterval.Value < MinSyncMinutes) {
                tbInterval.Value = Math.Min((int)tbInterval.Value * 60, tbInterval.Maximum - 1);
            } else if (cbIntervalUnit.Text == "Hours") {
                tbInterval.Value = Math.Ceiling((decimal)(int)tbInterval.Value / 60);
            }
            ActiveCalendarProfile.SyncIntervalUnit = cbIntervalUnit.Text;
            ActiveCalendarProfile.OgcsTimer.SetNextSync();
        }

        private void cbOutlookPush_CheckedChanged(object sender, EventArgs e) {
            ActiveCalendarProfile.OutlookPush = cbOutlookPush.Checked;
            if (!this.LoadingProfileConfig) {
                if (tbInterval.Value != 0) tbMinuteOffsets_ValueChanged(null, null);
                if (cbOutlookPush.Checked) ActiveCalendarProfile.RegisterForPushSync();
                else ActiveCalendarProfile.DeregisterForPushSync();
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
        private void lWhatExcludeInfo_MouseHover(object sender, EventArgs e) {
            showWhatPostit("AffectedExcludeItems");
        }
        private void lWhatExcludeInfo_MouseLeave(object sender, EventArgs e) {
            showWhatPostit("Description");
        }
        private void showWhatPostit(String info) {
            switch (info) {
                case "Description": {
                        tbWhatHelp.Text = "Google event descriptions don't support rich text (RTF) and truncate at 8Kb. So make sure you REALLY want to 2-way sync descriptions!";
                        Boolean visible = (ActiveCalendarProfile.AddDescription && ActiveCalendarProfile.SyncDirection.Id == Sync.Direction.Bidirectional.Id);
                        WhatPostit.Visible = visible && !ActiveCalendarProfile.AddDescription_OnlyToGoogle;
                        cbAddDescription_OnlyToGoogle.Visible = visible;
                        break;
                    }
                case "AffectedItems": {
                        tbWhatHelp.Text = "Changes will only affect items synced hereon in.\r" +
                            "To update ALL items, click the Sync button whilst pressing the shift key.";
                        WhatPostit.Visible = true;
                        break;
                    }
                case "AffectedExcludeItems": {
                        if (ActiveCalendarProfile.SyncDirection.Id == Sync.Direction.Bidirectional.Id) {
                            tbWhatHelp.Text = "Excluding items will only affect items synced hereon in.";
                        } else {
                            tbWhatHelp.Text = "Excluding items will delete those previously synced.";
                        }
                        tbWhatHelp.Text += "\rFor more fine-grained control, consider filtering on categories.";
                        WhatPostit.Visible = true;
                        break;
                    }
            }
            tbWhatHelp.SelectAll();
            tbWhatHelp.SelectionAlignment = HorizontalAlignment.Center;
            tbWhatHelp.DeselectAll();
        }

        private void cbLocation_CheckedChanged(object sender, EventArgs e) {
            ActiveCalendarProfile.AddLocation = cbLocation.Checked;
        }

        private void cbAddDescription_CheckedChanged(object sender, EventArgs e) {
            if (cbAddDescription.Checked && ActiveCalendarProfile.OutlookGalBlocked) {
                cbAddDescription.Checked = false;
                return;
            }
            ActiveCalendarProfile.AddDescription = cbAddDescription.Checked;
            cbAddDescription_OnlyToGoogle.Enabled = cbAddDescription.Checked;
            checkboxSoftRestrict(cbAddGMeet, !cbAddDescription.Checked);
            ToolTips.SetToolTip(cbAddGMeet, cbAddDescription.Checked ? "Sync conference details embedded in Outlook appointment body." : "Requires sync of Description (under Options > What)");
            showWhatPostit("Description");
        }
        private void cbAddDescription_OnlyToGoogle_CheckedChanged(object sender, EventArgs e) {
            ActiveCalendarProfile.AddDescription_OnlyToGoogle = cbAddDescription_OnlyToGoogle.Checked;
            showWhatPostit("Description");
        }

        private void cbAddReminders_CheckedChanged(object sender, EventArgs e) {
            if (!this.LoadingProfileConfig && sender != null) ActiveCalendarProfile.AddReminders = cbAddReminders.Checked;
            cbUseGoogleDefaultReminder.Enabled = ActiveCalendarProfile.SyncDirection.Id != Sync.Direction.GoogleToOutlook.Id;
            cbUseOutlookDefaultReminder.Enabled = ActiveCalendarProfile.SyncDirection.Id != Sync.Direction.OutlookToGoogle.Id;
            cbReminderDND.Enabled = cbAddReminders.Checked;
            dtDNDstart.Enabled = cbAddReminders.Checked;
            dtDNDend.Enabled = cbAddReminders.Checked;
            lDNDand.Enabled = cbAddReminders.Checked;
        }
        private void cbUseGoogleDefaultReminder_CheckedChanged(object sender, EventArgs e) {
            ActiveCalendarProfile.UseGoogleDefaultReminder = cbUseGoogleDefaultReminder.Checked;
        }
        private void cbUseOutlookDefaultReminder_CheckedChanged(object sender, EventArgs e) {
            ActiveCalendarProfile.UseOutlookDefaultReminder = cbUseOutlookDefaultReminder.Checked;
        }
        private void cbReminderDND_CheckedChanged(object sender, EventArgs e) {
            ActiveCalendarProfile.ReminderDND = cbReminderDND.Checked;
        }
        private void dtDNDstart_ValueChanged(object sender, EventArgs e) {
            ActiveCalendarProfile.ReminderDNDstart = dtDNDstart.Value;
        }
        private void dtDNDend_ValueChanged(object sender, EventArgs e) {
            ActiveCalendarProfile.ReminderDNDend = dtDNDend.Value;
        }

        private void cbAddAttendees_CheckedChanged(object sender, EventArgs e) {
            if (cbAddAttendees.Checked && ActiveCalendarProfile.OutlookGalBlocked) {
                cbAddAttendees.Checked = false;
                cbCloakEmail.Enabled = false;
                tbMaxAttendees.Enabled = false;
                return;
            }
            if (!this.LoadingProfileConfig && sender != null) ActiveCalendarProfile.AddAttendees = cbAddAttendees.Checked;
            tbMaxAttendees.Enabled = cbAddAttendees.Checked;
            cbCloakEmail.Visible = ActiveCalendarProfile.SyncDirection.Id != Sync.Direction.GoogleToOutlook.Id;
            cbCloakEmail.Enabled = cbAddAttendees.Checked;
            if (cbAddAttendees.Checked && string.IsNullOrEmpty(Outlook.Calendar.Instance.IOutlook.CurrentUserSMTP())) {
                Outlook.Calendar.Instance.IOutlook.GetCurrentUser(null);
            }
        }
        private void tbMaxAttendees_ValueChanged(object sender, EventArgs e) {
            ActiveCalendarProfile.MaxAttendees = (int)tbMaxAttendees.Value;
        }
        private void cbCloakEmail_CheckedChanged(object sender, EventArgs e) {
            ActiveCalendarProfile.CloakEmail = cbCloakEmail.Checked;
        }
        private void cbAddColours_CheckedChanged(object sender, EventArgs e) {
            ActiveCalendarProfile.AddColours = cbAddColours.Checked;
            btColourMap.Enabled = ActiveCalendarProfile.AddColours;
            cbSingleCategoryOnly.Enabled = ActiveCalendarProfile.AddColours;
        }
        private void btColourMap_Click(object sender, EventArgs e) {
            if (ActiveCalendarProfile.UseGoogleCalendar == null || string.IsNullOrEmpty(ActiveCalendarProfile.UseGoogleCalendar.Id)) {
                Ogcs.Extensions.MessageBox.Show("You need to select a Google Calendar first on the 'Settings' tab.", "Configuration Required", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            try {
                this.btColourMap.Enabled = false;
                using (Forms.ColourMap colourForm = new ColourMap()) {
                    colourForm.ShowDialog();
                }
            } catch (System.Exception ex) {
                Ogcs.Exception.Analyse(ex);
            }
        }
        private void cbSingleCategoryOnly_CheckedChanged(object sender, EventArgs e) {
            ActiveCalendarProfile.SingleCategoryOnly = cbSingleCategoryOnly.Checked;
        }

        private void cbExcludeFree_CheckedChanged(object sender, EventArgs e) {
            ActiveCalendarProfile.ExcludeFree = cbExcludeFree.Checked;
        }
        private void cbExcludeTentative_CheckedChanged(object sender, EventArgs e) {
            ActiveCalendarProfile.ExcludeTentative = cbExcludeTentative.Checked;
        }
        private void cbExcludePrivate_CheckedChanged(object sender, EventArgs e) {
            ActiveCalendarProfile.ExcludePrivate = cbExcludePrivate.Checked;
        }
        private void cbExcludeAllDays_CheckedChanged(object sender, EventArgs e) {
            ActiveCalendarProfile.ExcludeAllDays = cbExcludeAllDays.Checked;
            cbExcludeFreeAllDays.Enabled = cbExcludeAllDays.Checked;
            cbExcludeFreeAllDays.Checked = false;
        }
        private void cbExcludeFreeAllDays_CheckedChanged(object sender, EventArgs e) {
            ActiveCalendarProfile.ExcludeFreeAllDays = cbExcludeFreeAllDays.Checked;
        }
        private void cbExcludeSubject_CheckedChanged(object sender, EventArgs e) {
            ActiveCalendarProfile.ExcludeSubject =
            tbExcludeSubjectText.Enabled =
                cbExcludeSubject.Checked;
        }
        private void tbExcludeSubjectText_Leave(object sender, EventArgs e) {
            ActiveCalendarProfile.ExcludeSubjectText = tbExcludeSubjectText.Text;
        }
        #endregion
        #endregion
        #region Application settings
        private void cbStartOnStartup_CheckedChanged(object sender, EventArgs e) {
            Settings.Instance.StartOnStartup = cbStartOnStartup.Checked;
            tbStartupDelay.Enabled = cbStartOnStartup.Checked;
            cbStartOnStartupAllUsers.Enabled = cbStartOnStartup.Checked;
            try {
                Program.ManageStartupRegKey();
            } catch (System.Exception ex) {
                if (ex is System.Security.SecurityException) Ogcs.Exception.LogAsFail(ref ex); //User doesn't have rights to access registry
                ex.Analyse("Failed accessing registry for startup key(s).");
                if (this.Visible) {
                    Ogcs.Extensions.MessageBox.Show("You do not have permissions to access the system registry.\nThis setting cannot be used.",
                        "Registry access denied", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                }
                cbStartOnStartup.CheckedChanged -= cbStartOnStartup_CheckedChanged;
                cbStartOnStartup.Checked = false;
                Settings.Instance.StartOnStartup = false;
                tbStartupDelay.Enabled = false;
                cbStartOnStartup.CheckedChanged += cbStartOnStartup_CheckedChanged;
            }
        }
        private void cbStartOnStartupAllUsers_CheckedChanged(object sender, EventArgs e) {
            Settings.Instance.StartOnStartupAllUsers = cbStartOnStartupAllUsers.Checked;
            try {
                Program.ManageStartupRegKey();
            } catch (System.Exception ex) {
                if (ex is System.Security.SecurityException) Ogcs.Exception.LogAsFail(ref ex); //User doesn't have rights to access registry
                ex.Analyse("Failed accessing registry for HKLM startup key.");
                if (this.Visible) {
                    Ogcs.Extensions.MessageBox.Show("You do not have permissions to access the system registry.\nThis setting cannot be used.",
                        "Registry access denied", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                }
                cbStartOnStartupAllUsers.CheckedChanged -= cbStartOnStartupAllUsers_CheckedChanged;
                cbStartOnStartupAllUsers.Checked = false;
                Settings.Instance.StartOnStartupAllUsers = false;
                cbStartOnStartupAllUsers.CheckedChanged += cbStartOnStartupAllUsers_CheckedChanged;
            }
        }


        private void cbHideSplash_CheckedChanged(object sender, EventArgs e) {
            if (Settings.Instance.UserIsBenefactor()) {
                Settings.Instance.HideSplashScreen = cbHideSplash.Checked;
            } else {
                cbHideSplash.CheckedChanged -= cbHideSplash_CheckedChanged;
                cbHideSplash.Checked = false;
                cbHideSplash.CheckedChanged += cbHideSplash_CheckedChanged;
                ToolTips.SetToolTip(cbHideSplash, "Donate £10 or more to enable this feature.");
                ToolTips.Show(ToolTips.GetToolTip(cbHideSplash), cbHideSplash, 5000);
                Settings.Instance.HideSplashScreen = null;
            }
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

        private void cbShowSystemNotifications_CheckedChanged(object sender, EventArgs e) {
            Settings.Instance.ShowSystemNotifications = cbShowSystemNotifications.Checked;
            if (!cbShowSystemNotifications.Checked) cbShowSystemNotificationsIfChange.Checked = false;
            cbShowSystemNotificationsIfChange.Enabled = cbShowSystemNotifications.Checked;
        }
        private void cbShowSystemNotificationsIfChange_CheckedChanged(object sender, EventArgs e) {
            Settings.Instance.ShowSystemNotificationsIfChange = cbShowSystemNotificationsIfChange.Checked;
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
            if (Settings.AreApplied) {
                if (Program.StartedWithFileArgs)
                    Ogcs.Extensions.MessageBox.Show("It is not possible to change portability of OGCS when it is started with command line parameters.",
                        "Cannot change portability", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                else {
                    Settings.Instance.Portable = cbPortable.Checked;
                    Program.MakePortable(cbPortable.Checked);
                }
            }
        }

        private void pbExpandLogging_Click(object sender, EventArgs e) {
            groupboxSizing(gbAppBehaviour_Logging, pbExpandLogging);
        }

        private void pbExpandProxy_Click(object sender, EventArgs e) {
            groupboxSizing(gbAppBehaviour_Proxy, pbExpandProxy);
        }

        private void cbCreateFiles_CheckedChanged(object sender, EventArgs e) {
            Settings.Instance.CreateCSVFiles = cbCreateFiles.Checked;
        }

        private void cbLoggingLevel_SelectedIndexChanged(object sender, EventArgs e) {
            Settings.configureLoggingLevel(this.cbLoggingLevel.Text);
            if (!this.LoadingProfileConfig) Settings.Instance.LoggingLevel = this.cbLoggingLevel.Text.ToUpper();
        }

        private void btLogLocation_Click(object sender, EventArgs e) {
            try {
                log4net.Appender.IAppender[] appenders = log.Logger.Repository.GetAppenders();
                String logFileLocation = (((log4net.Appender.FileAppender)appenders[0]).File);
                logFileLocation = logFileLocation.Substring(0, logFileLocation.LastIndexOf("\\"));
                System.Diagnostics.Process.Start("explorer.exe", @logFileLocation);
            } catch {
                System.Diagnostics.Process.Start("explorer.exe", @Program.UserFilePath);
            }
        }

        private void cbCloudLogging_CheckStateChanged(object sender, EventArgs e) {
            if (!Settings.AreApplied) return;

            if (cbCloudLogging.CheckState == CheckState.Indeterminate)
                Settings.Instance.CloudLogging = null;
            else
                Settings.Instance.CloudLogging = cbCloudLogging.Checked;
        }

        private void cbAnonymiseLogs_CheckedChanged(object sender, EventArgs e) {
            if (!Settings.AreApplied) return;

            Settings.Instance.AnonymiseLogs = cbAnonymiseLogs.Checked;
        }

        private void cbTelemetryDisabled_CheckedChanged(object sender, EventArgs e) {
            if (!Settings.AreApplied) return;

            if (!cbTelemetryDisabled.Checked) {
                Settings.Instance.TelemetryDisabled = cbTelemetryDisabled.Checked;
                return;
            }
            DialogResult dr = MessageBox.Show("The telemetry only captures anonymised usage statistics, such as your version of OGCS and Outlook. " +
                "This helps focus ongoing improvements. Are you sure you wish to disable telemetry?", "OGCS Usage Statistics", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
            if (dr == DialogResult.No) {
                cbTelemetryDisabled.CheckedChanged -= cbTelemetryDisabled_CheckedChanged;
                cbTelemetryDisabled.Checked = false;
                cbTelemetryDisabled.CheckedChanged += cbTelemetryDisabled_CheckedChanged;
            }
            log.Info("Telemetry has been " + (cbTelemetryDisabled.Checked ? "dis" : "en") + "abled.");
            Settings.Instance.TelemetryDisabled = cbTelemetryDisabled.Checked;
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
                Helper.OpenBrowser(Program.OgcsWebsite + "/browseruseragent");
            } catch (System.Exception ex) {
                ex.Analyse("Failed to check browser's user agent.");
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
            Helper.OpenBrowser("https://github.com/phw198/OutlookGoogleCalendarSync/issues");
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
                    System.Diagnostics.Process.Start("explorer.exe", path);
                }
            } catch { }
        }

        private void lAboutURL_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e) {
            Helper.OpenBrowser(lAboutURL.Text);
        }

        private void pbDonate_Click(object sender, EventArgs e) {
            Program.Donate("About");
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
        private delegate Control getControlThreadSafeDelegate(Control control);

        private delegate object getControlPropertyThreadSafeDelegate(Control control, string propertyName);
        private delegate void setControlPropertyThreadSafeDelegate(Control control, string propertyName, object propertyValue);

        private delegate void callControlMethodThreadSafeDelegate(Control control, string methodName, object methodArgValue);

        public static Control GetControlThreadSafe(Control control) {
            if (control.InvokeRequired) {
                return (Control)control.Invoke(new getControlThreadSafeDelegate(GetControlThreadSafe), new object[] { control });
            } else {
                return control.GetType().InvokeMember("discarded", System.Reflection.BindingFlags.CreateInstance, null, control, null) as Control;
            }
        }

        public object GetControlPropertyThreadSafe(Control control, string propertyName) {
            if (control.InvokeRequired) {
                return control.Invoke(new getControlPropertyThreadSafeDelegate(GetControlPropertyThreadSafe), new object[] { control, propertyName });
            } else {
                return control.GetType().InvokeMember(propertyName, System.Reflection.BindingFlags.GetProperty, null, control, null);
            }
        }
        public void SetControlPropertyThreadSafe(Control control, string propertyName, object propertyValue) {
            if (control.InvokeRequired) {
                control.Invoke(new setControlPropertyThreadSafeDelegate(SetControlPropertyThreadSafe), new object[] { control, propertyName, propertyValue });
            } else {
                if (control is CheckedListBox && propertyValue is CheckedListBox.ObjectCollection) {
                    (control as CheckedListBox).Items.AddRange(propertyValue as CheckedListBox.ObjectCollection);
                    return;
                }
                var theObject = control.GetType().InvokeMember(propertyName, System.Reflection.BindingFlags.SetProperty, null, control, new object[] { propertyValue });
                if (control is TextBox) {
                    TextBox tb = control as TextBox;
                    tb.SelectionStart = tb.Text.Length;
                    tb.ScrollToCaret();
                }
            }
        }

        public void CallControlMethodThreadSafe(Control control, string methodName, object methodArgValue) {
            if (control.InvokeRequired) {
                control.Invoke(new callControlMethodThreadSafeDelegate(CallControlMethodThreadSafe), new object[] { control, methodName, methodArgValue });
            } else {
                var theObject = control.GetType().InvokeMember(methodName, System.Reflection.BindingFlags.InvokeMethod, null, control, new object[] { methodArgValue });
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
                    new Forms.Social().ShowDialog();
                }
            } catch (System.Exception ex) {
                log.Warn("Failed checking sync milestone.");
                Ogcs.Exception.Analyse(ex);
            }
        }

        private void btSocialTweet_Click(object sender, EventArgs e) {
            Social.Twitter_tweet();
        }
        private void pbSocialTwitterFollow_Click(object sender, EventArgs e) {
            Social.Twitter_follow();
        }

        private void btSocialFB_Click(object sender, EventArgs e) {
            Social.Facebook_share();
        }
        private void btFbLike_Click(object sender, EventArgs e) {
            Social.Facebook_like();
        }

        private void btSocialLinkedin_Click(object sender, EventArgs e) {
            Social.Linkedin_share();
        }

        private void btSocialRSSfeed_Click(object sender, EventArgs e) {
            Social.RSS_follow();
        }

        private void btSocialGitHub_Click(object sender, EventArgs e) {
            Social.GitHub();
        }
        #endregion
    }
}