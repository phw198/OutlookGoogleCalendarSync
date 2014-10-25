using System;
using System.IO;
using System.Collections.Generic;
using System.Drawing;
using System.Net;
using System.Windows.Forms;
using Microsoft.Office.Interop.Outlook;
using Google.Apis.Calendar.v3;
using Google.Apis.Calendar.v3.Data;

namespace OutlookGoogleSync {
    /// <summary>
    /// Description of MainForm.
    /// </summary>
    public partial class MainForm : Form {
        public static MainForm Instance;

        public const string FILENAME = "settings.xml";
        private string VERSION = "1.1.2";

        private Timer ogstimer;
        private List<int> MinuteOffsets = new List<int>();
        private DateTime lastSyncDate;
        private int currentTimerInterval = 0;

        public MainForm() {
            InitializeComponent();
            lAboutMain.Text = lAboutMain.Text.Replace("{version}", VERSION);

            Instance = this;

            //set system proxy
            WebProxy wp = (WebProxy)System.Net.GlobalProxySelection.Select;
            //http://www.dreamincode.net/forums/topic/160555-working-with-proxy-servers/
            //WebProxy wp = (WebProxy)WebRequest.DefaultWebProxy;
            wp.UseDefaultCredentials = true;
            System.Net.WebRequest.DefaultWebProxy = wp;

            //load settings/create settings file
            if (File.Exists(FILENAME)) {
                Settings.Instance = XMLManager.import<Settings>(FILENAME);
            } else {
                XMLManager.export(Settings.Instance, FILENAME);
            }

            //create the timer for the autosynchro 
            ogstimer = new Timer();
            ogstimer.Tick += new EventHandler(ogstimer_Tick);

            #region Update GUI from Settings
            this.SuspendLayout();
            #region Outlook box
            this.gbOutlook.SuspendLayout();
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
            if (OutlookCalendar.Instance.Accounts.Count == 1) {
                rbOutlookAltMB.Enabled = false;
                rbOutlookAltMB.Checked = false;
                ddMailboxName.Enabled = false;
            }
            for (int acc = 2; acc <= OutlookCalendar.Instance.Accounts.Count; acc++) {
                String mailbox = OutlookCalendar.Instance.Accounts[acc].SmtpAddress.ToLower();
                ddMailboxName.Items.Add(mailbox);
                if (Settings.Instance.MailboxName == mailbox) { ddMailboxName.SelectedIndex = acc-2; }
            }
            if (ddMailboxName.SelectedIndex==-1 && ddMailboxName.Items.Count>0) { ddMailboxName.SelectedIndex = 0; }

            cbOutlookCalendars.DataSource = new BindingSource(OutlookCalendar.Instance.CalendarFolders, null);
            cbOutlookCalendars.DisplayMember = "Key";
            cbOutlookCalendars.ValueMember = "Value";

            this.gbOutlook.ResumeLayout();
            #endregion
            #region Google box
            this.gbGoogle.SuspendLayout();
            cbGoogleCalendars.Items.Add(Settings.Instance.UseGoogleCalendar);
            cbGoogleCalendars.SelectedIndex = 0;
            this.gbGoogle.ResumeLayout();
            #endregion
            #region Sync Options box
            this.gbSyncOptions.SuspendLayout();
            tbDaysInThePast.Text = Settings.Instance.DaysInThePast.ToString();
            tbDaysInTheFuture.Text = Settings.Instance.DaysInTheFuture.ToString();
            tbInterval.Value = Settings.Instance.SyncInterval;
            cbIntervalUnit.Text = Settings.Instance.SyncIntervalUnit;
            cbAddDescription.Checked = Settings.Instance.AddDescription;
            cbAddAttendees.Checked = Settings.Instance.AddAttendees;
            cbAddReminders.Checked = Settings.Instance.AddReminders;
            cbDisableDeletion.Checked = Settings.Instance.DisableDelete;
            cbConfirmOnDelete.Enabled = !Settings.Instance.DisableDelete;
            cbConfirmOnDelete.Checked = Settings.Instance.ConfirmOnDelete;
            this.gbSyncOptions.ResumeLayout();
            #endregion
            #region Application behaviour
            this.gbAppBehaviour.SuspendLayout();
            cbShowBubbleTooltips.Checked = Settings.Instance.ShowBubbleTooltipWhenSyncing;
            cbStartInTray.Checked = Settings.Instance.StartInTray;
            cbMinimizeToTray.Checked = Settings.Instance.MinimizeToTray;
            cbCreateFiles.Checked = Settings.Instance.CreateTextFiles;
            this.gbAppBehaviour.ResumeLayout();
            #endregion
            lastSyncDate = Settings.Instance.LastSyncDate;
            cbVerboseOutput.Checked = Settings.Instance.VerboseOutput;
            this.ResumeLayout();
            #endregion

            //set up tooltips for some controls
            ToolTip toolTip1 = new ToolTip();
            toolTip1.AutoPopDelay = 10000;
            toolTip1.InitialDelay = 500;
            toolTip1.ReshowDelay = 200;
            toolTip1.ShowAlways = true;
            toolTip1.SetToolTip(cbOutlookCalendars,
                "The Outlook calendar to synchonize with. List includes subfolders of default calendar.");
            toolTip1.SetToolTip(cbGoogleCalendars,
                "The Google calendar to synchonize with.");
            toolTip1.SetToolTip(tbInterval,
                "Set to zero to disable");
            toolTip1.SetToolTip(cbAddAttendees,
                "While Outlook has fields for Organizer, RequiredAttendees and OptionalAttendees, Google has not.\n" +
                "If checked, this data is added at the end of the description as text.");
            toolTip1.SetToolTip(cbAddReminders,
                "If checked, the reminder set in outlook will be carried over to the Google Calendar entry (as a popup reminder).");
            toolTip1.SetToolTip(cbCreateFiles,
                "If checked, all entries found in Outlook/Google and identified for creation/deletion will be exported \n" +
                "to 4 separate text files in the application's directory (named \"export_*.txt\"). \n" +
                "Only for debug/diagnostic purposes.");
            toolTip1.SetToolTip(cbAddDescription,
                "The description may contain email addresses, which Outlook may complain about (PopUp-Message: \"Allow Access?\" etc.). \n" +
                "Turning this off allows OutlookGoogleSync to run without intervention in this case.");
            toolTip1.SetToolTip(rbOutlookAltMB,
                "Only choose this if you need to use an Outlook Calendar that is not in the default mailbox");

            //Refresh synchronizations (last and next)
            lLastSyncVal.Text = lastSyncDate.ToLongDateString() + " - " + lastSyncDate.ToLongTimeString();
            setNextSync(getResyncInterval());

            //Start in tray?
            if (cbStartInTray.Checked) {
                this.WindowState = FormWindowState.Minimized;
                notifyIcon1.Visible = true;
                this.Hide();
                this.ShowInTaskbar = false;
            }
        }

        #region Autosync functions
        int getResyncInterval() {
            int min = (int)tbInterval.Value;
            if (cbIntervalUnit.Text == "Hours") {
                min *= 60;
            }
            return min;
        }

        void ogstimer_Tick(object sender, EventArgs e) {
            if (cbShowBubbleTooltips.Checked) {
                notifyIcon1.ShowBalloonTip(
                    500,
                    "OutlookGoogleSync",
                    "Autosyncing calendar...",
                    ToolTipIcon.Info
                );
            }
            SyncNow_Click(null, null);
        }

        void setNextSync(int delay) {
            if (tbInterval.Value != 0) {
                DateTime nextSyncDate = lastSyncDate.AddMinutes(delay);
                if (currentTimerInterval != delay) {
                    ogstimer.Stop();
                    DateTime now = DateTime.Now;
                    TimeSpan diff = nextSyncDate - now;
                    currentTimerInterval = diff.Minutes;
                    if (currentTimerInterval < 1) { currentTimerInterval = 1; nextSyncDate = now.AddMinutes(currentTimerInterval); }
                    ogstimer.Interval = currentTimerInterval * 60000;
                    ogstimer.Start();
                }
                lNextSyncVal.Text = nextSyncDate.ToLongDateString() + " - " + nextSyncDate.ToLongTimeString();
            } else {
                lNextSyncVal.Text = "Inactive";
                ogstimer.Stop();
            }
        }
        #endregion

        private void SyncNow_Click(object sender, EventArgs e) {
            LogBox.Clear();
            
            if (Settings.Instance.UseGoogleCalendar.Id == "") {
                MessageBox.Show("You need to select a Google Calendar first on the 'Settings' tab.");
                return;
            }
            //Check network availability
            if (!System.Net.NetworkInformation.NetworkInterface.GetIsNetworkAvailable()) {
                Logboxout("There does not appear to be any network available! Sync aborted.");
                return;
            }
            bSyncNow.Enabled = false;
            lNextSyncVal.Text = "In progress...";

            DateTime SyncStarted = DateTime.Now;
            Logboxout("Sync started at " + SyncStarted.ToString());
            Logboxout("--------------------------------------------------");

            Boolean syncOk = false;
            int failedAttempts = 0;
            while (!syncOk) {
                if (failedAttempts > 0 &&
                    MessageBox.Show("The synchronisation failed. Do you want to try again?", "Sync Failed",
                    MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation) == System.Windows.Forms.DialogResult.No) { break; }
                syncOk = synchronize();
                failedAttempts += !syncOk ? 1 : 0;
            }

            Logboxout("--------------------------------------------------");
            Logboxout(syncOk ? "Sync finished with success!" : "Operation aborted after "+ failedAttempts +" failed attempts!");

            if (syncOk) {
                lastSyncDate = SyncStarted;
                Settings.Instance.LastSyncDate = lastSyncDate;
                lLastSyncVal.Text = SyncStarted.ToLongDateString() + " - " + SyncStarted.ToLongTimeString();
                setNextSync(getResyncInterval());
            } else {
                Logboxout("Another sync has been scheduled to automatically run in 5 minutes time.");
                setNextSync(5);
            }
            bSyncNow.Enabled = true;
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
                OutlookCalendar.Instance.Reset(); 
                return false;
            }
            Logboxout(OutlookEntries.Count + " Outlook calendar entries found.");
            Logboxout("--------------------------------------------------");
            #endregion

            #region Read Google items
            Logboxout("Reading Google Calendar Entries...");
            List<Event> googleEntries = null;
            try {
                googleEntries = GoogleCalendar.Instance.getCalendarEntriesInRange();
            } catch (System.Exception ex) {
                Logboxout("Unable to connect to the Google calendar. The following error occurred:");
                Logboxout(ex.Message + "\r\n => Check your network connection.");
                return false;
            }
            Logboxout(googleEntries.Count + " Google calendar entries found.");
            Logboxout("--------------------------------------------------");
            #endregion

            //  Make copies of each list of events (Not strictly needed)
            List<AppointmentItem> googleEntriesToBeCreated = new List<AppointmentItem>(OutlookEntries);
            List<Event> googleEntriesToBeDeleted = new List<Event>(googleEntries);
            Dictionary<AppointmentItem, Event> entriesToBeCompared = new Dictionary<AppointmentItem, Event>();

            //GoogleCalendar.Instance.ReclaimOrphanEntries(googleEntriesToBeCreated);

            GoogleCalendar.IdentifyEventDifferences(googleEntriesToBeCreated, googleEntriesToBeDeleted, entriesToBeCompared);

            if (Settings.Instance.DisableDelete) {
                googleEntriesToBeDeleted = new List<Event>();
            }
            Logboxout(googleEntriesToBeDeleted.Count + " Google calendar entries to be deleted.");
            Logboxout(googleEntriesToBeCreated.Count + " Google calendar entries to be created.");
            
            //Protect against very first syncs which may trample pre-existing non-Outlook events in Google
            if (!Settings.Instance.DisableDelete && !Settings.Instance.ConfirmOnDelete &&
                googleEntriesToBeDeleted.Count == googleEntries.Count) {
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
                    GoogleCalendar.Instance.deleteCalendarEntries(googleEntriesToBeDeleted);
                } catch (System.Exception ex) {
                    MainForm.Instance.Logboxout("Unable to delete obsolete entries out to the Google calendar. The following error occurred:");
                    MainForm.Instance.Logboxout(ex.Message + "\r\n => Check your network connection.");
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
                    GoogleCalendar.Instance.createCalendarEntries(googleEntriesToBeCreated);
                } catch (System.Exception ex) {
                    Logboxout("Unable to add new entries into the Google Calendar. The following error occurred:");
                    Logboxout(ex.Message + "\r\n => Check your network connection.");
                    return false;
                }
                Logboxout("Done.");
            }
            #endregion

            #region Update Google Entries
            if (entriesToBeCompared.Count > 0) {
                Logboxout("--------------------------------------------------");
                Logboxout("Comparing " + entriesToBeCompared.Count + " existing Google calendar entries...");
                int entriesUpdated = 0;
                try {
                    GoogleCalendar.Instance.updateCalendarEntries(entriesToBeCompared, ref entriesUpdated);
                } catch (System.Exception ex) {
                    Logboxout("Unable to update new entries into the Google calendar. The following error occurred:");
                    Logboxout(ex.Message + "\r\n => Check your network connection.");
                    return false;
                }
                Logboxout(entriesUpdated + " entries updated.");
            }
            #endregion
            
            return true;
        }
                
        #region Compare Event Attributes 
        public static Boolean CompareAttribute(String attrDesc, String googleAttr, String outlookAttr, System.Text.StringBuilder sb, ref int itemModified) {
            if (googleAttr == null) googleAttr = "";
            if (outlookAttr == null) outlookAttr = "";
            if (googleAttr != outlookAttr) {
                //Truncate long strings
                sb.AppendLine(attrDesc + ": " + 
                    ((googleAttr.Length>50)?googleAttr.Substring(0, 47)+"...":googleAttr) + " => " + 
                    ((outlookAttr.Length>50)?outlookAttr.Substring(0, 47)+"...":outlookAttr)
                    );
                itemModified++;
                return true;
            } else {
                return false;
            }
        }
        public static Boolean CompareAttribute(String attrDesc, Boolean googleAttr, Boolean outlookAttr, System.Text.StringBuilder sb, ref int itemModified) {
            if (googleAttr != outlookAttr) {
                sb.AppendLine(attrDesc + ": " + googleAttr + " => " + outlookAttr);
                itemModified++;
                return true;
            } else {
                return false;
            }
        }
        #endregion

        public void Logboxout(string s, bool newLine=true) {
            LogBox.Text += s + (newLine ? Environment.NewLine : "");
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
            XMLManager.export(Settings.Instance, FILENAME);
        }

        void NotifyIcon1Click(object sender, EventArgs e) {
            this.WindowState = FormWindowState.Normal;
            this.Show();
        }

        void MainFormResize(object sender, EventArgs e) {
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
        #region Outlook settings
        private void rbOutlookDefaultMB_CheckedChanged(object sender, EventArgs e) {
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
            Settings.Instance.MailboxName = ddMailboxName.Text;
            OutlookCalendar.Instance.Reset();
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
            KeyValuePair<String,MAPIFolder>calendar = (KeyValuePair<String,MAPIFolder>)cbOutlookCalendars.SelectedItem;
            OutlookCalendar.Instance.UseOutlookCalendar = calendar.Value;
        }
        #endregion
        #region Google settings
        void GetMyGoogleCalendars_Click(object sender, EventArgs e) {
            bGetGoogleCalendars.Enabled = false;
            cbGoogleCalendars.Enabled = false;
            List<MyCalendarListEntry> calendars = null;
            try {
                calendars = GoogleCalendar.Instance.getCalendars();
            } catch (System.Exception ex) {
                Logboxout("Unable to get the list of Google Calendars. The following error occurred:");
                Logboxout(ex.Message + "\r\n => Check your network connection.");
            }
            if (calendars != null) {
                cbGoogleCalendars.Items.Clear();
                foreach (MyCalendarListEntry mcle in calendars) {
                    cbGoogleCalendars.Items.Add(mcle);
                }
                cbGoogleCalendars.SelectedIndex = 0;
            }

            bGetGoogleCalendars.Enabled = true;
            cbGoogleCalendars.Enabled = true;
        }

        void cbGoogleCalendars_SelectedIndexChanged(object sender, EventArgs e) {
            Settings.Instance.UseGoogleCalendar = (MyCalendarListEntry)cbGoogleCalendars.SelectedItem;
        }
        #endregion
        #region Sync options
        private void tbDaysInThePast_ValueChanged(object sender, EventArgs e) {
            Settings.Instance.DaysInThePast = int.Parse(tbDaysInThePast.Text);
        }

        private void tbDaysInTheFuture_ValueChanged(object sender, EventArgs e) {
            Settings.Instance.DaysInTheFuture = int.Parse(tbDaysInTheFuture.Text);
        }

        private void tbMinuteOffsets_ValueChanged(object sender, EventArgs e) {
            Settings.Instance.SyncInterval = (int)tbInterval.Value;
            setNextSync(getResyncInterval());
        }

        private void cbIntervalUnit_SelectedIndexChanged(object sender, EventArgs e) {
            Settings.Instance.SyncIntervalUnit = cbIntervalUnit.Text;
            setNextSync(getResyncInterval());
        }

        void CbAddDescriptionCheckedChanged(object sender, EventArgs e) {
            Settings.Instance.AddDescription = cbAddDescription.Checked;
        }

        void CbAddRemindersCheckedChanged(object sender, EventArgs e) {
            Settings.Instance.AddReminders = cbAddReminders.Checked;
        }

        void cbAddAttendees_CheckedChanged(object sender, EventArgs e) {
            Settings.Instance.AddAttendees = cbAddAttendees.Checked;
        }

        void cbConfirmOnDelete_CheckedChanged(object sender, System.EventArgs e) {
            Settings.Instance.ConfirmOnDelete = cbConfirmOnDelete.Checked;
        }

        void cbDisableDeletion_CheckedChanged(object sender, System.EventArgs e) {
            Settings.Instance.DisableDelete = cbDisableDeletion.Checked;
            cbConfirmOnDelete.Enabled = !cbDisableDeletion.Checked;
        }
        #endregion
        #region Application settings
        void CbShowBubbleTooltipsCheckedChanged(object sender, System.EventArgs e) {
            Settings.Instance.ShowBubbleTooltipWhenSyncing = cbShowBubbleTooltips.Checked;
        }

        void CbStartInTrayCheckedChanged(object sender, System.EventArgs e) {
            Settings.Instance.StartInTray = cbStartInTray.Checked;
        }

        void CbMinimizeToTrayCheckedChanged(object sender, System.EventArgs e) {
            Settings.Instance.MinimizeToTray = cbMinimizeToTray.Checked;
        }

        void cbCreateFiles_CheckedChanged(object sender, EventArgs e) {
            Settings.Instance.CreateTextFiles = cbCreateFiles.Checked;
        }
        #endregion

        private void cbVerboseOutput_CheckedChanged(object sender, EventArgs e) {
            Settings.Instance.VerboseOutput = cbVerboseOutput.Checked;
        }
        #endregion

    }
}
