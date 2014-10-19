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
        private string VERSION = "1.1.1";

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
            cbAddAttendees.Checked = Settings.Instance.AddAttendeesToDescription;
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
            }
        }
        #endregion

        void SyncNow_Click(object sender, EventArgs e) {
            LogBox.Clear();
            
            if (Settings.Instance.UseGoogleCalendar.Id == "") {
                MessageBox.Show("You need to select a Google Calendar first on the 'Settings' tab.");
                return;
            }
            //Check network availability
            if (!System.Net.NetworkInformation.NetworkInterface.GetIsNetworkAvailable()) {
                logboxout("There does not appear to be any network available! Sync aborted.");
                return;
            }
            bSyncNow.Enabled = false;
            lNextSyncVal.Text = "In progress...";

            DateTime SyncStarted = DateTime.Now;
            logboxout("Sync started at " + SyncStarted.ToString());
            logboxout("--------------------------------------------------");

            Boolean syncOk = false;
            int failedAttempts = 0;
            while (!syncOk) {
                if (failedAttempts > 0 &&
                    MessageBox.Show("The synchronisation failed. Do you want to try again?", "Sync Failed",
                    MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation) == System.Windows.Forms.DialogResult.No) { break; }
                syncOk = synchronize();
                failedAttempts += !syncOk ? 1 : 0;
            }

            logboxout("--------------------------------------------------");
            logboxout(syncOk ? "Sync finished with success!" : "Operation aborted after "+ failedAttempts +" failed attempts!");

            if (syncOk) {
                lastSyncDate = SyncStarted;
                Settings.Instance.LastSyncDate = lastSyncDate;
                lLastSyncVal.Text = SyncStarted.ToLongDateString() + " - " + SyncStarted.ToLongTimeString();
                setNextSync(getResyncInterval());
            } else {
                logboxout("Another sync has been scheduled to automatically run in 5 minutes time.");
                setNextSync(5);
            }
            bSyncNow.Enabled = true;
        }

        Boolean synchronize() {
            logboxout("Reading Outlook Calendar Entries...");
            List<AppointmentItem> OutlookEntries = null;
            try {
                OutlookEntries = OutlookCalendar.Instance.getCalendarEntriesInRange();
            } catch (System.Exception ex) {
                logboxout("Unable to access the Outlook Calendar. The following error occurred:");
                logboxout(ex.Message + "\r\n => Retry later.");
                OutlookCalendar.Instance.Reset(); 
                return false;
            }
            if (cbCreateFiles.Checked) {
                TextWriter tw = new StreamWriter("export_found_in_outlook.txt");
                foreach (AppointmentItem ai in OutlookEntries) {
                    tw.WriteLine(signature(ai));
                }
                tw.Close();
            }
            logboxout("Found " + OutlookEntries.Count + " Outlook Calendar Entries.");
            logboxout("--------------------------------------------------");



            logboxout("Reading Google Calendar Entries...");
            List<Event> GoogleEntries = null;
            try {
                GoogleEntries = GoogleCalendar.Instance.getCalendarEntriesInRange();
            } catch (System.Exception ex) {
                logboxout("Unable to connect to the Google Calendar. The following error occurred:");
                logboxout(ex.Message + "\r\n => Check your network connection.");
                return false;
            }

            if (cbCreateFiles.Checked) {
                TextWriter tw = new StreamWriter("export_found_in_google.txt");
                foreach (Event ev in GoogleEntries) {
                    tw.WriteLine(signature(ev));
                }
                tw.Close();
            }
            logboxout("Found " + GoogleEntries.Count + " Google Calendar Entries.");
            logboxout("--------------------------------------------------");


            //  Make copies of each list of events (Not strictly needed)
            List<AppointmentItem> OutlookEntriesToBeCreated = new List<AppointmentItem>(OutlookEntries);
            List<Event> GoogleEntriesToBeDeleted = new List<Event>(GoogleEntries);
            IdentifyGoogleAddDeletes(OutlookEntriesToBeCreated, GoogleEntriesToBeDeleted);

            if (Settings.Instance.DisableDelete) {
                GoogleEntriesToBeDeleted = new List<Event>();
            } else {
                if (cbCreateFiles.Checked) {
                    TextWriter tw = new StreamWriter("export_to_be_deleted.txt");
                    foreach (Event ev in GoogleEntriesToBeDeleted) {
                        tw.WriteLine(signature(ev));
                    }
                    tw.Close();
                }
                logboxout(GoogleEntriesToBeDeleted.Count + " Google Calendar Entries to be deleted.");
            }

            //OutlookEntriesToBeCreated ...in Google!
            if (cbCreateFiles.Checked) {
                TextWriter tw = new StreamWriter("export_to_be_created.txt");
                foreach (AppointmentItem ai in OutlookEntriesToBeCreated) {
                    tw.WriteLine(signature(ai));
                }
                tw.Close();
            }
            logboxout(OutlookEntriesToBeCreated.Count + " Entries to be created in Google.");

            if (GoogleEntriesToBeDeleted.Count > 0) {
                logboxout("--------------------------------------------------");
                logboxout("Deleting " + GoogleEntriesToBeDeleted.Count + " Google Calendar Entries...");
                foreach (Event ev in GoogleEntriesToBeDeleted) {
                    String eventSummary = "";
                    Boolean delete = true;
                    
                    if (Settings.Instance.ConfirmOnDelete) {
                        eventSummary = DateTime.Parse(ev.Start.DateTime.ToString()).ToString("dd/MM/yyyy hh:mm") +" => ";
                        eventSummary += '"'+ ev.Summary +'"';
                        if (MessageBox.Show("Delete " + eventSummary + "?", "Deletion Confirmation", 
                            MessageBoxButtons.YesNo, MessageBoxIcon.Question) == System.Windows.Forms.DialogResult.No)
                        {
                            delete = false;
                            logboxout("Not deleted: " + eventSummary.ToString());
                        }
                    }
                    if (delete) {
                        try {
                            GoogleCalendar.Instance.deleteCalendarEntry(ev);
                            if (Settings.Instance.ConfirmOnDelete) logboxout("Deleted: " + eventSummary);
                        } catch (System.Exception ex) {
                            logboxout("Unable to delete obsolete entries out to the Google Calendar. The following error occurred:");
                            logboxout(ex.Message + "\r\n => Check your network connection.");
                            return false;
                        }
                    }
                }
                logboxout("Done.");
            }

            if (OutlookEntriesToBeCreated.Count > 0) {
                logboxout("--------------------------------------------------");
                logboxout("Creating " + OutlookEntriesToBeCreated.Count + " Entries in Google...");
                foreach (AppointmentItem ai in OutlookEntriesToBeCreated) {
                    Event ev = new Event();

                    //Add the Outlook appointment ID into Google event.
                    //This will make comparison more efficient and set the scene for 2-way sync.
                    ev.ExtendedProperties = new Event.ExtendedPropertiesData();
                    ev.ExtendedProperties.Private = new Event.ExtendedPropertiesData.PrivateData();
                    ev.ExtendedProperties.Private.Add("outlook_EntryID", ai.EntryID.ToString());

                    ev.Start = new EventDateTime();
                    ev.End = new EventDateTime();

                    if (ai.AllDayEvent) {
                        ev.Start.Date = ai.Start.ToString("yyyy-MM-dd");
                        ev.End.Date = ai.End.ToString("yyyy-MM-dd");
                    } else {
                        ev.Start.DateTime = GoogleCalendar.Instance.GoogleTimeFrom(ai.Start);
                        ev.End.DateTime = GoogleCalendar.Instance.GoogleTimeFrom(ai.End);
                    }
                    ev.Summary = ai.Subject;
                    if (cbAddDescription.Checked) ev.Description = ai.Body;
                    ev.Location = ai.Location;

                    ev.Organizer = new Event.OrganizerData();
                    ev.Organizer.Self = (ai.Recipients.Count == 0);

                    if (cbAddAttendees.Checked) {
                        ev.Attendees = new List<EventAttendee>();
                        foreach (Microsoft.Office.Interop.Outlook.Recipient recipient in ai.Recipients) {
                            Microsoft.Office.Interop.Outlook.PropertyAccessor pa = recipient.PropertyAccessor;
                            EventAttendee ea = new EventAttendee();
                            ea.DisplayName = recipient.Name;
                            ea.Email = pa.GetProperty(OutlookCalendar.PR_SMTP_ADDRESS).ToString();
                            ea.Optional = (ai.OptionalAttendees != null && ai.OptionalAttendees.Contains(recipient.Name));
                            if (ai.Organizer == recipient.Name) {
                                ea.Organizer = true;
                                ev.Organizer.Self = false;
                                ev.Organizer.DisplayName = ea.DisplayName;
                                ev.Organizer.Email = ea.Email;
                            }
                            ea.Self = (OutlookCalendar.Instance.CurrentUserName == recipient.Name);
                            switch (recipient.MeetingResponseStatus) {
                                case OlResponseStatus.olResponseNone: ea.ResponseStatus = "needsAction"; break;
                                case OlResponseStatus.olResponseAccepted: ea.ResponseStatus = "accepted"; break;
                                case OlResponseStatus.olResponseDeclined: ea.ResponseStatus = "declined"; break;
                                case OlResponseStatus.olResponseTentative: ea.ResponseStatus = "tentative"; break;
                            }
                            ev.Attendees.Add(ea);
                        }
                    }

                    //consider the reminder set in Outlook
                    if (cbAddReminders.Checked && ai.ReminderSet) {
                        ev.Reminders = new Event.RemindersData();
                        ev.Reminders.UseDefault = false;
                        EventReminder reminder = new EventReminder();
                        reminder.Method = "popup";
                        reminder.Minutes = ai.ReminderMinutesBeforeStart;
                        ev.Reminders.Overrides = new List<EventReminder>();
                        ev.Reminders.Overrides.Add(reminder);
                    }

                    try {
                        GoogleCalendar.Instance.addEntry(ev);
                    } catch (System.Exception ex) {
                        logboxout("Unable to add new entries into the Google Calendar. The following error occurred:");
                        logboxout(ex.Message + "\r\n => Check your network connection.");
                        return false;
                    }
                }

                logboxout("Done.");
            }
            return true;
        }

        //<summary>New logic for comparing Outlook and Google events works as follows:
  	    //      1.  Scan through both lists looking for duplicates
  	    //      2.  Remove found duplicates from both lists
  	    //      3.  Items remaining in Outlook list are new and need to be created
  	    //      4.  Items remaining in Google list need to be deleted
  	    //</summary>
        public void IdentifyGoogleAddDeletes(List<AppointmentItem> outlook, List<Event> google) {
            // Count backwards so that we can remove found items without affecting the order of remaining items
            for (int o = outlook.Count - 1; o >= 0; o--) {
                for (int g = google.Count - 1; g >= 0; g--) {
                    if (google[g].ExtendedProperties != null &&
                        google[g].ExtendedProperties.Private.ContainsKey("outlook_EntryID") &&
                        outlook[o].EntryID == google[g].ExtendedProperties.Private["outlook_EntryID"])
                    {
                        outlook.Remove(outlook[o]);
                        google.Remove(google[g]);
                        break;
                    }
                }
            }
        }

        //public List<Event> IdentifyGoogleEntriesToBeDeleted(List<AppointmentItem> outlook, List<Event> google) {
        //    List<Event> result = new List<Event>();
        //    foreach (Event g in google) {
        //        bool found = false;
        //        foreach (AppointmentItem o in outlook) {
        //            if (g.ExtendedProperties != null &&
        //                g.ExtendedProperties.Private.ContainsKey("outlook_EntryID") &&
        //                o.EntryID == g.ExtendedProperties.Private["outlook_EntryID"]) {
        //                found = true;
        //            }
        //        }
        //        if (!found) result.Add(g);
        //    }
        //    return result;
        //}

        //public List<AppointmentItem> IdentifyOutlookEntriesToBeCreated(List<AppointmentItem> outlook, List<Event> google) {
        //    List<AppointmentItem> result = new List<AppointmentItem>();
        //    foreach (AppointmentItem o in outlook) {
        //        bool found = false;
        //        foreach (Event g in google) {
        //            if (g.ExtendedProperties != null &&
        //                g.ExtendedProperties.Private.ContainsKey("outlook_EntryID") &&
        //                g.ExtendedProperties.Private.ContainsValue(o.EntryID)) {
        //                found = true;
        //            }
        //        }
        //        if (!found) result.Add(o);
        //    }
        //    return result;
        //}

        //creates a standardized summary string with the key attributes of a calendar entry for comparison
        public string signature(AppointmentItem ai) {
            return (GoogleCalendar.Instance.GoogleTimeFrom(ai.Start) + ";" + GoogleCalendar.Instance.GoogleTimeFrom(ai.End) + ";" + ai.Subject + ";" + ai.Location).Trim();
        }
        public string signature(Event ev) {
            if (ev.Start.DateTime == null) ev.Start.DateTime = GoogleCalendar.Instance.GoogleTimeFrom(DateTime.Parse(ev.Start.Date));
            if (ev.End.DateTime == null) ev.End.DateTime = GoogleCalendar.Instance.GoogleTimeFrom(DateTime.Parse(ev.End.Date));
            return (ev.Start.DateTime + ";" + ev.End.DateTime + ";" + ev.Summary + ";" + ev.Location).Trim();
        }

        void logboxout(string s) {
            LogBox.Text += s + Environment.NewLine;
        }

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

        public void HandleException(System.Exception ex) {
            MessageBox.Show(ex.ToString(), "Exception!", MessageBoxButtons.OK, MessageBoxIcon.Error);
            TextWriter tw = new StreamWriter("exception.txt");
            tw.WriteLine(ex.ToString());
            tw.Close();

            this.Close();
            System.Environment.Exit(-1);
            System.Windows.Forms.Application.Exit();
        }

        void LinkLabel1LinkClicked(object sender, LinkLabelLinkClickedEventArgs e) {
            System.Diagnostics.Process.Start(lAboutURL.Text);
        }

        #region EVENTS
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
                logboxout("Unable to get the list of Google Calendars. The following error occurred:");
                logboxout(ex.Message + "\r\n => Check your network connection.");
            }
            if (calendars != null) {
                cbGoogleCalendars.Items.Clear();
                foreach (MyCalendarListEntry mcle in calendars) {
                    cbGoogleCalendars.Items.Add(mcle);
                }
                MainForm.Instance.cbGoogleCalendars.SelectedIndex = 0;
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
            Settings.Instance.AddAttendeesToDescription = cbAddAttendees.Checked;
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
        #endregion

    }
}
