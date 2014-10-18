//TODO: consider description updates?
//TODO: optimize comparison algorithms
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
        public string VERSION = "1.0.11";

        public Timer ogstimer;
        public DateTime oldtime;
        public List<int> MinuteOffsets = new List<int>();
        DateTime lastSyncDate;
        int currentTimerInterval = 0;

        public MainForm() {
            InitializeComponent();
            label4.Text = label4.Text.Replace("{version}", VERSION);

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

            //update GUI from Settings
            tbDaysInThePast.Text = Settings.Instance.DaysInThePast.ToString();
            tbDaysInTheFuture.Text = Settings.Instance.DaysInTheFuture.ToString();
            tbMinuteOffsets.Text = Settings.Instance.MinuteOffsets;
            lastSyncDate = Settings.Instance.LastSyncDate;
            cbCalendars.Items.Add(Settings.Instance.UseGoogleCalendar);
            cbCalendars.SelectedIndex = 0;
            cbSyncEveryHour.Checked = Settings.Instance.SyncEveryHour;
            cbShowBubbleTooltips.Checked = Settings.Instance.ShowBubbleTooltipWhenSyncing;
            cbStartInTray.Checked = Settings.Instance.StartInTray;
            cbMinimizeToTray.Checked = Settings.Instance.MinimizeToTray;
            cbAddDescription.Checked = Settings.Instance.AddDescription;
            cbAddAttendees.Checked = Settings.Instance.AddAttendeesToDescription;
            cbAddReminders.Checked = Settings.Instance.AddReminders;
            cbCreateFiles.Checked = Settings.Instance.CreateTextFiles;

            //Mailboxes the user has access to
            this.ddMailboxName.SelectedIndexChanged -= ddMailboxName_SelectedIndexChanged;
            if (OutlookCalendar.Instance.Accounts.Count == 1) {
                cbAlternateMailbox.Enabled = false;
                cbAlternateMailbox.Checked = false;
                this.ddMailboxName.Enabled = false;
            } else {
                cbAlternateMailbox.Checked = Settings.Instance.AlternateMailbox;
            }

            for (int acc=2; acc<=OutlookCalendar.Instance.Accounts.Count; acc++) {
                String mailbox = OutlookCalendar.Instance.Accounts[acc].SmtpAddress.ToLower();
                this.ddMailboxName.Items.Add(mailbox);
                if (Settings.Instance.MailboxName == mailbox) { this.ddMailboxName.SelectedIndex = acc; }
            }
            if (!Settings.Instance.AlternateMailbox) { 
                this.ddMailboxName.SelectedIndex = 0;
                this.ddMailboxName.Enabled = false;
            }
            this.ddMailboxName.SelectedIndexChanged += ddMailboxName_SelectedIndexChanged;

            //Start in tray?
            if (cbStartInTray.Checked) {
                this.WindowState = FormWindowState.Minimized;
                notifyIcon1.Visible = true;
                this.Hide();
                this.ShowInTaskbar = false;
            }

            //set up tooltips for some controls
            ToolTip toolTip1 = new ToolTip();
            toolTip1.AutoPopDelay = 10000;
            toolTip1.InitialDelay = 500;
            toolTip1.ReshowDelay = 200;
            toolTip1.ShowAlways = true;
            toolTip1.SetToolTip(cbCalendars,
                "The Google Calendar to synchonize with.");
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
            toolTip1.SetToolTip(cbAlternateMailbox,
                "Only change this if you need to use an Outlook Calendar that is not in the default mailbox");

            //Refresh synchronizations (last and next)
            lLastSyncVal.Text = lastSyncDate.ToLongDateString() + " - " + lastSyncDate.ToLongTimeString();
            setNextSync(getResyncInterval());
        }

        int getResyncInterval() {
            int min = 0;
            int.TryParse(tbMinuteOffsets.Text, out min);
            if (min < 1) { min = 60; }
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
            if (cbSyncEveryHour.Checked) {
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

        void GetMyGoogleCalendars_Click(object sender, EventArgs e) {
            bGetMyCalendars.Enabled = false;
            cbCalendars.Enabled = false;
            List<MyCalendarListEntry> calendars = null;
            try {
                calendars = GoogleCalendar.Instance.getCalendars();
            } catch (System.Exception ex) {
                logboxout("Unable to get the list of Google Calendars. The following error occurred:");
                logboxout(ex.Message + "\r\n => Check your network connection.");
            }
            if (calendars != null) {
                cbCalendars.Items.Clear();
                foreach (MyCalendarListEntry mcle in calendars) {
                    cbCalendars.Items.Add(mcle);
                }
                MainForm.Instance.cbCalendars.SelectedIndex = 0;
            }

            bGetMyCalendars.Enabled = true;
            cbCalendars.Enabled = true;
        }

        void SyncNow_Click(object sender, EventArgs e) {
            bSyncNow.Enabled = false;

            lNextSyncVal.Text = "In progress...";

            LogBox.Clear();

            DateTime SyncStarted = DateTime.Now;

            logboxout("Sync started at " + SyncStarted.ToString());
            logboxout("--------------------------------------------------");

            Boolean syncOk = synchronize();
            logboxout("--------------------------------------------------");
            logboxout(syncOk ? "Sync finished with success!" : "Operation aborted!");

            if (syncOk) {
                lastSyncDate = SyncStarted;
                Settings.Instance.LastSyncDate = lastSyncDate;
                lLastSyncVal.Text = SyncStarted.ToLongDateString() + " - " + SyncStarted.ToLongTimeString();
                setNextSync(getResyncInterval());
            } else {
                setNextSync(5);
            }
            bSyncNow.Enabled = true;
        }

        Boolean synchronize() {
            if (Settings.Instance.UseGoogleCalendar.Id == "") {
                MessageBox.Show("You need to select a Google Calendar first on the 'Settings' tab.");
                return false;
            }

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


            List<Event> GoogleEntriesToBeDeleted = IdentifyGoogleEntriesToBeDeleted(OutlookEntries, GoogleEntries);
            if (cbCreateFiles.Checked) {
                TextWriter tw = new StreamWriter("export_to_be_deleted.txt");
                foreach (Event ev in GoogleEntriesToBeDeleted) {
                    tw.WriteLine(signature(ev));
                }
                tw.Close();
            }
            logboxout(GoogleEntriesToBeDeleted.Count + " Google Calendar Entries to be deleted.");

            //OutlookEntriesToBeCreated ...in Google!
            List<AppointmentItem> OutlookEntriesToBeCreated = IdentifyOutlookEntriesToBeCreated(OutlookEntries, GoogleEntries);
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
                try {
                    foreach (Event ev in GoogleEntriesToBeDeleted) GoogleCalendar.Instance.deleteCalendarEntry(ev);
                } catch (System.Exception ex) {
                    logboxout("Unable to delete obsolete entries out to the Google Calendar. The following error occurred:");
                    logboxout(ex.Message + "\r\n => Check your network connection.");
                    return false;
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
                            ea.Self = (OutlookCalendar.Instance.CalendarUserName == recipient.Name);
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

        public List<Event> IdentifyGoogleEntriesToBeDeleted(List<AppointmentItem> outlook, List<Event> google) {
            List<Event> result = new List<Event>();
            foreach (Event g in google) {
                bool found = false;
                foreach (AppointmentItem o in outlook) {
                    if (g.ExtendedProperties != null &&
                        g.ExtendedProperties.Private.ContainsKey("outlook_EntryID") &&
                        o.EntryID == g.ExtendedProperties.Private["outlook_EntryID"]) {
                        found = true;
                    }
                }
                if (!found) result.Add(g);
            }
            return result;
        }

        public List<AppointmentItem> IdentifyOutlookEntriesToBeCreated(List<AppointmentItem> outlook, List<Event> google) {
            List<AppointmentItem> result = new List<AppointmentItem>();
            foreach (AppointmentItem o in outlook) {
                bool found = false;
                foreach (Event g in google) {
                    if (g.ExtendedProperties != null &&
                        g.ExtendedProperties.Private.ContainsKey("outlook_EntryID") &&
                        g.ExtendedProperties.Private.ContainsValue(o.EntryID)) {
                        found = true;
                    }
                }
                if (!found) result.Add(o);
            }
            return result;
        }

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

        void ComboBox1SelectedIndexChanged(object sender, EventArgs e) {
            Settings.Instance.UseGoogleCalendar = (MyCalendarListEntry)cbCalendars.SelectedItem;
        }

        void TbDaysInThePastTextChanged(object sender, EventArgs e) {
            Settings.Instance.DaysInThePast = int.Parse(tbDaysInThePast.Text);
        }

        void TbDaysInTheFutureTextChanged(object sender, EventArgs e) {
            try {
                Settings.Instance.DaysInTheFuture = int.Parse(tbDaysInTheFuture.Text);
            } catch {
                Settings.Instance.DaysInTheFuture = 1;
            }
        }

        void TbMinuteOffsetsTextChanged(object sender, EventArgs e) {
            Settings.Instance.MinuteOffsets = tbMinuteOffsets.Text;
            setNextSync(getResyncInterval());
        }

        void CbSyncEveryHourCheckedChanged(object sender, System.EventArgs e) {
            Settings.Instance.SyncEveryHour = cbSyncEveryHour.Checked;
            setNextSync(getResyncInterval());
        }

        void CbShowBubbleTooltipsCheckedChanged(object sender, System.EventArgs e) {
            Settings.Instance.ShowBubbleTooltipWhenSyncing = cbShowBubbleTooltips.Checked;
        }

        void CbStartInTrayCheckedChanged(object sender, System.EventArgs e) {
            Settings.Instance.StartInTray = cbStartInTray.Checked;
        }

        void CbMinimizeToTrayCheckedChanged(object sender, System.EventArgs e) {
            Settings.Instance.MinimizeToTray = cbMinimizeToTray.Checked;
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

        void cbCreateFiles_CheckedChanged(object sender, EventArgs e) {
            Settings.Instance.CreateTextFiles = cbCreateFiles.Checked;
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
            System.Diagnostics.Process.Start(linkLabel1.Text);
        }

        private void cbAlternateMailbox_CheckedChanged(object sender, EventArgs e) {
            Settings.Instance.AlternateMailbox = cbAlternateMailbox.Checked;
            Settings.Instance.MailboxName = (cbAlternateMailbox.Checked ? ddMailboxName.Text : "");
            this.ddMailboxName.Enabled = cbAlternateMailbox.Checked;
            OutlookCalendar.Instance.Reset();
        }

        private void ddMailboxName_SelectedIndexChanged(object sender, EventArgs e) {
            Settings.Instance.MailboxName = ddMailboxName.Text;
            OutlookCalendar.Instance.Reset();
        }
    }
}
