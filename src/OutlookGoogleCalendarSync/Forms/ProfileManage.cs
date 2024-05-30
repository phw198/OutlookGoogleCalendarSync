using Ogcs = OutlookGoogleCalendarSync;
using log4net;
using System;
using System.Windows.Forms;

namespace OutlookGoogleCalendarSync.Forms {
    public partial class ProfileManage : Form {
        private static readonly ILog log = LogManager.GetLogger(typeof(ProfileManage));

        ComboBox ddProfile;

        public ProfileManage(String action, ComboBox ddProfile) {
            InitializeComponent();

            this.btOK.Text = action;
            this.ddProfile = ddProfile;

            if (this.btOK.Text == "Add") {
                this.txtProfileName.Text = "Profile #" + (Settings.Instance.Calendars.Count + 1);
            } else if (this.btOK.Text == "Rename") {
                this.txtProfileName.Text = ddProfile.Text;
            }

            if (Settings.Instance.UserIsBenefactor()) {
                panelDonationNote.Visible = false;
                this.Height = 185;
            } else {
                panelDonationNote.Visible = true;
                this.Height = 310;
            }
        }

        private void ProfileManage_FormClosing(object sender, FormClosingEventArgs e) {
            if ((sender as Form).DialogResult == DialogResult.Cancel) return;

            if (this.btOK.Text == "Add") {
                SettingsStore.Calendar newCalendar = null;
                try {
                    String profileName = this.txtProfileName.Text;
                    if (string.IsNullOrEmpty(profileName)) return;

                    newCalendar = new SettingsStore.Calendar();
                    newCalendar._ProfileName = profileName;
                    Settings.Instance.Calendars.Add(newCalendar);
                    log.Info("Added new calendar settings '" + profileName + "'.");
                    int addedIdx = ddProfile.Items.Add(profileName);
                    ddProfile.SelectedIndex = addedIdx;
                    Forms.Main.Instance.NotificationTray.AddProfileItem(profileName);

                    newCalendar.InitialiseTimer();
                    newCalendar.RegisterForPushSync();
                } catch (System.Exception ex) {
                    ex.Analyse("Failed to add new profile.");
                    throw;
                }

            } else if (this.btOK.Text == "Rename") {
                String currentProfileName = ddProfile.Text;
                String newProfileName = "";
                try {
                    newProfileName = this.txtProfileName.Text;
                    if (newProfileName == "") return;

                    Forms.Main.Instance.ActiveCalendarProfile._ProfileName = newProfileName;
                    int idx = ddProfile.SelectedIndex;
                    ddProfile.Items.RemoveAt(idx);
                    ddProfile.Items.Insert(idx, newProfileName);
                    ddProfile.SelectedIndex = idx;
                    log.Info("Renamed calendar settings from '" + currentProfileName + "' to '" + newProfileName + "'.");

                } catch (System.Exception ex) {
                    ex.Analyse("Failed to rename profile from '" + currentProfileName + "' to '" + newProfileName + "'.");
                    throw;
                }
                Forms.Main.Instance.NotificationTray.RenameProfileItem(currentProfileName, newProfileName);
            }

            Settings.Instance.Save();
        }

        private void pbDonate_Click(object sender, EventArgs e) {
            Program.Donate("Profiles");
        }

        /// <summary>
        /// Detect when F1 is pressed for help
        /// </summary>
        protected override bool ProcessCmdKey(ref Message msg, Keys keyData) {
            try {
                if (keyData == Keys.F1) {
                    try {
                        Helper.OpenBrowser(Program.OgcsWebsite + "/guide/settings");
                        return true; //This keystroke was handled, don't pass to the control with the focus
                    } catch (System.Exception ex) {
                        log.Warn("Failed to process captured F1 key.");
                        Ogcs.Exception.Analyse(ex);
                        System.Diagnostics.Process.Start(Program.OgcsWebsite + "/guide/settings");
                        return true;
                    }
                }
            } catch (System.Exception ex) {
                log.Warn("Failed to process captured command key.");
                Ogcs.Exception.Analyse(ex);
            }
            return base.ProcessCmdKey(ref msg, keyData);
        }
    }
}
