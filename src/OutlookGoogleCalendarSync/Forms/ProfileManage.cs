using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using log4net;

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

            if (!Settings.Instance.UserIsBenefactor()) {
                panelDonationNote.Visible = false;
                this.Height = 135;
            } else {
                panelDonationNote.Visible = true;
                this.Height = 255;
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
                    OGCSexception.Analyse("Failed to add new profile.", ex);
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
                    OGCSexception.Analyse("Failed to rename profile from '" + currentProfileName + "' to '" + newProfileName + "'.", ex);
                    throw;
                }
                Forms.Main.Instance.NotificationTray.RenameProfileItem(currentProfileName, newProfileName);
            }
        }

        private void pbDonate_Click(object sender, EventArgs e) {
            Program.Donate("Profiles");
        }
    }
}
