using System;
using System.Windows.Forms;

namespace OutlookGoogleCalendarSync {
    public partial class Splash : Form {
        
        private DateTime splashed;

        public Splash() {
            this.splashed = DateTime.Now;
            InitializeComponent();
            lVersion.Text = "v" + Application.ProductVersion;
            String completedSyncs = XMLManager.ImportElement("CompletedSyncs", Program.SettingsFile) ?? "0";
            if (completedSyncs == "0")
                lSyncCount.Visible = false;
            else {
                lSyncCount.Text = lSyncCount.Text.Replace("{syncs}", String.Format("{0:n0}", completedSyncs));
                lSyncCount.Left = (panel1.Width - (lSyncCount.Width)) / 2;
            }
        }

        public void Remove() {
            while (DateTime.Now < splashed.AddSeconds((System.Diagnostics.Debugger.IsAttached ? 1 : 8)) && !this.IsDisposed) {
                Application.DoEvents();
                System.Threading.Thread.Sleep(100);
            }
            if (!this.IsDisposed) this.Close();
        }

        private void pbDonate_Click(object sender, EventArgs e) {
            Social.Donate();
            this.Close();
        }

        private void pbSocialGplusCommunity_Click(object sender, EventArgs e) {
            Social.Google_goToCommunity();
            this.Close();
        }

        private void pbSocialTwitterFollow_Click(object sender, EventArgs e) {
            Social.Twitter_follow();
            this.Close();
        }
    }
}
