using System;
using System.Windows.Forms;

namespace OutlookGoogleCalendarSync.Forms {
    public partial class Social : Form {
        private ToolTip toolTips;
        public Social() {
            InitializeComponent();
            
            toolTips = new ToolTip();
            toolTips.AutoPopDelay = 10000;
            toolTips.InitialDelay = 500;
            toolTips.ReshowDelay = 200;
            toolTips.ShowAlways = true;
            if (Settings.Instance.UserIsBenefactor()) {
                pbDonate.Visible = false;
                lDonateTip.Visible = false;
            } else {
                toolTips.SetToolTip(cbSuppressSocialPopup, "Donate £10 or more to enable this feature.");
                if (Settings.Instance.SuppressSocialPopup) Settings.Instance.SuppressSocialPopup = false;
            }

            Int32 syncs = Settings.Instance.CompletedSyncs;
            lMilestone.Text = "You've completed " + syncs.ToString() + " syncs!";
            cbSuppressSocialPopup.Checked = Settings.Instance.SuppressSocialPopup;
        }

        #region Events
        #region Spread Word
        private void btSocialTweet_Click(object sender, EventArgs e) {
            Twitter_tweet();
        }

        private void btSocialFB_Click(object sender, EventArgs e) {
            Facebook_share();
        }

        private void btFbLike_Click(object sender, EventArgs e) {
            if (MessageBox.Show("Please click the 'Like' button on the project website, which will now open in your browser.",
                "Like on Facebook", MessageBoxButtons.OKCancel, MessageBoxIcon.Information) == DialogResult.OK) {
                System.Diagnostics.Process.Start("https://phw198.github.io/OutlookGoogleCalendarSync");
            }
        }

        private void btSocialGplus_Click_1(object sender, EventArgs e) {
            Google_share();
        }

        private void btSocialLinkedin_Click(object sender, EventArgs e) {
            Linkedin_share();
        }
        #endregion

        #region Keep in touch
        private void btSocialRSSfeed_Click(object sender, EventArgs e) {
            RSS_follow();
        }

        private void pbSocialTwitterFollow_Click(object sender, EventArgs e) {
            Twitter_follow();
        }

        private void btSocialGComm_Click(object sender, EventArgs e) {
            Google_goToCommunity();
        }

        private void btSocialGitHub_Click(object sender, EventArgs e) {
            gitHub();
        }
        #endregion
        #endregion

        private void btSocialGplus_Click(object sender, EventArgs e) {
            Social.Google_share();
        }
        private void pbSocialGplusCommunity_Click(object sender, EventArgs e) {
            Social.Google_goToCommunity();
        }

        public static void Google_goToCommunity() {
            System.Diagnostics.Process.Start("https://plus.google.com/communities/114412828247015553563");
        }
        public static void Google_share() {
            System.Diagnostics.Process.Start("https://plus.google.com/share?&url=http://bit.ly/OGCalSync");
        }

        public static void Twitter_tweet() {
            string text = "I'm using this Outlook-Google calendar sync tool - completely #free and feature loaded. #recommend";
            System.Diagnostics.Process.Start("http://twitter.com/intent/tweet?&url=http://bit.ly/OGCalSync&text=" + urlEncode(text) + "&via=ogcalsync");
        }
        public static void Twitter_follow() {
            System.Diagnostics.Process.Start("https://twitter.com/OGcalsync");
        }

        public static void Facebook_share() {
            System.Diagnostics.Process.Start("http://www.facebook.com/sharer/sharer.php?u=http://bit.ly/OGCalSync");
        }

        public static void RSS_follow() {
            System.Diagnostics.Process.Start("https://github.com/phw198/OutlookGoogleCalendarSync/releases.atom");
        }

        public static void Linkedin_share() {
            string text = "I'm using this Outlook-Google calendar sync tool - completely #free and feature loaded. #recommend";
            System.Diagnostics.Process.Start("http://www.linkedin.com/shareArticle?mini=true&url=http://bit.ly/OGCalSync&summary=" + urlEncode(text));
        }

        private static void gitHub() {
            System.Diagnostics.Process.Start("https://github.com/phw198/OutlookGoogleCalendarSync/");
        }

        private static String urlEncode(String text) {
            return text.Replace("#", "%23");
        }

        private void cbSuppressSocialPopup_CheckedChanged(object sender, EventArgs e) {
            if (!Settings.Instance.UserIsBenefactor()) {
                cbSuppressSocialPopup.CheckedChanged -= cbSuppressSocialPopup_CheckedChanged;
                cbSuppressSocialPopup.Checked = false;
                cbSuppressSocialPopup.CheckedChanged += cbSuppressSocialPopup_CheckedChanged;
                toolTips.SetToolTip(cbSuppressSocialPopup, "Donate £10 or more to enable this feature.");
            }
            Settings.Instance.SuppressSocialPopup = cbSuppressSocialPopup.Checked;
        }

        private void pbDonate_Click(object sender, EventArgs e) {
            Program.Donate();
        }

        private void btClose_Click(object sender, EventArgs e) {
            this.Close();
        }
    }
}
