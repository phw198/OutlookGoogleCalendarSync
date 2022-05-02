using System;
using System.Windows.Forms;

namespace OutlookGoogleCalendarSync.Forms {
    public partial class Social : Form {
        private ToolTip toolTips;
        public Social() {
            InitializeComponent();

            toolTips = new ToolTip {
                AutoPopDelay = 10000,
                InitialDelay = 500,
                ReshowDelay = 200,
                ShowAlways = true
            };
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
            Facebook_like();
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

        private void btSocialGitHub_Click(object sender, EventArgs e) {
            GitHub();
        }
        #endregion
        #endregion

        public static void Twitter_tweet() {
            string text = "I'm using this Outlook-Google calendar sync tool - completely #free and feature loaded. #recommend";
            Helper.OpenBrowser("http://twitter.com/intent/tweet?&url=http://bit.ly/OGCalSync&text=" + urlEncode(text) + "&via=ogcalsync");
        }
        public static void Twitter_follow() {
            Helper.OpenBrowser("https://twitter.com/OGcalsync");
        }

        public static void Facebook_share() {
            Helper.OpenBrowser("http://www.facebook.com/sharer/sharer.php?u=http://bit.ly/OGCalSync");
        }
        public static void Facebook_like() {
            if (OgcsMessageBox.Show("Please click the 'Like' button on the project website, which will now open in your browser.",
                "Like on Facebook", MessageBoxButtons.OKCancel, MessageBoxIcon.Information) == DialogResult.OK) {
                Helper.OpenBrowser("https://phw198.github.io/OutlookGoogleCalendarSync");
            }
        }

        public static void RSS_follow() {
            Helper.OpenBrowser("https://github.com/phw198/OutlookGoogleCalendarSync/releases.atom");
        }

        public static void Linkedin_share() {
            string text = "I'm using this Outlook-Google calendar sync tool - completely #free and feature loaded. #recommend";
            Helper.OpenBrowser("http://www.linkedin.com/shareArticle?mini=true&url=http://bit.ly/OGCalSync&summary=" + urlEncode(text));
        }

        public static void GitHub() {
            Helper.OpenBrowser("https://github.com/phw198/OutlookGoogleCalendarSync/");
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
            Program.Donate("Social");
        }

        private void btClose_Click(object sender, EventArgs e) {
            this.Close();
        }
    }
}
