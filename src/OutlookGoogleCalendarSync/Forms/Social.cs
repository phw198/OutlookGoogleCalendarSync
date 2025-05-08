using System;
using System.Windows.Forms;
using Ogcs = OutlookGoogleCalendarSync;

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
                if (Settings.Instance.SuppressSocialPopup) 
                    Forms.Main.Instance.SetControlPropertyThreadSafe(Forms.Main.Instance.cbSuppressSocialPopup, "Checked", false);
            }

            Int32 syncs = Settings.Instance.CompletedSyncs;
            lMilestone.Text = "You've completed " + syncs.ToString() + " syncs!";
            cbSuppressSocialPopup.Checked = Settings.Instance.SuppressSocialPopup;
        }

        #region Events
        #region Spread Word
        private void btSocialSkeet_Click(object sender, EventArgs e) {
            Bluesky_skeet();
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
            Bluesky_follow();
        }

        private void btSocialGitHub_Click(object sender, EventArgs e) {
            GitHub();
        }
        #endregion
        #endregion

        public static void Bluesky_skeet() {
            string text = "I'm using this Outlook-Google calendar sync software - completely #free and feature loaded! #recommended download from https://www.OutlookGoogleCalendarSync.com via @ogcalsync.bsky.social";
            Helper.OpenBrowser("http://bsky.app/intent/post?text=" + urlEncode(text));
        }
        public static void Bluesky_follow() {
            Helper.OpenBrowser("https://bsky.app/profile/ogcalsync.bsky.social");
        }

        public static void Facebook_share() {
            Helper.OpenBrowser("http://www.facebook.com/sharer/sharer.php?u=http://bit.ly/OGCalSync");
        }
        public static void Facebook_like() {
            if (Ogcs.Extensions.MessageBox.Show("Please click the 'Like' button on the project website, which will now open in your browser.",
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
            Forms.Main.Instance.SetControlPropertyThreadSafe(Forms.Main.Instance.cbSuppressSocialPopup, "Checked", cbSuppressSocialPopup.Checked);
        }

        private void pbDonate_Click(object sender, EventArgs e) {
            Program.Donate("Social");
        }

        private void btClose_Click(object sender, EventArgs e) {
            this.Close();
        }
    }
}
