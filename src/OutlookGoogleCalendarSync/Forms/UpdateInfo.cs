using Ogcs = OutlookGoogleCalendarSync;
using log4net;
using System;
using System.Drawing;
using System.Windows.Forms;

namespace OutlookGoogleCalendarSync.Forms {
    public partial class UpdateInfo : Form {
        private static readonly ILog log = LogManager.GetLogger(typeof(UpdateInfo));

        private String version = "";
        private String anchorRequested = "";
        private DialogResult optionChosen = DialogResult.None;
        private String htmlHead = @"
<html>
    <head>
        <meta http-equiv='X-UA-Compatible' content='IE=edge' />
        <link href='https://afeld.github.io/emoji-css/emoji.css' rel='stylesheet'>
        <style>
            body {
                font-family: Arial;
                font-size: 14px;
            }
        </style>
    </head>
    <body>";

        public UpdateInfo(String releaseVersion, String releaseType, String html, out DialogResult dr) {
            InitializeComponent();

            log.Info("Showing new release information.");
            version = releaseVersion;
            dr = DialogResult.Cancel;
            try {
                lTitle.Text = "A new " + (releaseType == "alpha" ? "alpha " : "") + "release of OGCS is available";
                lSummary.Text = "Would you like to upgrade to v" + releaseVersion + " now?";

                if (string.IsNullOrEmpty(html)) {
                    String githubReleaseNotes = Program.OgcsWebsite + "/release-notes";
                    anchorRequested = "v" + releaseVersion.Replace(".", "") + "---" + releaseType;
                    log.Debug("Browser anchor: " + anchorRequested);
                    llViewOnGithub.Tag = githubReleaseNotes +"#"+ anchorRequested;
                    llViewOnGithub.Visible = true;

                } else {
                    llViewOnGithub.Visible = false;
                    html = html.TrimStart("< ![CDATA[".ToCharArray());
                    html = html.TrimEnd("]]>".ToCharArray());
                    html = htmlHead + html + "</body></html>";
                    webBrowser.DocumentText = html;
                }
                if (html.Contains("<h2>Manual Upgrade Required</h2>")) btUpgrade.Enabled = false;
                dr = ShowDialog();

            } catch (System.Exception ex) {
                log.Debug("A problem was encountered showing the release notes.");
                Ogcs.Exception.Analyse(ex);
                dr = Ogcs.Extensions.MessageBox.Show("A new " + (releaseType == "alpha" ? "alpha " : "") + "release of OGCS is available.\nWould you like to upgrade to v" +
                               releaseVersion + " now?", "OGCS Update Available", MessageBoxButtons.YesNo, MessageBoxIcon.Information);
            } finally {
                optionChosen = dr;
            }
        }

        private void WebBrowser_DocumentCompleted(object sender, WebBrowserDocumentCompletedEventArgs e) {
            try {
                if (string.IsNullOrEmpty(anchorRequested)) return;

                HtmlElementCollection anchorElements = webBrowser.Document.Body.GetElementsByTagName("A");
                foreach (HtmlElement element in anchorElements) {
                    if (element.OuterHtml.Contains("user-content-"+ anchorRequested)) {
                        element.ScrollIntoView(true);
                        break;
                    }
                }
            } catch { }
        }

        /// <summary>
        /// Intercept clicked links and open in default browser
        /// </summary>
        private void webBrowser_Navigating(object sender, WebBrowserNavigatingEventArgs e) {
            if (!(e.Url.ToString().Equals("about:blank", StringComparison.InvariantCultureIgnoreCase))) {
                Helper.OpenBrowser(e.Url.ToString());
                e.Cancel = true;
            }
        }

        private void btSkipVersion_Click(object sender, EventArgs e) {
            log.Info("User has opted to skip upgrading to this version.");
            if (version.StartsWith("2") && Settings.Instance.SkipVersion != null && Settings.Instance.SkipVersion.StartsWith("3"))
                Settings.Instance.SkipVersion2 = version;
            else
                Settings.Instance.SkipVersion = version;
            this.Close();
        }

        private void llViewOnGithub_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e) {
            Helper.OpenBrowser(llViewOnGithub.Tag.ToString());
        }

        #region Upgrading
        public void PrepareForUpgrade() {
            btUpgrade.Text = "Upgrading";
            btSkipVersion.Visible = false;
            btLater.Visible = false;

            Point frmLocation = this.DesktopLocation;
            this.Visible = true;
            this.DesktopLocation = frmLocation;
            Application.DoEvents();

            //Copied from InitializeComponent()
            //Recreating webBrowser
            this.webBrowser = new WebBrowser();
            this.wbPanel.Controls.Add(this.webBrowser);
            this.webBrowser.Dock = System.Windows.Forms.DockStyle.Fill;
            this.webBrowser.Location = new System.Drawing.Point(0, 0);
            this.webBrowser.MinimumSize = new System.Drawing.Size(20, 20);
            this.webBrowser.Name = "webBrowser";
            this.webBrowser.ScriptErrorsSuppressed = true;
            this.webBrowser.Size = new System.Drawing.Size(465, 166);
            this.webBrowser.TabIndex = 0;
            this.webBrowser.WebBrowserShortcutsEnabled = false;
            this.webBrowser.Navigating += new System.Windows.Forms.WebBrowserNavigatingEventHandler(this.webBrowser_Navigating);
            this.Controls.Add(wbPanel);

            this.webBrowser.Navigate("about:blank");
            webBrowser.Document.Write(cachedWebPage);
            this.webBrowser.Refresh(WebBrowserRefreshOption.Completely);
            while (webBrowser.ReadyState != WebBrowserReadyState.Complete) {
                System.Threading.Thread.Sleep(250);
                Application.DoEvents();
            }
            Application.DoEvents();
        }

        private int previousProgress = 0;

        public void ShowUpgradeProgress(int i) {
            log.Debug($"Update progress: {i}%");

            Rectangle rect = new Rectangle(0, 0, 0, 0);
            for (int j = previousProgress; j <= (btUpgrade.Width * i / 100); j++) {
                Bitmap bmp = new Bitmap(btUpgrade.Width, btUpgrade.Height);
                Graphics g = Graphics.FromImage(bmp);
                rect = new Rectangle(0, 0, Math.Max(5, j), bmp.Height);
                using (var b1 = new System.Drawing.SolidBrush(Color.LimeGreen))
                    g.FillRectangle(b1, rect);
                btUpgrade.BackgroundImage = bmp;
                Application.DoEvents();
                System.Threading.Thread.Sleep(50);
            }
            previousProgress = rect.Width;
        }

        public void UpgradeCompleted() {
            if (this.Visible) btUpgrade.Text = "Restart OGCS";
            else
                Ogcs.Extensions.MessageBox.Show("The application has been updated and will now restart.",
                    "OGCS successfully updated!", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
        public Boolean AwaitingRestart { 
            get { return btUpgrade.Text == "Restart OGCS" && this.Visible; }
        }

        private String cachedWebPage;

        private void btUpgrade_Click(object sender, EventArgs e) {
            if (this.btUpgrade.Text == "Upgrading") return;
            
            if (this.AwaitingRestart && !this.IsDisposed) {
                this.btUpgrade.Text = "Restarting";
                this.Dispose();
            }
            
            cachedWebPage = webBrowser.DocumentText;
            this.webBrowser.Dispose();
        }

        private void UpdateInfo_FormClosed(object sender, FormClosedEventArgs e) {
            log.Info("Closed. " + e.CloseReason.ToString());
            this.optionChosen = DialogResult.Cancel;
        }

        private void UpdateInfo_FormClosing(object sender, FormClosingEventArgs e) {
            log.Info("Closing. " + e.CloseReason.ToString());
        }
        #endregion
    }
}
