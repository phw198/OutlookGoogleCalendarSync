using System;
using System.Windows.Forms;
using log4net;

namespace OutlookGoogleCalendarSync.Forms {
    public partial class UpdateInfo : Form {
        private static readonly ILog log = LogManager.GetLogger(typeof(UpdateInfo));

        private String version = "";
        private String anchorRequested = "";
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
                dr = ShowDialog();

            } catch (System.Exception ex) {
                log.Debug("A problem was encountered showing the release notes.");
                OGCSexception.Analyse(ex);
                dr = OgcsMessageBox.Show("A new " + (releaseType == "alpha" ? "alpha " : "") + "release of OGCS is available.\nWould you like to upgrade to v" +
                               releaseVersion + " now?", "OGCS Update Available", MessageBoxButtons.YesNo, MessageBoxIcon.Information);
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

        private void btSkipVersion_Click(object sender, EventArgs e) {
            log.Info("User has opted to skip upgrading to this version.");
            Settings.Instance.SkipVersion = version;
            this.Close();
        }

        private void llViewOnGithub_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e) {
            Helper.OpenBrowser(llViewOnGithub.Tag.ToString());
        }
    }
}
