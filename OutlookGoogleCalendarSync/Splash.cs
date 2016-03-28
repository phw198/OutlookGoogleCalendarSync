using log4net;
using System;
using System.Threading;
using System.Windows.Forms;

namespace OutlookGoogleCalendarSync {
    public partial class Splash : Form {
        private static readonly ILog log = LogManager.GetLogger(typeof(Splash));
        
        private static Thread splashThread;
        private static Splash splash;

        public Splash() {
            InitializeComponent();
        }

        public static void ShowMe() {
            if (splashThread == null) {
                splashThread = new Thread(new ThreadStart(doShowSplash));
                splashThread.IsBackground = true;
                splashThread.Start();
            }
        }
        private static void doShowSplash() {
            if (splash == null)
                splash = new Splash();

            splash.lVersion.Text = "v" + Application.ProductVersion;
            String completedSyncs = XMLManager.ImportElement("CompletedSyncs", Program.SettingsFile) ?? "0";
            if (completedSyncs == "0")
                splash.lSyncCount.Visible = false;
            else {
                splash.lSyncCount.Text = splash.lSyncCount.Text.Replace("{syncs}", String.Format("{0:n0}", completedSyncs));
                splash.lSyncCount.Left = (splash.panel1.Width - (splash.lSyncCount.Width)) / 2;
            }
            log.Debug("Showing splash screen.");
            Application.Run(splash);
            log.Debug("Disposed of splash screen.");
            splashThread.Abort();
        }

        public static void CloseMe() {
            if (splash.InvokeRequired) {
                splash.Invoke(new MethodInvoker(CloseMe));
            } else {
                if (!splash.IsDisposed) splash.Close();
            }
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

        private void Splash_Shown(object sender, EventArgs e) {
            splash.Tag = DateTime.Now;
            while (DateTime.Now < ((DateTime)splash.Tag).AddSeconds((System.Diagnostics.Debugger.IsAttached ? 1 : 8)) && !splash.IsDisposed) {
                Application.DoEvents();
                System.Threading.Thread.Sleep(100);
            }
            CloseMe();
        }
    }
}
