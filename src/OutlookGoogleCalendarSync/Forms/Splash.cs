using log4net;
using System;
using System.Threading;
using System.Windows.Forms;

namespace OutlookGoogleCalendarSync.Forms {
    public partial class Splash : Form {
        private static readonly ILog log = LogManager.GetLogger(typeof(Splash));
        
        private static Thread splashThread;
        private static Splash splash;
        private static ToolTip ToolTips;
        private static Boolean donor;
        private static DateTime subscribed;
        private static Boolean initialised = false;
        public static Boolean BeenAndGone { get; private set; }

        public Splash() {
            InitializeComponent();
            BeenAndGone = false;
        }

        public static void ShowMe() {
            if (splashThread == null) {
                splashThread = new Thread(new ThreadStart(doShowSplash)) { IsBackground = true };
                splashThread.Start();
                while (!initialised) {
                    //Stop the program continuing until splash screen has finished accessing settings.xml
                    Thread.Sleep(50);
                }
            }
        }
        private static void doShowSplash() {
            try {
                if (splash == null)
                    splash = new Splash();

                splash.lVersion.Text = "v" + Application.ProductVersion;
                String completedSyncs = XMLManager.ImportElement("CompletedSyncs", Settings.ConfigFile) ?? "0";
                if (completedSyncs == "0")
                    splash.lSyncCount.Visible = false;
                else {
                    splash.lSyncCount.Text = splash.lSyncCount.Text.Replace("{syncs}", String.Format("{0:n0}", completedSyncs));
                    splash.lSyncCount.Left = (splash.panel1.Width - (splash.lSyncCount.Width)) / 2;
                }
                //Load settings directly from XML
                donor = (XMLManager.ImportElement("Donor", Settings.ConfigFile) ?? "false") == "true";

                String subscribedDate = XMLManager.ImportElement("Subscribed", Settings.ConfigFile);
                if (!string.IsNullOrEmpty(subscribedDate)) subscribed = DateTime.Parse(subscribedDate); 
                else subscribed = GoogleOgcs.Authenticator.SubscribedNever;
                
                Boolean hideSplash = (XMLManager.ImportElement("HideSplashScreen", Settings.ConfigFile) ?? "false") == "true";
                initialised = true;

                splash.cbHideSplash.Checked = hideSplash;
                if (subscribed == GoogleOgcs.Authenticator.SubscribedNever && !donor) {
                    ToolTips = new ToolTip {
                        AutoPopDelay = 10000,
                        InitialDelay = 500,
                        ReshowDelay = 200,
                        ShowAlways = true
                    };

                    ToolTips.SetToolTip(splash.cbHideSplash, "Donate £10 or more to enable this feature.");
                } else if (hideSplash) {
                    log.Debug("Suppressing splash screen.");
                    return;
                }
                splash.TopLevel = true;
                splash.TopMost = true;
                log.Debug("Showing splash screen.");
                Application.Run(splash);
                log.Debug("Disposed of splash screen.");
                splashThread.Abort();
            } finally {
                initialised = true;
                BeenAndGone = true;
            }
        }

        public static void CloseMe() {
            if (splash == null) return;

            if (splash.InvokeRequired) {
                splash.Invoke(new MethodInvoker(CloseMe));
            } else {
                if (!splash.IsDisposed) splash.Close();
            }
            BeenAndGone = true;
        }

        private void pbDonate_Click(object sender, EventArgs e) {
            Program.Donate("Splash");
            this.Close();
        }

        private void pbSocialTwitterFollow_Click(object sender, EventArgs e) {
            Social.Twitter_follow();
            this.Close();
        }

        private void Splash_Shown(object sender, EventArgs e) {
            splash.Tag = DateTime.Now;
            while (DateTime.Now < ((DateTime)splash.Tag).AddSeconds((Program.InDeveloperMode ? 2 : 8)) && !splash.IsDisposed) {
                splash.BringToFront();
                splash.TopLevel = true;
                splash.TopMost = true;
                Application.DoEvents();
                System.Threading.Thread.Sleep(100);
            }
            CloseMe();
        }

        private void cbHideSplash_CheckedChanged(object sender, EventArgs e) {
            if (!this.Visible) return;

            if (subscribed == GoogleOgcs.Authenticator.SubscribedNever && !donor) {
                this.cbHideSplash.CheckedChanged -= cbHideSplash_CheckedChanged;
                cbHideSplash.Checked = false;
                this.cbHideSplash.CheckedChanged += cbHideSplash_CheckedChanged;
                ToolTips.Show(ToolTips.GetToolTip(cbHideSplash), cbHideSplash, 5000);
                return;
            }
            if (cbHideSplash.Checked) {
                this.Visible = false;
                while (!Settings.InstanceInitialiased() && !Forms.Main.Instance.IsHandleCreated) {
                    log.Debug("Waiting for settings and form to initialise in order to save HideSplashScreen preference.");
                    System.Threading.Thread.Sleep(2000);
                }
                Settings.Instance.HideSplashScreen = true;
                CloseMe();
            }
        }
    }
}
