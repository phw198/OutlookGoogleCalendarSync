using log4net;
using Squirrel;
using System;
using System.ComponentModel;
using System.Deployment.Application;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace OutlookGoogleCalendarSync {
    class Updater {
        private static readonly ILog log = LogManager.GetLogger(typeof(Updater));

        private Button bt;
        private Boolean isManualCheck {
            get { return bt != null; }
        }
        private Boolean isBusy = false;
        public Boolean IsBusy { 
            get { return isBusy; } 
        }
        private Boolean restartRequested = false;
        private String restartUpdateExe = "";
        private static String nonGitHubReleaseUri = null; //When testing, eg: "\\\\127.0.0.1\\Squirrel";

        public Updater() { }

        /// <summary>
        /// Check if there is a new release of the application.
        /// </summary>
        /// <param name="updateButton">The button that triggered this, if manually called.</param>
        public async void CheckForUpdate(Button updateButton = null) {
            if (System.Diagnostics.Debugger.IsAttached) return;

            bt = updateButton;
            log.Debug((isManualCheck ? "Manual" : "Automatic") + " update check requested.");
            if (isManualCheck) updateButton.Text = "Checking...";

            Settings.Instance.Proxy.Configure();

            if (await githubCheck()) {
                if (isManualCheck) updateButton.Text = "Check For Update";
                if (restartRequested) {
                    log.Debug("Restarting");
                    try {
                        //UpdateManager.RestartApp(restartExe); //Removes ClickOnce, but doesn't restart properly
                        System.Diagnostics.Process.Start(restartUpdateExe, "--processStartAndWait OutlookGoogleCalendarSync.exe");
                    } catch (System.Exception ex) {
                        OGCSexception.Analyse(ex, true);
                    }
                    MainForm.Instance.NotificationTray.ExitItem_Click(null, null);
                }
            } else {
                legacyCodeplexCheck();
            }
            if (isManualCheck) updateButton.Text = "Check for Update";
        }

        private async Task<Boolean> githubCheck() {
            Boolean isSquirrelInstall = false;
            isBusy = true;
            try {
                //*** Only required for final ClickOnce offboarding release - remove afterwards!
                if (!System.IO.File.Exists("..\\Update.exe") && System.IO.File.Exists("Update.exe")) {
                    log.Debug("Copying Update.exe to parent directory...");
                    System.IO.File.Copy("Update.exe", "..\\Update.exe");
                }
                using (var updateManager = new Squirrel.UpdateManager(null)) {
                    //This just checks if there is an Update.exe file in the parent directory of the OGCS executable
                    isSquirrelInstall = updateManager.IsInstalledApp;
                }
            } catch (System.Exception ex) {
                log.Error("Failed to determine if app is a Squirrel install. Assuming not.");
                OGCSexception.Analyse(ex);
                if (this.isManualCheck)
                    MessageBox.Show("Could not retrieve new version of the application.\n" + ex.Message, "Update Check Failed", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

            log.Info("This " + (isSquirrelInstall ? "is" : "is not") + " a Squirrel " + (Program.IsClickOnceInstall ? "aware ClickOnce " : "") + "install.");
            if (isSquirrelInstall) {
                log.Debug("Checking for Squirrel update...");
                UpdateManager updateManager = null;
                try {
                    String installRootDir = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);
                    if (string.IsNullOrEmpty(nonGitHubReleaseUri))
                        updateManager = await Squirrel.UpdateManager.GitHubUpdateManager("https://github.com/phw198/OutlookGoogleCalendarSync");
                    else
                        updateManager = new Squirrel.UpdateManager(nonGitHubReleaseUri, "OutlookGoogleCalendarSync", installRootDir);
                    
                    UpdateInfo updates = await updateManager.CheckForUpdate();
                    if (updates.ReleasesToApply.Any()) {
                        log.Info("Found " + updates.ReleasesToApply.Count() + " new releases available.");

                        foreach (ReleaseEntry update in updates.ReleasesToApply.OrderBy(x => x.Version).Reverse()) {
                            log.Info("Found a new " + update.Version.SpecialVersion + " version: " + update.Version.Version.ToString());
                            if (update.Version.SpecialVersion == "alpha" && !Settings.Instance.AlphaReleases) {
                                log.Debug("User doesn't want alpha releases.");
                                continue;
                            }
                            DialogResult dr = MessageBox.Show("A " + update.Version.SpecialVersion + " update for OGCS is available.\nWould you like to update the application to v" +
                                update.Version.Version.ToString() + " now?", "OGCS Update Available", MessageBoxButtons.YesNo, MessageBoxIcon.Information);
                            if (dr == DialogResult.Yes) {
                                /* 
                                new System.Net.WebClient().DownloadFile("https://github.com/phw198/OutlookGoogleCalendarSync/releases/download/v2.6-beta/OutlookGoogleCalendarSync-2.5.0-beta-full.nupkg", "OutlookGoogleCalendarSync-2.5.0-beta-full.nupkg");
                                String notes = update.GetReleaseNotes("");
                                //if (!string.IsNullOrEmpty(notes)) log.Debug(notes);
                                */
                                log.Info("Beginning the migration to Squirrel/github release...");
                                var migrator = new ClickOnceToSquirrelMigrator.InClickOnceAppMigrator(updateManager, Application.ProductName);
                                log.Info("RootAppDirectory: " + updateManager.RootAppDirectory);
                                await migrator.Execute();

                                log.Debug("Moving the Update.exe file");
                                System.IO.File.Move("..\\Update.exe", updateManager.RootAppDirectory + "\\Update.exe");

                                log.Info("The application has been successfully updated.");
                                MessageBox.Show("The application has been updated and will now restart.",
                                    "OGCS successfully updated!", MessageBoxButtons.OK, MessageBoxIcon.Information);
                                log.Info("Restarting OGCS.");
                                restartRequested = true;
                                restartUpdateExe = updateManager.RootAppDirectory + "\\Update.exe";
                            } else {
                                log.Info("User chose not to upgrade.");
                            }
                            break;
                        }
                        return true;
                    } else {
                        log.Info("Already running the latest version of OGCS.");
                        if (this.isManualCheck) { //Was a manual check, so give feedback
                            MessageBox.Show("You are already running the latest version of OGCS.", "Latest Version", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        }
                        return false;
                    }

                } catch (System.Exception ex) {
                    OGCSexception.Analyse(ex, true);
                    if (ex.InnerException != null) log.Error(ex.InnerException.Message);
                } finally {
                    isBusy = false;
                    updateManager.Dispose();
                }
            } 
            isBusy = false;
            return false;
        }

        /// <summary>
        /// Once we have made our first release on GitHub, this can be removed.
        /// Until then, let's keep the code just in case another CodePlex release is needed.
        /// </summary>
        private void legacyCodeplexCheck() {
            if (Program.IsClickOnceInstall) {
                ApplicationDeployment ad = ApplicationDeployment.CurrentDeployment;
                if (isManualCheck || ad.TimeOfLastUpdateCheck < DateTime.Now.AddDays(-1)) {
                    log.Debug("Checking for ClickOnce update...");
                    ad.CheckForUpdateCompleted -= new CheckForUpdateCompletedEventHandler(checkForUpdate_completed);
                    ad.CheckForUpdateCompleted += new CheckForUpdateCompletedEventHandler(checkForUpdate_completed);
                    ad.CheckForUpdateAsync();
                }
            } else {
                BackgroundWorker bwUpdater = new BackgroundWorker();
                bwUpdater.WorkerReportsProgress = false;
                bwUpdater.WorkerSupportsCancellation = false;
                bwUpdater.DoWork += new DoWorkEventHandler(checkForZip);
                bwUpdater.RunWorkerCompleted += new RunWorkerCompletedEventHandler(checkForZip_completed);
                bwUpdater.RunWorkerAsync();
            }
        }
        
        #region ClickOnce
        private void checkForUpdate_completed(object sender, CheckForUpdateCompletedEventArgs e) {
            if (e.Error != null) {
                log.Error("Could not retrieve new version of the application.");
                log.Error(e.Error.Message);
                if (this.isManualCheck)
                    MessageBox.Show("Could not retrieve new version of the application.\n" + e.Error.Message, "Update Check Failed", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            } else if (e.Cancelled == true) {
                log.Info("The update was cancelled");
                if (this.isManualCheck)
                    MessageBox.Show("The update was cancelled.", "Update Check Cancelled", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }

            if (e.UpdateAvailable) {
                log.Info("An update is available: v" + e.AvailableVersion);

                if (!e.IsUpdateRequired) {
                    log.Info("This is an optional update.");
                    DialogResult dr = MessageBox.Show("An update for OGCS is available. Would you like to update the application now?", "OGCS Update Available", MessageBoxButtons.YesNo);
                    if (dr == DialogResult.Yes) {
                        beginClickOnceUpdate();
                    }
                } else {
                    log.Info("This is a mandatory update.");
                    MessageBox.Show("A mandatory update for OGCS is required. The update will be installed now and the application restarted.", "OCGS Update Required", MessageBoxButtons.OK);
                    beginClickOnceUpdate();
                }
            } else {
                log.Info("Already running the latest ClickOnce version.");
                if (this.isManualCheck) { //Was a manual check, so give feedback
                    MessageBox.Show("You are already running the latest ClickOnce version.", "Latest Version", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
        }

        private void beginClickOnceUpdate() {
            log.Info("Beginning application update...");
            ApplicationDeployment ad = ApplicationDeployment.CurrentDeployment;
            ad.UpdateCompleted += new AsyncCompletedEventHandler(update_completed);
            ad.UpdateAsync();
        }
        
        private void update_completed(object sender, AsyncCompletedEventArgs e) {
            if (this.isManualCheck) MainForm.Instance.btCheckForUpdate.Text = "Check For Update";
            if (e.Cancelled) {
                log.Info("The update to the latest version was cancelled.");
                MessageBox.Show("The update to the latest version was cancelled.", "Installation Cancelled", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            } else if (e.Error != null) {
                log.Error("Could not install the latest version.\n" + e.Error.Message);
                MessageBox.Show("Could not install the latest version.\n" + e.Error.Message + "\n\nPlease download the update directly from CodePlex.", 
                    "Installation Failure", MessageBoxButtons.OK, MessageBoxIcon.Error);
                System.Diagnostics.Process.Start("https://outlookgooglecalendarsync.codeplex.com/downloads/get/clickOnce/OutlookGoogleCalendarSync.application");
                return;
            }

            DialogResult dr = MessageBox.Show("The application has been updated. Restart? (If you do not restart now, the new version will not take effect until after you quit and launch the application again.)", "Restart Application?", MessageBoxButtons.YesNo);
            if (dr == DialogResult.Yes) {
                log.Info("Restarting application following update.");
                Application.Restart();
            }
        }
        #endregion

        #region ZIP
        private void checkForZip(object sender, DoWorkEventArgs e) {
            string releaseURL = null;
            string releaseVersion = null;
            string releaseType = null;

            log.Debug("Checking for ZIP update...");
            string html = "";
            try {
                html = new System.Net.WebClient().DownloadString("https://outlookgooglecalendarsync.codeplex.com/wikipage?title=Latest%20Releases");
            } catch (Exception ex) {
                log.Error("Failed to retrieve data: " + ex.Message);
            }

            if (!string.IsNullOrEmpty(html)) {
                log.Debug("Finding Beta release...");
                MatchCollection release = getRelease(html, @"<b>Beta</b>: <a href=""(.*?)"">\r\nv([\d\.]+)");
                if (release.Count > 0) {
                    releaseType = "Beta";
                    releaseURL = release[0].Result("$1");
                    releaseVersion = release[0].Result("$2");
                }
                if (Settings.Instance.AlphaReleases) {
                    log.Debug("Finding Alpha release...");
                    release = getRelease(html, @"<b>Alpha</b>: <a href=""(.*?)"">\r\nv([\d\.]+)");
                    if (release.Count > 0) {
                        releaseType = "Alpha";
                        releaseURL = release[0].Result("$1");
                        releaseVersion = release[0].Result("$2");
                    }
                }
            }

            if (releaseVersion != null) {
                String paddedVersion = "";
                foreach (String versionBit in releaseVersion.Split('.')) {
                    paddedVersion += versionBit.PadLeft(2, '0');
                }
                Int32 releaseNum = Convert.ToInt32(paddedVersion);
                paddedVersion = "";
                foreach (String versionBit in Application.ProductVersion.Split('.')) {
                    paddedVersion += versionBit.PadLeft(2, '0');
                }
                Int32 myReleaseNum = Convert.ToInt32(paddedVersion);
                if (releaseNum > myReleaseNum) {
                    log.Info("New " + releaseType + " ZIP release found: " + releaseVersion);
                    DialogResult dr = MessageBox.Show("A new " + releaseType + " release is available for OGCS. Would you like to upgrade to v" + releaseVersion + "?", "New OGCS Release Available", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                    if (dr == DialogResult.Yes) {
                        System.Diagnostics.Process.Start(releaseURL);
                    }
                } else {
                    log.Info("Already on latest ZIP release.");
                    if (isManualCheck) MessageBox.Show("You are already on the latest release", "No Update Required", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            } else {
                log.Info("Did not find ZIP release.");
                if (isManualCheck) MessageBox.Show("Failed to check for ZIP release", "Update Check Failed", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        private void checkForZip_completed(object sender, RunWorkerCompletedEventArgs e) {
            if (isManualCheck)
                MainForm.Instance.btCheckForUpdate.Text = "Check For Update";
        }

        private static MatchCollection getRelease(string source, string pattern) {
            Regex rgx = new Regex(pattern, RegexOptions.IgnoreCase);
            return rgx.Matches(source);
        }
        #endregion
    }
}
