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
        private String restartUpdateExe = "";
        private static String nonGitHubReleaseUri = null; //When testing, eg: @"\\127.0.0.1\Squirrel";

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

            try {
                if (Program.IsInstalled) {
                    try {
                       if (await githubCheck()) {
                           log.Info("Restarting OGCS.");
                           try {
                               System.Diagnostics.Process.Start(restartUpdateExe, "--processStartAndWait OutlookGoogleCalendarSync.exe");
                           } catch (System.Exception ex) {
                               OGCSexception.Analyse(ex, true);
                           }
                           try {
                               Forms.Main.Instance.NotificationTray.ExitItem_Click(null, null);
                           } catch (System.Exception ex) {
                               log.Error("Failed to exit via the notification tray icon. " + ex.Message);
                               log.Debug("NotificationTray is " + (Forms.Main.Instance.NotificationTray == null ? "null" : "not null"));
                               Forms.Main.Instance.Close();
                           }
                       }
                    } finally {
                        if (isManualCheck) updateButton.Text = "Check For Update";
                    }
                } else {
                    zipChecker();
                }

            } catch (ApplicationException ex) {
                log.Error(ex.Message + " " + ex.InnerException.Message);
                if (MessageBox.Show("The upgrade failed.\nWould you like to get the latest version from the project website manually?", "Upgrade Failed", MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation) == DialogResult.Yes) {
                    System.Diagnostics.Process.Start("https://phw198.github.io/OutlookGoogleCalendarSync/");
                }

            } catch (System.Exception ex) {
                log.Error("Failure checking for update. " + ex.Message);
                if (isManualCheck) {
                    MessageBox.Show("Unable to check for new version.", "Check Failed", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }

        public static Boolean IsSquirrelInstall() {
            Boolean isSquirrelInstall = false;
            try {
                using (var updateManager = new Squirrel.UpdateManager(null)) {
                    //This just checks if there is an Update.exe file in the parent directory of the OGCS executable
                    isSquirrelInstall = updateManager.IsInstalledApp;
                }
            } catch (System.Exception ex) {
                if (OGCSexception.GetErrorCode(ex) == "0x80131500") //Update.exe not found
                    log.Debug(ex.Message);
                else {
                    log.Error("Failed to determine if app is a Squirrel install. Assuming not.");
                    OGCSexception.Analyse(ex);
                }
            }
            
            log.Info("This " + (isSquirrelInstall ? "is" : "is not") + " a Squirrel install.");
            return isSquirrelInstall;
        }

        /// <returns>True if the user has upgraded</returns>
        private async Task<Boolean> githubCheck() {
            log.Debug("Checking for Squirrel update...");
            UpdateManager updateManager = null;
            isBusy = true;
            try {
                String installRootDir = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);
                if (string.IsNullOrEmpty(nonGitHubReleaseUri))
                    updateManager = await Squirrel.UpdateManager.GitHubUpdateManager("https://github.com/phw198/OutlookGoogleCalendarSync", "OutlookGoogleCalendarSync", installRootDir,
                        prerelease: Settings.Instance.AlphaReleases);
                else
                    updateManager = new Squirrel.UpdateManager(nonGitHubReleaseUri, "OutlookGoogleCalendarSync", installRootDir);

                UpdateInfo updates = await updateManager.CheckForUpdate();
                if (updates.ReleasesToApply.Any()) {
                    if (updates.CurrentlyInstalledVersion != null)
                        log.Info("Currently installed version: " + updates.CurrentlyInstalledVersion.Version.ToString());
                    log.Info("Found " + updates.ReleasesToApply.Count() + " newer releases available.");

                    foreach (ReleaseEntry update in updates.ReleasesToApply.OrderBy(x => x.Version).Reverse()) {
                        log.Info("Found a new " + update.Version.SpecialVersion + " version: " + update.Version.Version.ToString());

                        if (!this.isManualCheck && update.Version.Version.ToString() == Settings.Instance.SkipVersion) {
                            log.Info("The user has previously requested to skip this version.");
                            break;
                        }

                        String releaseNotes = "";                        
                        if (nonGitHubReleaseUri != null) {
                            releaseNotes = update.GetReleaseNotes(nonGitHubReleaseUri);
                        } else {
                            //Somewhat annoyingly we have to download the release in order to get the release notes, as they are embedded in the .nupkg upgrade file(s)
                            try {
                                updateManager.DownloadReleases(new[] { update }).Wait(30 * 1000);
                                System.Collections.Generic.Dictionary<ReleaseEntry, String> allReleaseNotes = updates.FetchReleaseNotes();
                                releaseNotes = allReleaseNotes[update];
                            } catch (System.Exception ex) {
                                log.Error("Failed pre-fetching release notes. " + ex.Message);
                                releaseNotes = null;
                            }
                        }

                        DialogResult dr = DialogResult.Cancel;
                        if (!string.IsNullOrEmpty(releaseNotes)) log.Debug("Release notes retrieved.");
                        var t = new System.Threading.Thread(() => new Forms.UpdateInfo(update.Version.Version.ToString(), update.Version.SpecialVersion, releaseNotes, out dr));
                        t.SetApartmentState(System.Threading.ApartmentState.STA);
                        t.Start();
                        t.Join();

                        if (dr == DialogResult.OK || dr == DialogResult.Yes) {
                            log.Debug("Download started...");
                            if (!updateManager.DownloadReleases(new[] { update }).Wait(60 * 1000)) {
                                log.Warn("The download failed to completed within 60 seconds.");
                                if (MessageBox.Show("The update failed to download.", "Download timed out", MessageBoxButtons.RetryCancel, MessageBoxIcon.Exclamation) == DialogResult.Retry) {
                                    if (!updateManager.DownloadReleases(new[] { update }).Wait(60 * 1000)) {
                                        if (MessageBox.Show("The update failed to download again.\nTo download from the project website, click Yes.", "Download timed out", MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation) == DialogResult.Yes) {
                                            System.Diagnostics.Process.Start("https://phw198.github.io/OutlookGoogleCalendarSync/");
                                        }
                                        break;
                                    }
                                } else {
                                    if (MessageBox.Show("Would you like to download directly from the project website?", "Go to OGCS website", MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation) == DialogResult.Yes) {
                                        System.Diagnostics.Process.Start("https://phw198.github.io/OutlookGoogleCalendarSync/");
                                    }
                                    break;
                                }
                            }

                            try {
                                log.Debug("Download complete.");
                                log.Info("Applying the updated release...");
                                updateManager.ApplyReleases(updates).Wait();

                                log.Info("The application has been successfully updated.");
                                MessageBox.Show("The application has been updated and will now restart.",
                                    "OGCS successfully updated!", MessageBoxButtons.OK, MessageBoxIcon.Information);
                                restartUpdateExe = updateManager.RootAppDirectory + "\\Update.exe";
                                return true;

                            } catch (System.AggregateException ae) {
                                foreach (System.Exception ex in ae.InnerExceptions) {
                                    OGCSexception.Analyse(ex, true);
                                    throw new ApplicationException("Failed upgrading OGCS.", ex);
                                }
                            } catch (System.Exception ex) {
                                OGCSexception.Analyse(ex, true);
                                throw new ApplicationException("Failed upgrading OGCS.", ex);
                            }

                        } else {
                            log.Info("User chose not to upgrade.");
                        }
                        break;
                    }
                } else {
                    log.Info("Already running the latest version of OGCS.");
                    if (this.isManualCheck) { //Was a manual check, so give feedback
                        MessageBox.Show("You are already running the latest version of OGCS.", "Latest Version", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                }

            } catch (ApplicationException ex) {
                throw ex;
            } catch (System.AggregateException ae) {
                log.Error("Failed checking for update.");
                foreach (System.Exception ex in ae.InnerExceptions) {
                    OGCSexception.Analyse(ex, true);
                    throw ex;
                }
            } catch (System.Exception ex) {
                log.Error("Failed checking for update.");
                OGCSexception.Analyse(ex, true);
                throw ex;
            } finally {
                isBusy = false;
                updateManager.Dispose();
            }
            return false;
        }

        #region Squirrel Bits
        public static void MakeSquirrelAware() {
            try {
                log.Debug("Setting up Squirrel handlers.");
                Squirrel.SquirrelAwareApp.HandleEvents(
                     onFirstRun: onFirstRun,
                     onInitialInstall: v => onInitialInstall(v),
                     onAppUpdate: v => onAppUpdate(v),
                     onAppUninstall: v => onAppUninstall(v),
                     arguments: fixCliArgs()
                );
            } catch (System.Exception ex) {
                log.Error("SquirrelAwareApp.HandleEvents failed.");
                OGCSexception.Analyse(ex, true);
            }
        }
        private static void onFirstRun() {
            try {
                log.Debug("Removing ClickOnce install...");
                var migrator = new ClickOnceToSquirrelMigrator.InSquirrelAppMigrator(Application.ProductName);
                migrator.Execute().Wait();
                log.Info("ClickOnce install has been removed.");
            } catch (System.AggregateException ae) {
                foreach (System.Exception ex in ae.InnerExceptions) {
                    clickOnceUninstallError(ex);
                }
            } catch (System.Exception ex) {
                clickOnceUninstallError(ex);
            }
        }
        private static void onInitialInstall(Version version) {
            try {
                using (var mgr = new Squirrel.UpdateManager(null, "OutlookGoogleCalendarSync")) {
                    log.Info("Creating shortcuts.");
                    mgr.CreateShortcutsForExecutable(Path.GetFileName(System.Windows.Forms.Application.ExecutablePath),
                        Squirrel.ShortcutLocation.Desktop | Squirrel.ShortcutLocation.StartMenu, false);
                    log.Debug("Creating uninstaller registry keys.");
                    mgr.CreateUninstallerRegistryEntry().Wait();
                }
            } catch (System.Exception ex) {
                log.Error("Problem encountered on initiall install.");
                OGCSexception.Analyse(ex, true);
            }
            onFirstRun();
        }
        private static void onAppUpdate(Version version) {
            try {
                using (var mgr = new Squirrel.UpdateManager(null, "OutlookGoogleCalendarSync")) {
                    log.Info("Recreating shortcuts.");
                    mgr.CreateShortcutsForExecutable(Path.GetFileName(System.Windows.Forms.Application.ExecutablePath),
                        Squirrel.ShortcutLocation.Desktop | Squirrel.ShortcutLocation.StartMenu, false);
                }
            } catch (System.Exception ex) {
                log.Error("Problem encountered on app update.");
                OGCSexception.Analyse(ex, true);
            }
        }
        private static void onAppUninstall(Version version) {
            try {
                using (var mgr = new Squirrel.UpdateManager(null, "OutlookGoogleCalendarSync")) {
                    log.Info("Removing shortcuts.");
                    mgr.RemoveShortcutsForExecutable(Path.GetFileName(System.Windows.Forms.Application.ExecutablePath),
                        Squirrel.ShortcutLocation.Desktop | Squirrel.ShortcutLocation.StartMenu);
                    String startMenuFolder = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Programs), "Paul Woolcock");
                    Directory.Delete(startMenuFolder);
                    log.Debug("Removing registry uninstall keys.");
                    mgr.RemoveUninstallerRegistryEntry();
                }
                if (MessageBox.Show("Sorry to see you go!\nCould you spare 30 seconds for some feedback?", "Uninstalling OGCS",
                    MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes) {
                    log.Debug("User opted to give feedback.");
                    System.Diagnostics.Process.Start("https://docs.google.com/forms/d/e/1FAIpQLSfRWYFdgyfbFJBMQ0dz14patu195KSKxdLj8lpWvLtZn-GArw/viewform");
                } else {
                    log.Debug("User opted not to give feedback.");
                }
                log.Info("Deleting directory " + Path.GetDirectoryName(Program.SettingsFile));
                try {
                    log.Logger.Repository.Shutdown();
                    log4net.LogManager.Shutdown();
                    Directory.Delete(Path.GetDirectoryName(Program.SettingsFile), true);
                } catch (System.Exception ex) {
                    try { log.Error(ex.Message); } catch { }
                }
            } catch (System.Exception ex) {
                log.Error("Problem encountered on app uninstall.");
                OGCSexception.Analyse(ex, true);
            }
        }

        /// <summary>
        /// Prepares CLI arguments for use by SquirrelAwareApp.HandleEvents().
        /// </summary>
        /// <returns>SemVer string with any trailing "-prerelease" detail removed.</returns>
        private static String[] fixCliArgs() {
            //Seems to be a bug with SquirrelAwareApp: 2.5.0-beta is a valid SemanticVersion (semver.org), 
            //but HandleEvents() fails if eg "-beta" is present.
            //"C:\Users\username\AppData\Local\OutlookGoogleCalendarSync\app-2.5.0-beta\OutlookGoogleCalendarSync.exe" --squirrel-uninstall 2.5.0-beta
            try {
                String[] cliArgs = Environment.GetCommandLineArgs().Skip(1).ToArray();
                if (cliArgs.Length == 2 && cliArgs[0].ToLower().StartsWith("--squirrel")) {
                    log.Debug("CLI arguments: " + string.Join(" ", cliArgs));
                    cliArgs[1] = cliArgs[1].Split('-')[0];
                }
                return cliArgs;
            } catch (System.Exception ex) {
                log.Error("Failed processing CLI arguments. " + ex.Message);
                return null;
            }
        }
        private static void clickOnceUninstallError(System.Exception ex) {
            if (OGCSexception.GetErrorCode(ex) == "0x80131509") {
                log.Debug("No ClickOnce install found.");
            } else {
                log.Error("Failed removing ClickOnce install.");
                OGCSexception.Analyse(ex, true);
            }
        }
        #endregion

        #region ZIP
        private void zipChecker() {
            BackgroundWorker bwUpdater = new BackgroundWorker();
            bwUpdater.WorkerReportsProgress = false;
            bwUpdater.WorkerSupportsCancellation = false;
            bwUpdater.DoWork += new DoWorkEventHandler(checkForZip);
            bwUpdater.RunWorkerCompleted += new RunWorkerCompletedEventHandler(checkForZip_completed);
            bwUpdater.RunWorkerAsync();
        }
        
        private void checkForZip(object sender, DoWorkEventArgs e) {
            string releaseURL = null;
            string releaseVersion = null;
            string releaseType = null;

            log.Debug("Checking for ZIP update...");
            string html = "";
            try {
                html = new System.Net.WebClient().DownloadString("https://github.com/phw198/OutlookGoogleCalendarSync/blob/master/docs/latest_zip_release.md");
            } catch (System.Exception ex) {
                log.Error("Failed to retrieve data: " + ex.Message);
            }

            if (!string.IsNullOrEmpty(html)) {
                log.Debug("Finding Beta release...");
                MatchCollection release = getRelease(html, @"<strong>Beta</strong>: <a href=""(.*?)"">v([\d\.]+)</a>");
                if (release.Count > 0) {
                    releaseType = "Beta";
                    releaseURL = release[0].Result("$1");
                    releaseVersion = release[0].Result("$2");
                }
                if (Settings.Instance.AlphaReleases) {
                    log.Debug("Finding Alpha release...");
                    release = getRelease(html, @"<strong>Alpha</strong>: <a href=""(.*?)"">v([\d\.]+)</a>");
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
                Forms.Main.Instance.btCheckForUpdate.Text = "Check For Update";
        }

        private static MatchCollection getRelease(string source, string pattern) {
            Regex rgx = new Regex(pattern, RegexOptions.IgnoreCase);
            return rgx.Matches(source);
        }
        #endregion
    }
}
