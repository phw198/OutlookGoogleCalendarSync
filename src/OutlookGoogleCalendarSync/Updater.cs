using log4net;
using Squirrel;
using System;
using System.ComponentModel;
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
            if (string.IsNullOrEmpty(nonGitHubReleaseUri) && Program.InDeveloperMode) return;

            bt = updateButton;
            log.Debug((isManualCheck ? "Manual" : "Automatic") + " update check requested.");
            if (isManualCheck) updateButton.Text = "Checking...";
            
            try {
                if (!string.IsNullOrEmpty(nonGitHubReleaseUri) || Program.IsInstalled) {
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
                if (OgcsMessageBox.Show("The upgrade failed.\nWould you like to get the latest version from the project website manually?", "Upgrade Failed", MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation) == DialogResult.Yes) {
                    Helper.OpenBrowser(Program.OgcsWebsite);
                }

            } catch (System.Exception ex) {
                log.Fail("Failure checking for update. " + ex.Message);
                if (isManualCheck) {
                    OgcsMessageBox.Show("Unable to check for new version.", "Check Failed", MessageBoxButtons.OK, MessageBoxIcon.Error);
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
                        new Squirrel.FileDownloader(new Extensions.OgcsWebClient()), prerelease: Settings.Instance.AlphaReleases);
                else
                    updateManager = new Squirrel.UpdateManager(nonGitHubReleaseUri, "OutlookGoogleCalendarSync", installRootDir);

                UpdateInfo updates = await updateManager.CheckForUpdate();
                if ((Settings.Instance.AlphaReleases && updates.ReleasesToApply.Any()) ||
                    updates.ReleasesToApply.Any(r => r.Version.SpecialVersion != "alpha")) {

                    if (updates.CurrentlyInstalledVersion != null)
                        log.Info("Currently installed version: " + updates.CurrentlyInstalledVersion.Version.ToString());
                    log.Info("Found " + updates.ReleasesToApply.Count() + " newer releases available.");
                    log.Info("Download directory = " + updates.PackageDirectory);

                    DialogResult dr = DialogResult.Cancel;
                    String squirrelAnalyticsLabel = "";
                    String releaseNotes = "";
                    String releaseVersion = "";
                    String releaseType = "";

                    foreach (ReleaseEntry update in updates.ReleasesToApply.OrderBy(x => x.Version).Reverse()) {
                        log.Info("New " + update.Version.SpecialVersion + " version available: " + update.Version.Version.ToString());

                        if (!this.isManualCheck && update.Version.Version.ToString() == Settings.Instance.SkipVersion && update == updates.ReleasesToApply.Last()) {
                            log.Info("The user has previously requested to skip this version.");
                            return false;
                        }

                        String localFile = updates.PackageDirectory + "\\" + update.Filename;
                        if (updateManager.CheckIfAlreadyDownloaded(update, localFile)) {
                            log.Debug("This has already been downloaded.");
                        } else {
                            try {
                                //"https://github.com/phw198/OutlookGoogleCalendarSync/releases/download/v2.8.6-alpha"
                                if (string.IsNullOrEmpty(nonGitHubReleaseUri)) {
                                    String nupkgUrl = "https://github.com/phw198/OutlookGoogleCalendarSync/releases/download/v" + update.Version + "/" + update.Filename;
                                    log.Debug("Downloading " + nupkgUrl);
                                    new Extensions.OgcsWebClient().DownloadFile(nupkgUrl, localFile); 
                                } else {
                                    String nupkgUrl = nonGitHubReleaseUri + "\\" + update.Filename;
                                    log.Debug("Downloading " + nupkgUrl);
                                    new System.Net.WebClient().DownloadFile(nupkgUrl, localFile);
                                }
                                log.Debug("Download complete.");
                            } catch (System.Exception ex) {
                                OGCSexception.Analyse("Failed downloading release file " + update.Filename + " for " + update.Version, ex);
                                ex.Data.Add("analyticsLabel", "from=" + Application.ProductVersion + ";download_file=" + update.Filename + ";" + ex.Message);
                                throw new ApplicationException("Failed upgrading OGCS.", ex);
                            }
                        }

                        if (string.IsNullOrEmpty(releaseNotes)) {
                            log.Debug("Retrieving release notes.");
                            releaseNotes = update.GetReleaseNotes(updates.PackageDirectory);
                            releaseVersion = update.Version.Version.ToString();
                            releaseType = update.Version.SpecialVersion;
                            squirrelAnalyticsLabel = "from=" + Application.ProductVersion + ";to=" + update.Version.Version.ToString();
                        }
                    }
                    
                    var t = new System.Threading.Thread(() => new Forms.UpdateInfo(releaseVersion, releaseType, releaseNotes, out dr));
                    t.SetApartmentState(System.Threading.ApartmentState.STA);
                    t.Start();
                    t.Join();

                    if (dr == DialogResult.No) {
                        log.Info("User chose not to upgrade right now.");
                        Telemetry.Send(Analytics.Category.squirrel, Analytics.Action.upgrade, squirrelAnalyticsLabel + ";later");

                    } else if (dr == DialogResult.Ignore) {
                        Telemetry.Send(Analytics.Category.squirrel, Analytics.Action.upgrade, squirrelAnalyticsLabel + ";skipped");

                    } else if (dr == DialogResult.Yes) {
                        try {
                            Telemetry.Send(Analytics.Category.squirrel, Analytics.Action.download, squirrelAnalyticsLabel + ";successful");
                            log.Info("Applying the updated release(s)...");
                            //updateManager.UpdateApp().Wait();

                            int ApplyAttempt = 1;
                            while (ApplyAttempt <= 5) {
                                try {
                                    updateManager.ApplyReleases(updates).Wait();
                                    break;
                                } catch (System.AggregateException ex) {
                                    ApplyAttempt++;
                                    if (OGCSexception.GetErrorCode(ex.InnerException) == "0x80070057") { //File does not exist
                                        //File does not exist: C:\Users\Paul\AppData\Local\OutlookGoogleCalendarSync\packages\OutlookGoogleCalendarSync-2.8.4-alpha-full.nupkg
                                        //Extract the nupkg filename
                                        String regexMatch = ".*" + updates.PackageDirectory.Replace(@"\", @"\\") + @"\\(.*?([\d\.]+-\w+).*)$";
                                        System.Text.RegularExpressions.Match match = System.Text.RegularExpressions.Regex.Match(ex.InnerException.Message, regexMatch);

                                        if (match?.Groups?.Count == 3) {
                                            log.Warn("Could not update as missing file " + match.Groups[1]);
                                            String nupkgUrl = "https://github.com/phw198/OutlookGoogleCalendarSync/releases/download/v" + match.Groups[2] + "/" + match.Groups[1];
                                            log.Debug("Downloading " + nupkgUrl);
                                            new Extensions.OgcsWebClient().DownloadFile(nupkgUrl, updates.PackageDirectory + "\\" + match.Groups[1]);
                                            log.Debug("Download complete.");
                                        }
                                    } else throw;
                                }
                            }

                            log.Info("The application has been successfully updated.");
                            OgcsMessageBox.Show("The application has been updated and will now restart.",
                                "OGCS successfully updated!", MessageBoxButtons.OK, MessageBoxIcon.Information);
                            restartUpdateExe = updateManager.RootAppDirectory + "\\Update.exe";
                            return true;

                        } catch (System.AggregateException ae) {
                            foreach (System.Exception ex in ae.InnerExceptions) {
                                OGCSexception.Analyse(ex, true);
                                ex.Data.Add("analyticsLabel", squirrelAnalyticsLabel);
                                throw new ApplicationException("Failed upgrading OGCS.", ex);
                            }
                        } catch (System.Exception ex) {
                            OGCSexception.Analyse(ex, true);
                            ex.Data.Add("analyticsLabel", squirrelAnalyticsLabel);
                            throw new ApplicationException("Failed upgrading OGCS.", ex);
                        }
                    }

                } else {
                    log.Info("Already running the latest version of OGCS.");
                    if (this.isManualCheck) { //Was a manual check, so give feedback
                        OgcsMessageBox.Show("You are already running the latest version of OGCS.", "Latest Version", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                }

            } catch (ApplicationException ex) {
                Telemetry.Send(Analytics.Category.squirrel, Analytics.Action.download, ex.Data["analyticsLabel"] + ";failed");
                throw;
            } catch (System.AggregateException ae) {
                log.Fail("Failed checking for update.");
                foreach (System.Exception ex in ae.InnerExceptions) {
                    OGCSexception.Analyse(OGCSexception.LogAsFail(ex), true);
                    throw;
                }
            } catch (System.Exception ex) {
                OGCSexception.Analyse("Failed checking for update.", OGCSexception.LogAsFail(ex), true);
                throw;
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
                Telemetry.Send(Analytics.Category.squirrel, Analytics.Action.uninstall, "clickonce");
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
            Telemetry.Send(Analytics.Category.squirrel, Analytics.Action.install, version.ToString());
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
                Telemetry.Send(Analytics.Category.squirrel, Analytics.Action.uninstall, version.ToString());
                if (OgcsMessageBox.Show("Sorry to see you go!\nCould you spare 30 seconds for some feedback?", "Uninstalling OGCS",
                    MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes) {
                    log.Debug("User opted to give feedback.");
                    Telemetry.Send(Analytics.Category.squirrel, Analytics.Action.uninstall, Application.ProductVersion + "-feedback");
                    Helper.OpenBrowser("https://docs.google.com/forms/d/e/1FAIpQLSfRWYFdgyfbFJBMQ0dz14patu195KSKxdLj8lpWvLtZn-GArw/viewform?entry.1161230174=v" + Application.ProductVersion);
                } else {
                    log.Debug("User opted not to give feedback.");
                }
                log.Info("Deleting directory " + Path.GetDirectoryName(Settings.ConfigFile));
                try {
                    log.Logger.Repository.Shutdown();
                    log4net.LogManager.Shutdown();
                    Directory.Delete(Path.GetDirectoryName(Settings.ConfigFile), true);
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
                String[] cliArgs = null;
                if (Program.StartedWithSquirrelArgs) {
                    cliArgs = Environment.GetCommandLineArgs().Skip(1).ToArray();
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
            BackgroundWorker bwUpdater = new BackgroundWorker {
                WorkerReportsProgress = false,
                WorkerSupportsCancellation = false
            };
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
            String errorDetails = "";
            try {
                html = new Extensions.OgcsWebClient().DownloadString("https://github.com/phw198/OutlookGoogleCalendarSync/blob/master/docs/latest_zip_release.md");
            } catch (System.Net.WebException ex) {
                if (OGCSexception.GetErrorCode(ex) == "0x80131509")
                    log.Warn("Failed to retrieve data (no network?): " + ex.Message);
                else
                    OGCSexception.Analyse("Failed to retrieve data", ex);
            } catch (System.Exception ex) {
                OGCSexception.Analyse("Failed to retrieve data: ", ex);
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
                Int32 releaseNum = Program.VersionToInt(releaseVersion);
                Int32 myReleaseNum = Program.VersionToInt(Application.ProductVersion);
                if (releaseNum > myReleaseNum) {
                    log.Info("New " + releaseType + " ZIP release found: " + releaseVersion);
                    DialogResult dr = OgcsMessageBox.Show("A new " + releaseType + " release is available for OGCS. Would you like to upgrade to v" + releaseVersion + "?", "New OGCS Release Available", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                    if (dr == DialogResult.Yes) {
                        Helper.OpenBrowser(releaseURL);
                    }
                } else {
                    log.Info("Already on latest ZIP release.");
                    if (isManualCheck) OgcsMessageBox.Show("You are already on the latest release", "No Update Required", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            } else {
                log.Info("Did not find ZIP release.");
                if (isManualCheck) OgcsMessageBox.Show("Failed to check for ZIP release." + (string.IsNullOrEmpty(errorDetails) ? "" : "\r\n" + errorDetails),
                    "Update Check Failed", MessageBoxButtons.OK, MessageBoxIcon.Information);
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
