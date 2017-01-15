using System;
using System.ComponentModel;
using System.Deployment.Application;
using System.Linq;
using System.IO;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using log4net;
using log4net.Config;

namespace OutlookGoogleCalendarSync {
    /// <summary>
    /// Class with program entry point.
    /// </summary>
    internal sealed class Program {
        /// <summary>
        /// Program entry point.
        /// </summary>
        public static string UserFilePath;
        private static readonly ILog log = LogManager.GetLogger(typeof(Program));
        private const string logFile = "logger.xml";
        //log4net.Core.Level.Fine == log4net.Core.Level.Debug (30000), so manually changing its value
        public static log4net.Core.Level MyFineLevel = new log4net.Core.Level(25000, "FINE");
        public static log4net.Core.Level MyUltraFineLevel = new log4net.Core.Level(24000, "ULTRA-FINE"); //Logs email addresses

        private const String settingsFilename = "settings.xml";
        private static String settingsFile;
        public static String SettingsFile {
            get { return settingsFile; }
        }
        private static String startingTab = null;
        private static String roamingOGCS;

        [STAThread]
        private static void Main(string[] args) {
            initialiseFiles();

            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);

            Splash.ShowMe();
            
            log.Debug("Loading settings from file.");
            Settings.Load();
            isNewVersion();
            checkForUpdate();
            TimezoneDB.Instance.CheckForUpdate();

            try {
                Application.Run(new MainForm(startingTab));
            } catch (ApplicationException ex) {
                log.Fatal(ex.Message);
                MessageBox.Show(ex.Message, "Application terminated!", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            } catch (System.Runtime.InteropServices.COMException ex) {
                if (OGCSexception.GetErrorCode(ex, 0x000FFFFF) == "0x00040115") {
                    log.Error(ex.Message);
                    MessageBox.Show("OGCS is not able to run as Outlook is not properly connected to the Exchange server.\r\n" +
                        "Please try again later.", "Application cannot initialise!", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                } else OGCSexception.Analyse(ex);
                
            } catch (System.Exception ex) {
                OGCSexception.Analyse(ex);
                log.Fatal("Application unexpectedly terminated!");
                MessageBox.Show(ex.Message, "Application unexpectedly terminated!", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            try { OutlookCalendar.Instance.IOutlook.Disconnect(); } catch { }
            Splash.CloseMe();
            GC.Collect();
            GC.WaitForPendingFinalizers();
            log.Info("Application closed.");
        }

        private static void initialiseFiles() {
            string appFilePath = System.Windows.Forms.Application.StartupPath;
            string roamingAppData = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
            roamingOGCS = Path.Combine(roamingAppData, Application.ProductName);

            //Don't know where to write log file to yet. If settings.xml exists in Roaming profile, 
            //then log should go there too.
            if (File.Exists(Path.Combine(roamingOGCS, settingsFilename))) {
                UserFilePath = roamingOGCS;
                initialiseLogger(UserFilePath, true);
                log.Info("Storing user files in roaming directory: " + UserFilePath);
            } else {
                UserFilePath = appFilePath;
                initialiseLogger(UserFilePath, true);
                log.Info("Storing user files in application directory: " + appFilePath);

                if (!File.Exists(Path.Combine(appFilePath, settingsFilename))) {
                    log.Info("No settings.xml file found in " + appFilePath);
                    Settings.Instance.Save(Path.Combine(appFilePath, settingsFilename));
                    log.Info("New blank template created.");
                    startingTab = "Help";
                }
            }
            log.Info("Running from " + System.Windows.Forms.Application.ExecutablePath);

            //Now let's confirm the actual setting
            settingsFile = Path.Combine(UserFilePath, settingsFilename);
            Boolean keepPortable = (XMLManager.ImportElement("Portable", settingsFile) ?? "false").Equals("true");
            if (keepPortable) {
                if (UserFilePath != appFilePath) {
                    log.Info("File storage location is incorrect according to " + settingsFilename);
                    MakePortable(true);
                }
            } else {
                if (UserFilePath != roamingOGCS) {
                    log.Info("File storage location is incorrect according to " + settingsFilename);
                    MakePortable(false);
                }
            }

            string logLevel = XMLManager.ImportElement("LoggingLevel", settingsFile);
            Settings.configureLoggingLevel(logLevel ?? "FINE");
            purgeLogFiles(30);
        }

        private static void initialiseLogger(string logPath, Boolean bootstrap = false) {
            log4net.GlobalContext.Properties["LogPath"] = logPath + "\\";
            log4net.LogManager.GetRepository().LevelMap.Add(MyFineLevel);
            log4net.LogManager.GetRepository().LevelMap.Add(MyUltraFineLevel);
            XmlConfigurator.Configure(new System.IO.FileInfo(logFile));

            if (bootstrap) log.Info("Program started: v" + Application.ProductVersion);
        }

        private static void purgeLogFiles(Int16 retention) {
            log.Info("Purging log files older than "+ retention +" days...");
            foreach (String file in System.IO.Directory.GetFiles(UserFilePath, "*.log.????-??-??", SearchOption.TopDirectoryOnly)) {
                if (System.IO.File.GetLastWriteTime(file) < DateTime.Now.AddDays(-retention)) {
                    log.Debug("Deleted "+ file);
                    System.IO.File.Delete(file);
                }
            }
            log.Info("Purge complete.");
        }

        #region Application Behaviour
        #region Startup Registry Key
        private static String startupKeyPath = "Software\\Microsoft\\Windows\\CurrentVersion\\Run";

        public static void ManageStartupRegKey(Boolean recreate = false) {
            //Check for legacy Startup menu shortcut <=v2.1.4
            Boolean startupConfigExists = Program.CheckShortcut(Environment.SpecialFolder.Startup);
            if (startupConfigExists) 
                Program.RemoveShortcut(Environment.SpecialFolder.Startup);

            startupConfigExists = checkRegKey();
            
            if (Settings.Instance.StartOnStartup && !startupConfigExists)
                addRegKey();
            else if (!Settings.Instance.StartOnStartup && startupConfigExists)
                removeRegKey();
            else if (startupConfigExists && recreate) {
                log.Debug("Forcing update of startup registry key.");
                addRegKey();
            }
        }

        private static Boolean checkRegKey() {
            String[] regKeys = Microsoft.Win32.Registry.CurrentUser.OpenSubKey(startupKeyPath).GetValueNames();
            return regKeys.Contains(Application.ProductName);
        }

        private static void addRegKey() {
            Microsoft.Win32.RegistryKey startupKey = Microsoft.Win32.Registry.CurrentUser.OpenSubKey(startupKeyPath, true);
            String keyValue = startupKey.GetValue(Application.ProductName, "").ToString();
            if (keyValue == "" || keyValue != Application.ExecutablePath) {
                log.Debug("Startup registry key "+ (keyValue == "" ? "created" : "updated") +".");
                startupKey.SetValue(Application.ProductName, Application.ExecutablePath);
            }
            startupKey.Close();
        }

        private static void removeRegKey() {
            log.Debug("Startup registry key being removed.");
            Microsoft.Win32.RegistryKey startupKey = Microsoft.Win32.Registry.CurrentUser.OpenSubKey(startupKeyPath, true);
            startupKey.DeleteValue(Application.ProductName, false);
        }
        #endregion

        #region Legacy Start Menu Shortcut
        //public static void AddShortcut(Environment.SpecialFolder directory, String subdir = "") {
        //    log.Debug("AddShortcut: directory=" + directory.ToString() + "; subdir=" + subdir);
        //    String appPath = Application.ExecutablePath;
        //    if (subdir != "") subdir = "\\" + subdir;
        //    String shortcutDir = Environment.GetFolderPath(directory) + subdir;

        //    if (!System.IO.Directory.Exists(shortcutDir)) {
        //        log.Debug("Creating directory " + shortcutDir);
        //        System.IO.Directory.CreateDirectory(shortcutDir);
        //    }

        //    string shortcutLocation = System.IO.Path.Combine(shortcutDir, Application.ProductName + ".lnk");
        //    #region "Windows Script Host Object Model - 32bit only"
        //    /*
        //    IWshRuntimeLibrary.WshShell shell = new IWshRuntimeLibrary.WshShell();
        //    IWshRuntimeLibrary.IWshShortcut shortcut = shell.CreateShortcut(shortcutLocation) as IWshRuntimeLibrary.WshShortcut;
        //    shortcut.Description = "Synchronise Outlook and Google calendars";
        //    shortcut.IconLocation = appPath.ToLower().Replace("OutlookGoogleCalendarSync.exe", "icon.ico");
        //    shortcut.TargetPath = appPath;
        //    shortcut.WorkingDirectory = Application.StartupPath;
        //    shortcut.Save();
        //    */
        //    #endregion
        //    #region "C:\Windows\System32\shell32.dll - works but only compiles in VSE 2010?!"
        //    /*
        //    System.IO.File.WriteAllBytes(shortcutLocation, new byte[] { });
            
        //    //Initialize a ShellLinkObject for that .lnk file
        //    Shell32.Shell shl = new Shell32.ShellClass();
        //    Shell32.Folder dir = shl.NameSpace(shortcutDir);
        //    Shell32.FolderItem itm = dir.Items().Item(Application.ProductName + ".lnk");
        //    Shell32.ShellLinkObject lnk = (Shell32.ShellLinkObject)itm.GetLink;

        //    lnk.Path = appPath;
        //    lnk.Description = "Synchronise Outlook and Google calendars";
        //    lnk.WorkingDirectory = Application.StartupPath;
        //    lnk.SetIconLocation(appPath.ToLower().Replace("OutlookGoogleCalendarSync.exe", "icon.ico"), 0);
            
        //    lnk.Save(shortcutLocation);
        //    */
        //    #endregion
        //    log.Info("Created shortcut in \"" + shortcutDir + "\"");
        //}

        public static Boolean CheckShortcut(Environment.SpecialFolder directory, String subdir = "") {
            log.Debug("CheckShortcut: directory=" + directory.ToString() + "; subdir=" + subdir);
            Boolean foundShortcut = false;
            if (subdir != "") subdir = "\\" + subdir;
            String shortcutDir = Environment.GetFolderPath(directory) + subdir;

            if (!System.IO.Directory.Exists(shortcutDir)) return false;

            foreach (String file in System.IO.Directory.GetFiles(shortcutDir)) {
                if (file.EndsWith("\\OutlookGoogleCalendarSync.lnk") || //legacy name <=v2.1.0.0
                    file.EndsWith("\\" + Application.ProductName + ".lnk")) {
                    foundShortcut = true;
                    break;
                }
            }
            return foundShortcut;
        }

        public static void RemoveShortcut(Environment.SpecialFolder directory, String subdir = "") {
            try {
                log.Debug("RemoveShortcut: directory=" + directory.ToString() + "; subdir=" + subdir);
                if (subdir != "") subdir = "\\" + subdir;
                String shortcutDir = Environment.GetFolderPath(directory) + subdir;

                if (!System.IO.Directory.Exists(shortcutDir)) {
                    log.Info("Failed to delete shortcut in \"" + shortcutDir + "\" - directory does not exist.");
                    return;
                }
                foreach (String file in System.IO.Directory.GetFiles(shortcutDir)) {
                    if (file.EndsWith("\\OutlookGoogleCalendarSync.lnk") || //legacy name <=v2.1.0.0
                        file.EndsWith("\\" + Application.ProductName + ".lnk")) {
                        System.IO.File.Delete(file);
                        log.Info("Deleted shortcut in \"" + shortcutDir + "\"");
                        break;
                    }
                }
            } catch (System.Exception ex) {
                log.Warn("Problem trying to remove legacy Start Menu shortcut.");
                log.Error(ex.Message);
            }
        }
        #endregion

        public static void MakePortable(Boolean portable) {
            if (portable) {
                log.Info("Making the application portable...");
                string appFilePath = System.Windows.Forms.Application.StartupPath;
                if (appFilePath == UserFilePath) {
                    log.Info("It already is!");
                    return;
                }
                moveFiles(UserFilePath, appFilePath);

            } else {
                log.Info("Making the application non-portable...");
                if (roamingOGCS == UserFilePath) {
                    log.Info("It already is!");
                    return;
                }
                if (!Directory.Exists(roamingOGCS))
                    Directory.CreateDirectory(roamingOGCS);

                moveFiles(UserFilePath, roamingOGCS);
            }
        }

        private static void moveFiles(string srcDir, string dstDir) {
            log.Info("Moving files from " + srcDir + " to " + dstDir + ":-");
            if (!Directory.Exists(dstDir)) Directory.CreateDirectory(dstDir);

            string dstFile = Path.Combine(dstDir, settingsFilename);
            File.Delete(dstFile);
            log.Debug("  " + settingsFilename);
            File.Move(SettingsFile, dstFile);
            settingsFile = Path.Combine(dstDir, settingsFilename);

            foreach (string file in Directory.GetFiles(srcDir)) {
                if (Path.GetFileName(file).StartsWith("OGcalsync.log") || file.EndsWith(".csv")) {
                    dstFile = Path.Combine(dstDir, Path.GetFileName(file));
                    File.Delete(dstFile);
                    log.Debug("  " + Path.GetFileName(file));
                    if (file.EndsWith(".log")) {
                        log.Logger.Repository.Shutdown();
                        log4net.LogManager.Shutdown();
                        LogManager.GetRepository().ResetConfiguration();
                        File.Move(file, dstFile);
                        initialiseLogger(dstDir);
                    } else {
                        File.Move(file, dstFile);
                    }
                }
            }
            try {
                log.Debug("Deleting directory " + srcDir);
                Directory.Delete(srcDir);
            } catch (System.Exception ex) {
                log.Debug(ex.Message);
            }
            UserFilePath = dstDir;
        }
        #endregion

        #region Update Checking
        private static Boolean isManualCheck = false;
        
        public static Boolean isClickOnceInstall() {
            return ApplicationDeployment.IsNetworkDeployed;
        }
        public static void checkForUpdate(Boolean isManualCheck = false) {
            Settings.Instance.Proxy.Configure();
            if (System.Diagnostics.Debugger.IsAttached) return;

            Program.isManualCheck = isManualCheck;
            log.Debug((isManualCheck ? "Manual" : "Automatic") + " update check requested.");
            if (isManualCheck) MainForm.Instance.btCheckForUpdate.Text = "Checking...";

            if (isClickOnceInstall()) {
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
        private static void checkForUpdate_completed(object sender, CheckForUpdateCompletedEventArgs e) {
            if (e.Error != null) {
                log.Error("Could not retrieve new version of the application.");
                log.Error(e.Error.Message);
                if (Program.isManualCheck)
                    MessageBox.Show("Could not retrieve new version of the application.\n" + e.Error.Message, "Update Check Failed", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            } else if (e.Cancelled == true) {
                log.Info("The update was cancelled");
                if (Program.isManualCheck)
                    MessageBox.Show("The update was cancelled.", "Update Check Cancelled", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }

            if (e.UpdateAvailable) {
                log.Info("An update is available: v" + e.AvailableVersion);

                if (!e.IsUpdateRequired) {
                    log.Info("This is an optional update.");
                    DialogResult dr = MessageBox.Show("An update for OGCS is available. Would you like to update the application now?", "OGCS Update Available", MessageBoxButtons.YesNo);
                    if (dr == DialogResult.Yes) {
                        beginUpdate();
                    }
                } else {
                    log.Info("This is a mandatory update.");
                    MessageBox.Show("A mandatory update for OGCS is required. The update will be installed now and the application restarted.", "OCGS Update Required", MessageBoxButtons.OK);
                    beginUpdate();
                }
            } else {
                log.Info("Already running the latest version.");
                if (Program.isManualCheck) { //Was a manual check, so give feedback
                    MessageBox.Show("You are already running the latest version.", "Latest Version", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
        }

        private static void beginUpdate() {
            log.Info("Beginning application update...");
            ApplicationDeployment ad = ApplicationDeployment.CurrentDeployment;
            ad.UpdateCompleted += new AsyncCompletedEventHandler(update_completed);
            ad.UpdateAsync();
        }
        private static void update_completed(object sender, AsyncCompletedEventArgs e) {
            if (isManualCheck) MainForm.Instance.btCheckForUpdate.Text = "Check For Update";
            if (e.Cancelled) {
                log.Info("The update to the latest version was cancelled.");
                MessageBox.Show("The update to the latest version was cancelled.", "Installation Cancelled", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            } else if (e.Error != null) {
                log.Error("Could not install the latest version.\n" + e.Error.Message);
                MessageBox.Show("Could not install the latest version.\n" + e.Error.Message, "Installation Failure", MessageBoxButtons.OK, MessageBoxIcon.Error);
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
        private static void checkForZip(object sender, DoWorkEventArgs e) {
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
        
        private static void checkForZip_completed(object sender, RunWorkerCompletedEventArgs e) {
            if (isManualCheck)
                MainForm.Instance.btCheckForUpdate.Text = "Check For Update";
        }
        
        private static MatchCollection getRelease(string source, string pattern) {
            Regex rgx = new Regex(pattern, RegexOptions.IgnoreCase);
            return rgx.Matches(source);
        }
        #endregion

        private static void isNewVersion() {
            string settingsVersion = string.IsNullOrEmpty(Settings.Instance.Version) ? "Unknown" : Settings.Instance.Version;
            if (settingsVersion != Application.ProductVersion) {
                log.Info("New version detected - upgraded from " + settingsVersion + " to " + Application.ProductVersion);
                Program.ManageStartupRegKey(recreate: true);
                Settings.Instance.Version = Application.ProductVersion;
                System.Diagnostics.Process.Start("https://outlookgooglecalendarsync.codeplex.com/wikipage?title=Release Notes");
            }
        }
        #endregion
    }
}
