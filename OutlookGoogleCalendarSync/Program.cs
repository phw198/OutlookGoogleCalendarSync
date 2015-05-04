using System;
using System.IO;
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

            #region SplashScreen
            Form splash = new Splash();
            splash.Show();
            DateTime splashed = DateTime.Now;
            while (DateTime.Now < splashed.AddSeconds((System.Diagnostics.Debugger.IsAttached ? 1 :8)) && !splash.IsDisposed) {
                Application.DoEvents();
                System.Threading.Thread.Sleep(100);
            }
            if (!splash.IsDisposed) splash.Close();
            #endregion 

            try {
                Application.Run(new MainForm(startingTab));
            } catch (Exception ex) {
                log.Fatal("Application unexpectedly terminated!");
                log.Fatal(ex.Message);
                MessageBox.Show(ex.Message, "Application unexpectedly terminated!", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
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
        }

        private static void initialiseLogger(string logPath, Boolean bootstrap = false) {
            log4net.GlobalContext.Properties["LogPath"] = logPath + "\\";
            log4net.LogManager.GetRepository().LevelMap.Add(MyFineLevel);
            XmlConfigurator.Configure(new System.IO.FileInfo(logFile));

            if (bootstrap) log.Info("Program started: v" + Application.ProductVersion);
        }

        #region Application Behaviour
        public static void AddShortcut(Environment.SpecialFolder directory, String subdir = "") {
            log.Debug("AddShortcut: directory=" + directory.ToString() + "; subdir=" + subdir);
            String appPath = Application.ExecutablePath;
            if (subdir != "") subdir = "\\" + subdir;
            String shortcutDir = Environment.GetFolderPath(directory) + subdir;

            if (!System.IO.Directory.Exists(shortcutDir)) {
                log.Debug("Creating directory " + shortcutDir);
                System.IO.Directory.CreateDirectory(shortcutDir);
            }

            string shortcutLocation = System.IO.Path.Combine(shortcutDir, "OutlookGoogleCalendarSync.lnk");
            IWshRuntimeLibrary.WshShell shell = new IWshRuntimeLibrary.WshShell();
            IWshRuntimeLibrary.IWshShortcut shortcut = shell.CreateShortcut(shortcutLocation) as IWshRuntimeLibrary.WshShortcut;

            shortcut.Description = "Synchronise Outlook and Google calendars";
            shortcut.IconLocation = appPath.ToLower().Replace("OutlookGoogleCalendarSync.exe", "icon.ico");
            shortcut.TargetPath = appPath;
            shortcut.WorkingDirectory = Application.StartupPath;
            shortcut.Save();
            log.Info("Created shortcut in \"" + shortcutDir + "\"");
        }

        public static Boolean CheckShortcut(Environment.SpecialFolder directory, String subdir = "") {
            log.Debug("CheckShortcut: directory=" + directory.ToString() + "; subdir=" + subdir);
            Boolean foundShortcut = false;
            if (subdir != "") subdir = "\\" + subdir;
            String shortcutDir = Environment.GetFolderPath(directory) + subdir;

            if (!System.IO.Directory.Exists(shortcutDir)) return false;

            foreach (String file in System.IO.Directory.GetFiles(shortcutDir)) {
                if (file.EndsWith("\\OutlookGoogleCalendarSync.lnk")) {
                    foundShortcut = true;
                    break;
                }
            }
            return foundShortcut;
        }

        public static void RemoveShortcut(Environment.SpecialFolder directory, String subdir = "") {
            log.Debug("RemoveShortcut: directory=" + directory.ToString() + "; subdir=" + subdir);
            if (subdir != "") subdir = "\\" + subdir;
            String shortcutDir = Environment.GetFolderPath(directory) + subdir;

            if (!System.IO.Directory.Exists(shortcutDir)) {
                log.Info("Failed to delete shortcut in \"" + shortcutDir + "\" - directory does not exist.");
                return;
            }
            foreach (String file in System.IO.Directory.GetFiles(shortcutDir)) {
                if (file.EndsWith("\\OutlookGoogleCalendarSync.lnk")) {
                    System.IO.File.Delete(file);
                    log.Info("Deleted shortcut in \"" + shortcutDir + "\"");
                    break;
                }
            }
        }

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
            log.Debug("  "+ settingsFilename);
            File.Move(SettingsFile, dstFile);
            settingsFile = Path.Combine(dstDir, settingsFilename);

            foreach (string file in Directory.GetFiles(srcDir)) {
                if (Path.GetFileName(file).StartsWith("OGcalsync.log") || file.EndsWith(".csv")) {
                    dstFile = Path.Combine(dstDir, Path.GetFileName(file));
                    File.Delete(dstFile);
                    log.Debug("  " + Path.GetFileName(file));
                    if (file.EndsWith(".log")) {
                        log4net.LogManager.Shutdown();
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
    }
}
