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
            
        [STAThread]
        private static void Main(string[] args) {
            #region User File Management
            string localAppData = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);
            UserFilePath = Path.Combine(localAppData, Application.ProductName);
            settingsFile = Path.Combine(UserFilePath, settingsFilename);
            
            if (!Directory.Exists(UserFilePath))
                Directory.CreateDirectory(UserFilePath);

            log.Debug("Checking existance of settings.xml file.");
            if (!File.Exists(settingsFile)) {
                log.Info("User settings.xml file does not exist in "+ settingsFile);
                //Try and copy from where the application.exe is - this is to support legacy versions <= v1.2.4
                string sourceFilePath = Path.Combine(System.Windows.Forms.Application.StartupPath, settingsFilename);
                if (!File.Exists(sourceFilePath)) {
                    log.Info("No settings.xml file found in " + sourceFilePath);
                    Settings.Instance.Save(sourceFilePath);
                    log.Info("New blank template created.");
                }
                log.Info("Copying settings.xml to user's local appdata store.");
                File.Copy(sourceFilePath, settingsFile);
            }
            #endregion
            
            log4net.LogManager.GetRepository().LevelMap.Add(MyFineLevel);
            XmlConfigurator.Configure(new System.IO.FileInfo(logFile));
            log.Info("Program started: v"+ Application.ProductVersion);
            
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

            Application.Run(new MainForm());
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
        #endregion
    }
}
