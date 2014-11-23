using System;
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
        private static readonly ILog log = LogManager.GetLogger(typeof(Program));
        private static string logFile = "logger.xml";

        [STAThread]
        private static void Main(string[] args) {
            XmlConfigurator.Configure(new System.IO.FileInfo(logFile));
            log.Info("Program started: v"+ Application.ProductVersion);
            
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
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
