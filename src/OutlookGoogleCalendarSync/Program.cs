using log4net;
using log4net.Config;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows.Forms;

namespace OutlookGoogleCalendarSync {
    /// <summary>
    /// Class with program entry point.
    /// </summary>
    internal sealed class Program {
        public static string UserFilePath;
        private static readonly ILog log = LogManager.GetLogger(typeof(Program));
        public const string OgcsWebsite = "https://phw198.github.io/OutlookGoogleCalendarSync";
        private const string logSettingsFile = "logger.xml";
        private const string defaultLogFilename = "OGcalsync.log";
        public static String WorkingFilesDirectory;
        public static log4net.Core.Level MyFailLevel = new log4net.Core.Level(65000, "FAIL"); //An error but not one for reporting
        //log4net.Core.Level.Fine == log4net.Core.Level.Debug (30000), so manually changing its value
        public static log4net.Core.Level MyFineLevel = new log4net.Core.Level(25000, "FINE");
        public static log4net.Core.Level MyUltraFineLevel = new log4net.Core.Level(24000, "ULTRA-FINE"); //Logs email addresses

        public static Boolean StartedWithFileArgs = false;
        public static String Title { get; private set; }
        public static Boolean StartedWithSquirrelArgs {
            get {
                String[] cliArgs = Environment.GetCommandLineArgs().Skip(1).ToArray();
                return (cliArgs.Length == 2 && cliArgs[0].ToLower().StartsWith("--squirrel"));
            }
        }
        /// <summary>
        /// The OGCS directory within user's roaming profile
        /// </summary>
        public static String RoamingProfileOGCS;

        private static Boolean? isInstalled = null;
        public static Boolean IsInstalled {
            get {
                isInstalled = isInstalled ?? Updater.IsSquirrelInstall();
                return (Boolean)isInstalled;
            }
        }
        private static Boolean isHotFix {
            get {
                return !Application.ProductVersion.EndsWith(".0");
            }
        }
        public static Updater Updater;

        [STAThread]
        private static void Main(string[] args) {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);

            try {
                setSecurityProtocols();
                GoogleOgcs.ErrorReporting.Initialise();

                RoamingProfileOGCS = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), Application.ProductName);
                parseArgumentsAndInitialise(args);

                Updater.MakeSquirrelAware();
                Program.instancesRunning();
                Forms.Splash.ShowMe();

                SettingsStore.Upgrade.Check();
                log.Debug("Loading settings from file.");
                Settings.Load();
                Settings.Instance.Proxy.Configure();

                new Telemetry.GA4Event(Telemetry.GA4Event.Event.Name.application_started).Send();
                
                Updater = new Updater();
                isNewVersion(Program.IsInstalled);
                Updater.CheckForUpdate();

                TimezoneDB.Instance.CheckForUpdate();

                try {
                    String startingTab = Settings.Instance.CompletedSyncs == 0 ? "Help" : null;
                    Application.Run(new Forms.Main(startingTab));
                } catch (ApplicationException ex) {
                    String reportError = ex.Message;
                    log.Fatal(reportError);
                    if (ex.InnerException != null) {
                        reportError = ex.InnerException.Message;
                        log.Fatal(reportError);
                    }
                    MessageBox.Show(reportError, "Application terminated!", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    throw new ApplicationException(ex.Message.StartsWith("COM error") ? "Suggest startup delay" : "");

                } catch (System.Runtime.InteropServices.COMException ex) {
                    OGCSexception.Analyse(ex);
                    MessageBox.Show(ex.Message, "Application terminated!", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    throw new ApplicationException("Suggest startup delay");
                }

            } catch (ApplicationException aex) {
                if (aex.Message == "Suggest startup delay") {
                    if (isCLIstartup() && Settings.Instance.StartOnStartup) {
                        log.Debug("Suggesting to set a startup delay.");
                        MessageBox.Show("If this error only happens when logging in to Windows, try " +
                            ((Settings.Instance.StartupDelay == 0) ? "setting a" : "increasing the") + " delay for OGCS on startup.",
                            "Set a delay on startup", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                } else if (!string.IsNullOrEmpty(aex.Message))
                    MessageBox.Show(aex.Message, "Application terminated!", MessageBoxButtons.OK, MessageBoxIcon.Error);
                log.Warn("OGCS has crashed out.");

            } catch (System.Exception ex) {
                OGCSexception.Analyse(ex, true);
                log.Fatal("Application unexpectedly terminated!");
                MessageBox.Show(ex.Message, "Application unexpectedly terminated!", MessageBoxButtons.OK, MessageBoxIcon.Error);
                log.Warn("OGCS has crashed out.");

            } finally {
                log.Debug("Shutting down application.");
                OutlookOgcs.Calendar.Disconnect();
                Forms.Splash.CloseMe();
                GC.Collect();
                GC.WaitForPendingFinalizers();
                while (Updater != null && Updater.IsBusy) {
                    Application.DoEvents();
                    System.Threading.Thread.Sleep(100);
                }
                log.Info("Application closed.");
            }
        }

        private static void parseArgumentsAndInitialise(string[] args) {
            //We're interested in non-Squirrel arguments here, ie ones which don't start with Linux-esque dashes (--squirrel)
            StartedWithFileArgs = (args.Length != 0 && args.Count(a => a.StartsWith("/") && !a.StartsWith("/d")) != 0);

            if (args.Contains("/?") || args.Contains("/help", StringComparer.OrdinalIgnoreCase)) {
                OgcsMessageBox.Show("Command line parameters:-\r\n" +
                    "  /?\t\tShow options\r\n" +
                    "  /l:OGcalsync.log\tFile to log to\r\n" +
                    "  /s:settings.xml\tSettings file to use.\r\n\t\tFile created with defaults if it doesn't exist\r\n" +
                    "  /d:60\t\tSeconds startup delay\r\n" +
                    "  /t:\"Config A\"\tAppend custom text to application title",
                    "OGCS command line parameters", MessageBoxButtons.OK, MessageBoxIcon.Information);
                Environment.Exit(0);
            }

            Dictionary<String, String> loggingArg = parseArgument(args, 'l');
            initialiseLogger(loggingArg["Filename"], loggingArg["Directory"], bootstrap: true);

            Dictionary<String, String> settingsArg = parseArgument(args, 's');
            Settings.InitialiseConfigFile(settingsArg["Filename"], settingsArg["Directory"]);

            log.Info("Storing user files in directory: " + MaskFilePath(UserFilePath));

            //Before settings have been loaded, early config of cloud logging
            GoogleOgcs.ErrorReporting.UpdateLogUuId();
            Boolean cloudLogSetting = false;
            String cloudLogXmlSetting = XMLManager.ImportElement("CloudLogging", Settings.ConfigFile);
            if (!string.IsNullOrEmpty(cloudLogXmlSetting)) cloudLogSetting = Boolean.Parse(cloudLogXmlSetting);
            GoogleOgcs.ErrorReporting.SetThreshold(cloudLogSetting);

            if (!StartedWithFileArgs) {
                //Now let's confirm files are actually in the right place
                Boolean keepPortable = (XMLManager.ImportElement("Portable", Settings.ConfigFile) ?? "false").Equals("true");
                if (keepPortable) {
                    if (UserFilePath != System.Windows.Forms.Application.StartupPath) {
                        log.Info("File storage location is incorrect according to " + Settings.ConfigFile);
                        MakePortable(true);
                    }
                } else {
                    if (UserFilePath != Program.RoamingProfileOGCS) {
                        log.Info("File storage location is incorrect according to " + Settings.ConfigFile);
                        MakePortable(false);
                    }
                }
            }

            string logLevel = XMLManager.ImportElement("LoggingLevel", Settings.ConfigFile);
            Settings.configureLoggingLevel(logLevel ?? "FINE");

            if (args.Contains("--delay")) { //Format up to and including v2.7.1
                log.Info("Converting old --delay parameter to /d");
                try {
                    String delay = args[Array.IndexOf(args, "--delay") + 1];
                    log.Debug("Delay of " + delay + "s being migrated.");
                    addRegKey(Microsoft.Win32.Registry.CurrentUser, delay);
                    delayStartup(delay);
                } catch (System.Exception ex) {
                    log.Error(ex.Message);
                }
            }
            Dictionary<String, String> delayArg = parseArgument(args, 'd');
            if (delayArg["Value"] != null) delayStartup(delayArg["Value"]);

            Dictionary<String, String> titleArg = parseArgument(args, 't');
            Title = titleArg["Value"];
        }

        private static Dictionary<String, String> parseArgument(String[] args, char arg) {
            Dictionary<String, String> details = new Dictionary<String, String>();
            details.Add("Value", null);
            details.Add("Directory", null);
            details.Add("Filename", null);

            try {
                String argVal = args.Where(a => a.ToLower().StartsWith("/" + arg + ":")).FirstOrDefault();
                if (argVal != null) {
                    details["Value"] = argVal.Split(':')[1];
                    if (arg == 'l' || arg == 's') {
                        details["Filename"] = System.IO.Path.GetFileName(argVal);
                        if (string.IsNullOrEmpty(details["Filename"]) || !Path.HasExtension(details["Filename"])) {
                            throw new ApplicationException("The /" + arg + " parameter must be used with a filename.");
                        }
                        details["Directory"] = System.IO.Path.GetDirectoryName(argVal.TrimStart(("/" + arg + ":").ToCharArray()));
                        if (!string.IsNullOrEmpty(details["Directory"]) && !System.IO.Directory.Exists(details["Directory"])) {
                            throw new ApplicationException("The specified directory '" + details["Directory"] + "' does not exist.\r\n" +
                                "Please correct the parameter value passed or create the directory.");
                        }
                    }
                }
            } catch (System.Exception ex) {
                throw new ApplicationException("Failed processing /" + arg + " parameter.\r\n" + ex.Message);
            }
            return details;
        }

        private static void initialiseLogger(string logFilename, string logPath = null, Boolean bootstrap = false) {
            if (string.IsNullOrEmpty(logFilename)) logFilename = defaultLogFilename;
            log4net.GlobalContext.Properties["LogFilename"] = logFilename;
            if (string.IsNullOrEmpty(logPath)) {
                if (Program.IsInstalled || File.Exists(Path.Combine(RoamingProfileOGCS, logFilename)))
                    logPath = RoamingProfileOGCS;
                else
                    logPath = Application.StartupPath;
            }
            UserFilePath = logPath;
            log4net.GlobalContext.Properties["LogPath"] = logPath + "\\";
            log4net.LogManager.GetRepository().LevelMap.Add(MyFailLevel);
            log4net.LogManager.GetRepository().LevelMap.Add(MyFineLevel);
            log4net.LogManager.GetRepository().LevelMap.Add(MyUltraFineLevel);

            GoogleOgcs.ErrorReporting.LogId = "v" + Application.ProductVersion;
            GoogleOgcs.ErrorReporting.UpdateLogUuId();

            XmlConfigurator.Configure(new System.IO.FileInfo(
                Path.Combine(Path.GetDirectoryName(System.Windows.Forms.Application.ExecutablePath), logSettingsFile)
            ));

            GoogleOgcs.ErrorReporting.SetThreshold(false);

            if (bootstrap) {
                log.Info("Program started: v" + Application.ProductVersion);
                log.Info("Started " + (isCLIstartup() ? "automatically" : "interactively") + ".");
                if (Environment.GetCommandLineArgs().Count() > 1)
                    log.Info("Invoked with arguments: " + string.Join(" ", Environment.GetCommandLineArgs().Skip(1).ToArray()));
            }
            log.Info("Logging to: " + MaskFilePath(UserFilePath) + "\\" + logFilename);
            purgeLogFiles(30);
        }

        private static void purgeLogFiles(Int16 retention) {
            log.Info("Purging log files older than " + retention + " days...");
            foreach (String file in System.IO.Directory.GetFiles(UserFilePath, "*.log.????-??-??", SearchOption.TopDirectoryOnly)) {
                if (System.IO.File.GetLastWriteTime(file) < DateTime.Now.AddDays(-retention)) {
                    try {
                        System.IO.File.Delete(file);
                        log.Debug("Deleted " + MaskFilePath(file));
                    } catch (System.Exception ex) {
                        OGCSexception.Analyse("Could not delete file " + file, OGCSexception.LogAsFail(ex));
                    }
                }
            }
            log.Info("Purge complete.");
        }

        #region Application Behaviour
        #region Startup Registry Key
        private static Microsoft.Win32.RegistryKey openStartupRegKey(Microsoft.Win32.RegistryKey hive, Boolean forWriting = false) {
            String path = null;
            if (hive == Microsoft.Win32.Registry.CurrentUser) path = @"Software\Microsoft\Windows\CurrentVersion\Run";
            else if (hive == Microsoft.Win32.Registry.LocalMachine) path = @"Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Run";
            else throw new ApplicationException("Unexpected registry hive: " + hive.ToString());

            Microsoft.Win32.RegistryKey openedKey = hive.OpenSubKey(path, forWriting);
            if (openedKey == null) {
                log.Warn("The startup registry path does not exist in " + hive.ToString() + @"\" + path);
                if (forWriting) {
                    log.Info("Creating startup registry path " + hive.ToString() + @"\" + path);
                    openedKey = hive.CreateSubKey(path, Microsoft.Win32.RegistryKeyPermissionCheck.ReadWriteSubTree);
                }
            }
            return openedKey;
        }
        public static void ManageStartupRegKey() {
            //Check for legacy Startup menu shortcut <=v2.1.4
            Boolean startupConfigExists = Program.CheckShortcut(Environment.SpecialFolder.Startup);
            if (startupConfigExists)
                Program.RemoveShortcut(Environment.SpecialFolder.Startup);

            Boolean startupConfigExistsHKCU = checkRegKey(Microsoft.Win32.Registry.CurrentUser);
            Boolean startupConfigExistsHKLM = checkRegKey(Microsoft.Win32.Registry.LocalMachine);

            if (Settings.Instance.StartOnStartup) {
                if (startupConfigExistsHKCU) log.Debug("Forcing update of HKCU startup registry key.");
                addRegKey(Microsoft.Win32.Registry.CurrentUser);
                if (Settings.Instance.StartOnStartupAllUsers) {
                    if (startupConfigExistsHKLM) log.Debug("Forcing update of HKLM startup registry key.");
                    addRegKey(Microsoft.Win32.Registry.LocalMachine);
                } else {
                    if (startupConfigExistsHKLM) removeRegKey(Microsoft.Win32.Registry.LocalMachine);
                    else log.Debug("No HKLM startup registry key to remove.");
                }
            } else {
                if (startupConfigExistsHKCU) removeRegKey(Microsoft.Win32.Registry.CurrentUser);
                else log.Debug("No HKCU startup registry key to remove.");
                if (startupConfigExistsHKLM) removeRegKey(Microsoft.Win32.Registry.LocalMachine);
                else log.Debug("No HKLM startup registry key to remove.");
            }
        }

        private static Boolean checkRegKey(Microsoft.Win32.RegistryKey hive) {
            Microsoft.Win32.RegistryKey startupKey = null;
            try {
                startupKey = openStartupRegKey(hive);
                String[] regKeys = startupKey?.GetValueNames();
                return regKeys?.Contains(Application.ProductName) ?? false;
            } finally {
                startupKey?.Close();
            }
        }

        private static void addRegKey(Microsoft.Win32.RegistryKey hive, String startupDelay = null) {
            Microsoft.Win32.RegistryKey startupKey = openStartupRegKey(hive, true);
            String keyValue = startupKey.GetValue(Application.ProductName, "").ToString();
            String delayedStartup = "";
            if (Convert.ToInt16(startupDelay ?? Settings.Instance.StartupDelay.ToString()) > 0)
                delayedStartup = " /d:" + (startupDelay ?? Settings.Instance.StartupDelay.ToString());

            String cliArgs = string.Join(" ", Environment.GetCommandLineArgs().Skip(1).Where(a => "l,s".Contains(a.Substring(1, 1).ToLower())));
            cliArgs = (" " + cliArgs).TrimEnd();

            if (keyValue == "" || keyValue != (Application.ExecutablePath + delayedStartup + cliArgs)) {
                log.Debug("Startup " + hive.ToString() + " registry key " + (keyValue == "" ? "created" : "updated") + ".");
                try {
                    startupKey.SetValue(Application.ProductName, Application.ExecutablePath + delayedStartup + cliArgs);
                } catch (System.UnauthorizedAccessException ex) {
                    log.Warn("Could not create/update " + hive.ToString() + " registry key. " + ex.Message);
                    Settings.Instance.StartOnStartup = false;
                    if (OgcsMessageBox.Show("You don't have permission to update the registry, so the application can't be set to run on startup.\r\n" +
                        "Try manually adding a shortcut to the 'Startup' folder in Windows instead?", "Permission denied", MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation)
                        == DialogResult.Yes) {
                        System.Diagnostics.Process.Start(System.Windows.Forms.Application.StartupPath);
                        System.Diagnostics.Process.Start(Environment.GetFolderPath(Environment.SpecialFolder.Startup));
                    }
                }
            }
            startupKey.Close();
        }

        private static void removeRegKey(Microsoft.Win32.RegistryKey hive) {
            log.Debug("Startup registry key being removed from " + hive.ToString());
            Microsoft.Win32.RegistryKey startupKey = null;
            try {
                startupKey = openStartupRegKey(hive, true);
                startupKey.DeleteValue(Application.ProductName, false);
            } finally {
                startupKey?.Close();
            }
        }
        #endregion
        private static void delayStartup(String seconds) {
            try {
                DateTime delayUntil = DateTime.Now.AddSeconds(Convert.ToInt32(seconds));
                log.Info("Startup delay configured until " + delayUntil.ToString("HH:mm:ss"));
                while (DateTime.Now < delayUntil) {
                    System.Threading.Thread.Sleep(250);
                }
            } catch (System.Exception ex) {
                log.Warn("Failure in delayStartup(). Seconds: " + seconds);
                log.Error(ex.Message);
            }
        }

        #region Legacy Start Menu Shortcut
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
            if (StartedWithFileArgs) {
                log.Warn("Cannot move user files when OGCS is started with CLI arguments.");
                return;
            }

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
                if (RoamingProfileOGCS == UserFilePath) {
                    log.Info("It already is!");
                    return;
                }
                if (!Directory.Exists(RoamingProfileOGCS))
                    Directory.CreateDirectory(RoamingProfileOGCS);

                moveFiles(UserFilePath, RoamingProfileOGCS);
            }
        }

        private static void moveFiles(string srcDir, string dstDir) {
            log.Info("Moving files from " + srcDir + " to " + dstDir + ":-");
            if (!Directory.Exists(dstDir)) Directory.CreateDirectory(dstDir);

            string dstFile = Path.Combine(dstDir, Settings.ConfigFilename);
            File.Delete(dstFile);
            log.Debug("  " + Settings.ConfigFilename);
            File.Move(Settings.ConfigFile, dstFile);
            WorkingFilesDirectory = dstDir;

            foreach (string file in Directory.GetFiles(srcDir)) {
                if (Path.GetFileName(file).StartsWith("OGcalsync.log") || file.EndsWith(".csv") || file.EndsWith(".json") || file == GoogleOgcs.Authenticator.TokenFile) {
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

        private static void isNewVersion(Boolean isSquirrelInstall) {
            string settingsVersion = string.IsNullOrEmpty(Settings.Instance.Version) ? "Unknown" : Settings.Instance.Version;
            if (settingsVersion != Application.ProductVersion) {
                if (settingsVersion == "Unknown") log.Info("New install and/or brand new settings file detected.");
                else log.Info("New upgraded version detected: from " + settingsVersion + " to " + Application.ProductVersion);
                try {
                    Program.ManageStartupRegKey();
                } catch (System.Exception ex) {
                    if (ex is System.Security.SecurityException) OGCSexception.LogAsFail(ref ex); //User doesn't have rights to access registry
                    OGCSexception.Analyse("Failed accessing registry for startup key.", ex);
                }
                Settings.Instance.Version = Application.ProductVersion;
                if (isHotFix) {
                    if (!(Settings.Instance.CloudLogging ?? false) | Settings.Instance.TelemetryDisabled) {
                        String disabledSetting = (!(Settings.Instance.CloudLogging ?? false) ? "automatic feedback of errors" : "");
                        if (Settings.Instance.TelemetryDisabled) {
                            if (!String.IsNullOrEmpty(disabledSetting)) disabledSetting += " and ";
                            disabledSetting += "telemetry";
                        }
                        if (OgcsMessageBox.Show("As you are running a hotfix release, it would be helpful if you could enable " + disabledSetting + ".",
                            "OGCS hotfix release troubleshooting", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes) {
                            Settings.Instance.TelemetryDisabled = false;
                            Settings.Instance.CloudLogging = true;
                        }
                    }
                } else { //Release notes not updated for hotfixes.
                    String releaseNotesUrl = "/release-notes.html";
                    if (!String.IsNullOrEmpty(Settings.Instance.GaccountEmail)) {
                        byte[] plainTextBytes = System.Text.Encoding.UTF8.GetBytes(Settings.Instance.GaccountEmail);
                        releaseNotesUrl += "?id=" + System.Convert.ToBase64String(plainTextBytes);
                    }
                    Helper.OpenBrowser(OgcsWebsite + releaseNotesUrl);
                    if (isSquirrelInstall) {
                        Telemetry.Send(Analytics.Category.squirrel, Analytics.Action.upgrade, "from=" + settingsVersion + ";to=" + Application.ProductVersion);
                        Telemetry.GA4Event.Event squirrelGaEv = new(Telemetry.GA4Event.Event.Name.squirrel);
                        squirrelGaEv.AddParameter(GA4.Squirrel.upgraded_from, settingsVersion);
                    }
                }
            }

            //Check upgrade to Squirrel release went OK
            try {
                if (isSquirrelInstall) {
                    Int32 upgradedFrom = Int16.MaxValue;
                    String expectedInstallDir = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);
                    expectedInstallDir = Path.Combine(expectedInstallDir, "OutlookGoogleCalendarSync");
                    if (settingsVersion != "Unknown") {
                        upgradedFrom = Program.VersionToInt(settingsVersion);
                    }
                    if (!Program.InDeveloperMode && (settingsVersion == "Unknown" || upgradedFrom < 2050000) &&
                        !System.Windows.Forms.Application.ExecutablePath.ToString().StartsWith(expectedInstallDir))
                    {
                        log.Warn("OGCS is running from " + System.Windows.Forms.Application.ExecutablePath.ToString());
                        OgcsMessageBox.Show("A suspected improper install location has been detected.\r\n" +
                            "Click 'OK' for further details.", "Improper Install Location",
                            MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        Helper.OpenBrowser("https://github.com/phw198/OutlookGoogleCalendarSync/issues/265");
                    }
                }
            } catch (System.Exception ex) {
                log.Warn("Failed to determine if OGCS is installed in the correct location.");
                log.Error(ex.Message);
            }
        }

        private static Boolean isCLIstartup() {
            try {
                if (File.Exists(logSettingsFile)) return false;
                else if (File.Exists(Path.Combine(Path.GetDirectoryName(System.Windows.Forms.Application.ExecutablePath), logSettingsFile))) return true;
                else return false;
            } catch (System.Exception ex) {
                log.Error("Failed to determine if OGCS was started by CLI.");
                OGCSexception.Analyse(ex);
                return false;
            }
        }

        public static void Donate(String source) {
            try {
                Telemetry.Send(Analytics.Category.ogcs, Analytics.Action.donate, source);
                Telemetry.Send(Analytics.Category.ogcs, Analytics.Action.donate, Application.ProductVersion);
                
                Telemetry.GA4Event.Event donateGa4Ev = new(Telemetry.GA4Event.Event.Name.donate);
                donateGa4Ev.AddParameter("source", source);
                donateGa4Ev.AddParameter(GA4.General.sync_count, Settings.Instance.CompletedSyncs);
                donateGa4Ev.AddParameter("account_present", !String.IsNullOrEmpty(Settings.Instance.GaccountEmail));
                donateGa4Ev.Send();

            } finally {
                Helper.OpenBrowser("https://www.paypal.com/cgi-bin/webscr?cmd=_s-xclick&hosted_button_id=44DUQ7UT6WE2C&item_name=Outlook Google Calendar Sync from " + Settings.Instance.GaccountEmail);
            }
        }

        /// <summary>
        /// Convert a semantic version number string to an integer.
        /// </summary>
        /// <param name="semanticVersion">The semantic version number.</param>
        /// <returns>The converted integer version number.</returns>
        public static Int32 VersionToInt(String semanticVersion) {
            String paddedVersion = "";
            foreach (String versionBit in semanticVersion.Split('.')) {
                paddedVersion += versionBit.PadLeft(2, '0');
            }
            return Convert.ToInt32(paddedVersion);
        }

        public static Boolean InDeveloperMode {
            get { return System.Diagnostics.Debugger.IsAttached; }
        }

        /// <summary>
        /// Replace the %USERNAME% element, if present in a file path, with <userid>
        /// </summary>
        /// <param name="path">The path to check</param>
        /// <returns>The maskes path</returns>
        public static string MaskFilePath(String path) {
            try {
                String userProfile = Environment.GetEnvironmentVariable("USERPROFILE");
                if (path.StartsWith(userProfile)) {
                    String username = Environment.GetEnvironmentVariable("USERNAME");
                    if (username == null) {
                        log.Debug("User:    " + Environment.GetEnvironmentVariable("USERNAME", EnvironmentVariableTarget.User));
                        log.Debug("Process: " + Environment.GetEnvironmentVariable("USERNAME", EnvironmentVariableTarget.Process));
                        log.Debug("Machine: " + Environment.GetEnvironmentVariable("USERNAME", EnvironmentVariableTarget.Machine));
                        log.Error("%USERNAME% environment variable not available. This may well fix itself with a reboot #1282");
                        return path;
                    }
                    String userProfileMasked = userProfile.Replace(username, "<userid>");
                    return path.Replace(userProfile, userProfileMasked);
                } else
                    return path;
            } catch (System.Exception ex) {
                OGCSexception.Analyse("Problems accessing environment variables.", ex);
                return path;
            }
        }

        private static void setSecurityProtocols() {
            //Enable TSL1.1,1.2
            System.Net.ServicePointManager.SecurityProtocol |= System.Net.SecurityProtocolType.Tls11 | System.Net.SecurityProtocolType.Tls12;
            //Disable SSL3?
            //System.Net.ServicePointManager.SecurityProtocol &= ~System.Net.SecurityProtocolType.Ssl3;
        }

        /// <summary>
        /// Determine what process is in the current call stack
        /// </summary>
        /// <param name="callingProcessNames">A comma-separated list of process names</param>
        /// <returns>True if the call stack contains any of the process names</returns>
        public static Boolean CalledByProcess(String callingProcessNames) {
            String[] processNames = callingProcessNames.Split(',');
            System.Diagnostics.StackTrace stackTrace = new System.Diagnostics.StackTrace();
            foreach (System.Diagnostics.StackFrame frame in stackTrace.GetFrames().Reverse()) {
                if (processNames.Contains(frame.GetMethod().Name, StringComparer.OrdinalIgnoreCase)) {
                    return true;
                }
            }
            return false;
        }

        public static void StackTraceToString() {
            try {
                String stackString = "";
                List<System.Diagnostics.StackFrame> stackFrames = new System.Diagnostics.StackTrace().GetFrames().ToList();
                stackFrames.ForEach(sf => stackString += sf.GetMethod().Name + " < ");
                log.Warn("StackTrace path: " + stackString);
            } catch (System.Exception ex) {
                OGCSexception.Analyse(ex);
            }
        }

        /// <summary>Check how many OGCS processes we have running</summary>
        private static void instancesRunning() {
            try {
                System.Diagnostics.Process currentProcess = System.Diagnostics.Process.GetCurrentProcess();
                System.Diagnostics.Process[] processes = System.Diagnostics.Process.GetProcessesByName(currentProcess.ProcessName);
                
                if (processes.Count() > 1) {
                    log.Warn("There are " + processes.Count() + " " + currentProcess.ProcessName + " processes currently running.");
                    List<System.Linq.IGrouping<string, System.Diagnostics.Process>> sameExe = processes.GroupBy(p => p.MainModule.FileName).Where(e => e.Count() > 1).ToList();
                    log.Debug(sameExe.Count() + " executables have more than one process attached; checking runtime arguments");
                    log.Debug("Current process command line:-");
                    String currentCmdLine = getProcessCommandLine(currentProcess.Id);

                    foreach (System.Linq.IGrouping<string, System.Diagnostics.Process> exe in sameExe) {
                        log.Debug("Checking other processes running the same executable:-");
                        log.Debug(exe.Key);
                        foreach (System.Diagnostics.Process process in exe) {
                            if (process.Id == currentProcess.Id) continue;

                            String cmdLine = getProcessCommandLine(process.Id);
                            if (cmdLine == currentCmdLine) {
                                OgcsMessageBox.Show("You already have an instance of OGCS running using the same configuration.\r\n" +
                                    "This is not recommended and may cause problems if they sync at the same time.",
                                    "Multiple OGCS instances running", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                                return;
                            }
                        }
                        log.Debug("OK - they are running with different configurations.");
                    }
                }

            } catch (System.Exception ex) {
                OGCSexception.Analyse("Unable to check for concurrent OGCS processes.", ex);
            }
        }

        private static String getProcessCommandLine(int processId) {
            System.Management.ManagementObjectSearcher commandLineSearcher = new System.Management.ManagementObjectSearcher("SELECT CommandLine FROM Win32_Process WHERE ProcessId = " + processId);
            String commandLine = "";
            foreach (System.Management.ManagementObject commandLineObject in commandLineSearcher.Get()) {
                commandLine += (String)commandLineObject["CommandLine"];
            }
            log.Debug(" " + commandLine);

            return commandLine;
        }
    }
}
