using log4net;
using System;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Runtime.Serialization;

namespace OutlookGoogleCalendarSync {
    /// <summary>
    /// Description of Settings.
    /// </summary>

    [DataContract]
    public class Settings {
        private static readonly ILog log = LogManager.GetLogger(typeof(Settings));

        private static String configFilename = "settings.xml";
        public static String ConfigFilename {
            get { return configFilename; }
        }
        /// <summary>
        /// Absolute path to config file, eg C:\foo\bar\settings.xml
        /// </summary>
        public static String ConfigFile {
            get { return Path.Combine(Program.WorkingFilesDirectory, ConfigFilename); }
        }

        public static void InitialiseConfigFile(String filename, String directory = null) {
            if (!string.IsNullOrEmpty(filename)) configFilename = filename;
            Program.WorkingFilesDirectory = directory;

            if (string.IsNullOrEmpty(Program.WorkingFilesDirectory)) {
                if (Program.IsInstalled || File.Exists(Path.Combine(Program.RoamingProfileOGCS, ConfigFilename)))
                    Program.WorkingFilesDirectory = Program.RoamingProfileOGCS;
                else
                    Program.WorkingFilesDirectory = System.Windows.Forms.Application.StartupPath;
            }

            if (!File.Exists(ConfigFile)) {
                log.Info("No settings.xml file found in " + Program.WorkingFilesDirectory);
                Settings.Instance.Save(ConfigFile);
                log.Info("New blank template created.");
                if (!Program.IsInstalled)
                    XMLManager.ExportElement("Portable", true, ConfigFile);
            }

            log.Info("Running OGCS from " + System.Windows.Forms.Application.ExecutablePath);
        }

        private static Settings instance;
        //Settings saved immediately
        private Boolean apiLimit_inEffect;
        private DateTime apiLimit_lastHit;
        private DateTime lastSyncDate;
        private Int32 completedSyncs;
        private Boolean portable;
        private Boolean alphaReleases;
        private String version;
        private Boolean donor;
        private Boolean hideSplashScreen;

        public Settings() {
            setDefaults();
        }

        //Default values before loading from xml and attribute not yet serialized
        [OnDeserializing]
        void OnDeserializing(StreamingContext context) {
            setDefaults();
        }

        private void setDefaults() {
            //Default values
            assignedClientIdentifier = "";
            assignedClientSecret = "";
            PersonalClientIdentifier = "";
            PersonalClientSecret = "";
            OutlookService = OutlookOgcs.Calendar.Service.DefaultMailbox;
            MailboxName = "";
            SharedCalendar = "";
            UseOutlookCalendar = new OutlookCalendarListEntry();
            CategoriesRestrictBy = RestrictBy.Exclude;
            Categories = new System.Collections.Generic.List<String>();
            OnlyRespondedInvites = false;
            OutlookDateFormat = "g";
            outlookGalBlocked = false;

            UseGoogleCalendar = new GoogleCalendarListEntry();
            apiLimit_inEffect = false;
            apiLimit_lastHit = DateTime.Parse("01-Jan-2000");
            GaccountEmail = "";
            CloakEmail = true;

            SyncDirection = Sync.Direction.OutlookToGoogle;
            DaysInThePast = 1;
            DaysInTheFuture = 60;
            SyncInterval = 0;
            SyncIntervalUnit = "Hours";
            OutlookPush = false;
            AddLocation = true;
            AddDescription = true;
            AddDescription_OnlyToGoogle = true;
            AddReminders = false;
            UseGoogleDefaultReminder = false;
            UseOutlookDefaultReminder = false;
            ReminderDND = false;
            ReminderDNDstart = DateTime.Now.Date.AddHours(22);
            ReminderDNDend = DateTime.Now.Date.AddDays(1).AddHours(6);
            AddAttendees = false;
            AddColours = false;
            MergeItems = true;
            DisableDelete = true;
            ConfirmOnDelete = true;
            TargetCalendar = Sync.Direction.OutlookToGoogle;
            CreatedItemsOnly = true;
            SetEntriesPrivate = false;
            SetEntriesAvailable = false;
            SetEntriesColour = false;
            SetEntriesColourValue = Microsoft.Office.Interop.Outlook.OlCategoryColor.olCategoryColorNone.ToString();
            SetEntriesColourName = "None";
            Obfuscation = new Obfuscate();

            MuteClickSounds = false;
            ShowBubbleTooltipWhenSyncing = true;
            StartOnStartup = false;
            StartupDelay = 0;
            StartInTray = false;
            MinimiseToTray = false;
            MinimiseNotClose = false;
            ShowBubbleWhenMinimising = true;

            CreateCSVFiles = false;
            LoggingLevel = "DEBUG";
            portable = false;
            Proxy = new SettingsProxy();

            alphaReleases = !System.Windows.Forms.Application.ProductVersion.EndsWith("0.0");
            SkipVersion = null;
            Subscribed = DateTime.Parse("01-Jan-2000");
            donor = false;
            hideSplashScreen = false;
            
            lastSyncDate = new DateTime(0);
            completedSyncs = 0;
            VerboseOutput = true;
        }

        public static Boolean InstanceInitialiased() {
            return (instance != null);
        }

        public static Settings Instance {
            get {
                if (instance == null) instance = new Settings();
                return instance;
            }
            set {
                instance = value;
            }
        }
        
        #region Outlook
        public enum RestrictBy {
            Include, Exclude
        }
        [DataMember] public OutlookOgcs.Calendar.Service OutlookService { get; set; }
        [DataMember] public string MailboxName { get; set; }
        [DataMember] public string SharedCalendar { get; set; }
        [DataMember] public OutlookCalendarListEntry UseOutlookCalendar { get; set; }
        [DataMember] public RestrictBy CategoriesRestrictBy { get; set; }
        [DataMember] public System.Collections.Generic.List<string> Categories { get; set; }
        [DataMember] public Boolean OnlyRespondedInvites { get; set; }
        [DataMember] public string OutlookDateFormat { get; set; }
        private Boolean outlookGalBlocked;
        [DataMember] public Boolean OutlookGalBlocked {
            get { return outlookGalBlocked; }
            set {
                outlookGalBlocked = value;
                if (!loading() && Forms.Main.Instance.IsHandleCreated) Forms.Main.Instance.FeaturesBlockedByCorpPolicy(value);
            }
        }
        #endregion
        #region Google
        private String assignedClientIdentifier;
        [DataMember] public String AssignedClientIdentifier {
            get { return assignedClientIdentifier; }
            set {
                assignedClientIdentifier = value.Trim();
                if (!loading()) XMLManager.ExportElement("AssignedClientIdentifier", value.Trim(), ConfigFile);
            }
        }
        private String assignedClientSecret;
        [DataMember] public String AssignedClientSecret {
            get { return assignedClientSecret; }
            set {
                assignedClientSecret = value.Trim();
                if (!loading()) XMLManager.ExportElement("AssignedClientSecret", value.Trim(), ConfigFile);
            }
        }
        private String personalClientIdentifier;
        private String personalClientSecret;
        [DataMember] public String PersonalClientIdentifier {
            get { return personalClientIdentifier; }
            set { personalClientIdentifier = value.Trim(); }
        }
        [DataMember] public String PersonalClientSecret {
            get { return personalClientSecret; }
            set { personalClientSecret = value.Trim(); }
        }
        public Boolean UsingPersonalAPIkeys() {
            return !string.IsNullOrEmpty(PersonalClientIdentifier) && !string.IsNullOrEmpty(PersonalClientSecret);
        }
        [DataMember] public GoogleCalendarListEntry UseGoogleCalendar { get; set; }
        [DataMember] public Boolean APIlimit_inEffect {
            get { return apiLimit_inEffect; }
            set {
                apiLimit_inEffect = value;
                if (!loading()) XMLManager.ExportElement("APIlimit_inEffect", value, ConfigFile);
            }
        }
        [DataMember] public DateTime APIlimit_lastHit {
            get { return apiLimit_lastHit; }
            set {
                apiLimit_lastHit = value;
                if (!loading()) XMLManager.ExportElement("APIlimit_lastHit", value, ConfigFile);
            }
        }
        [DataMember] public String GaccountEmail { get; set; }
        public String GaccountEmail_masked() {
            if (string.IsNullOrWhiteSpace(GaccountEmail)) return "<null>";
            return EmailAddress.MaskAddress(GaccountEmail);
        }
        [DataMember] public Boolean CloakEmail { get; set; }
        #endregion
        #region Sync Options
        //Main
        public DateTime SyncStart { get { return DateTime.Today.AddDays(-DaysInThePast); } }
        public DateTime SyncEnd { get { return DateTime.Today.AddDays(+DaysInTheFuture + 1); } }
        [DataMember] public Sync.Direction SyncDirection { get; set; }
        [DataMember] public int DaysInThePast { get; set; }
        [DataMember] public int DaysInTheFuture { get; set; }
        [DataMember] public int SyncInterval { get; set; }
        [DataMember] public String SyncIntervalUnit { get; set; }
        [DataMember] public bool OutlookPush { get; set; }
        [DataMember] public bool AddLocation { get; set; }
        [DataMember] public bool AddDescription { get; set; }
        [DataMember] public bool AddDescription_OnlyToGoogle { get; set; }
        [DataMember] public bool AddReminders { get; set; }
        [DataMember] public bool UseGoogleDefaultReminder { get; set; }
        [DataMember] public bool UseOutlookDefaultReminder { get; set; }
        [DataMember] public bool ReminderDND { get; set; }
        [DataMember] public DateTime ReminderDNDstart { get; set; }
        [DataMember] public DateTime ReminderDNDend { get; set; }
        [DataMember] public bool AddAttendees { get; set; }
        [DataMember] public bool AddColours { get; set; }
        [DataMember] public bool MergeItems { get; set; }
        [DataMember] public bool DisableDelete { get; set; }
        [DataMember] public bool ConfirmOnDelete { get; set; }
        [DataMember] public Sync.Direction TargetCalendar { get; set; }
        [DataMember] public Boolean CreatedItemsOnly { get; set; }
        [DataMember] public bool SetEntriesPrivate { get; set; }
        [DataMember] public bool SetEntriesAvailable { get; set; }
        [DataMember] public bool SetEntriesColour { get; set; }
        [DataMember] public String SetEntriesColourValue { get; set; }
        [DataMember] public String SetEntriesColourName { get; set; }
        
        //Obfuscation
        [DataMember] public Obfuscate Obfuscation { get; set; }

        #endregion
        #region App behaviour
        [DataMember] public bool HideSplashScreen {
            get { return hideSplashScreen; }
            set {
                if (!loading() && hideSplashScreen != value) {
                    XMLManager.ExportElement("HideSplashScreen", value, ConfigFile);
                    if (Forms.Main.Instance != null) Forms.Main.Instance.cbHideSplash.Checked = value;
                }
                hideSplashScreen = value;
            }
        }

        [DataMember] public bool ShowBubbleTooltipWhenSyncing { get; set; }
        [DataMember] public bool StartOnStartup { get; set; }
        [DataMember] public Int32 StartupDelay { get; set; }
        [DataMember] public bool StartInTray { get; set; }
        [DataMember] public bool MinimiseToTray { get; set; }
        [DataMember] public bool MinimiseNotClose { get; set; }
        [DataMember] public bool ShowBubbleWhenMinimising { get; set; }
        [DataMember] public bool Portable {
            get { return portable; }
            set {
                portable = value;
                if (!loading()) XMLManager.ExportElement("Portable", value, ConfigFile);
            }
        }

        [DataMember] public bool CreateCSVFiles { get; set; }
        [DataMember] public String LoggingLevel { get; set; }
        private bool? cloudLogging;
        [DataMember] public bool? CloudLogging {
            get { return cloudLogging; }
            set {
                cloudLogging = value;
                GoogleOgcs.ErrorReporting.SetThreshold(value ?? false);
            }
        }
        //Proxy
        [DataMember] public SettingsProxy Proxy { get; set; }
        #endregion
        #region About
        [DataMember] public string Version {
            get { return version; }
            set {
                if (version != null && version != value) {
                    XMLManager.ExportElement("Version", value, ConfigFile);
                }
                version = value;
            }
        }
        [DataMember] public bool AlphaReleases {
            get { return alphaReleases; }
            set {
                alphaReleases = value;
                if (!loading()) XMLManager.ExportElement("AlphaReleases", value, ConfigFile);
            }
        }
        [DataMember] public DateTime Subscribed { get; set; }
        [DataMember] public Boolean Donor {
            get { return donor; }
            set {
                donor = value;
                if (!loading()) XMLManager.ExportElement("Donor", value, ConfigFile);
            }
        }
        #endregion

        [DataMember] public DateTime LastSyncDate {
            get { return lastSyncDate; }
            set {
                lastSyncDate = value;
                if (!loading()) XMLManager.ExportElement("LastSyncDate", value, ConfigFile);
            }
        }
        [DataMember] public Int32 CompletedSyncs {
            get { return completedSyncs; }
            set {
                completedSyncs = value;
                if (!loading()) XMLManager.ExportElement("CompletedSyncs", value, ConfigFile);
            }
        }
        [DataMember] public bool VerboseOutput { get; set; }
        [DataMember] public bool MuteClickSounds { get; set; }
        [DataMember] public String SkipVersion { get; set; }

        private static Boolean isLoaded = false;
        public static Boolean IsLoaded {
            get { return isLoaded; }
        }

        public static void Load(string XMLfile = null) {
            try {
                Settings.Instance = XMLManager.Import<Settings>(XMLfile ?? ConfigFile);
                log.Fine("User settings loaded.");
                Settings.isLoaded = true;
            } catch (ApplicationException ex) {
                log.Error(ex.Message);
                System.Windows.Forms.MessageBox.Show("Your OGCS settings appear to be corrupt and will have to be reset.",
                    "Corrupt OGCS Settings", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Exclamation);
                log.Warn("Resetting settings.xml file to defaults.");
                System.IO.File.Delete(XMLfile ?? ConfigFile);
                Settings.Instance.Save(XMLfile ?? ConfigFile);
                try {
                    Settings.Instance = XMLManager.Import<Settings>(XMLfile ?? ConfigFile);
                    log.Debug("User settings loaded successfully this time.");
                } catch (System.Exception ex2) {
                    log.Error("Still failed to load settings!");
                    OGCSexception.Analyse(ex2);
                }
            }
        }

        public void Save(string XMLfile = null) {
            log.Info("Saving settings.");
            XMLManager.Export(this, XMLfile ?? ConfigFile);
        }

        private Boolean loading() {
            StackTrace stackTrace = new StackTrace();
            foreach (StackFrame frame in stackTrace.GetFrames().Reverse()) {
                if (new String[] {"Load","isNewVersion"}.Contains(frame.GetMethod().Name)) {
                    return true;
                }
            }
            return false;
        }

        public void LogSettings() {
            log.Info(ConfigFile);
            log.Info("OUTLOOK SETTINGS:-");
            log.Info("  Service: "+ OutlookService.ToString());
            if (OutlookService == OutlookOgcs.Calendar.Service.SharedCalendar) {
                log.Info("  Shared Calendar: " + SharedCalendar);
            } else {
                log.Info("  Mailbox/FolderStore Name: " + MailboxName);
            }
            log.Info("  Calendar: "+ (UseOutlookCalendar.Name=="Calendar"?"Default ":"") + UseOutlookCalendar.Name);
            log.Info("  Category Filter: " + CategoriesRestrictBy.ToString());
            log.Info("  Categories: " + String.Join(",", Categories.ToArray()));
            log.Info("  Only Responded Invites: " + OnlyRespondedInvites);
            log.Info("  Filter String: " + OutlookDateFormat);
            log.Info("  GAL Blocked: " + OutlookGalBlocked);
            
            log.Info("GOOGLE SETTINGS:-");
            log.Info("  Calendar: " + UseGoogleCalendar.Name);
            log.Info("  Personal API Keys: " + UsingPersonalAPIkeys());
            log.Info("    Client Identifier: " + PersonalClientIdentifier);
            log.Info("    Client Secret: " + (PersonalClientSecret.Length < 5
                ? "".PadLeft(PersonalClientSecret.Length, '*')
                : PersonalClientSecret.Substring(0, PersonalClientSecret.Length - 5).PadRight(5, '*')));
            log.Info("  API attendee limit in effect: " + APIlimit_inEffect);
            log.Info("  API attendee limit last reached: " + APIlimit_lastHit);
            log.Info("  Assigned API key: " + AssignedClientIdentifier);
            log.Info("  Cloak Email: " + CloakEmail);
        
            log.Info("SYNC OPTIONS:-");
            log.Info(" How");
            log.Info("  SyncDirection: "+ SyncDirection.Name);
            log.Info("  MergeItems: " + MergeItems);
            log.Info("  DisableDelete: " + DisableDelete);
            log.Info("  ConfirmOnDelete: " + ConfirmOnDelete);
            log.Info("  SetEntriesPrivate: " + SetEntriesPrivate);
            log.Info("  SetEntriesAvailable: " + SetEntriesAvailable);
            log.Info("  SetEntriesColour: " + SetEntriesColour + (SetEntriesColour ? "; " + SetEntriesColourValue + "; \"" + SetEntriesColourName + "\"" : ""));
            if ((SetEntriesPrivate || SetEntriesAvailable || SetEntriesColour) && SyncDirection == Sync.Direction.Bidirectional) {
                log.Info("    TargetCalendar: " + TargetCalendar.Name);
                log.Info("    CreatedItemsOnly: " + CreatedItemsOnly);
            }
            log.Info("  Obfuscate Words: " + Obfuscation.Enabled);
            if (Obfuscation.Enabled) {
                if (Settings.Instance.Obfuscation.FindReplace.Count == 0) log.Info("    No regex defined.");
                else {
                    foreach (FindReplace findReplace in Settings.Instance.Obfuscation.FindReplace) {
                        log.Info("    '" + findReplace.find + "' -> '" + findReplace.replace + "'");
                    }
                }
            }
            log.Info(" When");
            log.Info("  DaysInThePast: "+ DaysInThePast);
            log.Info("  DaysInTheFuture:" + DaysInTheFuture);
            log.Info("  SyncInterval: " + SyncInterval);
            log.Info("  SyncIntervalUnit: " + SyncIntervalUnit);
            log.Info("  Push Changes: " + OutlookPush);
            log.Info(" What");
            log.Info("  AddLocation: " + AddLocation);
            log.Info("  AddDescription: " + AddDescription + "; OnlyToGoogle: " + AddDescription_OnlyToGoogle);
            log.Info("  AddAttendees: " + AddAttendees);
            log.Info("  AddColours: " + AddColours);
            log.Info("  AddReminders: " + AddReminders);
            log.Info("    UseGoogleDefaultReminder: " + UseGoogleDefaultReminder);
            log.Info("    UseOutlookDefaultReminder: " + UseOutlookDefaultReminder);
            log.Info("    ReminderDND: " + ReminderDND + " (" + ReminderDNDstart.ToString("HH:mm") + "-" + ReminderDNDend.ToString("HH:mm") + ")");
            
            log.Info("PROXY:-");
            log.Info("  Type: " + Proxy.Type);
            if (Proxy.BrowserUserAgent != Proxy.DefaultBrowserAgent)
                log.Info("  Browser Agent: " + Proxy.BrowserUserAgent);
            if (Proxy.Type == "Custom") {
                log.Info("  Server Name: " + Proxy.ServerName);
                log.Info("  Port: " + Proxy.Port.ToString());
                log.Info("  Authentication Required: " + Proxy.AuthenticationRequired);
                log.Info("  UserName: " + Proxy.UserName);
                log.Info("  Password: " + (string.IsNullOrEmpty(Proxy.Password) ? "" : "*********"));
            } 
        
            log.Info("APPLICATION BEHAVIOUR:-");
            log.Info("  ShowBubbleTooltipWhenSyncing: " + ShowBubbleTooltipWhenSyncing);
            log.Info("  StartOnStartup: " + StartOnStartup + "; DelayedStartup: "+ StartupDelay.ToString());
            log.Info("  HideSplashScreen: " + ((Subscribed != DateTime.Parse("01-Jan-2000") || Donor) ? HideSplashScreen.ToString() : "N/A"));
            log.Info("  StartInTray: " + StartInTray);
            log.Info("  MinimiseToTray: " + MinimiseToTray);
            log.Info("  MinimiseNotClose: " + MinimiseNotClose);
            log.Info("  ShowBubbleWhenMinimising: " + ShowBubbleWhenMinimising);
            log.Info("  Portable: " + Portable);
            log.Info("  CreateCSVFiles: " + CreateCSVFiles);

            log.Info("  VerboseOutput: " + VerboseOutput);
            log.Info("  MuteClickSounds: " + MuteClickSounds);
            //To pick up from settings.xml file:
            //((log4net.Repository.Hierarchy.Hierarchy)log.Logger.Repository).Root.Level.Name);
            log.Info("  Logging Level: "+ LoggingLevel);
            log.Info("  Error Reporting: " + CloudLogging ?? "Undefined");

            log.Info("ABOUT:-");
            log.Info("  Alpha Releases: " + alphaReleases);
            log.Info("  Skip Version: " + SkipVersion);
            log.Info("  Subscribed: " + Subscribed.ToString("dd-MMM-yyyy"));
            log.Info("  Timezone Database: " + TimezoneDB.Instance.Version);
            
            log.Info("ENVIRONMENT:-");
            log.Info("  Current Locale: " + System.Globalization.CultureInfo.CurrentCulture.Name);
            log.Info("  Short Date Format: "+ System.Globalization.CultureInfo.CurrentCulture.DateTimeFormat.ShortDatePattern);
            log.Info("  Short Time Format: "+ System.Globalization.CultureInfo.CurrentCulture.DateTimeFormat.ShortTimePattern);
            log.Info("  Completed Syncs: "+ CompletedSyncs);
        }

        public static void configureLoggingLevel(string logLevel) {
            log.Info("Logging level configured to '" + logLevel + "'");
            ((log4net.Repository.Hierarchy.Hierarchy)LogManager.GetRepository()).Root.Level = log.Logger.Repository.LevelMap[logLevel];
            ((log4net.Repository.Hierarchy.Hierarchy)LogManager.GetRepository()).RaiseConfigurationChanged(EventArgs.Empty);
        }
    }
}
