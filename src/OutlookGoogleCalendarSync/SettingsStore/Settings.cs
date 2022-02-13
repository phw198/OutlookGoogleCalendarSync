using log4net;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Runtime.Serialization;

namespace OutlookGoogleCalendarSync {
    /// <summary>
    /// The main Settings class.
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
                log.Info("No settings.xml file found in " + Program.MaskFilePath(Program.WorkingFilesDirectory));
                Settings.Instance.Save(ConfigFile);
                log.Info("New blank template created.");
                if (!Program.IsInstalled)
                    XMLManager.ExportElement(Settings.Instance, "Portable", true, ConfigFile);
            }

            log.Info("Running OGCS from " + Program.MaskFilePath(System.Windows.Forms.Application.ExecutablePath));
        }

        private static Settings instance;
        //Settings saved immediately
        private String assignedClientIdentifier;
        private String assignedClientSecret;
        private Boolean apiLimit_inEffect;
        private DateTime apiLimit_lastHit;
        private Int32 completedSyncs;
        private Boolean portable;
        private Boolean alphaReleases;
        private String version;
        private Boolean donor;
        private DateTime subscribed;
        private bool? hideSplashScreen;
        private Boolean suppressSocialPopup;
        private bool? cloudLogging;

        private Settings() {
            Settings.AreLoaded = false;
            Settings.AreApplied = false;
        }

        //Default values before Loading() from xml and attribute not yet serialized
        [OnDeserializing]
        void OnDeserializing(StreamingContext context) {
            SetDefaults();
        }

        public void SetDefaults() {
            //Default values
            assignedClientIdentifier = "";
            assignedClientSecret = "";
            PersonalClientIdentifier = "";
            PersonalClientSecret = "";
            DisconnectOutlookBetweenSync = false;
            TimezoneMaps = new TimezoneMappingDictionary();
            DisconnectOutlookBetweenSync = false;
            TimezoneMaps = new TimezoneMappingDictionary();

            apiLimit_inEffect = false;
            apiLimit_lastHit = DateTime.Parse("01-Jan-2000");
            GaccountEmail = "";

            Calendars = new List<SettingsStore.Calendar>();
            Calendars.Add(new SettingsStore.Calendar());

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
            cloudLogging = null;
            portable = false;
            Proxy = new SettingsStore.Proxy();

            alphaReleases = !System.Windows.Forms.Application.ProductVersion.EndsWith("0.0");
            SkipVersion = null;
            subscribed = GoogleOgcs.Authenticator.SubscribedNever;
            donor = false;
            hideSplashScreen = null;
            suppressSocialPopup = false;

            completedSyncs = 0;
            VerboseOutput = true;
        }

        public static Boolean InstanceInitialiased() {
            return (instance != null);
        }

        public static Settings Instance {
            get {
                if (instance == null) {
                    instance = new Settings();
                    instance.SetDefaults();
                }
                return instance;
            }
            set {
                instance = value;
            }
        }
        
        #region Outlook
        [DataMember] public Boolean DisconnectOutlookBetweenSync { get; set; }
        [DataMember] public TimezoneMappingDictionary TimezoneMaps { get; private set; }
        [CollectionDataContract(
            ItemName = "TimeZoneMap",
            KeyName = "OrganiserTz",
            ValueName = "SystemTz",
            Namespace = "http://schemas.datacontract.org/2004/07/OutlookGoogleCalendarSync"
        )]
        public class TimezoneMappingDictionary : Dictionary<String, String> { }
        #endregion
        #region Google
        [DataMember] public String AssignedClientIdentifier {
            get { return assignedClientIdentifier; }
            set {
                assignedClientIdentifier = value.Trim();
                if (!Loading()) XMLManager.ExportElement(this, "AssignedClientIdentifier", value.Trim(), ConfigFile);
            }
        }
        [DataMember] public String AssignedClientSecret {
            get { return assignedClientSecret; }
            set {
                assignedClientSecret = value.Trim();
                if (!Loading()) XMLManager.ExportElement(this, "AssignedClientSecret", value.Trim(), ConfigFile);
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
        [DataMember] public Boolean APIlimit_inEffect {
            get { return apiLimit_inEffect; }
            set {
                apiLimit_inEffect = value;
                if (!Loading()) XMLManager.ExportElement(this, "APIlimit_inEffect", value, ConfigFile);
            }
        }
        [DataMember] public DateTime APIlimit_lastHit {
            get { return apiLimit_lastHit; }
            set {
                apiLimit_lastHit = value;
                if (!Loading()) XMLManager.ExportElement(this, "APIlimit_lastHit", value, ConfigFile);
            }
        }
        [DataMember] public String GaccountEmail { get; set; }
        public String GaccountEmail_masked() {
            if (string.IsNullOrWhiteSpace(GaccountEmail)) return "<null>";
            return EmailAddress.MaskAddress(GaccountEmail);
        }
        #endregion
        #region App behaviour
        [DataMember] public bool? HideSplashScreen {
            get { return hideSplashScreen; }
            set {
                if (!Loading() && hideSplashScreen != value) {
                    XMLManager.ExportElement(this, "HideSplashScreen", value, ConfigFile);
                    if (Forms.Main.Instance != null) Forms.Main.Instance.cbHideSplash.Checked = value ?? false;
                }
                hideSplashScreen = value;
            }
        }

        [DataMember] public bool SuppressSocialPopup {
            get { return suppressSocialPopup; }
            set {
                if (!Loading() && suppressSocialPopup != value) {
                    XMLManager.ExportElement(this, "SuppressSocialPopup", value, ConfigFile);
                    if (Forms.Main.Instance != null) Forms.Main.Instance.cbSuppressSocialPopup.Checked = value;
                }
                suppressSocialPopup = value;
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
                if (!Loading()) XMLManager.ExportElement(this, "Portable", value, ConfigFile);
            }
        }

        [DataMember] public bool CreateCSVFiles { get; set; }
        [DataMember] public String LoggingLevel { get; set; }
        [DataMember] public bool? CloudLogging {
            get { return cloudLogging; }
            set {
                cloudLogging = value;
                GoogleOgcs.ErrorReporting.SetThreshold(value ?? false);
                if (value == null) GoogleOgcs.ErrorReporting.ErrorOccurred = false;
                if (!Loading()) XMLManager.ExportElement(this, "CloudLogging", value, ConfigFile);
            }
        }
        [DataMember] public bool TelemetryDisabled { get; set; }
        //Proxy
        [DataMember] public SettingsStore.Proxy Proxy { get; set; }
        [DataMember] public List<SettingsStore.Calendar> Calendars { get; set; }
        #endregion
        #region About
        [DataMember] public string Version {
            get { return version; }
            set {
                if (version != null && version != value) {
                    XMLManager.ExportElement(this, "Version", value, ConfigFile);
                }
                version = value;
            }
        }
        [DataMember] public bool AlphaReleases {
            get { return alphaReleases; }
            set {
                alphaReleases = value;
                if (!Loading()) XMLManager.ExportElement(this, "AlphaReleases", value, ConfigFile);
            }
        }
        public Boolean UserIsBenefactor() {
            return Subscribed != GoogleOgcs.Authenticator.SubscribedNever || donor;
        }
        [DataMember] public DateTime Subscribed {
            get { return subscribed; }
            set {
                subscribed = value;
                if (!Loading()) XMLManager.ExportElement(this, "Subscribed", value, ConfigFile);
            }
        }
        [DataMember] public Boolean Donor {
            get { return donor; }
            set {
                donor = value;
                if (!Loading()) XMLManager.ExportElement(this, "Donor", value, ConfigFile);
            }
        }
        #endregion

        [DataMember] public Int32 CompletedSyncs {
            get { return completedSyncs; }
            set {
                completedSyncs = value;
                if (!Loading()) XMLManager.ExportElement(this, "CompletedSyncs", value, ConfigFile);
            }
        }
        [DataMember] public bool VerboseOutput { get; set; }
        [DataMember] public bool MuteClickSounds { get; set; }
        [DataMember] public String SkipVersion { get; set; }

        public static Boolean AreLoaded { get; protected set; }

        /// <summary>
        /// The settings file has been loaded and configuration applied
        /// </summary>
        public static Boolean AreApplied { get; set; }

        /// <summary>
        /// Load all OGCS settings as defined in the configuration file.
        /// </summary>
        public static void Load(String XMLfile = null) {
            try {
                Settings.Instance = XMLManager.Import<Settings>(XMLfile ?? ConfigFile);
                log.Fine("User settings loaded.");
                Settings.AreLoaded = true;

            } catch (ApplicationException ex) {
                log.Error("Failed to load settings file '" + (XMLfile ?? ConfigFile) + "'. " + ex.Message);
                ResetFile(XMLfile ?? ConfigFile);
                try {
                    Settings.Instance = XMLManager.Import<Settings>(XMLfile ?? ConfigFile);
                    log.Debug("User settings loaded successfully this time.");
                } catch (System.Exception ex2) {
                    log.Error("Still failed to load settings!");
                    OGCSexception.Analyse(ex2);
                }
            }
        }

        public static void ResetFile(String XMLfile = null) {
            System.Windows.Forms.OgcsMessageBox.Show("Your OGCS settings appear to be corrupt and will have to be reset.",
                    "Corrupt OGCS Settings", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Exclamation);
            log.Warn("Resetting settings.xml file to defaults.");
            System.IO.File.Delete(XMLfile ?? ConfigFile);
            Settings.Instance.Save(XMLfile ?? ConfigFile);
        }

        public void Save(String XMLfile = null) {
            log.Info("Saving settings.");
            XMLManager.Export(this, XMLfile ?? ConfigFile);
        }

        public Boolean Loading() {
            return Program.CalledByProcess("Load,isNewVersion");
        }

        public void LogSettings() {
            log.Info(Program.MaskFilePath(ConfigFile));
            log.Info("OUTLOOK SETTINGS:-");
            log.Info("  Disconnect Between Sync: " + DisconnectOutlookBetweenSync);
            if (TimezoneMaps.Count > 0) {
                log.Info("  Custom Timezone Mapping:-");
                TimezoneMaps.ToList().ForEach(tz => log.Info("    " + tz.Key + " => " + tz.Value));
            }
            log.Info("GOOGLE SETTINGS:-");
            log.Info("  Personal API Keys: " + UsingPersonalAPIkeys());
            log.Info("    Client Identifier: " + PersonalClientIdentifier);
            log.Info("    Client Secret: " + (PersonalClientSecret.Length < 5
                ? "".PadLeft(PersonalClientSecret.Length, '*')
                : PersonalClientSecret.Substring(0, PersonalClientSecret.Length - 5).PadRight(5, '*')));
            log.Info("  API attendee limit in effect: " + APIlimit_inEffect);
            log.Info("  API attendee limit last reached: " + APIlimit_lastHit);
            log.Info("  Assigned API key: " + AssignedClientIdentifier);
        
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
            log.Info("  HideSplashScreen: " + (UserIsBenefactor() ? HideSplashScreen.ToString() : "N/A"));
            log.Info("  SuppressSocialPopup: " + (UserIsBenefactor() ? SuppressSocialPopup.ToString() : "N/A"));
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
            TimeZone curTimeZone = TimeZone.CurrentTimeZone;
            log.Info("  System Time Zone: " + curTimeZone.StandardName + "; DST=" + curTimeZone.IsDaylightSavingTime(DateTime.Now));
            log.Info("  Completed Syncs: "+ CompletedSyncs);
        }

        public static void configureLoggingLevel(string logLevel) {
            log.Info("Logging level configured to '" + logLevel + "'");
            ((log4net.Repository.Hierarchy.Hierarchy)LogManager.GetRepository()).Root.Level = log.Logger.Repository.LevelMap[logLevel];
            ((log4net.Repository.Hierarchy.Hierarchy)LogManager.GetRepository()).RaiseConfigurationChanged(EventArgs.Empty);
        }

        /// <summary>
        /// Deregister all profiles from Push Sync
        /// </summary>
        public void DeregisterAllForPushSync() {
            foreach (SettingsStore.Calendar calendar in Settings.Instance.Calendars) {
                calendar.DeregisterForPushSync();
            }
        }

        public class Profile {
            public enum Type {
                Calendar,
                Global,
                None,
                Unknown
            }
            public static Type GetType(Object settingsStore) {
                if (settingsStore == null) return Type.None;

                switch (settingsStore?.GetType().ToString()) {
                    case "OutlookGoogleCalendarSync.Settings": return Type.Global;
                    case "OutlookGoogleCalendarSync.SettingsStore.Calendar": return Type.Calendar;
                }
                log.Warn("Unknown profile type: " + settingsStore?.GetType().ToString());
                return Type.Unknown;
            }

            public static String Name(Object settingsStore) {
                Type settingsStoreType = GetType(settingsStore);
                switch (settingsStoreType) {
                    case Type.Calendar: return (settingsStore as SettingsStore.Calendar)._ProfileName;
                    default: return settingsStoreType.ToString();
                }
            }

            /// <summary>
            /// Dynamically determine which profile is being used.
            /// </summary>
            /// <returns>Currently hard-coded to a Calendar profile</returns>
            public static SettingsStore.Calendar InPlay() {
                SettingsStore.Calendar aProfile;

                if (Program.CalledByProcess("manualSynchronize,Sync_Click,updateGUIsettings,UpdateGUIsettings_Profile,miCatRefresh_Click," +
                    "GetMyGoogleCalendars_Click,btColourMap_Click,btTestOutlookFilter_Click,ColourPicker_Enter,OnSelectedIndexChanged,OnCheckedChanged")) {
                    aProfile = Forms.Main.Instance.ActiveCalendarProfile;
                    log.Fine("Using profile Forms.Main.Instance.ActiveCalendarProfile");
                
                } else if (Program.CalledByProcess("synchronize,OnTick")) {
                    aProfile = Sync.Engine.Calendar.Instance.Profile;
                    log.Fine("Using profile Sync.Engine.Calendar.Instance.Profile");

                } else {
                    Program.StackTraceToString();
                    log.Error("Unknown profile being referenced.");
                    aProfile = Forms.Main.Instance.ActiveCalendarProfile;
                }
                if (aProfile == null) log.Warn("The profile in play is NULL!");
                return aProfile;
            }
        }
    }
}
