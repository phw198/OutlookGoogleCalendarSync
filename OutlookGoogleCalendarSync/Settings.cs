using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Reflection;
using System.Runtime.Serialization;
using log4net;

namespace OutlookGoogleCalendarSync {
    /// <summary>
    /// Description of Settings.
    /// </summary>
    
    [DataContract]
    public class Settings {
        private static readonly ILog log = LogManager.GetLogger(typeof(Settings));
        private static Settings instance;
        //Settings saved immediately
        private Boolean apiLimit_inEffect;
        private DateTime apiLimit_lastHit;
        private DateTime syncStart;
        private DateTime syncEnd;
        private DateTime lastSyncDate;
        private Int32 completedSyncs;
        private Boolean portable;
        private Boolean alphaReleases;
        private String version;

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
            OutlookService = OutlookCalendar.Service.DefaultMailbox;
            MailboxName = "";
            EWSuser = "";
            EWSpassword = "";
            EWSserver = "";
            UseOutlookCalendar = new MyOutlookCalendarListEntry();
            OutlookDateFormat = "g";

            UseGoogleCalendar = new MyGoogleCalendarListEntry();
            RefreshToken = "";
            apiLimit_inEffect = false;
            apiLimit_lastHit = DateTime.Parse("01-Jan-2000");

            SyncDirection = new SyncDirection();
            DaysInThePast = 1;
            DaysInTheFuture = 60;
            SyncInterval = 0;
            SyncIntervalUnit = "Hours";
            OutlookPush = false;
            AddDescription = true;
            AddDescription_OnlyToGoogle = true;
            AddReminders = false;
            AddAttendees = true;
            MergeItems = true;
            DisableDelete = true;
            ConfirmOnDelete = true;
            Obfuscation = new Obfuscate();

            ShowBubbleTooltipWhenSyncing = true;
            StartOnStartup = false;
            StartInTray = false;
            MinimiseToTray = false;
            MinimiseNotClose = false;
            ShowBubbleWhenMinimising = true;

            CreateCSVFiles = false;
            LoggingLevel = "DEBUG";
            portable = false;
            Proxy = new SettingsProxy();

            alphaReleases = false;
            
            lastSyncDate = new DateTime(0);
            completedSyncs = 0;
            VerboseOutput = false;
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
        [DataMember] public OutlookCalendar.Service OutlookService { get; set; }
        [DataMember] public string MailboxName { get; set; }
        [DataMember] public string EWSuser { get; set; }
        [DataMember] public string EWSpassword { get; set; }
        [DataMember] public string EWSserver { get; set; }
        [DataMember] public MyOutlookCalendarListEntry UseOutlookCalendar { get; set; }
        [DataMember] public string OutlookDateFormat { get; set; }
        #endregion
        #region Google
        [DataMember] public MyGoogleCalendarListEntry UseGoogleCalendar { get; set; }
        [DataMember] public string RefreshToken { get; set; }
        [DataMember] public Boolean APIlimit_inEffect {
            get { return apiLimit_inEffect; }
            set {
                apiLimit_inEffect = value;
                if (!loading()) XMLManager.ExportElement("APIlimit_inEffect", value, Program.SettingsFile);
            } 
        }
        [DataMember] public DateTime APIlimit_lastHit { 
            get { return apiLimit_lastHit; }
            set { 
                apiLimit_lastHit = value;
                if (!loading()) XMLManager.ExportElement("APIlimit_lastHit", value, Program.SettingsFile);
            }
        }
        #endregion
        #region Sync Options
        //Main
        private int daysInThePast;
        private int daysInTheFuture;
        public DateTime SyncStart { get { return this.syncStart; } }
        public DateTime SyncEnd { get { return this.syncEnd; } }
        [DataMember] public SyncDirection SyncDirection { get; set; }
        [DataMember] public int DaysInThePast {
            get { return daysInThePast; }
            set {
                this.daysInThePast = value;
                this.syncStart = DateTime.Today.AddDays(-value); 
            } 
        }
        [DataMember] public int DaysInTheFuture {
            get { return daysInTheFuture; }
            set {
                this.daysInTheFuture = value; 
                this.syncEnd = DateTime.Today.AddDays(+value + 1);
            }
        }
        [DataMember] public int SyncInterval { get; set; }
        [DataMember] public String SyncIntervalUnit { get; set; }
        [DataMember] public bool OutlookPush { get; set; }
        [DataMember] public bool AddDescription { get; set; }
        [DataMember] public bool AddDescription_OnlyToGoogle { get; set; }
        [DataMember] public bool AddReminders { get; set; }
        [DataMember] public bool AddAttendees { get; set; }
        [DataMember] public bool MergeItems { get; set; }
        [DataMember] public bool DisableDelete { get; set; }
        [DataMember] public bool ConfirmOnDelete { get; set; }

        //Obfuscation
        [DataMember] public Obfuscate Obfuscation { get; set; }

        #endregion
        #region App behaviour
        [DataMember] public bool ShowBubbleTooltipWhenSyncing { get; set; }
        [DataMember] public bool StartOnStartup { get; set; }
        [DataMember] public bool StartInTray { get; set; }
        [DataMember] public bool MinimiseToTray { get; set; }
        [DataMember] public bool MinimiseNotClose { get; set; }
        [DataMember] public bool ShowBubbleWhenMinimising { get; set; }
        [DataMember] public bool Portable {
            get { return portable; }
            set {
                portable = value;
                if (!loading()) XMLManager.ExportElement("Portable", value, Program.SettingsFile);
            }
        }

        [DataMember] public bool CreateCSVFiles { get; set; }
        [DataMember] public String LoggingLevel { get; set; }
        //Proxy
        [DataMember] public SettingsProxy Proxy { get; set; }
        #endregion
        #region About
        [DataMember] public string Version {
            get { return version; }
            set {
                version = value;
                if (!loading()) XMLManager.ExportElement("Version", value, Program.SettingsFile);
            }
        }
        [DataMember] public bool AlphaReleases {
            get { return alphaReleases; }
            set {
                alphaReleases = value;
                if (!loading()) XMLManager.ExportElement("AlphaReleases", value, Program.SettingsFile);
            }
        }
        #endregion

        [DataMember] public DateTime LastSyncDate {
            get { return lastSyncDate; }
            set {
                lastSyncDate = value;
                if (!loading()) XMLManager.ExportElement("LastSyncDate", value, Program.SettingsFile);
            }
        }
        [DataMember] public Int32 CompletedSyncs {
            get { return completedSyncs; }
            set {
                completedSyncs = value;
                if (!loading()) XMLManager.ExportElement("CompletedSyncs", value, Program.SettingsFile);
            }
        }
        [DataMember] public bool VerboseOutput { get; set; }

        public static void Load(string XMLfile = null) {
            Settings.Instance = XMLManager.Import<Settings>(XMLfile ?? Program.SettingsFile);
            log.Fine("User settings loaded.");
        }

        public void Save(string XMLfile = null) {
            XMLManager.Export(this, XMLfile ?? Program.SettingsFile);
        }

        private Boolean loading() {
            StackTrace stackTrace = new StackTrace();
            foreach (StackFrame frame in stackTrace.GetFrames().Reverse()) {
                if (frame.GetMethod().Name == "Load") {
                    return true;
                }
            }
            return false;
        }

        public void LogSettings() {
            log.Info(Program.SettingsFile);
            log.Info("OUTLOOK SETTINGS:-");
            log.Info("  Service: "+ OutlookService.ToString());
            log.Info("  Calendar: "+ (UseOutlookCalendar.Name=="Calendar"?"Default ":"") + UseOutlookCalendar.Name);
            log.Info("  Filter String: " + OutlookDateFormat);
            
            log.Info("GOOGLE SETTINGS:-");
            log.Info("  Calendar: "+ UseGoogleCalendar.Name);
            log.Info("  API attendee limit in effect: " + APIlimit_inEffect);
            log.Info("  API attendee limit last reached: " + APIlimit_lastHit);
        
            log.Info("SYNC OPTIONS:-");
            log.Info(" Main");
            log.Info("  SyncDirection: "+ SyncDirection.Name);
            log.Info("  DaysInThePast: "+ DaysInThePast);
            log.Info("  DaysInTheFuture:" + DaysInTheFuture);
            log.Info("  SyncInterval: " + SyncInterval);
            log.Info("  SyncIntervalUnit: " + SyncIntervalUnit);
            log.Info("  Push Changes: " + OutlookPush);
            log.Info("  AddDescription: " + AddDescription + "; OnlyToGoogle: " + AddDescription_OnlyToGoogle);
            log.Info("  AddReminders: " + AddReminders);
            log.Info("  AddAttendees: " + AddAttendees);
            log.Info("  MergeItems: " + MergeItems);
            log.Info("  DisableDelete: " + DisableDelete);
            log.Info("  ConfirmOnDelete: " + ConfirmOnDelete);
            log.Info("  Obfuscate Words: "+ Obfuscation.Enabled);
            if (Obfuscation.Enabled) {
                if (Settings.Instance.Obfuscation.FindReplace.Count == 0) log.Info("    No regex defined.");
                else {
                    foreach (FindReplace findReplace in Settings.Instance.Obfuscation.FindReplace) {
                        log.Info("    '" + findReplace.find + "' -> '" + findReplace.replace + "'");
                    }
                }
            }

            log.Info("PROXY:-");
            log.Info("  Type: " + Proxy.Type);
            if (Proxy.Type == "Custom") {
                log.Info("  Server Name: " + Proxy.ServerName);
                log.Info("  Port: " + Proxy.Port.ToString());
                log.Info("  UserName: " + Proxy.UserName);
                log.Info("  Password: " + (string.IsNullOrEmpty(Proxy.Password) ? "" : "*********"));
            } 
        
            log.Info("APPLICATION BEHAVIOUR:-");
            log.Info("  ShowBubbleTooltipWhenSyncing: " + ShowBubbleTooltipWhenSyncing);
            log.Info("  StartOnStartup: " + StartOnStartup);
            log.Info("  StartInTray: " + StartInTray);
            log.Info("  MinimiseToTray: " + MinimiseToTray);
            log.Info("  MinimiseNotClose: " + MinimiseNotClose);
            log.Info("  ShowBubbleWhenMinimising: " + ShowBubbleWhenMinimising);
            log.Info("  Portable: " + Portable);
            log.Info("  CreateCSVFiles: " + CreateCSVFiles);

            log.Info("  VerboseOutput: " + VerboseOutput);
            //To pick up from settings.xml file:
            //((log4net.Repository.Hierarchy.Hierarchy)log.Logger.Repository).Root.Level.Name);
            log.Info("  Logging Level: "+ LoggingLevel);

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
