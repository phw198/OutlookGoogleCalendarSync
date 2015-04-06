using System;
using log4net;

namespace OutlookGoogleCalendarSync {
    /// <summary>
    /// Description of Settings.
    /// </summary>
    public class Settings {
        private static readonly ILog log = LogManager.GetLogger(typeof(Settings));
        private static Settings instance;

        public static Settings Instance {
            get {
                if (instance == null) instance = new Settings();
                return instance;
            }
            set {
                instance = value;
            }

        }
        //Outlook
        public OutlookCalendar.Service OutlookService = OutlookCalendar.Service.DefaultMailbox;
        public string MailboxName = "";
        public string EWSuser = "";
        public string EWSpassword = "";
        public string EWSserver = "";
        public MyOutlookCalendarListEntry UseOutlookCalendar = new MyOutlookCalendarListEntry();

        //Google
        public MyGoogleCalendarListEntry UseGoogleCalendar = new MyGoogleCalendarListEntry();
        public string RefreshToken = "";
        
        //Sync Options
        public SyncDirection SyncDirection = new SyncDirection();
        public int DaysInThePast = 1;
        public int DaysInTheFuture = 60;
        public int SyncInterval = 0;
        public String SyncIntervalUnit = "Hours";
        public bool OutlookPush = false;
        public bool AddDescription = true;
        public bool AddReminders = false;
        public bool AddAttendees = true;
        public bool MergeItems = true;
        public bool DisableDelete = true;
        public bool ConfirmOnDelete = true;

        //Proxy
        public SettingsProxy Proxy = new SettingsProxy();
        
        //App behaviour
        public bool ShowBubbleTooltipWhenSyncing = false;
        public bool StartOnStartup = true;
        public bool StartInTray = false;
        public bool MinimizeToTray = false;
        public bool CreateCSVFiles = true;
        public String LoggingLevel = "DEBUG";

        public DateTime LastSyncDate = new DateTime(0);

        public bool VerboseOutput = false;

        public Settings() {
        }

        public void LogSettings() {
            log.Info(Program.SettingsFile);
            log.Info("OUTLOOK SETTINGS:-");
            log.Info("  Service: "+ OutlookService.ToString());
            log.Info("  Calendar: "+ (UseOutlookCalendar.Name=="Calendar"?"Default ":"") + UseOutlookCalendar.Name);
            
            log.Info("GOOGLE SETTINGS:-");
            log.Info("  Calendar: "+ UseGoogleCalendar.Name);
        
            log.Info("SYNC OPTIONS:-");
            log.Info("  SyncDirection: "+ SyncDirection.Name);
            log.Info("  DaysInThePast: "+ DaysInThePast);
            log.Info("  DaysInTheFuture:" + DaysInTheFuture);
            log.Info("  SyncInterval: " + SyncInterval);
            log.Info("  SyncIntervalUnit: " + SyncIntervalUnit);
            log.Info("  Push Changes: " + OutlookPush);
            log.Info("  AddDescription: " + AddDescription);
            log.Info("  AddReminders: " + AddReminders);
            log.Info("  AddAttendees: " + AddAttendees);
            log.Info("  MergeItems: " + MergeItems);
            log.Info("  DisableDelete: " + DisableDelete);
            log.Info("  ConfirmOnDelete: " + ConfirmOnDelete);

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
            log.Info("  MinimizeToTray: " + MinimizeToTray);
            log.Info("  CreateCSVFiles: " + CreateCSVFiles);

            log.Info("  VerboseOutput: " + VerboseOutput);
            //To pick up from settings.xml file:
            //((log4net.Repository.Hierarchy.Hierarchy)log.Logger.Repository).Root.Level.Name);
            log.Info("  Logging Level: "+ LoggingLevel);

            log.Info("ENVIRONMENT:-");
            log.Info("  Current Locale: " + System.Globalization.CultureInfo.CurrentCulture.Name);
            log.Info("  Short Date Format: "+ System.Globalization.CultureInfo.CurrentCulture.DateTimeFormat.ShortDatePattern);
            log.Info("  Short Time Format: "+ System.Globalization.CultureInfo.CurrentCulture.DateTimeFormat.ShortTimePattern);
        }
    }
}
