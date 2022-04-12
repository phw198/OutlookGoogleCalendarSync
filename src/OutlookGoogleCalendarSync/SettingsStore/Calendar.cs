using log4net;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.Serialization;

namespace OutlookGoogleCalendarSync.SettingsStore {
    [DataContract(Namespace = "http://schemas.datacontract.org/2004/07/OutlookGoogleCalendarSync")]
    public class Calendar {
        private static readonly ILog log = LogManager.GetLogger(typeof(Calendar));

        public Sync.SyncTimer OgcsTimer;
        public Sync.PushSyncTimer OgcsPushTimer;

        //Settings saved immediately
        private DateTime lastSyncDate;

        public Calendar() {
            setDefaults();
        }

        public override String ToString() {
            return this._ProfileName + ": O[" + this.UseOutlookCalendar.Name + "] " + this.SyncDirection.ToString() + " G[" + this.UseGoogleCalendar.ToString() + "]";
        }

        //Default values before loading from xml and attribute not yet serialized
        [OnDeserializing]
        void OnDeserializing(StreamingContext context) {
            setDefaults();
        }

        private void setDefaults() {
            _ProfileName = "Default";

            //Outlook
            OutlookService = OutlookOgcs.Calendar.Service.DefaultMailbox;
            MailboxName = "";
            SharedCalendar = "";
            UseOutlookCalendar = new OutlookCalendarListEntry();
            CategoriesRestrictBy = RestrictBy.Exclude;
            Categories = new List<String>();
            OnlyRespondedInvites = false;
            OutlookDateFormat = "g";
            outlookGalBlocked = false;

            //Google
            UseGoogleCalendar = new GoogleCalendarListEntry();
            CloakEmail = true;
            ExcludeDeclinedInvites = true;
            ExcludeGoals = true;

            //Sync Options
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
            MaxAttendees = 200;
            AddColours = false;
            MergeItems = true;
            DisableDelete = true;
            ConfirmOnDelete = true;
            TargetCalendar = Sync.Direction.OutlookToGoogle;
            CreatedItemsOnly = true;
            SetEntriesPrivate = false;
            SetEntriesAvailable = false;
            AvailabilityStatus = Microsoft.Office.Interop.Outlook.OlBusyStatus.olFree.ToString();
            SetEntriesColour = false;
            SetEntriesColourValue = Microsoft.Office.Interop.Outlook.OlCategoryColor.olCategoryColorNone.ToString();
            SetEntriesColourName = "None";
            SetEntriesColourGoogleId = "0";
            ColourMaps = new ColourMappingDictionary();
            Obfuscation = new Obfuscate();
            
            ExtirpateOgcsMetadata = false;
            lastSyncDate = new DateTime(0);
        }

        [DataMember] public string _ProfileName { get; set; }

        #region Outlook
        public enum RestrictBy {
            Include, Exclude
        }
        [DataMember] public OutlookOgcs.Calendar.Service OutlookService { get; set; }
        [DataMember] public string MailboxName { get; set; }
        [DataMember] public string SharedCalendar { get; set; }
        [DataMember] public OutlookCalendarListEntry UseOutlookCalendar { get; set; }
        [DataMember] public RestrictBy CategoriesRestrictBy { get; set; }
        [DataMember] public List<string> Categories { get; set; }
        /// <summary>Only allow Outlook to have one category assigned</summary>
        [DataMember] public Boolean SingleCategoryOnly { get; set; }
        [DataMember] public Boolean OnlyRespondedInvites { get; set; }
        [DataMember] public string OutlookDateFormat { get; set; }
        private Boolean outlookGalBlocked;
        [DataMember] public Boolean OutlookGalBlocked {
            get { return outlookGalBlocked; }
            set {
                outlookGalBlocked = value;
                if (!Settings.Instance.Loading() && Forms.Main.Instance.IsHandleCreated) Forms.Main.Instance.FeaturesBlockedByCorpPolicy(value);
            }
        }
        #endregion
        #region Google
        [DataMember] public GoogleCalendarListEntry UseGoogleCalendar { get; set; }
        [DataMember] public Boolean CloakEmail { get; set; }
        [DataMember] public Boolean ExcludeDeclinedInvites { get; set; }
        [DataMember] public Boolean ExcludeGoals { get; set; }
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
        [DataMember] public int MaxAttendees { get; set; }
        [DataMember] public bool AddColours { get; set; }
        [DataMember] public bool MergeItems { get; set; }
        [DataMember] public bool DisableDelete { get; set; }
        [DataMember] public bool ConfirmOnDelete { get; set; }
        [DataMember] public Sync.Direction TargetCalendar { get; set; }
        [DataMember] public Boolean CreatedItemsOnly { get; set; }
        [DataMember] public bool SetEntriesPrivate { get; set; }
        [DataMember] public bool SetEntriesAvailable { get; set; }
        [DataMember] public String AvailabilityStatus { get; set; }
        [DataMember] public bool SetEntriesColour { get; set; }

        /// <summary>Set all Outlook appointments to this OlCategoryColor</summary>
        [DataMember] public String SetEntriesColourValue { get; set; }
        /// <summary>Set all Outlook appointments to this custom category name</summary>
        [DataMember] public String SetEntriesColourName { get; set; }
        /// <summary>Set all Google events to this colour ID</summary>
        [DataMember] public String SetEntriesColourGoogleId { get; set; }
        [DataMember]
        public ColourMappingDictionary ColourMaps { get; private set; }
        [CollectionDataContract(
            ItemName = "ColourMap",
            KeyName = "OutlookCategoryName",
            ValueName = "GoogleColourId",
            Namespace = "http://schemas.datacontract.org/2004/07/OutlookGoogleCalendarSync"
        )]
        public class ColourMappingDictionary : Dictionary<String, String> { }

        //Obfuscation
        [DataMember] public Obfuscate Obfuscation { get; set; }
        #endregion

        #region Advanced - Non GUI
        [DataMember] public Boolean ExtirpateOgcsMetadata { get; private set; }
        #endregion

        [DataMember] public DateTime LastSyncDate {
            get { return lastSyncDate; }
            set {
                lastSyncDate = value;
                if (!Settings.Instance.Loading()) {
                    XMLManager.ExportElement(this, "LastSyncDate", value, Settings.ConfigFile);
                    if (Forms.Main.Instance.ProfileVal == this._ProfileName)
                        Forms.Main.Instance.LastSyncVal = this.LastSyncDateText;
                }
            }
        }

        public String LastSyncDateText {
            get { return lastSyncDate.ToLongDateString() + " @ " + lastSyncDate.ToLongTimeString(); }
        }

        /// <summary>
        /// Make this calendar profile display settings in GUI
        /// </summary>
        public void SetActive() {
            if (Forms.Main.Instance.ActiveCalendarProfile != null &&
                Forms.Main.Instance.ActiveCalendarProfile == this) return;

            log.Debug("Changing active settings profile: " + this._ProfileName);
            Forms.Main.Instance.ActiveCalendarProfile = this;

            if (Forms.Main.Instance.Visible) 
                Forms.Main.Instance?.UpdateGUIsettings_Profile();
        }

        public void InitialiseTimer() {
            log.Debug("Creating the calendar timer for auto synchronisation on profile: " + this._ProfileName);
            OgcsTimer = new Sync.SyncTimer(this);
        }

        #region Push Sync
        public void RegisterForPushSync() {
            if (!this.OutlookPush || this.SyncDirection.Id == Sync.Direction.GoogleToOutlook.Id) return;

            log.Info("Start monitoring for Outlook appointments changes on profile: " + this._ProfileName);
            if (this.OgcsPushTimer == null)
                this.OgcsPushTimer = new Sync.PushSyncTimer(this);
            if (!this.OgcsPushTimer.Running())
                this.OgcsPushTimer.Activate(true);
        }

        public void DeregisterForPushSync() {
            log.Info("Stop monitoring for Outlook appointment changes on profile: " + this._ProfileName);
            if (this.OgcsPushTimer != null && this.OgcsPushTimer.Running())
                this.OgcsPushTimer.Activate(false);
        }
        #endregion

        public void LogSettings() {
            log.Info("CALENDAR SYNC SETTINGS");
            log.Info("Profile: " + _ProfileName);
            log.Info("Last Synced: " + LastSyncDate);

            log.Info("OUTLOOK SETTINGS:-");
            log.Info("  Service: " + OutlookService.ToString());
            if (OutlookService == OutlookOgcs.Calendar.Service.SharedCalendar) {
                log.Info("  Shared Calendar: " + SharedCalendar);
            } else {
                log.Info("  Mailbox/FolderStore Name: " + MailboxName);
            }
            log.Info("  Calendar: " + (UseOutlookCalendar.Name == "Calendar" ? "Default " : "") + UseOutlookCalendar.ToString());
            log.Info("  Category Filter: " + CategoriesRestrictBy.ToString());
            log.Info("  Categories: " + String.Join(",", Categories.ToArray()));
            log.Info("  Only Responded Invites: " + OnlyRespondedInvites);
            log.Info("  Filter String: " + OutlookDateFormat);
            log.Info("  GAL Blocked: " + OutlookGalBlocked);

            log.Info("GOOGLE SETTINGS:-");
            log.Info("  Calendar: " + (UseGoogleCalendar == null ? "" : UseGoogleCalendar.ToString(true)));
            log.Info("  Exclude Declined Invites: " + ExcludeDeclinedInvites);
            log.Info("  Exclude Goals: " + ExcludeGoals);
            log.Info("  Cloak Email: " + CloakEmail);

            log.Info("SYNC OPTIONS:-");
            log.Info(" How");
            log.Info("  SyncDirection: " + SyncDirection.Name);
            log.Info("  MergeItems: " + MergeItems);
            log.Info("  DisableDelete: " + DisableDelete);
            log.Info("  ConfirmOnDelete: " + ConfirmOnDelete);
            log.Info("  SetEntriesPrivate: " + SetEntriesPrivate);
            log.Info("  SetEntriesAvailable: " + SetEntriesAvailable + (SetEntriesAvailable ? "; " + AvailabilityStatus : ""));
            log.Info("  SetEntriesColour: " + SetEntriesColour + (SetEntriesColour ? "; " + SetEntriesColourValue + "; \"" + SetEntriesColourName + "\"" : ""));
            if ((SetEntriesPrivate || SetEntriesAvailable || SetEntriesColour) && SyncDirection.Id == Sync.Direction.Bidirectional.Id) {
                log.Info("    TargetCalendar: " + TargetCalendar.Name);
                log.Info("    CreatedItemsOnly: " + CreatedItemsOnly);
            }
            if (ColourMaps.Count > 0) {
                log.Info("  Custom Colour/Category Mapping:-");
                if (OutlookOgcs.Factory.OutlookVersionName == OutlookOgcs.Factory.OutlookVersionNames.Outlook2003)
                    log.Fail("    Using Outlook2003 - categories not supported, although mapping exists");
                else
                    ColourMaps.ToList().ForEach(c => log.Info("    " + OutlookOgcs.Calendar.Categories.OutlookColour(c.Key) + ":" + c.Key + " <=> " +
                        c.Value + ":" + GoogleOgcs.EventColour.Palette.GetColourName(c.Value)));
            }
            log.Info("  SingleCategoryOnly: " + SingleCategoryOnly);
            log.Info("  Obfuscate Words: " + Obfuscation.Enabled);
            if (Obfuscation.Enabled) {
                if (Obfuscation.FindReplace.Count == 0) log.Info("    No regex defined.");
                else {
                    foreach (FindReplace findReplace in Obfuscation.FindReplace) {
                        log.Info("    '" + findReplace.find + "' -> '" + findReplace.replace + "'");
                    }
                }
            }
            log.Info(" When");
            log.Info("  DaysInThePast: " + DaysInThePast);
            log.Info("  DaysInTheFuture:" + DaysInTheFuture);
            log.Info("  SyncInterval: " + SyncInterval);
            log.Info("  SyncIntervalUnit: " + SyncIntervalUnit);
            log.Info("  Push Changes: " + OutlookPush);
            log.Info(" What");
            log.Info("  AddLocation: " + AddLocation);
            log.Info("  AddDescription: " + AddDescription + "; OnlyToGoogle: " + AddDescription_OnlyToGoogle);
            log.Info("  AddAttendees: " + AddAttendees + " <" + MaxAttendees);
            log.Info("  AddColours: " + AddColours);
            log.Info("  AddReminders: " + AddReminders);
            log.Info("    UseGoogleDefaultReminder: " + UseGoogleDefaultReminder);
            log.Info("    UseOutlookDefaultReminder: " + UseOutlookDefaultReminder);
            log.Info("    ReminderDND: " + ReminderDND + " (" + ReminderDNDstart.ToString("HH:mm") + "-" + ReminderDNDend.ToString("HH:mm") + ")");
        }

        public static SettingsStore.Calendar GetCalendarProfile(Object settingsStore) {
            if (settingsStore is SettingsStore.Calendar)
                return settingsStore as SettingsStore.Calendar;
            else throw new ArgumentException("Expected calendar settings, received " + Settings.Profile.GetType(settingsStore));
        }

        #region Override Methods
        public override bool Equals(Object calendarProfile) {
            return (calendarProfile is SettingsStore.Calendar && this._ProfileName == (calendarProfile as SettingsStore.Calendar)._ProfileName);
        }
        public override int GetHashCode() { return 0; } //Suppress compiler warning CS0659
        #endregion
    }
}
