using System;
using System.Collections.Generic;
using System.Windows.Forms;
using log4net;
using Microsoft.Office.Interop.Outlook;

namespace OutlookGoogleCalendarSync {
    class OutlookNew : OutlookInterface {
        private static readonly ILog log = LogManager.GetLogger(typeof(OutlookNew));
        
        private Microsoft.Office.Interop.Outlook.Application oApp;
        private String currentUserSMTP;  //SMTP of account owner that has Outlook open
        private String currentUserName;  //Name of account owner - used to determine if attendee is "self"
        private Accounts accounts;
        private MAPIFolder useOutlookCalendar;
        private Dictionary<string, MAPIFolder> calendarFolders = new Dictionary<string, MAPIFolder>();
        private OlExchangeConnectionMode exchangeConnectionMode;

        public void Connect() {
            log.Debug("Setting up Outlook connection.");
            
            // Create the Outlook application.
            oApp = new Microsoft.Office.Interop.Outlook.Application();

            // Get the NameSpace and Logon information.
            NameSpace oNS = oApp.GetNamespace("mapi");

            //Log on by using a dialog box to choose the profile.
            //oNS.Logon("", Type.Missing, true, true); 

            //Implicit logon to default profile, with no dialog box
            //If 1< profile, a dialogue is forced unless implicit login used
            exchangeConnectionMode = oNS.ExchangeConnectionMode;
            if (exchangeConnectionMode != OlExchangeConnectionMode.olNoExchange) {
                log.Info("Exchange server version: " + oNS.ExchangeMailboxServerVersion.ToString());
            }
            
            //Logon using a specific profile. Can't see a use case for this when using OGsync
            //If using this logon method, change the profile name to an appropriate value:
            //HKEY_CURRENT_USER\Software\Microsoft\Windows NT\CurrentVersion\Windows Messaging Subsystem\Profiles
            //oNS.Logon("YourValidProfile", Type.Missing, false, true); 

            log.Info("Exchange connection mode: " + exchangeConnectionMode.ToString());
            currentUserSMTP = GetRecipientEmail(oNS.CurrentUser);
            currentUserName = oNS.CurrentUser.Name;
            if (currentUserName == "Unknown") {
                log.Info("Current username is \"Unknown\"");
                if (Settings.Instance.AddAttendees) {
                    System.Windows.Forms.MessageBox.Show("It appears you do not have an Email Account configured in Outlook.\r\n" +
                        "You should set one up now (Tools > Email Accounts) to avoid problems syncing meeting attendees.",
                        "No Email Account Found", System.Windows.Forms.MessageBoxButtons.OK,
                        System.Windows.Forms.MessageBoxIcon.Warning);
                }
            }
            
            //Get the accounts configured in Outlook
            accounts = oNS.Accounts;
            
            // Get the Calendar folders
            useOutlookCalendar = getDefaultCalendar(oNS);

            // Done. Log off.
            oNS.Logoff();
        }
        
        public List<String> Accounts() {
            List<String> accs = new List<String>();
            foreach (Account acc in accounts) {
                if (acc.SmtpAddress != null)
                    accs.Add(acc.SmtpAddress.ToLower());
            }
            return accs;
        }
        public Dictionary<string, MAPIFolder> CalendarFolders() { 
            return calendarFolders;
        }
        public MAPIFolder UseOutlookCalendar() {
            return useOutlookCalendar;
        }
        public void UseOutlookCalendar(MAPIFolder set) {
            useOutlookCalendar = set;
        }
        public String CurrentUserSMTP() {
            return currentUserSMTP;
        }
        public String CurrentUserName() {
            return currentUserName;
        }
        public Boolean Offline() {
            try {
                return oApp.GetNamespace("mapi").Offline;
            } catch {
                OutlookCalendar.Instance.Reset();
                return oApp.GetNamespace("mapi").Offline;
            }
        }
        public OlExchangeConnectionMode ExchangeConnectionMode() {
            return exchangeConnectionMode;
        }

        private const String gEventID = "googleEventID";
        public const String PR_SMTP_ADDRESS = "http://schemas.microsoft.com/mapi/proptag/0x39FE001E";
        private const String EMAIL1ADDRESS = "http://schemas.microsoft.com/mapi/id/{00062004-0000-0000-C000-000000000046}/8084001F";

        private MAPIFolder getDefaultCalendar(NameSpace oNS) {
            MAPIFolder defaultCalendar = null;
            if (Settings.Instance.OutlookService == OutlookCalendar.Service.AlternativeMailbox && Settings.Instance.MailboxName != "") {
                log.Debug("Finding Alternative Mailbox calendar folders");
                findCalendars(oNS.Folders[Settings.Instance.MailboxName].Folders, calendarFolders, defaultCalendar);

                //Default to first calendar in drop down
                foreach (KeyValuePair<String, MAPIFolder> calendar in calendarFolders) {
                    defaultCalendar = calendar.Value;
                    break;
                }
                if (defaultCalendar == null) {
                    log.Info("Could not find Alternative mailbox Calendar folder. Reverting to the default mailbox calendar.");
                    System.Windows.Forms.MessageBox.Show("Unable to find a Calendar folder in the alternative mailbox.\r\n" +
                        "Reverting to the default mailbox calendar", "Calendar not found", System.Windows.Forms.MessageBoxButtons.OK);
                    MainForm.Instance.rbOutlookDefaultMB.CheckedChanged -= MainForm.Instance.rbOutlookDefaultMB_CheckedChanged;
                    MainForm.Instance.rbOutlookDefaultMB.Checked = true;
                    MainForm.Instance.rbOutlookDefaultMB.CheckedChanged += MainForm.Instance.rbOutlookDefaultMB_CheckedChanged;
                    defaultCalendar = oNS.GetDefaultFolder(OlDefaultFolders.olFolderCalendar);
                    calendarFolders.Add("Default " + defaultCalendar.Name, defaultCalendar);
                    findCalendars(oNS.DefaultStore.GetRootFolder().Folders, calendarFolders, defaultCalendar);
                }

            } else {
                log.Debug("Finding default Mailbox calendar folders");
                defaultCalendar = oNS.GetDefaultFolder(OlDefaultFolders.olFolderCalendar);
                calendarFolders.Add("Default " + defaultCalendar.Name, defaultCalendar);
                findCalendars(oNS.DefaultStore.GetRootFolder().Folders, calendarFolders, defaultCalendar);
            }
            log.Debug("Default Calendar folder: " + defaultCalendar.Name);
            return defaultCalendar;
        }

        private void findCalendars(Folders folders, Dictionary<string, MAPIFolder> calendarFolders, MAPIFolder defaultCalendar) {
            string excludeDeletedFolder = folders.Application.Session.GetDefaultFolder(OlDefaultFolders.olFolderDeletedItems).EntryID;
            foreach (MAPIFolder folder in folders) {
                try {
                    OlItemType defaultItemType = folder.DefaultItemType;
                    if (defaultItemType == OlItemType.olAppointmentItem) {
                        if (defaultCalendar == null ||
                            (folder.EntryID != defaultCalendar.EntryID))
                            calendarFolders.Add(folder.Name, folder);
                    }
                    if (folder.EntryID != excludeDeletedFolder && folder.Folders.Count > 0) {
                        findCalendars(folder.Folders, calendarFolders, defaultCalendar);
                    }

                } catch (System.Exception ex) {
                    if (oApp.Session.ExchangeConnectionMode.ToString().Contains("Disconnected") &&
                        ex.Message.StartsWith("Network problems are preventing connection to Microsoft Exchange.")) {
                            log.Info("Currently disconnected from Exchange - unable to retreive MAPI folders.");
                        MainForm.Instance.ToolTips.SetToolTip(MainForm.Instance.cbOutlookCalendars,
                            "The Outlook calendar to synchonize with.\nSome may not be listed as you are currently disconnected.");
                    } else {
                        log.Error("Failed to recurse MAPI folders.");
                        log.Error(ex.Message);
                        MessageBox.Show("A problem was encountered when searching for Outlook calendar folders.",
                            "Calendar Folders", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    }
                }
            }
        }

        public String GetRecipientEmail(Recipient recipient) {
            String retEmail = "";
            log.Fine("Determining email of recipient: " + recipient.Name);
            try {
                AddressEntry addressEntry = recipient.AddressEntry;
            } catch {
                log.Warn("Can't resolve this recipient!");
                recipient.AddressEntry = null;
            }
            if (recipient.AddressEntry == null) {
                log.Warn("No AddressEntry exists!");
                retEmail = EmailAddress.BuildFakeEmailAddress(recipient.Name);
                EmailAddress.IsValidEmail(retEmail);
                return retEmail;
            }
            log.Fine("AddressEntry Type: " + recipient.AddressEntry.Type);
            if (recipient.AddressEntry.Type == "EX") { //Exchange
                log.Fine("Address is from Exchange");
                if (recipient.AddressEntry.AddressEntryUserType == OlAddressEntryUserType.olExchangeUserAddressEntry ||
                    recipient.AddressEntry.AddressEntryUserType == OlAddressEntryUserType.olExchangeRemoteUserAddressEntry) {
                    ExchangeUser eu = recipient.AddressEntry.GetExchangeUser();
                    if (eu != null && eu.PrimarySmtpAddress != null)
                        retEmail = eu.PrimarySmtpAddress;
                    else {
                        log.Warn("Exchange does not have an email for recipient: "+ recipient.Name);
                        try {
                            //Should I try PR_EMS_AB_PROXY_ADDRESSES next to cater for cached mode?
                            Microsoft.Office.Interop.Outlook.PropertyAccessor pa = recipient.PropertyAccessor;
                            retEmail = pa.GetProperty(OutlookNew.PR_SMTP_ADDRESS).ToString();
                            log.Debug("Retrieved from PropertyAccessor instead.");
                        } catch {
                            log.Warn("Also failed to retrieve email from PropertyAccessor.");
                            retEmail = EmailAddress.BuildFakeEmailAddress(recipient.Name);
                        }
                    }

                } else if (recipient.AddressEntry.AddressEntryUserType == OlAddressEntryUserType.olOutlookContactAddressEntry) {
                    log.Fine("This is an Outlook contact");
                    ContactItem contact = null;
                    try {
                        contact = recipient.AddressEntry.GetContact();
                    } catch {
                        log.Warn("Doesn't seem to be a valid contact object. Maybe this account is not longer in Exchange.");
                        retEmail = EmailAddress.BuildFakeEmailAddress(recipient.Name);
                    }
                    if (contact != null) {
                        if (contact.Email1AddressType == "EX") {
                            log.Fine("Address is from Exchange.");
                            log.Fine("Using PropertyAccessor to get email address.");
                            Microsoft.Office.Interop.Outlook.PropertyAccessor pa = contact.PropertyAccessor;
                            retEmail = pa.GetProperty(EMAIL1ADDRESS).ToString();
                        } else {
                            retEmail = contact.Email1Address;
                        }
                    }
                } else {
                    log.Fine("Exchange type: " + recipient.AddressEntry.AddressEntryUserType.ToString());
                    log.Fine("Using PropertyAccessor to get email address.");
                    Microsoft.Office.Interop.Outlook.PropertyAccessor pa = recipient.PropertyAccessor;
                    retEmail = pa.GetProperty(OutlookNew.PR_SMTP_ADDRESS).ToString();
                }

            } else if (recipient.AddressEntry.Type.ToUpper() == "NOTES") {
                log.Fine("From Lotus Notes");
                //Migrated contacts from notes, have weird "email addresses" eg: "James T. Kirk/US-Corp03/enterprise/US"
                retEmail = EmailAddress.BuildFakeEmailAddress(recipient.Name);

            } else {
                log.Fine("Not from Exchange");
                retEmail = recipient.AddressEntry.Address;
            }

            if (retEmail.IndexOf("<") > 0) {
                retEmail = retEmail.Substring(retEmail.IndexOf("<") + 1);
                retEmail = retEmail.TrimEnd(Convert.ToChar(">"));
            }
            log.Fine("Email address: " + retEmail);
            EmailAddress.IsValidEmail(retEmail);
            return retEmail;
        }
        
    }
}
