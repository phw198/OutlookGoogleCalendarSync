using System;
using System.Collections.Generic;
using log4net;
using Microsoft.Office.Interop.Outlook;

namespace OutlookGoogleCalendarSync {
    class OutlookNew : OutlookInterface {
        private static readonly ILog log = LogManager.GetLogger(typeof(OutlookNew));
        
        private Application oApp;
        private String currentUserSMTP;  //SMTP of account owner that has Outlook open
        private String currentUserName;  //Name of account owner - used to determine if attendee is "self"
        private Accounts accounts;
        private MAPIFolder useOutlookCalendar;
        private Dictionary<string, MAPIFolder> calendarFolders = new Dictionary<string, MAPIFolder>();

        public void Connect() {
            log.Debug("Setting up Outlook connection.");
            
            // Create the Outlook application.
            oApp = new Application();

            // Get the NameSpace and Logon information.
            NameSpace oNS = oApp.GetNamespace("mapi");

            //Log on by using a dialog box to choose the profile.
            oNS.Logon("", "", true, true);
            currentUserSMTP = GetRecipientEmail(oNS.CurrentUser);
            currentUserName = oNS.CurrentUser.Name;

            //Alternate logon method that uses a specific profile.
            // If you use this logon method, 
            // change the profile name to an appropriate value.
            //oNS.Logon("YourValidProfile", Missing.Value, false, true); 

            //Get the accounts configured in Outlook
            accounts = oNS.Accounts;

            // Get the Default Calendar folder
            if (Settings.Instance.OutlookService == OutlookCalendar.Service.AlternativeMailbox && Settings.Instance.MailboxName != "") {
                useOutlookCalendar = oNS.Folders[Settings.Instance.MailboxName].Folders["Calendar"];
            } else {
                // Use the logged in user's Calendar folder.
                useOutlookCalendar = oNS.GetDefaultFolder(OlDefaultFolders.olFolderCalendar);
            }
            calendarFolders.Add("Default " + useOutlookCalendar.Name, useOutlookCalendar);
            //Get any subfolders - note, this isn't recursive
            foreach (MAPIFolder calendar in useOutlookCalendar.Folders) {
                if (calendar.DefaultItemType == OlItemType.olAppointmentItem) {
                    calendarFolders.Add(calendar.Name, calendar);
                }
            }

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

        private const String gEventID = "googleEventID";
        public const String PR_SMTP_ADDRESS = "http://schemas.microsoft.com/mapi/proptag/0x39FE001E";

        public String GetRecipientEmail(Recipient recipient) {
            String retEmail = "";
            log.Fine("Determining email of recipient: " + recipient.Name);
            if (recipient.AddressEntry == null) {
                log.Debug("No AddressEntry exists!");
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
                        log.Error("Exchange does not have an email for this recipient's account!");
                        try {
                            Microsoft.Office.Interop.Outlook.PropertyAccessor pa = recipient.PropertyAccessor;
                            retEmail = pa.GetProperty(OutlookNew.PR_SMTP_ADDRESS).ToString();
                            log.Debug("Retrieved from PropertyAccessor instead.");
                        } catch {
                            log.Error("Also failed to retrieve email from PropertyAccessor.");
                            String buildFakeEmail = recipient.Name.Replace(",", "");
                            buildFakeEmail = buildFakeEmail.Replace(" ", "");
                            buildFakeEmail += "@unknownemail.com";
                            log.Debug("Built a fake email for them: " + buildFakeEmail);
                            retEmail = buildFakeEmail;
                        }
                    }
                } else {
                    log.Fine("Exchange type: " + recipient.AddressEntry.AddressEntryUserType.ToString());
                    log.Fine("Using PropertyAccessor to get email address.");
                    Microsoft.Office.Interop.Outlook.PropertyAccessor pa = recipient.PropertyAccessor;
                    retEmail = pa.GetProperty(OutlookNew.PR_SMTP_ADDRESS).ToString();
                }
            } else {
                log.Fine("Not from Exchange");
                retEmail = recipient.AddressEntry.Address;
            }
            log.Fine("Email address: " + retEmail);
            return retEmail;
        }
        
    }
}
