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
                foreach (Folder folder in oNS.Folders[Settings.Instance.MailboxName].Folders) {
                    if (folder.DefaultItemType == OlItemType.olAppointmentItem) {
                        log.Fine("Alternate mailbox default Calendar folder: " + folder.Name);
                        useOutlookCalendar = folder;
                    }
                }
                if (useOutlookCalendar == null) {
                    System.Windows.Forms.MessageBox.Show("Unable to find a Calendar folder in the alternative mailbox.\r\n" +
                        "Reverting to the default mailbox calendar", "Calendar not found", System.Windows.Forms.MessageBoxButtons.OK);
                    MainForm.Instance.rbOutlookDefaultMB.CheckedChanged -= MainForm.Instance.rbOutlookDefaultMB_CheckedChanged;
                    MainForm.Instance.rbOutlookDefaultMB.Checked = true;
                    MainForm.Instance.rbOutlookDefaultMB.CheckedChanged += MainForm.Instance.rbOutlookDefaultMB_CheckedChanged;
                    useOutlookCalendar = oNS.GetDefaultFolder(OlDefaultFolders.olFolderCalendar);
                }
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
        private const String PR_ORIGINAL_DISPLAY_NAME = "http://schemas.microsoft.com/mapi/proptag/0x3A13001E";

        public String GetRecipientEmail(Recipient recipient) {
            String retEmail = "";
            log.Fine("Determining email of recipient: " + recipient.Name);
            try {
                AddressEntry addressEntry = recipient.AddressEntry;
            } catch {
                log.Warn("Can't resolve this recipient!");
                return OutlookCalendar.BuildFakeEmailAddress(recipient.Name);
            }
            if (recipient.AddressEntry == null) {
                log.Warn("No AddressEntry exists!");
                return OutlookCalendar.BuildFakeEmailAddress(recipient.Name);
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
                        log.Warn("Exchange does not have an email for this recipient's account!");
                        try {
                            //Should I try PR_EMS_AB_PROXY_ADDRESSES next to cater for cached mode?
                            Microsoft.Office.Interop.Outlook.PropertyAccessor pa = recipient.PropertyAccessor;
                            retEmail = pa.GetProperty(OutlookNew.PR_SMTP_ADDRESS).ToString();
                            log.Debug("Retrieved from PropertyAccessor instead.");
                        } catch {
                            log.Warn("Also failed to retrieve email from PropertyAccessor.");
                            retEmail = OutlookCalendar.BuildFakeEmailAddress(recipient.Name);
                        }
                    }

                } else if (recipient.AddressEntry.AddressEntryUserType == OlAddressEntryUserType.olOutlookContactAddressEntry) {
                    log.Fine("This is an Exchange contact");
                    ContactItem contact = null;
                    try {
                        contact = recipient.AddressEntry.GetContact();
                    } catch {
                        log.Warn("Doesn't seem to be a valid contact object. Maybe this account is not longer in Exchange.");
                        retEmail = OutlookCalendar.BuildFakeEmailAddress(recipient.Name);
                    }
                    if (contact != null) {
                        Microsoft.Office.Interop.Outlook.PropertyAccessor pa = contact.PropertyAccessor;
                        retEmail = pa.GetProperty(OutlookNew.PR_ORIGINAL_DISPLAY_NAME).ToString();
                        retEmail = contact.Email1Address;
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
