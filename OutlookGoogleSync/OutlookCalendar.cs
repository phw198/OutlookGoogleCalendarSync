using System;
using System.Collections.Generic;
using Microsoft.Office.Interop.Outlook;
using System.IO;

namespace OutlookGoogleSync {
    /// <summary>
    /// Description of OutlookCalendar.
    /// </summary>
    public class OutlookCalendar {
        public const String PR_SMTP_ADDRESS = "http://schemas.microsoft.com/mapi/proptag/0x39FE001E";

        private static OutlookCalendar instance;
        private String currentUserSMTP;  //SMTP of account owner that has Outlook open
        private String currentUserName;  //Name of account owner - used to determine if attendee is "self"
        private MAPIFolder useOutlookCalendar;
        private Accounts accounts;
        private Dictionary<string, MAPIFolder> calendarFolders = new Dictionary<string, MAPIFolder>();

        public static OutlookCalendar Instance {
            get {
                if (instance == null) instance = new OutlookCalendar();
                return instance;
            }
        }
        public String CurrentUserSMTP {
            get { return currentUserSMTP; }
        }
        public String CurrentUserName {
            get { return currentUserName; }
        }
        public MAPIFolder UseOutlookCalendar { get; set; }
        public Accounts Accounts {
            get { return accounts; }
        }
        public Dictionary<string, MAPIFolder> CalendarFolders {
            get { return calendarFolders; }
        }
        public enum Service {
            DefaultMailbox,
            AlternativeMailbox,
            EWS
        }

        public OutlookCalendar() {

            // Create the Outlook application.
            Application oApp = new Application();

            // Get the NameSpace and Logon information.
            NameSpace oNS = oApp.GetNamespace("mapi");

            //Log on by using a dialog box to choose the profile.
            oNS.Logon("", "", true, true);
            currentUserSMTP = ((oNS.CurrentUser as Recipient).PropertyAccessor as PropertyAccessor).GetProperty(OutlookCalendar.PR_SMTP_ADDRESS).ToString().ToLower();
            currentUserName = oNS.CurrentUser.Name;

            //Alternate logon method that uses a specific profile.
            // If you use this logon method, 
            // change the profile name to an appropriate value.
            //oNS.Logon("YourValidProfile", Missing.Value, false, true); 

            //Get the accounts configured in Outlook
            accounts = oNS.Accounts;

            // Get the Default Calendar folder
            if (Settings.Instance.OutlookService == Service.AlternativeMailbox && Settings.Instance.MailboxName != "") {
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

        public void Reset() {
            instance = new OutlookCalendar();
        }

        //public List<AppointmentItem> getCalendarEntries() {
        //    Items OutlookItems = UseOutlookCalendar.Items;
        //    if (OutlookItems != null) {
        //        List<AppointmentItem> result = new List<AppointmentItem>();
        //        foreach (AppointmentItem ai in OutlookItems) {
        //            result.Add(ai);
        //        }
        //        return result;
        //    }
        //    return null;
        //}

        public List<AppointmentItem> getCalendarEntriesInRange() {
            List<AppointmentItem> result = new List<AppointmentItem>();

            Items OutlookItems = UseOutlookCalendar.Items;
            OutlookItems.Sort("[Start]", Type.Missing);
            OutlookItems.IncludeRecurrences = true;

            if (OutlookItems != null) {
                DateTime min = DateTime.Now.AddDays(-Settings.Instance.DaysInThePast);
                DateTime max = DateTime.Now.AddDays(+Settings.Instance.DaysInTheFuture + 1);
                string filter = "[End] >= '" + min.ToString("g") + "' AND [Start] < '" + max.ToString("g") + "'";

                foreach (AppointmentItem ai in OutlookItems.Restrict(filter)) {
                    result.Add(ai);
                }
            }

            if (Settings.Instance.CreateTextFiles) {
                TextWriter tw = new StreamWriter("export_found_in_outlook.txt");
                foreach (AppointmentItem ai in result) {
                    tw.WriteLine(signature(ai));
                }
                tw.Close();
            }

            return result;
        }

        #region STATIC functions
        public static string signature(AppointmentItem ai) {
            return (GoogleCalendar.GoogleTimeFrom(ai.Start) + ";" + GoogleCalendar.GoogleTimeFrom(ai.End) + ";" + ai.Subject + ";" + ai.Location).Trim();
        }
        #endregion
    }
}
