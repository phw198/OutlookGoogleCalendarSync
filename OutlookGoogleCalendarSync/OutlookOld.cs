using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using log4net;
using Microsoft.Office.Interop.Outlook;
using Google.Apis.Calendar.v3.Data;
using NodaTime;

namespace OutlookGoogleCalendarSync {
    class OutlookOld : OutlookInterface {
        private static readonly ILog log = LogManager.GetLogger(typeof(OutlookOld));
        
        private Microsoft.Office.Interop.Outlook.Application oApp;
        private String currentUserSMTP;  //SMTP of account owner that has Outlook open
        private String currentUserName;  //Name of account owner - used to determine if attendee is "self"
        private MAPIFolder useOutlookCalendar;
        private Dictionary<string, MAPIFolder> calendarFolders = new Dictionary<string, MAPIFolder>();
        private OlExchangeConnectionMode exchangeConnectionMode;

        public void Connect() {
            oApp = OutlookCalendar.AttachToOutlook();
            log.Debug("Setting up Outlook connection.");

            // Get the NameSpace and Logon information.
            NameSpace oNS = oApp.GetNamespace("mapi");

            //Implicit logon to default profile, with no dialog box
            //If 1< profile, a dialogue is forced unless implicit login used
            exchangeConnectionMode = oNS.ExchangeConnectionMode;
            if (exchangeConnectionMode != OlExchangeConnectionMode.olNoExchange) {
                log.Info("Exchange server version: Unknown");
            }
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

            // Get the Calendar folders
            useOutlookCalendar = getDefaultCalendar(oNS);
            if (MainForm.Instance.IsHandleCreated) { //resetting connection, so pick up selected calendar from GUI dropdown
                MainForm.Instance.cbOutlookCalendars.DataSource = new BindingSource(calendarFolders, null);
                KeyValuePair<String, MAPIFolder> calendar = (KeyValuePair<String, MAPIFolder>)MainForm.Instance.cbOutlookCalendars.SelectedItem;
                calendar = (KeyValuePair<String, MAPIFolder>)MainForm.Instance.cbOutlookCalendars.SelectedItem;
                useOutlookCalendar = calendar.Value;
            }
            
            // Done. Log off.
            oNS.Logoff();
        }
        public void Disconnect() {
            Marshal.FinalReleaseComObject(oApp);
            oApp = null;
        }
        
        public List<String> Accounts() {
            List<String> accs = new List<String>();
            accs.Add(currentUserSMTP.ToLower());
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
            return oApp.GetNamespace("mapi").Offline;
        }
        public OlExchangeConnectionMode ExchangeConnectionMode() {
            return exchangeConnectionMode;
        }
        private const String gEventID = "googleEventID";

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
                    findCalendars(((MAPIFolder)defaultCalendar.Parent).Folders, calendarFolders, defaultCalendar);
                }

            } else {
                log.Debug("Finding default Mailbox calendar folders");
                defaultCalendar = oNS.GetDefaultFolder(OlDefaultFolders.olFolderCalendar);
                calendarFolders.Add("Default " + defaultCalendar.Name, defaultCalendar);
                findCalendars(((MAPIFolder)defaultCalendar.Parent).Folders, calendarFolders, defaultCalendar);
            }
            log.Debug("Default Calendar folder: " + defaultCalendar.Name);
            return defaultCalendar;
        }

        private void findCalendars(Folders folders, Dictionary<string, MAPIFolder> calendarFolders, MAPIFolder defaultCalendar) {
            foreach (MAPIFolder folder in folders) {
                try {
                    OlItemType defaultItemType = folder.DefaultItemType;
                    if (defaultItemType == OlItemType.olAppointmentItem) {
                        if (defaultCalendar == null ||
                            (folder.EntryID != defaultCalendar.EntryID))
                            calendarFolders.Add(folder.Name, folder);
                    }
                    if (folder.Folders.Count > 0) {
                        findCalendars(folder.Folders, calendarFolders, defaultCalendar);
                    }

                } catch (System.Exception ex) {
                    if (oApp.Session.ExchangeConnectionMode.ToString().Contains("Disconnected") &&
                        ex.Message == "Network problems are preventing connection to Microsoft Exchange.") {
                            log.Info("Currently disconnected from Exchange - unable to retrieve MAPI folders.");
                        MainForm.Instance.ToolTips.SetToolTip(MainForm.Instance.cbOutlookCalendars,
                            "The Outlook calendar to synchonize with.\nSome may not be listed as you are currently disconnected.");
                    } else {
                        log.Error("Failed to recurse MAPI folders.");
                        log.Error(ex.Message);
                        MessageBox.Show("An problem was encountered when searching for Outlook calendar folders.",
                            "Calendar Folders", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    }
                } 
            }
        }

        public String GetRecipientEmail(Recipient recipient) {
            String retEmail = "";
            log.Fine("Determining email of recipient: " + recipient.Name);
            AddressEntry addressEntry;
            try {
                addressEntry = recipient.AddressEntry;
            } catch {
                log.Warn("Can't resolve this recipient!");
                addressEntry = null;
            }
            if (addressEntry == null) {
                log.Warn("No AddressEntry exists!");
                retEmail = EmailAddress.BuildFakeEmailAddress(recipient.Name);
                EmailAddress.IsValidEmail(retEmail);
                return retEmail;
            }
            log.Fine("AddressEntry Type: " + recipient.AddressEntry.Type);
            if (recipient.AddressEntry.Type == "EX") { //Exchange
                log.Fine("Address is from Exchange");
                retEmail = ADX_GetSMTPAddress(recipient.AddressEntry.Address);
            } else if (recipient.AddressEntry.Type.ToUpper() == "NOTES") {
                log.Fine("From Lotus Notes");
                //Migrated contacts from notes, have weird "email addresses" eg: "James T. Kirk/US-Corp03/enterprise/US"
                retEmail = EmailAddress.BuildFakeEmailAddress(recipient.Name);

            } else {
                log.Fine("Not from Exchange");
                retEmail = recipient.AddressEntry.Address;
            }
            if (retEmail == null || retEmail == String.Empty || retEmail == "" || retEmail == "Unknown") {
                log.Error("Failed to get email address through Addin MAPI access!");
                retEmail = EmailAddress.BuildFakeEmailAddress(recipient.Name);
            }


            if (retEmail.IndexOf("<") > 0) {
                retEmail = retEmail.Substring(retEmail.IndexOf("<") + 1);
                retEmail = retEmail.TrimEnd(Convert.ToChar(">"));
            }
            log.Fine("Email address: " + retEmail, retEmail);
            EmailAddress.IsValidEmail(retEmail);
            return retEmail;
        }

        public String GetGlobalApptID(AppointmentItem ai) {
            try {
                if (ai.GlobalAppointmentID == null) throw new System.Exception();
                else return ai.GlobalAppointmentID;
            } catch {
                return ai.EntryID;
            } 
        }

        #region Addin Express Code
        //This code has been sourced from:
        //https://www.add-in-express.com/creating-addins-blog/2009/05/08/outlook-exchange-email-address-smtp/
        //https://www.add-in-express.com/files/howtos/blog/adx-ol-smtp-address-cs.zip
        public static string ADX_GetSMTPAddress(string exchangeAddress) {
            string smtpAddress = string.Empty;
            IAddrBook addrBook = ADX_GetAddrBook();
            if (addrBook != null)
                try {
                    IntPtr szPtr = IntPtr.Zero;
                    IntPtr propValuePtr = Marshal.AllocHGlobal(16);
                    IntPtr adrListPtr = Marshal.AllocHGlobal(16);

                    Marshal.WriteInt32(propValuePtr, (int)MAPI.PR_DISPLAY_NAME);
                    Marshal.WriteInt32(new IntPtr(propValuePtr.ToInt32() + 4), 0);
                    szPtr = Marshal.StringToHGlobalAnsi(exchangeAddress);
                    Marshal.WriteInt64(new IntPtr(propValuePtr.ToInt32() + 8), szPtr.ToInt32());

                    Marshal.WriteInt32(adrListPtr, 1);
                    Marshal.WriteInt32(new IntPtr(adrListPtr.ToInt32() + 4), 0);
                    Marshal.WriteInt32(new IntPtr(adrListPtr.ToInt32() + 8), 1);
                    Marshal.WriteInt32(new IntPtr(adrListPtr.ToInt32() + 12), propValuePtr.ToInt32());
                    try {
                        if (addrBook.ResolveName(0, MAPI.MAPI_DIALOG, null, adrListPtr) == MAPI.S_OK) {
                            SPropValue spValue = new SPropValue();
                            int pcount = Marshal.ReadInt32(new IntPtr(adrListPtr.ToInt32() + 8));
                            IntPtr props = new IntPtr(Marshal.ReadInt32(new IntPtr(adrListPtr.ToInt32() + 12)));
                            for (int i = 0; i < pcount; i++) {
                                spValue = (SPropValue)Marshal.PtrToStructure(
                                    new IntPtr(props.ToInt32() + (16 * i)), typeof(SPropValue));
                                if (spValue.ulPropTag == MAPI.PR_ENTRYID) {
                                    IntPtr addrEntryPtr = IntPtr.Zero;
                                    IntPtr propAddressPtr = IntPtr.Zero;
                                    uint objType = 0;
                                    uint cb = (uint)(spValue.Value & 0xFFFFFFFF);
                                    IntPtr entryID = new IntPtr((int)(spValue.Value >> 32));
                                    if (addrBook.OpenEntry(cb, entryID, IntPtr.Zero, 0, out objType, out addrEntryPtr) == MAPI.S_OK)
                                        try {
                                            if (MAPI.HrGetOneProp(addrEntryPtr, MAPI.PR_EMS_AB_PROXY_ADDRESSES, out propAddressPtr) == MAPI.S_OK) {
                                                IntPtr emails = IntPtr.Zero;
                                                SPropValue addrValue = (SPropValue)Marshal.PtrToStructure(propAddressPtr, typeof(SPropValue));
                                                int acount = (int)(addrValue.Value & 0xFFFFFFFF);
                                                IntPtr pemails = new IntPtr((int)(addrValue.Value >> 32));
                                                for (int j = 0; j < acount; j++) {
                                                    emails = new IntPtr(Marshal.ReadInt32(new IntPtr(pemails.ToInt32() + (4 * j))));
                                                    smtpAddress = Marshal.PtrToStringAnsi(emails);
                                                    if (smtpAddress.IndexOf("SMTP:") == 0) {
                                                        smtpAddress = smtpAddress.Substring(5, smtpAddress.Length - 5);
                                                        break;
                                                    }
                                                }
                                            }
                                        }
                                        finally {
                                            if (propAddressPtr != IntPtr.Zero)
                                                Marshal.Release(propAddressPtr);
                                            if (addrEntryPtr != IntPtr.Zero)
                                                Marshal.Release(addrEntryPtr);
                                        }
                                }
                            }
                        }
                    }
                    finally {
                        Marshal.FreeHGlobal(szPtr);
                        Marshal.FreeHGlobal(propValuePtr);
                        Marshal.FreeHGlobal(adrListPtr);
                    }
                }
                finally {
                    Marshal.ReleaseComObject(addrBook);
                }
            return smtpAddress;
        }

        private static IAddrBook ADX_GetAddrBook() {
            if (MAPI.MAPIInitialize(IntPtr.Zero) == MAPI.S_OK) {
                IntPtr sessionPtr = IntPtr.Zero;
                MAPI.MAPILogonEx(0, null, null, MAPI.MAPI_EXTENDED | MAPI.MAPI_ALLOW_OTHERS, out sessionPtr);
                if (sessionPtr == IntPtr.Zero)
                    MAPI.MAPILogonEx(0, null, null, MAPI.MAPI_EXTENDED | MAPI.MAPI_NEW_SESSION | MAPI.MAPI_USE_DEFAULT, out sessionPtr);
                if (sessionPtr != IntPtr.Zero)
                    try {
                        object sessionObj = Marshal.GetObjectForIUnknown(sessionPtr);
                        if (sessionObj != null)
                            try {
                                IMAPISession session = sessionObj as IMAPISession;
                                if (session != null) {
                                    IntPtr addrBookPtr = IntPtr.Zero;
                                    session.OpenAddressBook(0, IntPtr.Zero, MAPI.AB_NO_DIALOG, out addrBookPtr);
                                    if (addrBookPtr != IntPtr.Zero)
                                        try {
                                            object addrBookObj = Marshal.GetObjectForIUnknown(addrBookPtr);
                                            if (addrBookObj != null)
                                                return addrBookObj as IAddrBook;
                                        }
                                        finally {
                                            Marshal.Release(addrBookPtr);
                                        }
                                }
                            }
                            finally {
                                Marshal.ReleaseComObject(sessionObj);
                            }
                    }
                    finally {
                        Marshal.Release(sessionPtr);
                    }
            } else
                throw new ApplicationException("MAPI can not be initialized.");
            return null;
        }

        #region Extended MAPI routines

        internal class MAPI {
            public const int S_OK = 0;

            public const uint MV_FLAG = 0x1000;

            public const uint PT_UNSPECIFIED = 0;
            public const uint PT_NULL = 1;
            public const uint PT_I2 = 2;
            public const uint PT_LONG = 3;
            public const uint PT_R4 = 4;
            public const uint PT_DOUBLE = 5;
            public const uint PT_CURRENCY = 6;
            public const uint PT_APPTIME = 7;
            public const uint PT_ERROR = 10;
            public const uint PT_BOOLEAN = 11;
            public const uint PT_OBJECT = 13;
            public const uint PT_I8 = 20;
            public const uint PT_STRING8 = 30;
            public const uint PT_UNICODE = 31;
            public const uint PT_SYSTIME = 64;
            public const uint PT_CLSID = 72;
            public const uint PT_BINARY = 258;
            public const uint PT_MV_TSTRING = (MV_FLAG | PT_STRING8);

            public const uint PR_SENDER_ADDRTYPE = (PT_STRING8 | (0x0C1E << 16));
            public const uint PR_SENDER_EMAIL_ADDRESS = (PT_STRING8 | (0x0C1F << 16));
            public const uint PR_SENDER_NAME = (PT_STRING8 | (0x0C1A << 16));
            public const uint PR_ADDRTYPE = (PT_STRING8 | (0x3002 << 16));
            public const uint PR_ADDRTYPE_W = (PT_UNICODE | (0x3002 << 16));
            public const uint PR_EMAIL_ADDRESS = (PT_STRING8 | (0x3003 << 16));
            public const uint PR_EMAIL_ADDRESS_W = (PT_UNICODE | (0x3003 << 16));
            public const uint PR_DISPLAY_NAME = (PT_STRING8 | (0x3001 << 16));
            public const uint PR_DISPLAY_NAME_W = (PT_UNICODE | (0x3001 << 16));
            public const uint PR_ENTRYID = (PT_BINARY | (0x0FFF << 16));
            public const uint PR_EMS_AB_PROXY_ADDRESSES = unchecked((uint)(PT_MV_TSTRING | (0x800F << 16)));

            public const uint PR_SMTP_ADDRESS = (PT_STRING8 | (0x39FE << 16));
            public const uint PR_SMTP_ADDRESS_W = (PT_UNICODE | (0x39FE << 16));

            public const uint MAPI_NEW_SESSION = 0x00000002;
            public const uint MAPI_FORCE_DOWNLOAD = 0x00001000;
            public const uint MAPI_LOGON_UI = 0x00000001;
            public const uint MAPI_ALLOW_OTHERS = 0x00000008;
            public const uint MAPI_EXPLICIT_PROFILE = 0x00000010;
            public const uint MAPI_EXTENDED = 0x00000020;
            public const uint MAPI_SERVICE_UI_ALWAYS = 0x00002000;
            public const uint MAPI_NO_MAIL = 0x00008000;
            public const uint MAPI_USE_DEFAULT = 0x00000040;

            public const uint AB_NO_DIALOG = 0x00000001;
            public const uint MAPI_DIALOG = 0x00000008;

            public const string IID_IMAPIProp = "00020303-0000-0000-C000-000000000046";

            [DllImport("MAPI32.DLL", CharSet = CharSet.Ansi, EntryPoint = "HrGetOneProp@12")]
            public static extern int HrGetOneProp(IntPtr pmp, uint ulPropTag, out IntPtr ppProp);

            [DllImport("MAPI32.DLL", CharSet = CharSet.Ansi, EntryPoint = "MAPIFreeBuffer@4")]
            public static extern void MAPIFreeBuffer(IntPtr lpBuffer);

            [DllImport("MAPI32.DLL", CharSet = CharSet.Ansi, EntryPoint = "MAPIInitialize@4")]
            public static extern int MAPIInitialize(IntPtr lpMapiInit);

            [DllImport("MAPI32.DLL", CharSet = CharSet.Ansi, EntryPoint = "MAPILogonEx@20")]
            public static extern int MAPILogonEx(uint ulUIParam, [MarshalAs(UnmanagedType.LPWStr)] string lpszProfileName,
                [MarshalAs(UnmanagedType.LPWStr)] string lpszPassword, uint flFlags, out IntPtr lppSession);
        }

        [StructLayout(LayoutKind.Sequential)]
        public struct SPropValue {
            public uint ulPropTag;
            public uint dwAlignPad;
            public long Value;
        }

        [ComImport, ComVisible(false), InterfaceType(ComInterfaceType.InterfaceIsIUnknown),
        Guid("00020300-0000-0000-C000-000000000046")]
        public interface IMAPISession {
            int GetLastError(int hResult, uint ulFlags, out IntPtr lppMAPIError);
            int GetMsgStoresTable(uint ulFlags, out IntPtr lppTable);
            int OpenMsgStore(uint ulUIParam, uint cbEntryID, IntPtr lpEntryID, ref Guid lpInterface, uint ulFlags, out IntPtr lppMDB);
            int OpenAddressBook(uint ulUIParam, IntPtr lpInterface, uint ulFlags, out IntPtr lppAdrBook);
            int OpenProfileSection(ref Guid lpUID, ref Guid lpInterface, uint ulFlags, out IntPtr lppProfSect);
            int GetStatusTable(uint ulFlags, out IntPtr lppTable);
            int OpenEntry(uint cbEntryID, IntPtr lpEntryID, ref Guid lpInterface, uint ulFlags, out uint lpulObjType, out IntPtr lppUnk);
            int CompareEntryIDs(uint cbEntryID1, IntPtr lpEntryID1, uint cbEntryID2, IntPtr lpEntryID2, uint ulFlags, out uint lpulResult);
            int Advise(uint cbEntryID, IntPtr lpEntryID, uint ulEventMask, IntPtr lpAdviseSink, out uint lpulConnection);
            int Unadvise(uint ulConnection);
            int MessageOptions(uint ulUIParam, uint ulFlags, [MarshalAs(UnmanagedType.LPWStr)] string lpszAdrType, IntPtr lpMessage);
            int QueryDefaultMessageOpt([MarshalAs(UnmanagedType.LPWStr)] string lpszAdrType, uint ulFlags, out uint lpcValues, out IntPtr lppOptions);
            int EnumAdrTypes(uint ulFlags, out uint lpcAdrTypes, out IntPtr lpppszAdrTypes);
            int QueryIdentity(out uint lpcbEntryID, out IntPtr lppEntryID);
            int Logoff(uint ulUIParam, uint ulFlags, uint ulReserved);
            int SetDefaultStore(uint ulFlags, uint cbEntryID, IntPtr lpEntryID);
            int AdminServices(uint ulFlags, out IntPtr lppServiceAdmin);
            int ShowForm(uint ulUIParam, IntPtr lpMsgStore, IntPtr lpParentFolder, ref Guid lpInterface, uint ulMessageToken,
                IntPtr lpMessageSent, uint ulFlags, uint ulMessageStatus, uint ulMessageFlags, uint ulAccess, [MarshalAs(UnmanagedType.LPWStr)] string lpszMessageClass);
            int PrepareForm(ref Guid lpInterface, IntPtr lpMessage, out uint lpulMessageToken);
        }

        [ComImport, ComVisible(false), InterfaceType(ComInterfaceType.InterfaceIsIUnknown),
        Guid("00020309-0000-0000-C000-000000000046")]
        public interface IAddrBook {
            int GetLastError(int hResult, uint ulFlags, out IntPtr lppMAPIError);
            int SaveChanges(uint ulFlags);
            int GetProps(IntPtr lpPropTagArray, uint ulFlags, out uint lpcValues, out IntPtr lppPropArray);
            int GetPropList(uint ulFlags, out IntPtr lppPropTagArray);
            int OpenProperty(uint ulPropTag, ref Guid lpiid, uint ulInterfaceOptions, uint ulFlags, out IntPtr lppUnk);
            int SetProps(uint cValues, IntPtr lpPropArray, out IntPtr lppProblems);
            int DeleteProps(IntPtr lpPropTagArray, out IntPtr lppProblems);
            int CopyTo(uint ciidExclude, ref Guid rgiidExclude, IntPtr lpExcludeProps, uint ulUIParam,
                IntPtr lpProgress, ref Guid lpInterface, IntPtr lpDestObj, uint ulFlags, out IntPtr lppProblems);
            int CopyProps(IntPtr lpIncludeProps, uint ulUIParam, IntPtr lpProgress, ref Guid lpInterface,
                IntPtr lpDestObj, uint ulFlags, out IntPtr lppProblems);
            int GetNamesFromIDs(out IntPtr lppPropTags, ref Guid lpPropSetGuid, uint ulFlags,
                out uint lpcPropNames, out IntPtr lpppPropNames);
            int GetIDsFromNames(uint cPropNames, ref IntPtr lppPropNames, uint ulFlags, out IntPtr lppPropTags);
            int OpenEntry(uint cbEntryID, IntPtr lpEntryID, IntPtr lpInterface, uint ulFlags, out uint lpulObjType, out IntPtr lppUnk);
            int CompareEntryIDs(uint cbEntryID1, IntPtr lpEntryID1, uint cbEntryID2, IntPtr lpEntryID2, uint ulFlags, out uint lpulResult);
            int Advise(uint cbEntryID, IntPtr lpEntryID, uint ulEventMask, IntPtr lpAdviseSink, out uint lpulConnection);
            int Unadvise(uint ulConnection);
            int CreateOneOff([MarshalAs(UnmanagedType.LPWStr)] string lpszName, [MarshalAs(UnmanagedType.LPWStr)] string lpszAdrType,
                [MarshalAs(UnmanagedType.LPWStr)] string lpszAddress, uint ulFlags, out uint lpcbEntryID, out IntPtr lppEntryID);
            int NewEntry(uint ulUIParam, uint ulFlags, uint cbEIDContainer, IntPtr lpEIDContainer, uint cbEIDNewEntryTpl, IntPtr lpEIDNewEntryTpl, out uint lpcbEIDNewEntry, out IntPtr lppEIDNewEntry);
            int ResolveName(uint ulUIParam, uint ulFlags, [MarshalAs(UnmanagedType.LPWStr)] string lpszNewEntryTitle, IntPtr lpAdrList);
            int Address(out uint lpulUIParam, IntPtr lpAdrParms, out IntPtr lppAdrList);
            int Details(out uint lpulUIParam, IntPtr lpfnDismiss, IntPtr lpvDismissContext, uint cbEntryID, IntPtr lpEntryID,
                IntPtr lpfButtonCallback, IntPtr lpvButtonContext, [MarshalAs(UnmanagedType.LPWStr)] string lpszButtonText, uint ulFlags);
            int RecipOptions(uint ulUIParam, uint ulFlags, IntPtr lpRecip);
            int QueryDefaultRecipOpt([MarshalAs(UnmanagedType.LPWStr)] string lpszAdrType, uint ulFlags, out uint lpcValues, out IntPtr lppOptions);
            int GetPAB(out uint lpcbEntryID, out IntPtr lppEntryID);
            int SetPAB(uint cbEntryID, IntPtr lpEntryID);
            int GetDefaultDir(out uint lpcbEntryID, out IntPtr lppEntryID);
            int SetDefaultDir(uint cbEntryID, IntPtr lpEntryID);
            int GetSearchPath(uint ulFlags, out IntPtr lppSearchPath);
            int SetSearchPath(uint ulFlags, IntPtr lpSearchPath);
            int PrepareRecips(uint ulFlags, IntPtr lpSPropTagArray, IntPtr lpRecipList);
        }

        #endregion

        #endregion

        #region TimeZone Stuff
        public Event IANAtimezone_set(Event ev, AppointmentItem ai) {
            ev.Start.TimeZone = IANAtimezone("UTC", "(UTC) Coordinated Universal Time");
            ev.End.TimeZone = IANAtimezone("UTC", "(UTC) Coordinated Universal Time");
            return ev;
        }

        private String IANAtimezone(String oTZ_id, String oTZ_name) {
            //Convert from Windows Timezone to Iana
            //Eg "(UTC) Dublin, Edinburgh, Lisbon, London" => "Europe/London"
            //http://unicode.org/repos/cldr/trunk/common/supplemental/windowsZones.xml
            if (oTZ_id.Equals("UTC", StringComparison.OrdinalIgnoreCase)) {
                log.Fine("Windows Timezone \"" + oTZ_name + "\" mapped to \"Etc/UTC\"");
                return "Etc/UTC";
            }

            NodaTime.TimeZones.TzdbDateTimeZoneSource tzDBsource = NodaTime.TimeZones.TzdbDateTimeZoneSource.Default;
            TimeZoneInfo tzi = TimeZoneInfo.FindSystemTimeZoneById(oTZ_id);
            String tzID = tzDBsource.MapTimeZoneId(tzi);
            log.Fine("Windows Timezone \"" + oTZ_name + "\" mapped to \"" + tzDBsource.CanonicalIdMap[tzID] + "\"");
            return tzDBsource.CanonicalIdMap[tzID];
        }

        public AppointmentItem WindowsTimeZone_set(AppointmentItem ai, Event ev) {
            ai.Start = WindowsTimeZone(ev.Start);
            ai.End = WindowsTimeZone(ev.End);
            return ai;
        }

        private DateTime WindowsTimeZone(EventDateTime time) {
            DateTime theDate = DateTime.Parse(time.DateTime ?? time.Date);
            if (time.TimeZone == null) return theDate;

            LocalDateTime local = new LocalDateTime(theDate.Year, theDate.Month, theDate.Day, theDate.Hour, theDate.Minute);
            DateTimeZone zone = DateTimeZoneProviders.Tzdb[time.TimeZone];
            ZonedDateTime zonedTime = local.InZoneLeniently(zone);
            DateTime zonedUTC = zonedTime.ToDateTimeUtc();
            log.Fine("IANA Timezone \"" + time.TimeZone + "\" mapped to \""+ zone.Id.ToString() +"\" with a UTC of "+ zonedUTC.ToString("dd/MM/yyyy HH:mm:ss"));
            return zonedUTC;
        }
        #endregion

    }
}
