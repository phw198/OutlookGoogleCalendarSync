using System;
using System.Collections.Generic;
using log4net;
using Microsoft.Office.Interop.Outlook;
using System.Runtime.InteropServices;

namespace OutlookGoogleCalendarSync {
    class OutlookOld : OutlookInterface {
        private static readonly ILog log = LogManager.GetLogger(typeof(OutlookOld));
        
        private Application oApp;
        private String currentUserSMTP;  //SMTP of account owner that has Outlook open
        private String currentUserName;  //Name of account owner - used to determine if attendee is "self"
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
            currentUserSMTP = (oNS.CurrentUser as Recipient).Address.ToString().ToLower();
            currentUserName = oNS.CurrentUser.Name;

            //Alternate logon method that uses a specific profile.
            // If you use this logon method, 
            // change the profile name to an appropriate value.
            //oNS.Logon("YourValidProfile", Missing.Value, false, true); 

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

        private const String gEventID = "googleEventID";

        public String GetRecipientEmail(Recipient recipient) {
            String retEmail = "";
            log.Fine("Determining email of recipient: " + recipient.Name);
            if (recipient.AddressEntry == null) {
                log.Warn("No AddressEntry exists!");
                return retEmail;
            }
            log.Fine("AddressEntry Type: " + recipient.AddressEntry.Type);
            if (recipient.AddressEntry.Type == "EX") { //Exchange
                log.Fine("Address is from Exchange");
                retEmail = ADX_GetSMTPAddress(recipient.AddressEntry.Address);
            } else {
                log.Fine("Not from Exchange");
                retEmail = recipient.AddressEntry.Address;
            }
            if (retEmail == null || retEmail == String.Empty || retEmail == "" || retEmail == "Unknown") {
                log.Error("Failed to get email address through Addin MAPI access!");
                String buildFakeEmail = recipient.Name.Replace(",", "");
                buildFakeEmail = buildFakeEmail.Replace(" ", "");
                buildFakeEmail += "@unknownemail.com";
                log.Debug("Built a fake email for them: " + buildFakeEmail);
                retEmail = buildFakeEmail;
            }                
            
            log.Fine("Email address: " + retEmail);
            return retEmail;
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
        
    }
}
