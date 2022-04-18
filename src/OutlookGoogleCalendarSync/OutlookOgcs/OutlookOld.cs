using Google.Apis.Calendar.v3.Data;
using log4net;
using Microsoft.Office.Interop.Outlook;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace OutlookGoogleCalendarSync.OutlookOgcs {
    class OutlookOld : Interface {
        private static readonly ILog log = LogManager.GetLogger(typeof(OutlookOld));

        private Microsoft.Office.Interop.Outlook.Application oApp;
        private String currentUserSMTP;  //SMTP of account owner that has Outlook open
        private String currentUserName;  //Name of account owner - used to determine if attendee is "self"
        private Folders folders;
        private MAPIFolder useOutlookCalendar;
        private Dictionary<string, MAPIFolder> calendarFolders = new Dictionary<string, MAPIFolder>();
        private OlExchangeConnectionMode exchangeConnectionMode;

        public void Connect() {
            if (!OutlookOgcs.Calendar.InstanceConnect) return;

            OutlookOgcs.Calendar.AttachToOutlook(ref oApp, openOutlookOnFail: true, withSystemCall: false);
            log.Debug("Setting up Outlook connection.");

            // Get the NameSpace and Logon information.
            NameSpace oNS = null;
            try {
                oNS = oApp.GetNamespace("mapi");

                //Implicit logon to default profile, with no dialog box
                //If 1< profile, a dialogue is forced unless implicit login used
                exchangeConnectionMode = oNS.ExchangeConnectionMode;
                if (exchangeConnectionMode != OlExchangeConnectionMode.olNoExchange) {
                    log.Info("Exchange server version: Unknown");
                }
                log.Info("Exchange connection mode: " + exchangeConnectionMode.ToString());

                oNS = GetCurrentUser(oNS);
                SettingsStore.Calendar profile = Settings.Profile.InPlay();

                if (!profile.OutlookGalBlocked && currentUserName == "Unknown") {
                    log.Info("Current username is \"Unknown\"");
                    if (profile.AddAttendees) {
                        System.Windows.Forms.OgcsMessageBox.Show("It appears you do not have an Email Account configured in Outlook.\r\n" +
                            "You should set one up now (Tools > Email Accounts) to avoid problems syncing meeting attendees.",
                            "No Email Account Found", System.Windows.Forms.MessageBoxButtons.OK,
                            System.Windows.Forms.MessageBoxIcon.Warning);
                    }
                }

                log.Debug("Get the folders configured in Outlook");
                folders = oNS.Folders;

                // Get the Calendar folders
                useOutlookCalendar = getCalendarStore(oNS);
                if (Forms.Main.Instance.IsHandleCreated && profile.Equals(Forms.Main.Instance.ActiveCalendarProfile)) {
                    log.Fine("Resetting connection, so re-selecting calendar from GUI dropdown");

                    Forms.Main.Instance.cbOutlookCalendars.SelectedIndexChanged -= Forms.Main.Instance.cbOutlookCalendar_SelectedIndexChanged;
                    Forms.Main.Instance.cbOutlookCalendars.DataSource = new BindingSource(calendarFolders, null);

                    //Select the right calendar
                    int c = 0;
                    foreach (KeyValuePair<String, MAPIFolder> calendarFolder in calendarFolders) {
                        if (calendarFolder.Value.EntryID == profile.UseOutlookCalendar.Id) {
                            Forms.Main.Instance.SetControlPropertyThreadSafe(Forms.Main.Instance.cbOutlookCalendars, "SelectedIndex", c);
                        }
                        c++;
                    }
                    if ((int)Forms.Main.Instance.GetControlPropertyThreadSafe(Forms.Main.Instance.cbOutlookCalendars, "SelectedIndex") == -1)
                        Forms.Main.Instance.SetControlPropertyThreadSafe(Forms.Main.Instance.cbOutlookCalendars, "SelectedIndex", 0);

                    KeyValuePair<String, MAPIFolder> calendar = (KeyValuePair<String, MAPIFolder>)Forms.Main.Instance.GetControlPropertyThreadSafe(Forms.Main.Instance.cbOutlookCalendars, "SelectedItem");
                    useOutlookCalendar = calendar.Value;

                    Forms.Main.Instance.cbOutlookCalendars.SelectedIndexChanged += Forms.Main.Instance.cbOutlookCalendar_SelectedIndexChanged;
                }

                OutlookOgcs.Calendar.Categories = null;

            } finally {
                // Done. Log off.
                if (oNS != null) oNS.Logoff();
                oNS = (NameSpace)OutlookOgcs.Calendar.ReleaseObject(oNS);
            }
        }
        public void Disconnect(Boolean onlyWhenNoGUI = false) {
            if (Settings.Instance.DisconnectOutlookBetweenSync ||
                !onlyWhenNoGUI ||
                (onlyWhenNoGUI && NoGUIexists()))
            {
                log.Debug("De-referencing all Outlook application objects.");
                try {
                    folders = (Folders)OutlookOgcs.Calendar.ReleaseObject(folders);
                    useOutlookCalendar = (MAPIFolder)OutlookOgcs.Calendar.ReleaseObject(useOutlookCalendar);
                    for (int fld = calendarFolders.Count - 1; fld >= 0; fld--) {
                        MAPIFolder mFld = calendarFolders.ElementAt(fld).Value;
                        mFld = (MAPIFolder)OutlookOgcs.Calendar.ReleaseObject(mFld);
                        calendarFolders.Remove(calendarFolders.ElementAt(fld).Key);
                    }
                    calendarFolders = new Dictionary<string, MAPIFolder>();
                } catch (System.Exception ex) {
                    log.Debug(ex.Message);
                }

                log.Info("Disconnecting from Outlook application.");
                System.Runtime.InteropServices.Marshal.FinalReleaseComObject(oApp);
                oApp = null;
                GC.Collect();
            }
        }

        public Boolean NoGUIexists() {
            Boolean retVal = (oApp == null);
            if (!retVal) {
                Explorers explorers = null;
                try {
                    explorers = oApp.Explorers;
                    retVal = (explorers.Count == 0);
                } catch (System.Exception) {
                    if (System.Diagnostics.Process.GetProcessesByName("OUTLOOK").Count() == 0) {
                        log.Fine("No running outlook.exe process found.");
                        retVal = true;
                    } else {
                        OutlookOgcs.Calendar.AttachToOutlook(ref oApp, openOutlookOnFail: false);
                        try {
                            explorers = oApp.Explorers;
                            retVal = (explorers.Count == 0);
                        } catch {
                            log.Warn("Failed to reattach to Outlook instance.");
                            retVal = true;
                        }
                    }
                } finally {
                    explorers = (Explorers)OutlookOgcs.Calendar.ReleaseObject(explorers);
                }
            }
            if (retVal) log.Fine("No Outlook GUI detected.");
            return retVal;
        }

        public Folders Folders() { return folders; }
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
            if (string.IsNullOrEmpty(currentUserName)) {
                GetCurrentUser(null);
            }
            return currentUserName;
        }
        public Boolean Offline() {
            try {
                return oApp.GetNamespace("mapi").Offline;
            } catch {
                OutlookOgcs.Calendar.Instance.Reset();
                return false;
            }
        }
        public OlExchangeConnectionMode ExchangeConnectionMode() {
            return exchangeConnectionMode;
        }

        private const String gEventID = "googleEventID";
        private const String PR_IPM_WASTEBASKET_ENTRYID = "http://schemas.microsoft.com/mapi/proptag/0x35E30102";

        public NameSpace GetCurrentUser(NameSpace oNS) {
            SettingsStore.Calendar profile = Settings.Profile.InPlay();

            //We only need the current user details when syncing meeting attendees.
            //If GAL had previously been detected as blocked, let's always try one attempt to see if it's been opened up
            if (!profile.OutlookGalBlocked && !profile.AddAttendees) return oNS;

            Boolean releaseNamespace = (oNS == null);
            if (releaseNamespace) oNS = oApp.GetNamespace("mapi");

            Recipient currentUser = null;
            try {
                DateTime triggerOOMsecurity = DateTime.Now;
                try {
                    currentUser = oNS.CurrentUser;
                    if (!Forms.Main.Instance.IsHandleCreated && (DateTime.Now - triggerOOMsecurity).TotalSeconds > 1) {
                        log.Warn(">1s delay possibly due to Outlook security popup.");
                        OutlookOgcs.Calendar.OOMsecurityInfo = true;
                    }
                } catch (System.Exception ex) {
                    OGCSexception.Analyse(ex);
                    if (profile.OutlookGalBlocked) { //Fail fast
                        log.Debug("Corporate policy is still blocking access to GAL.");
                        return oNS;
                    }
                    log.Warn("We seem to have a faux connection to Outlook! Forcing starting it with a system call :-/");
                    oNS = (NameSpace)OutlookOgcs.Calendar.ReleaseObject(oNS);
                    Disconnect();
                    OutlookOgcs.Calendar.AttachToOutlook(ref oApp, openOutlookOnFail: true, withSystemCall: true);
                    oNS = oApp.GetNamespace("mapi");

                    int maxDelay = 5;
                    int delay = 1;
                    while (delay <= maxDelay) {
                        log.Debug("Sleeping..." + delay + "/" + maxDelay);
                        System.Threading.Thread.Sleep(10000);
                        try {
                            currentUser = oNS.CurrentUser;
                            delay = maxDelay;
                        } catch (System.Exception ex2) {
                            if (delay == maxDelay) {
                                log.Warn("OGCS is unable to obtain CurrentUser from Outlook.");
                                if (OGCSexception.GetErrorCode(ex2) == "0x80004004") { //E_ABORT
                                    log.Warn("Corporate policy or possibly anti-virus is blocking access to GAL.");
                                } else OGCSexception.Analyse(OGCSexception.LogAsFail(ex2));
                                profile.OutlookGalBlocked = true;
                                return oNS;
                            }
                            OGCSexception.Analyse(ex2);
                        }
                        delay++;
                    }
                }
                if (profile.OutlookGalBlocked) log.Debug("GAL is no longer blocked!");
                profile.OutlookGalBlocked = false;
                currentUserSMTP = GetRecipientEmail(currentUser);
                currentUserName = currentUser.Name;
            } finally {
                currentUser = (Recipient)OutlookOgcs.Calendar.ReleaseObject(currentUser);
                if (releaseNamespace) oNS = (NameSpace)OutlookOgcs.Calendar.ReleaseObject(oNS);
            }
            return oNS;
        }

        private MAPIFolder getCalendarStore(NameSpace oNS) {
            MAPIFolder defaultCalendar = null;
            SettingsStore.Calendar profile = Settings.Profile.InPlay();
            if (profile.OutlookService == OutlookOgcs.Calendar.Service.DefaultMailbox) {
                getDefaultCalendar(oNS, ref defaultCalendar);
            }
            log.Debug("Default Calendar folder: " + defaultCalendar.Name);
            return defaultCalendar;
        }

        private MAPIFolder getSharedCalendar(NameSpace oNS, String sharedURI) {
            if (string.IsNullOrEmpty(sharedURI)) return null;

            Recipient sharer = null;
            MAPIFolder sharedCalendar = null;
            try {
                sharer = oNS.CreateRecipient(sharedURI);
                sharer.Resolve();
                if (sharer.DisplayType == OlDisplayType.olDistList)
                    throw new System.Exception("User selected a distribution list!");

                sharedCalendar = oNS.GetSharedDefaultFolder(sharer, OlDefaultFolders.olFolderCalendar);
                if (sharedCalendar.DefaultItemType != OlItemType.olAppointmentItem) {
                    log.Debug(sharer.Name + " does not have a calendar shared.");
                    throw new System.Exception("Wrong default item type.");
                }
                calendarFolders.Add(sharer.Name, sharedCalendar);
                return sharedCalendar;

            } catch (System.Exception ex) {
                log.Error("Failed to get shared calendar from " + sharedURI + ". " + ex.Message);
                OgcsMessageBox.Show("Could not find a shared calendar for '" + sharer.Name + "'.", "No shared calendar found",
                        MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return null;
            } finally {
                sharer = (Recipient)OutlookOgcs.Calendar.ReleaseObject(sharer);
            }
        }

        private void getDefaultCalendar(NameSpace oNS, ref MAPIFolder defaultCalendar) {
            log.Debug("Finding default Mailbox calendar folders");
            try {
                SettingsStore.Calendar profile = Settings.Profile.InPlay();
                Boolean updateGUI = profile.Equals(Forms.Main.Instance.ActiveCalendarProfile);
                if (updateGUI) {
                    Forms.Main.Instance.rbOutlookDefaultMB.CheckedChanged -= Forms.Main.Instance.rbOutlookDefaultMB_CheckedChanged;
                    Forms.Main.Instance.rbOutlookDefaultMB.Checked = true;
                }
                profile.OutlookService = OutlookOgcs.Calendar.Service.DefaultMailbox;
                if (updateGUI)
                    Forms.Main.Instance.rbOutlookDefaultMB.CheckedChanged += Forms.Main.Instance.rbOutlookDefaultMB_CheckedChanged;

                defaultCalendar = oNS.GetDefaultFolder(OlDefaultFolders.olFolderCalendar);
                calendarFolders.Add("Default " + defaultCalendar.Name, defaultCalendar);
                string excludeDeletedFolder = folders.Application.Session.GetDefaultFolder(OlDefaultFolders.olFolderDeletedItems).EntryID;

                if (updateGUI) {
                    Forms.Main.Instance.lOutlookCalendar.BackColor = System.Drawing.Color.Yellow;
                    Forms.Main.Instance.SetControlPropertyThreadSafe(Forms.Main.Instance.lOutlookCalendar, "Text", "Getting calendars");
                }
                findCalendars(((MAPIFolder)defaultCalendar.Parent).Folders, calendarFolders, excludeDeletedFolder, defaultCalendar);
                if (updateGUI) {
                    Forms.Main.Instance.lOutlookCalendar.BackColor = System.Drawing.Color.White;
                    Forms.Main.Instance.SetControlPropertyThreadSafe(Forms.Main.Instance.lOutlookCalendar, "Text", "Select calendar");
                }
            } catch (System.Exception ex) {
                OGCSexception.Analyse(ex, true);
                throw;
            }
        }

        private void findCalendars(Folders folders, Dictionary<string, MAPIFolder> calendarFolders, String excludeDeletedFolder, MAPIFolder defaultCalendar = null) {
            //Initiate progress bar (red line underneath "Getting calendars" text)
            System.Drawing.Graphics g = Forms.Main.Instance.tabOutlook.CreateGraphics();
            System.Drawing.Pen p = new System.Drawing.Pen(System.Drawing.Color.Red, 3);
            System.Drawing.Point startPoint = new System.Drawing.Point(Forms.Main.Instance.lOutlookCalendar.Location.X,
                Forms.Main.Instance.lOutlookCalendar.Location.Y + Forms.Main.Instance.lOutlookCalendar.Size.Height + 3);
            double stepSize = Forms.Main.Instance.lOutlookCalendar.Size.Width / folders.Count;

            int fldCnt = 0;
            foreach (MAPIFolder folder in folders) {
                fldCnt++;
                System.Drawing.Point endPoint = new System.Drawing.Point(Forms.Main.Instance.lOutlookCalendar.Location.X + Convert.ToInt16(fldCnt * stepSize),
                    Forms.Main.Instance.lOutlookCalendar.Location.Y + Forms.Main.Instance.lOutlookCalendar.Size.Height + 3);
                try { g.DrawLine(p, startPoint, endPoint); } catch { /*May get GDI+ error if g has been repainted*/ }
                System.Windows.Forms.Application.DoEvents();
                try {
                    OlItemType defaultItemType = folder.DefaultItemType;
                    if (defaultItemType == OlItemType.olAppointmentItem) {
                        if (defaultCalendar == null || (folder.EntryID != defaultCalendar.EntryID))
                            calendarFolderAdd(folder.Name, folder);
                    }
                    if (folder.EntryID != excludeDeletedFolder && folder.Folders.Count > 0) {
                        findCalendars(folder.Folders, calendarFolders, excludeDeletedFolder, defaultCalendar);
                    }
                } catch (System.Exception ex) {
                    if (oApp.Session.ExchangeConnectionMode.ToString().Contains("Disconnected") ||
                        OGCSexception.GetErrorCode(ex) == "0xC204011D" || ex.Message.StartsWith("Network problems are preventing connection to Microsoft Exchange.") ||
                        OGCSexception.GetErrorCode(ex, 0x000FFFFF) == "0x00040115") {
                        log.Warn(ex.Message);
                        log.Info("Currently disconnected from Exchange - unable to retrieve MAPI folders.");
                        Forms.Main.Instance.ToolTips.SetToolTip(Forms.Main.Instance.cbOutlookCalendars,
                            "The Outlook calendar to synchonize with.\nSome may not be listed as you are currently disconnected.");
                    } else {
                        OGCSexception.Analyse("Failed to recurse MAPI folders.", ex);
                        OgcsMessageBox.Show("A problem was encountered when searching for Outlook calendar folders.\r\n" + ex.Message,
                            "Calendar Folders", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    }
                }
            }
            p.Dispose();
            try { g.Clear(System.Drawing.Color.White); } catch { }
            g.Dispose();
            System.Windows.Forms.Application.DoEvents();
        }

        /// <summary>
        /// Handles non-unique calendar names by recursively adding parent folders to the name
        /// </summary>
        /// <param name="name">Name/path of calendar (dictionary key)</param>
        /// <param name="folder">The target folder (dictionary value)</param>
        /// <param name="parentFolder">Recursive parent folder - leave null on initial call</param>
        private void calendarFolderAdd(String name, MAPIFolder folder, MAPIFolder parentFolder = null) {
            try {
                calendarFolders.Add(name, folder);
            } catch (System.ArgumentException ex) {
                if (OGCSexception.GetErrorCode(ex) == "0x80070057") {
                    //An item with the same key has already been added.
                    //Let's recurse up to the parent folder, looking to make it unique
                    object parentObj = (parentFolder != null ? parentFolder.Parent : folder.Parent);
                    if (parentObj is NameSpace) {
                        //We've traversed all the way up the folder path to the root and still not unique
                        log.Warn("MAPIFolder " + name + " does not have a unique name - so cannot use!");
                    } else if (parentObj is MAPIFolder) {
                        String parentFolderName = (parentObj as MAPIFolder).FolderPath.Split('\\').Last();
                        calendarFolderAdd(System.IO.Path.Combine(parentFolderName, name), folder, parentObj as MAPIFolder);
                    }
                }
            }
        }

        public List<Object> FilterItems(Items outlookItems, String filter) {
            List<Object> restrictedItems = new List<Object>();
            foreach (Object obj in outlookItems.Restrict(filter)) {
                restrictedItems.Add(obj);
            }

            // Recurring items with start dates before the synced date range are excluded incorrectly - this retains them
            List<Object> o2003recurring = new List<Object>();
            try {
                for (int i = 1; i <= outlookItems.Count; i++) {
                    AppointmentItem ai = null;
                    if (outlookItems[i] is AppointmentItem) {
                        ai = outlookItems[i] as AppointmentItem;
                        SettingsStore.Calendar profile = Settings.Profile.InPlay();
                        if (ai.IsRecurring && ai.Start.Date < profile.SyncStart && ai.End.Date < profile.SyncStart)
                            o2003recurring.Add(outlookItems[i]);
                    }
                }
                log.Info(o2003recurring.Count + " recurring items successfully kept for Outlook 2003.");
            } catch (System.Exception ex) {
                OGCSexception.Analyse("Unable to iterate Outlook items.", ex);
            }
            restrictedItems.AddRange(o2003recurring);

            return restrictedItems;
        }

        public MAPIFolder GetFolderByID(String entryID) {
            NameSpace ns = null;
            try {
                ns = oApp.GetNamespace("mapi");
                return ns.GetFolderFromID(entryID);
            } finally {
                ns = (NameSpace)OutlookOgcs.Calendar.ReleaseObject(ns);
            }
        }
        
        public void GetAppointmentByID(String entryID, out AppointmentItem ai) {
            NameSpace ns = null;
            try {
                oApp.GetNamespace("mapi");
                ai = ns.GetItemFromID(entryID) as AppointmentItem;
            } finally {
                ns = (NameSpace)OutlookOgcs.Calendar.ReleaseObject(ns);
            }
        }

        public String GetRecipientEmail(Recipient recipient) {
            String retEmail = "";
            Boolean builtFakeEmail = false;

            log.Fine("Determining email of recipient: " + recipient.Name);
            AddressEntry addressEntry = null;
            String addressEntryType = "";
            try {
                try {
                    addressEntry = recipient.AddressEntry;
                } catch {
                    log.Warn("Can't resolve this recipient!");
                    addressEntry = null;
                }
                if (addressEntry == null) {
                    log.Warn("No AddressEntry exists!");
                    retEmail = EmailAddress.BuildFakeEmailAddress(recipient.Name, out builtFakeEmail);
                } else {
                    try {
                        addressEntryType = addressEntry.Type;
                    } catch {
                        log.Warn("Cannot access addressEntry.Type!");
                    }
                    log.Fine("AddressEntry Type: " + addressEntryType);
                    if (addressEntryType == "EX") { //Exchange
                        log.Fine("Address is from Exchange");
                        retEmail = ADX_GetSMTPAddress(addressEntry.Address);
                    } else if (addressEntry.Type != null && addressEntry.Type.ToUpper() == "NOTES") {
                        log.Fine("From Lotus Notes");
                        //Migrated contacts from notes, have weird "email addresses" eg: "James T. Kirk/US-Corp03/enterprise/US"
                        retEmail = EmailAddress.BuildFakeEmailAddress(recipient.Name, out builtFakeEmail);

                    } else {
                        log.Fine("Not from Exchange");
                        try {
                            if (string.IsNullOrEmpty(addressEntry.Address)) {
                                log.Warn("addressEntry.Address is empty.");
                                retEmail = EmailAddress.BuildFakeEmailAddress(recipient.Name, out builtFakeEmail);
                            } else {
                                retEmail = addressEntry.Address;
                            }
                        } catch (System.Exception ex) {
                            log.Error("Failed accessing addressEntry.Address");
                            log.Error(ex.Message);
                            retEmail = EmailAddress.BuildFakeEmailAddress(recipient.Name, out builtFakeEmail);
                        }
                    }
                }

                if (string.IsNullOrEmpty(retEmail) || retEmail == "Unknown") {
                    log.Error("Failed to get email address through Addin MAPI access!");
                    retEmail = EmailAddress.BuildFakeEmailAddress(recipient.Name, out builtFakeEmail);
                }

                if (retEmail != null && retEmail.IndexOf("<") > 0) {
                    retEmail = retEmail.Substring(retEmail.IndexOf("<") + 1);
                    retEmail = retEmail.TrimEnd(Convert.ToChar(">"));
                }
                log.Fine("Email address: " + retEmail, retEmail);
                if (!EmailAddress.IsValidEmail(retEmail) && !builtFakeEmail) {
                    retEmail = EmailAddress.BuildFakeEmailAddress(recipient.Name, out builtFakeEmail);
                    if (!EmailAddress.IsValidEmail(retEmail)) {
                        Forms.Main.Instance.Console.Update("Recipient \"" + recipient.Name + "\" with email address \"" + retEmail + "\" is invalid.", Console.Markup.error, notifyBubble: true);
                        Forms.Main.Instance.Console.Update("This must be manually resolved in order to sync this appointment.");
                        throw new ApplicationException("Invalid recipient email for \"" + recipient.Name + "\"");
                    }
                }
                return retEmail;
            } finally {
                addressEntry = (AddressEntry)OutlookOgcs.Calendar.ReleaseObject(addressEntry);
            }
        }

        public String GetGlobalApptID(AppointmentItem ai) {
            return ai.EntryID;
        }

        public void RefreshCategories() { }

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
                                        } finally {
                                            if (propAddressPtr != IntPtr.Zero)
                                                Marshal.Release(propAddressPtr);
                                            if (addrEntryPtr != IntPtr.Zero)
                                                Marshal.Release(addrEntryPtr);
                                        }
                                }
                            }
                        }
                    } finally {
                        Marshal.FreeHGlobal(szPtr);
                        Marshal.FreeHGlobal(propValuePtr);
                        Marshal.FreeHGlobal(adrListPtr);
                    }
                } finally {
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
                                        } finally {
                                            Marshal.Release(addrBookPtr);
                                        }
                                }
                            } finally {
                                Marshal.ReleaseComObject(sessionObj);
                            }
                    } finally {
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
                log.Fine("Timezone \"" + oTZ_name + "\" mapped to \"Etc/UTC\"");
                return "Etc/UTC";
            }

            NodaTime.TimeZones.TzdbDateTimeZoneSource tzDBsource = TimezoneDB.Instance.Source;
            String retVal = null;
            if (!tzDBsource.WindowsMapping.PrimaryMapping.TryGetValue(oTZ_id, out retVal) || retVal == null)
                log.Fail("Could not find mapping for \"" + oTZ_name + "\"");
            else {
                retVal = tzDBsource.CanonicalIdMap[retVal];
                log.Fine("Timezone \"" + oTZ_name + "\" mapped to \"" + retVal + "\"");
            }
            return retVal;
        }

        public void WindowsTimeZone_get(AppointmentItem ai, out String startTz, out String endTz) {
            startTz = "UTC";
            endTz = "UTC";
        }

        public AppointmentItem WindowsTimeZone_set(AppointmentItem ai, Event ev, String attr = "Both", Boolean onlyTZattribute = false) {
            ai.Start = WindowsTimeZone(ev.Start);
            ai.End = WindowsTimeZone(ev.End);
            return ai;
        }

        private DateTime WindowsTimeZone(EventDateTime time) {
            DateTime theDate = time.DateTime ?? DateTime.Parse(time.Date);
            /*if (time.TimeZone == null)*/
            return theDate;

            /*Issue #713: It appears Outlook will calculate the UTC time itself, based on the system's timezone
            LocalDateTime local = new LocalDateTime(theDate.Year, theDate.Month, theDate.Day, theDate.Hour, theDate.Minute);
            DateTimeZone zone = DateTimeZoneProviders.Tzdb[TimezoneDB.FixAlexa(time.TimeZone)];
            ZonedDateTime zonedTime = local.InZoneLeniently(zone);
            DateTime zonedUTC = zonedTime.ToDateTimeUtc();
            log.Fine("IANA Timezone \"" + time.TimeZone + "\" mapped to \""+ zone.Id.ToString() +"\" with a UTC of "+ zonedUTC.ToString("dd/MM/yyyy HH:mm:ss"));
            return zonedUTC;
            */
        }
        #endregion
    }
}
