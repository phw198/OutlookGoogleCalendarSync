﻿using Google.Apis.Calendar.v3.Data;
using log4net;
using Microsoft.Office.Interop.Outlook;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;

namespace OutlookGoogleCalendarSync.OutlookOgcs {
    class OutlookNew : Interface {
        private static readonly ILog log = LogManager.GetLogger(typeof(OutlookNew));
        
        private Microsoft.Office.Interop.Outlook.Application oApp;
        private ExplorerWatcher explorerWatcher;
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
                
                oNS = GetCurrentUser(oNS);

                if (!Settings.Instance.OutlookGalBlocked && currentUserName == "Unknown") {
                    log.Info("Current username is \"Unknown\"");
                    if (Settings.Instance.AddAttendees) {
                        System.Windows.Forms.MessageBox.Show("It appears you do not have an Email Account configured in Outlook.\r\n" +
                            "You should set one up now (Tools > Email Accounts) to avoid problems syncing meeting attendees.",
                            "No Email Account Found", System.Windows.Forms.MessageBoxButtons.OK,
                            System.Windows.Forms.MessageBoxIcon.Warning);
                    }
                }

                log.Debug("Get the folders configured in Outlook");
                folders = oNS.Folders;
                
                // Get the Calendar folders
                useOutlookCalendar = getCalendarStore(oNS);
                if (Forms.Main.Instance.IsHandleCreated) {
                    log.Fine("Resetting connection, so re-selecting calendar from GUI dropdown");

                    Forms.Main.Instance.cbOutlookCalendars.SelectedIndexChanged -= Forms.Main.Instance.cbOutlookCalendar_SelectedIndexChanged;
                    Forms.Main.Instance.cbOutlookCalendars.DataSource = new BindingSource(calendarFolders, null);

                    //Select the right calendar
                    int c = 0;
                    foreach (KeyValuePair<String, MAPIFolder> calendarFolder in calendarFolders) {
                        if (calendarFolder.Value.EntryID == Settings.Instance.UseOutlookCalendar.Id) {
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
                
                OutlookOgcs.Calendar.Categories = new OutlookOgcs.Categories();
                Calendar.Categories.Get(oApp, useOutlookCalendar);
                
                //Set up event handlers
                explorerWatcher = new ExplorerWatcher(oApp);
                
            } catch (System.Runtime.InteropServices.COMException ex) {
                if (OGCSexception.GetErrorCode(ex) == "0x84120009") { //Cannot complete the operation. You are not connected. [Issue #514, occurs on GetNamespace("mapi")]
                    log.Warn(ex.Message);
                    throw new ApplicationException("A problem was encountered with your Office install.\r\n" +
                            "Please perform an Office Repair or reinstall Outlook and then try running OGCS again.");
                } else throw ex;

            } finally {
                // Done. Log off.
                if (oNS != null) oNS.Logoff();
                oNS = (NameSpace)OutlookOgcs.Calendar.ReleaseObject(oNS);
            }
        }
        public void Disconnect(Boolean onlyWhenNoGUI = false) {
            if (!onlyWhenNoGUI ||
                (onlyWhenNoGUI && oApp.Explorers.Count == 0)) 
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
                    Calendar.Categories.Dispose();
                    explorerWatcher = (ExplorerWatcher)Calendar.ReleaseObject(explorerWatcher);
                } catch (System.Exception ex) {
                    log.Debug(ex.Message);
                }

                log.Info("Disconnecting from Outlook application.");
                System.Runtime.InteropServices.Marshal.FinalReleaseComObject(oApp);
                oApp = null;
                GC.Collect();
            }
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
        private const String PR_SMTP_ADDRESS = "http://schemas.microsoft.com/mapi/proptag/0x39FE001E";
        private const String EMAIL1ADDRESS = "http://schemas.microsoft.com/mapi/id/{00062004-0000-0000-C000-000000000046}/8084001F";
        private const String PR_IPM_WASTEBASKET_ENTRYID = "http://schemas.microsoft.com/mapi/proptag/0x35E30102";
        private const String PR_ORGANISER_TIMEZONE = "http://schemas.microsoft.com/mapi/id/{00062002-0000-0000-C000-000000000046}/8234001E";

        public NameSpace GetCurrentUser(NameSpace oNS) {
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
                    if (Settings.Instance.OutlookGalBlocked) { //Fail fast
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
                                if (OGCSexception.GetErrorCode(ex2) == "0x80004004") { //E_ABORT
                                    log.Warn("Corporate policy or possibly anti-virus is blocking access to GAL.");
                                } else OGCSexception.Analyse(ex2);
                                log.Warn("OGCS is unable to obtain CurrentUser from Outlook.");
                                Settings.Instance.OutlookGalBlocked = true;
                                return oNS;
                            }
                            OGCSexception.Analyse(ex2);
                        }
                        delay++;
                    }
                }
                if (Settings.Instance.OutlookGalBlocked) log.Debug("GAL is no longer blocked!");
                Settings.Instance.OutlookGalBlocked = false;

                //Issue 402
                log.Debug("Getting active window inspector");
                Inspector inspector = oApp.ActiveInspector();
                inspector = (Inspector)OutlookOgcs.Calendar.ReleaseObject(inspector);
                log.Debug("Done.");

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
            if (Settings.Instance.OutlookService == OutlookOgcs.Calendar.Service.AlternativeMailbox && Settings.Instance.MailboxName != "") {
                log.Debug("Finding Alternative Mailbox calendar folders");
                Folders binFolders = null;
                Store binStore = null;
                PropertyAccessor pa = null;
                try {
                    binFolders = oNS.Folders;
                    binStore = binFolders[Settings.Instance.MailboxName].Store;
                    pa = binStore.PropertyAccessor;
                    object bin = pa.GetProperty(PR_IPM_WASTEBASKET_ENTRYID);
                    string excludeDeletedFolder = pa.BinaryToString(bin); //EntryID

                    Forms.Main.Instance.lOutlookCalendar.Text = "Getting calendars";
                    Forms.Main.Instance.lOutlookCalendar.BackColor = System.Drawing.Color.Yellow;
                    findCalendars(oNS.Folders[Settings.Instance.MailboxName].Folders, calendarFolders, excludeDeletedFolder);
                    Forms.Main.Instance.lOutlookCalendar.BackColor = System.Drawing.Color.White;
                    Forms.Main.Instance.lOutlookCalendar.Text = "Select calendar";
                } catch (System.Exception ex) {
                    OGCSexception.Analyse("Failed to find calendar folders in alternate mailbox '" + Settings.Instance.MailboxName + "'.", ex, true);
                    if (!(Forms.Main.Instance.Visible && Forms.Main.Instance.ActiveControl.Name == "rbOutlookAltMB"))
                        throw new System.Exception("Failed to access alternate mailbox calendar.", ex);
                } finally {
                    pa = (PropertyAccessor)OutlookOgcs.Calendar.ReleaseObject(pa);
                    binStore = (Store)OutlookOgcs.Calendar.ReleaseObject(binStore);
                    binFolders = (Folders)OutlookOgcs.Calendar.ReleaseObject(binFolders);
                }

                //Default to first calendar in drop down
                defaultCalendar = calendarFolders.FirstOrDefault().Value;
                if (defaultCalendar == null) {
                    log.Info("Could not find Alternative mailbox Calendar folder. Reverting to the default mailbox calendar.");
                    System.Windows.Forms.MessageBox.Show("Unable to find a Calendar folder in the alternative mailbox.\r\n" +
                        "Reverting to the default mailbox calendar", "Calendar not found", System.Windows.Forms.MessageBoxButtons.OK);
                    getDefaultCalendar(oNS, ref defaultCalendar);
                    Forms.Main.Instance.ddMailboxName.Text = "";
                }

            } else if (Settings.Instance.OutlookService == OutlookOgcs.Calendar.Service.SharedCalendar) {
                log.Debug("Finding shared calendar");
                if (Forms.Main.Instance.Visible && Forms.Main.Instance.ActiveControl.Name == "rbOutlookSharedCal") {
                    SelectNamesDialog snd;
                    try {
                        snd = oNS.GetSelectNamesDialog();
                        snd.NumberOfRecipientSelectors = OlRecipientSelectors.olShowNone;
                        snd.ForceResolution = true;
                        snd.AllowMultipleSelection = false;
                        snd.Display();
                        if (snd.Recipients.Count == 0) {
                            log.Info("No shared calendar selected.");
                            getDefaultCalendar(oNS, ref defaultCalendar);
                        } else {
                            String sharedURI = snd.Recipients[1].Address;
                            MAPIFolder sharedCalendar = getSharedCalendar(oNS, sharedURI, true);
                            if (sharedCalendar == null) getDefaultCalendar(oNS, ref defaultCalendar);
                            else {
                                Settings.Instance.SharedCalendar = sharedURI;
                                return sharedCalendar;
                            }
                        }
                    } finally {
                        snd = null;
                    }
                } else {
                    defaultCalendar = getSharedCalendar(oNS, Settings.Instance.SharedCalendar, false);
                    return defaultCalendar;
                }

            } else {
                getDefaultCalendar(oNS, ref defaultCalendar);
            }
            log.Debug("Default Calendar folder: " + defaultCalendar.Name);
            log.Debug("Folder type: " + defaultCalendar.Store.ExchangeStoreType.ToString());
            return defaultCalendar;
        }

        private MAPIFolder getSharedCalendar(NameSpace oNS, String sharedURI, Boolean interactive) {
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
                if (interactive) {
                    String sharerName = ".";
                    if (sharer != null) sharerName = " for '" + sharer.Name + "'.";
                    MessageBox.Show("Could not find shared calendar" + sharerName, "No shared calendar found",
                        MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    return null;
                } else {
                    throw new System.Exception("Failed to access shared calendar.", ex);
                }
            } finally {
                sharer = (Recipient)OutlookOgcs.Calendar.ReleaseObject(sharer);
            }
        }

        private void getDefaultCalendar(NameSpace oNS, ref MAPIFolder defaultCalendar) {
            log.Debug("Finding default Mailbox calendar folders");
            try {
                Forms.Main.Instance.rbOutlookDefaultMB.CheckedChanged -= Forms.Main.Instance.rbOutlookDefaultMB_CheckedChanged;
                Forms.Main.Instance.rbOutlookDefaultMB.Checked = true;
                Settings.Instance.OutlookService = OutlookOgcs.Calendar.Service.DefaultMailbox;
                Forms.Main.Instance.rbOutlookDefaultMB.CheckedChanged += Forms.Main.Instance.rbOutlookDefaultMB_CheckedChanged;

                defaultCalendar = oNS.GetDefaultFolder(OlDefaultFolders.olFolderCalendar);
                calendarFolders.Add("Default " + defaultCalendar.Name, defaultCalendar);
                string excludeDeletedFolder = folders.Application.Session.GetDefaultFolder(OlDefaultFolders.olFolderDeletedItems).EntryID;

                Forms.Main.Instance.lOutlookCalendar.BackColor = System.Drawing.Color.Yellow;
                Forms.Main.Instance.lOutlookCalendar.Text = "Getting calendars";

                findCalendars(oNS.DefaultStore.GetRootFolder().Folders, calendarFolders, excludeDeletedFolder, defaultCalendar);

                Forms.Main.Instance.lOutlookCalendar.BackColor = System.Drawing.Color.White;
                Forms.Main.Instance.lOutlookCalendar.Text = "Select calendar";
            } catch (System.Exception ex) {
                OGCSexception.Analyse(ex, true);
                throw ex;
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
                    OGCSexception.Analyse(ex);
                    if (oApp.Session.ExchangeConnectionMode.ToString().Contains("Disconnected") ||
                        ex.Message.StartsWith("Network problems are preventing connection to Microsoft Exchange.") ||
                        OGCSexception.GetErrorCode(ex, 0x000FFFFF) == "0x00040115") {
                        log.Info("Currently disconnected from Exchange - unable to retrieve MAPI folders.");
                        Forms.Main.Instance.ToolTips.SetToolTip(Forms.Main.Instance.cbOutlookCalendars,
                            "The Outlook calendar to synchonize with.\nSome may not be listed as you are currently disconnected.");
                    } else {
                        log.Error("Failed to recurse MAPI folders.");
                        log.Error(ex.Message);
                        MessageBox.Show("A problem was encountered when searching for Outlook calendar folders.",
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

        public void GetAppointmentByID(String entryID, out AppointmentItem ai) {
            NameSpace ns = oApp.GetNamespace("mapi");
            ai = ns.GetItemFromID(entryID) as AppointmentItem;
            ns = (NameSpace)OutlookOgcs.Calendar.ReleaseObject(ns);
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
                        if (addressEntry.AddressEntryUserType == OlAddressEntryUserType.olExchangeUserAddressEntry ||
                            addressEntry.AddressEntryUserType == OlAddressEntryUserType.olExchangeRemoteUserAddressEntry) {
                            ExchangeUser eu = null;
                            try {
                                eu = addressEntry.GetExchangeUser();
                                if (eu != null && eu.PrimarySmtpAddress != null)
                                    retEmail = eu.PrimarySmtpAddress;
                                else {
                                    log.Warn("Exchange does not have an email for recipient: " + recipient.Name);
                                    Microsoft.Office.Interop.Outlook.PropertyAccessor pa = null;
                                    try {
                                        //Should I try PR_EMS_AB_PROXY_ADDRESSES next to cater for cached mode?
                                        pa = recipient.PropertyAccessor;
                                        retEmail = pa.GetProperty(OutlookNew.PR_SMTP_ADDRESS).ToString();
                                        log.Debug("Retrieved from PropertyAccessor instead.");
                                    } catch {
                                        log.Warn("Also failed to retrieve email from PropertyAccessor.");
                                        retEmail = EmailAddress.BuildFakeEmailAddress(recipient.Name, out builtFakeEmail);
                                    } finally {
                                        pa = (Microsoft.Office.Interop.Outlook.PropertyAccessor)OutlookOgcs.Calendar.ReleaseObject(pa);
                                    }
                                }
                            } finally {
                                eu = (ExchangeUser)OutlookOgcs.Calendar.ReleaseObject(eu);
                            }

                        } else if (addressEntry.AddressEntryUserType == OlAddressEntryUserType.olOutlookContactAddressEntry) {
                            log.Fine("This is an Outlook contact");
                            ContactItem contact = null;
                            try {
                                try {
                                    contact = addressEntry.GetContact();
                                } catch {
                                    log.Warn("Doesn't seem to be a valid contact object. Maybe this account is no longer in Exchange.");
                                    retEmail = EmailAddress.BuildFakeEmailAddress(recipient.Name, out builtFakeEmail);
                                }
                                if (contact != null) {
                                    if (contact.Email1AddressType == "EX") {
                                        log.Fine("Address is from Exchange.");
                                        log.Fine("Using PropertyAccessor to get email address.");
                                        Microsoft.Office.Interop.Outlook.PropertyAccessor pa = null;
                                        try {
                                            pa = contact.PropertyAccessor;
                                            retEmail = pa.GetProperty(EMAIL1ADDRESS).ToString();
                                        } finally {
                                            pa = (Microsoft.Office.Interop.Outlook.PropertyAccessor)OutlookOgcs.Calendar.ReleaseObject(pa);
                                        }
                                    } else {
                                        retEmail = contact.Email1Address;
                                    }
                                }
                            } finally {
                                contact = (ContactItem)OutlookOgcs.Calendar.ReleaseObject(contact);
                            }
                        } else {
                            log.Fine("Exchange type: " + addressEntry.AddressEntryUserType.ToString());
                            log.Fine("Using PropertyAccessor to get email address.");
                            Microsoft.Office.Interop.Outlook.PropertyAccessor pa = null;
                            try {
                                pa = recipient.PropertyAccessor;
                                retEmail = pa.GetProperty(OutlookNew.PR_SMTP_ADDRESS).ToString();
                            } finally {
                                pa = (Microsoft.Office.Interop.Outlook.PropertyAccessor)OutlookOgcs.Calendar.ReleaseObject(pa);
                            }
                        }

                    } else if (addressEntryType != null && addressEntryType.ToUpper() == "NOTES") {
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
                            log.Fail("Failed accessing addressEntry.Address");
                            log.Fail(ex.Message);
                            retEmail = EmailAddress.BuildFakeEmailAddress(recipient.Name, out builtFakeEmail);
                        }
                    }
                }

                if (retEmail != null && retEmail.IndexOf("<") > 0) {
                    retEmail = retEmail.Substring(retEmail.IndexOf("<") + 1);
                    retEmail = retEmail.TrimEnd(Convert.ToChar(">"));
                }
                log.Fine("Email address: " + retEmail, retEmail);
                if (!EmailAddress.IsValidEmail(retEmail) && !builtFakeEmail) {
                    retEmail = EmailAddress.BuildFakeEmailAddress(recipient.Name, out builtFakeEmail);
                    if (!EmailAddress.IsValidEmail(retEmail)) {
                        Forms.Main.Instance.Console.Update("Recipient \"" + recipient.Name + "\" with email address \"" + retEmail + "\" is invalid.<br/>" +
                            "This must be manually resolved in order to sync this appointment.", Console.Markup.error, notifyBubble: true);
                        throw new ApplicationException("Invalid recipient email for \"" + recipient.Name + "\"");
                    }
                }
                return retEmail;
            } finally {
                addressEntry = (AddressEntry)OutlookOgcs.Calendar.ReleaseObject(addressEntry);
            }
        }

        public String GetGlobalApptID(AppointmentItem ai) {
            try {
                if (ai.GlobalAppointmentID == null)
                    throw new System.Exception("GlobalAppointmentID is null - this shouldn't happen! Falling back to EntryID.");
                return ai.GlobalAppointmentID;
            } catch (System.Exception ex) {
                log.Warn(ex.Message);
                return ai.EntryID;
            }
        }

        public void RefreshCategories() {
            log.Debug("Refreshing categories...");
            OutlookOgcs.Calendar.Categories.Get(oApp, useOutlookCalendar);
            Forms.Main.Instance.ddCategoryColour.AddCategoryColours();
            foreach (Extensions.ColourPicker.ColourInfo cInfo in Forms.Main.Instance.ddCategoryColour.Items) {
                if (cInfo.OutlookCategory.ToString() == Settings.Instance.SetEntriesColourValue &&
                    cInfo.Text == Settings.Instance.SetEntriesColourName) {
                    Forms.Main.Instance.ddCategoryColour.SelectedItem = cInfo;
                }
            }
        }

        #region TimeZone Stuff
        //http://stackoverflow.com/questions/17348807/how-to-translate-between-windows-and-iana-time-zones
        //https://en.wikipedia.org/wiki/List_of_tz_database_time_zones
        //https://blogs.technet.microsoft.com/dst2007/ - MS timezone updates

        public Event IANAtimezone_set(Event ev, AppointmentItem ai) {
            String organiserTZname = null;
            String organiserTZid = null;
            try {
                if (ai.Organizer != CurrentUserName()) {
                log.Fine("Meeting organiser is someone else - checking their timezone.");
                    PropertyAccessor pa = null;
                    try {
                        pa = ai.PropertyAccessor;
                        organiserTZname = pa.GetProperty(PR_ORGANISER_TIMEZONE).ToString();
                    } finally {
                        pa = (PropertyAccessor)OutlookOgcs.Calendar.ReleaseObject(pa);
                    }
                    if (organiserTZname != ai.StartTimeZone.Name) {
                        log.Fine("Appointment's timezone: " + ai.StartTimeZone.Name);
                        log.Fine("Organiser's timezone:   " + organiserTZname);
                        if (organiserTZname == "Customized Time Zone") {
                            System.Exception ex = new System.Exception("Cannot translate " + organiserTZname + " to a timezone ID.");
                            ex.Data.Add("OGCS", "");
                            throw ex;
                        } 
                        log.Debug("Retrieving the meeting organiser's timezone ID.");
                        TimeZoneInfo tzi = getWindowsTimezoneFromDescription(organiserTZname);
                        if (tzi == null) log.Error("No timezone ID exists for organiser's timezone " + organiserTZname);
                        else organiserTZid = tzi.Id;
                    }
                }
            } catch (System.Exception ex) {
                Forms.Main.Instance.Console.Update(OutlookOgcs.Calendar.GetEventSummary(ai) +
                    "<br/>Could not determine the organiser's timezone. Google Event may have incorrect time.", Console.Markup.warning);
                if (ex.Data.Contains("OGCS")) log.Warn(ex.Message);
                else OGCSexception.Analyse(ex);
                organiserTZname = null;
                organiserTZid = null;
            }

            try {
                try {
                    ev.Start.TimeZone = IANAtimezone(organiserTZid ?? ai.StartTimeZone.ID, organiserTZname ?? ai.StartTimeZone.Name);
                } catch (System.Exception ex) {
                    log.Debug(ex.Message);
                    throw new ApplicationException("Failed to set start timezone. [" + ai.StartTimeZone.ID + ", " + ai.StartTimeZone.Name + "]");
                }
                try {
                    ev.End.TimeZone = IANAtimezone(organiserTZid ?? ai.EndTimeZone.ID, organiserTZname ?? ai.EndTimeZone.Name);
                } catch (System.Exception ex) {
                    log.Debug(ex.Message);
                    throw new ApplicationException("Failed to set end timezone. [" + ai.EndTimeZone.ID + ", " + ai.EndTimeZone.Name + "]");
                }
            } catch (ApplicationException ex) {
                log.Warn(ex.Message);
            }
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
            TimeZoneInfo tzi = TimeZoneInfo.FindSystemTimeZoneById(oTZ_id);
            String tzID = tzDBsource.MapTimeZoneId(tzi);
            log.Fine("Timezone \"" + oTZ_name + "\" mapped to \"" + tzDBsource.CanonicalIdMap[tzID] + "\"");
            return tzDBsource.CanonicalIdMap[tzID];
        }

        public void WindowsTimeZone_get(AppointmentItem ai, out String startTz, out String endTz) {
            Microsoft.Office.Interop.Outlook.TimeZone _startTz = null;
            Microsoft.Office.Interop.Outlook.TimeZone _endTz = null;
            try {
                _startTz = ai.StartTimeZone;
                _endTz = ai.EndTimeZone;
                startTz = _startTz.ID;
                endTz = _endTz.ID;
            } finally {
                _startTz = (Microsoft.Office.Interop.Outlook.TimeZone)OutlookOgcs.Calendar.ReleaseObject(_startTz);
                _endTz = (Microsoft.Office.Interop.Outlook.TimeZone)OutlookOgcs.Calendar.ReleaseObject(_endTz);
            }
        }

        public AppointmentItem WindowsTimeZone_set(AppointmentItem ai, Event ev, String attr = "Both", Boolean onlyTZattribute = false) {
            if ("Both,Start".Contains(attr)) {
                if (!String.IsNullOrEmpty(ev.Start.TimeZone)) {
                    log.Fine("Has starting timezone: " + ev.Start.TimeZone);
                    ai.StartTimeZone = WindowsTimeZone(ev.Start.TimeZone);
                }
                if (!onlyTZattribute) ai.Start = ev.Start.DateTime ?? DateTime.Parse(ev.Start.Date);
            }
            if ("Both,End".Contains(attr)) {
                if (!String.IsNullOrEmpty(ev.End.TimeZone)) {
                    log.Fine("Has ending timezone: " + ev.End.TimeZone);
                    ai.EndTimeZone = WindowsTimeZone(ev.End.TimeZone);
                }
                if (!onlyTZattribute) ai.End = ev.End.DateTime ?? DateTime.Parse(ev.End.Date);
            }
            return ai;
        }

        private Microsoft.Office.Interop.Outlook.TimeZone WindowsTimeZone(string ianaZoneId) {
            ianaZoneId = TimezoneDB.FixAlexa(ianaZoneId);

            Microsoft.Office.Interop.Outlook.TimeZones tzs = oApp.TimeZones;
            var utcZones = new[] { "Etc/UTC", "Etc/UCT", "UTC", "Etc/GMT" };
            if (utcZones.Contains(ianaZoneId, StringComparer.OrdinalIgnoreCase)) {
                log.Fine("Timezone \"" + ianaZoneId + "\" mapped to \"UTC\"");
                return tzs["UTC"];
            }

            NodaTime.TimeZones.TzdbDateTimeZoneSource tzDBsource = TimezoneDB.Instance.Source;
            
            // resolve any link, since the CLDR doesn't necessarily use canonical IDs
            var links = tzDBsource.CanonicalIdMap
              .Where(x => x.Value.Equals(ianaZoneId, StringComparison.OrdinalIgnoreCase))
              .Select(x => x.Key);

            // resolve canonical zones, and include original zone as well
            var possibleZones = tzDBsource.CanonicalIdMap.ContainsKey(ianaZoneId)
                ? links.Concat(new[] { tzDBsource.CanonicalIdMap[ianaZoneId], ianaZoneId })
                : links;

            // map the windows zone
            var mappings = tzDBsource.WindowsMapping.MapZones;
            var item = mappings.FirstOrDefault(x => x.TzdbIds.Any(possibleZones.Contains));
            if (item == null) {
                throw new System.ApplicationException("Timezone \"" + ianaZoneId + "\" has no mapping.");
            }
            log.Fine("Timezone \"" + ianaZoneId + "\" mapped to \"" + item.WindowsId + "\"");

            return tzs[item.WindowsId];
        }

        private TimeZoneInfo getWindowsTimezoneFromDescription(String tzDescription) {
            try {
                System.Collections.ObjectModel.ReadOnlyCollection<TimeZoneInfo> sysTZ = TimeZoneInfo.GetSystemTimeZones();
                
                //First let's just search with what we've got
                TimeZoneInfo tzi = sysTZ.FirstOrDefault(t => t.DisplayName == tzDescription|| t.StandardName == tzDescription || t.Id == tzDescription);
                if (tzi != null) return tzi;

                log.Warn("Could not find timezone ID based on given description. Attempting some fuzzy logic...");
                if (tzDescription.StartsWith("(GMT")) {
                    log.Fine("Replace GMT with UTC");
                    String modTzDescription = tzDescription.Replace("(GMT", "(UTC");
                    tzi = sysTZ.FirstOrDefault(t => t.DisplayName == modTzDescription || t.StandardName == modTzDescription || t.Id == modTzDescription);
                    if (tzi != null) return tzi;

                    log.Fine("Removing offset prefix");
                    modTzDescription = System.Text.RegularExpressions.Regex.Replace(modTzDescription, @"^\(UTC[+-]\d{1,2}:\d{0,2}\)\s+", "").Trim();
                    tzi = sysTZ.FirstOrDefault(t => t.StandardName == modTzDescription || t.Id == modTzDescription);
                    if (tzi != null) return tzi;
                }


                //Try searching just by timezone offset. This would at least get the right time for the appointment, eg if the tzDescription doesn't match
                //because they it is in a different language that the user's system data.
                Int16? offset = null;
                offset = TimezoneDB.GetTimezoneOffset(tzDescription);
                if (offset != null) {
                    List<TimeZoneInfo> tzis = sysTZ.Where(t => t.BaseUtcOffset.Hours == offset).ToList();
                    if (tzis.Count == 0)
                        log.Warn("No timezone ID exists for organiser's GMT offset timezone " + tzDescription);
                    else if (tzis.Count == 1)
                        return tzis.First();
                    else {
                        String tzCountry = tzDescription.Substring(tzDescription.LastIndexOf("/") + 1);
                        if (string.IsNullOrEmpty(tzCountry)) {
                            log.Warn("Could not determine country; and multiple timezones exist with same GMT offset of " + offset + ". Picking the first.");
                            return tzis.FirstOrDefault();
                        } else {
                            List<TimeZoneInfo> countryTzis = tzis.Where(t => t.DisplayName.Contains(tzCountry)).ToList();
                            if (countryTzis.Count == 0) {
                                log.Warn("Could not find timezone with GMT offset of " + offset + " for country " + tzCountry + ". Picking the first offset match regardless of country.");
                                return tzis.FirstOrDefault();
                            } else if (countryTzis.Count == 1)
                                return countryTzis.First();
                            else {
                                log.Warn("Could not find unique timezone with GMT offset of " + offset + " for country " + tzCountry + ". Picking the first.");
                                return countryTzis.FirstOrDefault();
                            }
                        }
                    }

                } else {
                    //Check if it's already an IANA value
                    NodaTime.TimeZones.TzdbDateTimeZoneSource tzDBsource = TimezoneDB.Instance.Source;
                    IEnumerable<NodaTime.TimeZones.TzdbZoneLocation> a = tzDBsource.ZoneLocations.Where(l => l.ZoneId == tzDescription);
                    if (a.Count() >= 1) {
                        log.Debug("It appears to be an IANA timezone already!");
                        Microsoft.Office.Interop.Outlook.TimeZone tz = WindowsTimeZone(tzDescription);
                        if (tz != null)
                            return TimeZoneInfo.FindSystemTimeZoneById(tz.ID);
                    }
                }

            } catch (System.Exception ex) {
                log.Warn("Failed to get the organiser's timezone ID for " + tzDescription);
                OGCSexception.Analyse(ex);
            }
            return null;
        }

        #endregion
    }
}
