using log4net;
using System;
using System.IO;
using System.Windows.Forms;
using System.Xml.Linq;

namespace OutlookGoogleCalendarSync.SettingsStore {
    public static class Upgrade {
        private static readonly ILog log = LogManager.GetLogger(typeof(Upgrade));

        //OGCS releases that require the settings XML to be upgraded
        private static Int32 multipleCalendars = Program.VersionToInt("2.9.3.0");

        private static String settingsVersion;
        private static Int32 settingsVersionNum;


        public static void Check() {
            settingsVersion = XMLManager.ImportElement("Version", Settings.ConfigFile);
            settingsVersionNum = Program.VersionToInt(settingsVersion);

            while (upgradePerformed()) {
            }
        }


        private static Boolean upgradePerformed() {
            try {
                if (settingsVersionNum > 0 && settingsVersionNum < multipleCalendars) {
                    upgradeToMultiCalendar();
                    settingsVersionNum = multipleCalendars;
                    return true;
                } else
                    return false;
            } catch {
                log.Warn("Upgrade(s) didn't complete successfully. The user will likely need to reset their settings.");
                return false;
            }
        }

        private static void backupSettingsFile() {
            if (string.IsNullOrEmpty(settingsVersion)) {
                log.Debug("The settings file is a vanilla template. No need to back it up.");
                return;
            }
            
            String backupFile = "";
            try {
                log.Info("Backing up '" + Settings.ConfigFile + "' for v" + settingsVersion);
                backupFile = System.Text.RegularExpressions.Regex.Replace(Settings.ConfigFile, @"(\.\w+)$", "-v" + settingsVersion + "$1");
                File.Copy(Settings.ConfigFile, backupFile, false);
                log.Info(backupFile + " created.");
            } catch (System.IO.IOException ex) {
                if (OGCSexception.GetErrorCode(ex) == "0x80070050") { //File already exists
                    log.Warn("The backfile already exists, not overwriting " + backupFile);
                }
            } catch (System.Exception ex) {
                OGCSexception.Analyse("Failed to create backup settings file", ex);
            }
        }

        private static void upgradeToMultiCalendar() {
            backupSettingsFile();

            XDocument xml = null;
            try {
                xml = XDocument.Load(Settings.ConfigFile);
                XElement settingsElement = XMLManager.GetElement("Settings", xml);
                XElement calendarsElement = XMLManager.AddElement("Calendars", settingsElement);
                XElement calendarElement = XMLManager.AddElement("Calendar", calendarsElement);

                //If a source element doesn't exist, the XML is not changed
                //Manually add Profile Name - it's critical to be able to select the right profile later on and a Settings.Save() might not have happened.
                XMLManager.AddElement("_ProfileName", calendarElement).Value = "Default";

                XMLManager.MoveElement("OutlookService", settingsElement, calendarElement);
                XMLManager.MoveElement("MailboxName", settingsElement, calendarElement);
                XMLManager.MoveElement("SharedCalendar", settingsElement, calendarElement);
                XMLManager.MoveElement("UseOutlookCalendar", settingsElement, calendarElement);
                XMLManager.MoveElement("CategoriesRestrictBy", settingsElement, calendarElement);
                XMLManager.MoveElement("Categories", settingsElement, calendarElement);
                XMLManager.MoveElement("OnlyRespondedInvites", settingsElement, calendarElement);
                XMLManager.MoveElement("OutlookDateFormat", settingsElement, calendarElement);
                XMLManager.MoveElement("OutlookGalBlocked", settingsElement, calendarElement);
                
                XMLManager.MoveElement("UseGoogleCalendar", settingsElement, calendarElement);
                XMLManager.MoveElement("CloakEmail", settingsElement, calendarElement);
                XMLManager.MoveElement("ExcludeDeclinedInvites", settingsElement, calendarElement);
                XMLManager.MoveElement("ExcludeGoals", settingsElement, calendarElement);

                XMLManager.MoveElement("SyncDirection", settingsElement, calendarElement);
                XMLManager.MoveElement("DaysInThePast", settingsElement, calendarElement);
                XMLManager.MoveElement("DaysInTheFuture", settingsElement, calendarElement);
                XMLManager.MoveElement("SyncInterval", settingsElement, calendarElement);
                XMLManager.MoveElement("SyncIntervalUnit", settingsElement, calendarElement);
                XMLManager.MoveElement("OutlookPush", settingsElement, calendarElement);
                XMLManager.MoveElement("AddLocation", settingsElement, calendarElement);
                XMLManager.MoveElement("AddDescription", settingsElement, calendarElement);
                XMLManager.MoveElement("AddDescription_OnlyToGoogle", settingsElement, calendarElement);
                XMLManager.MoveElement("AddReminders", settingsElement, calendarElement);
                XMLManager.MoveElement("UseGoogleDefaultReminder", settingsElement, calendarElement);
                XMLManager.MoveElement("UseOutlookDefaultReminder", settingsElement, calendarElement);
                XMLManager.MoveElement("ReminderDND", settingsElement, calendarElement);
                XMLManager.MoveElement("ReminderDNDstart", settingsElement, calendarElement);
                XMLManager.MoveElement("ReminderDNDend", settingsElement, calendarElement);
                XMLManager.MoveElement("AddAttendees", settingsElement, calendarElement);
                XMLManager.MoveElement("MaxAttendees", settingsElement, calendarElement);
                XMLManager.MoveElement("AddColours", settingsElement, calendarElement);
                XMLManager.MoveElement("MergeItems", settingsElement, calendarElement);
                XMLManager.MoveElement("DisableDelete", settingsElement, calendarElement);
                XMLManager.MoveElement("ConfirmOnDelete", settingsElement, calendarElement);
                XMLManager.MoveElement("TargetCalendar", settingsElement, calendarElement);
                XMLManager.MoveElement("CreatedItemsOnly", settingsElement, calendarElement);
                XMLManager.MoveElement("SetEntriesPrivate", settingsElement, calendarElement);
                XMLManager.MoveElement("SetEntriesAvailable", settingsElement, calendarElement);
                XMLManager.MoveElement("AvailabilityStatus", settingsElement, calendarElement);
                XMLManager.MoveElement("SetEntriesColour", settingsElement, calendarElement);
                XMLManager.MoveElement("SetEntriesColourValue", settingsElement, calendarElement);
                XMLManager.MoveElement("SetEntriesColourName", settingsElement, calendarElement);
                XMLManager.MoveElement("SetEntriesColourGoogleId", settingsElement, calendarElement);
                XMLManager.MoveElement("ColourMaps", settingsElement, calendarElement);
                XMLManager.MoveElement("SingleCategoryOnly", settingsElement, calendarElement);
                XMLManager.MoveElement("Obfuscation", settingsElement, calendarElement);
                
                XMLManager.MoveElement("ExtirpateOgcsMetadata", settingsElement, calendarElement);
                XMLManager.MoveElement("LastSyncDate", settingsElement, calendarElement);

            } catch (System.Exception ex) {
                OGCSexception.Analyse("Problem encountered whilst upgrading " + Settings.ConfigFilename, ex);
                throw ex;
            } finally {
                if (xml != null) {
                    xml.Root.Sort();
                    while (true) {
                        try {
                            xml.Save(Settings.ConfigFile);
                            break;
                        } catch (System.IO.IOException ex) {
                            log.Fail("Another process has locked file " + Settings.ConfigFile);
                            if (MessageBox.Show("Another program is using the settings file " + Settings.ConfigFile +
                                "\r\nPlease close any other instance of OGCS that may be using it.",
                                "Settings Cannot Be Saved", MessageBoxButtons.RetryCancel, MessageBoxIcon.Exclamation) == DialogResult.Cancel) 
                            {
                                log.Warn("User cancelled attempt to save new settings file.");
                                OGCSexception.Analyse("Could not save upgraded settings file " + Settings.ConfigFile, ex);
                                throw;
                            }
                        } catch (System.Exception ex) {
                            OGCSexception.Analyse("Could not save upgraded settings file " + Settings.ConfigFile, ex);
                            throw;
                        }
                    }
                }
            }
        }
    }
}
