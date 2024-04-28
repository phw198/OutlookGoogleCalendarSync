using Ogcs = OutlookGoogleCalendarSync;
using log4net;
using Microsoft.Office.Interop.Outlook;
using System;
using System.Collections.Generic;

namespace OutlookGoogleCalendarSync.Outlook {
    class ExplorerWatcher {
        private static readonly ILog log = LogManager.GetLogger(typeof(ExplorerWatcher));

        private Explorers explorers;
        private List<Explorer> watchExplorers = new List<Explorer>();

        public ExplorerWatcher(Application oApp) {
            Explorers explorers = null;
            try {
                explorers = oApp.Explorers;

                log.Info("Setting up Explorer event watchers.");
                log.Debug("Watcher needed for " + explorers.Count + " existing explorers.");
                for (int e = 1; e <= explorers.Count; e++) {
                    watchForPasteEvents(explorers[e]);
                }
                log.Fine("Watch for new explorers.");
                this.explorers = explorers;
                this.explorers.NewExplorer += new ExplorersEvents_NewExplorerEventHandler(explorers_NewExplorer);
            } finally {
                explorers = (Explorers)Calendar.ReleaseObject(explorers);
            }
        }

        private void watchForPasteEvents(Explorer newExplorer) {
            if (!watchExplorers.Contains(newExplorer)) {
                log.Debug("Adding listener for Explorer '" + EmailAddress.MaskAddressWithinText(newExplorer.Caption) + "'");
                newExplorer.BeforeItemPaste += new ExplorerEvents_10_BeforeItemPasteEventHandler(beforeItemPaste);
                watchExplorers.Add(newExplorer);
            }
        }

        private void explorers_NewExplorer(Explorer Explorer) {
            log.Info("Detected new Explorer window.");
            watchForPasteEvents(Explorer);
        }

        /// <summary>
        /// Detects paste events in order to remove OGCS properties from pasted item.
        /// 
        /// However, the clipboard is a reference to the copied item 
        /// and the pasted object is not available yet until AFTER this function!
        /// We can't short-circuit the paste event by setting "Cancel = true" and performing the Copy()
        /// because it pastes to the same DateTime as the copied item.
        /// In Outlook2010 the (Explorer.View as CalendarView).SelectedStartTime exists, but not in 2007,
        /// so there's no way of knowing the time to paste the item in to.
        /// 
        /// So the workaround is to temporarily doctor the original copied item (ie remove OGCS properties),
        /// which the pasted item inherits. A delayed background thread then post-processes the original item
        /// to revert any changes and repopulate values.
        /// </summary>
        private void beforeItemPaste(ref object ClipboardContent, MAPIFolder Target, ref bool Cancel) {
            log.Info("Item paste event caught.");
            try {
                Selection selection = ClipboardContent as Selection;
                if (selection == null) {
                    log.Warn("Clipboard selection returned nothing.");
                    return;
                }
                log.Debug("We've got " + selection.Count + " items selected for copy.");

                foreach (Object item in selection) {
                    AppointmentItem copiedAi = null;
                    try {
                        if (item is AppointmentItem) {
                            copiedAi = item as AppointmentItem;
                        } else {
                            throw new ApplicationException("The item is not an appointment item.");
                        }
                        log.Debug(Calendar.GetEventSummary(copiedAi));
                        String entryID = copiedAi.EntryID;
                        if (CustomProperty.AnyStartsWith(copiedAi, CustomProperty.MetadataId.gEventID)) {
                            Dictionary<String, object> propertyBackup = cleanIDs(ref copiedAi);
                            CustomProperty.Add(ref copiedAi, CustomProperty.MetadataId.originalStartDate, copiedAi.Start);
                            copiedAi.Save();
                            System.Threading.Thread repopIDsThrd = new System.Threading.Thread(() => repopulateIDs(entryID, propertyBackup));
                            repopIDsThrd.Start();

                        } else {
                            log.Debug("This item isn't managed by OGCS.");
                            //But we still need to tag the pasted item as a "copied" item to avoid bad matches on Google events.
                            CustomProperty.Add(ref copiedAi, CustomProperty.MetadataId.locallyCopied, true.ToString());
                            copiedAi.Save();
                            //Untag the original copied item
                            System.Threading.Thread untagAsCopiedThrd = new System.Threading.Thread(() => untagAsCopied(entryID));
                            untagAsCopiedThrd.Start();
                        }

                    } catch (System.ApplicationException ex) {
                        log.Debug(ex.Message);

                    } catch (System.Exception ex) {
                        log.Warn("Not able to process copy and pasted event.");
                        Ogcs.Exception.Analyse(ex);

                    } finally {
                        copiedAi = (AppointmentItem)Calendar.ReleaseObject(copiedAi);
                    }
                }
            } catch (System.Exception ex) {
                Ogcs.Exception.Analyse(ex);
            }
        }

        private Dictionary<String, object> cleanIDs(ref AppointmentItem copiedAi) {
            log.Info("Temporarily removing OGCS properties from copied Outlook appointment item.");

            Dictionary<String, object> propertyBackup = new Dictionary<String, object>();
            UserProperties ups = null;
            try {
                object backupValue = null;
                ups = copiedAi.UserProperties;
                for (int p = 1; p <= ups.Count; p++) {
                    UserProperty up = null;
                    try {
                        up = ups[p];
                        String metaDataId = up.Name;
                        log.Fine("Backing up " + metaDataId.ToString());
                        backupValue = up.Value;
                        if (backupValue == null || (backupValue is DateTime time && time == new DateTime())) {
                            continue;
                        }
                        log.Fine("Property value: " + backupValue);
                        propertyBackup.Add(metaDataId, backupValue);
                    } finally {
                        up = (UserProperty)Calendar.ReleaseObject(up);
                    }
                }
                CustomProperty.Extirpate(ref copiedAi);
                CustomProperty.Add(ref copiedAi, CustomProperty.MetadataId.locallyCopied, true.ToString());
                copiedAi.Save();

            } catch (System.Exception ex) {
                log.Warn("Failed to clean OGCS properties from copied item.");
                Ogcs.Exception.Analyse(ex);
            } finally {
                ups = (UserProperties)Calendar.ReleaseObject(ups);
            }
            return propertyBackup;
        }

        private void repopulateIDs(String entryID, Dictionary<String, object> propertyValues) {
            //Allow time for pasted item to complete
            System.Threading.Thread.Sleep(2000);
            log.Debug("Repopulating IDs to original copied Outlook item");

            AppointmentItem copiedAi = null;
            try {
                untagAsCopied(entryID);
                Calendar.Instance.IOutlook.GetAppointmentByID(entryID, out copiedAi);
                if (copiedAi == null) {
                    throw new System.Exception("Could not find Outlook item with entryID " + entryID + " for post-processing.");
                }

                log.Debug(Calendar.GetEventSummary(copiedAi));
                foreach (KeyValuePair<String, object> property in propertyValues) {
                    if (property.Value is DateTime)
                        addOutlookCustomProperty(ref copiedAi, property.Key, OlUserPropertyType.olDateTime, property.Value);
                    else
                        addOutlookCustomProperty(ref copiedAi, property.Key, OlUserPropertyType.olText, property.Value);
                }
                log.Fine("Restored properties:-");
                CustomProperty.LogProperties(copiedAi, log4net.Core.Level.Debug);
                copiedAi.Save();

            } catch (System.Exception ex) {
                if (ex is System.Runtime.InteropServices.COMException && (
                    ex.GetErrorCode() == "0x8004010F" || //The message you specified cannot be found
                    ex.GetErrorCode() == "0x8004010A"))  //The operation cannot be performed because the object has been deleted
                {
                    log.Warn("Could not find Outlook item with entryID " + entryID + " for post-processing.");
                    Ogcs.Exception.LogAsFail(ref ex);
                }
                ex.Analyse("Failed to repopulate OGCS properties back to copied item.");
            } finally {
                copiedAi = (AppointmentItem)Calendar.ReleaseObject(copiedAi);
            }
        }

        private void untagAsCopied(String entryID) {
            //Allow time for pasted item to complete
            System.Threading.Thread.Sleep(2000);
            log.Debug("Untagging copied Outlook item");

            AppointmentItem copiedAi = null;
            try {
                Calendar.Instance.IOutlook.GetAppointmentByID(entryID, out copiedAi);
                if (copiedAi == null) {
                    throw new System.Exception("Could not find Outlook item with entryID " + entryID + " for post-processing.");
                }
                log.Debug(Calendar.GetEventSummary(copiedAi));
                String deletedPropVal = deleteOutlookCustomProperty(ref copiedAi, CustomProperty.MetadataId.locallyCopied.ToString());
                deletedPropVal = deleteOutlookCustomProperty(ref copiedAi, CustomProperty.MetadataId.originalStartDate.ToString());
                copiedAi.Save();

                if (!String.IsNullOrEmpty(deletedPropVal)) {
                    DateTime origStartDate = DateTime.Parse(deletedPropVal);

                    if (origStartDate != copiedAi.Start) { /* Item moved, not copied */
                        foreach (SettingsStore.Calendar profile in Settings.Instance.Calendars) {

                            if (origStartDate < profile.SyncStart && copiedAi.Start >= profile.SyncStart) {
                                Int16 newDaysInPast = (Int16)(profile.SyncStart.Date - origStartDate.Date).TotalDays;
                                Ogcs.Extensions.MessageBox.Show("Sync profile affected: " + profile._ProfileName + "\r\n" +
                                    "An already synced appointment has been moved back into the synced date range.\r\n" +
                                    "In order to avoid it being deleted, configuration has automatically been updated to " + (profile.DaysInThePast + newDaysInPast) + " days in the past.\r\n" +
                                    "After the next sync you may revert it to " + profile.DaysInThePast + ".", "Appointment moved into synced date range",
                                    System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Exclamation);
                                Forms.Main.Instance.SetControlPropertyThreadSafe(Forms.Main.Instance.tbDaysInThePast, "Text", (profile.DaysInThePast + newDaysInPast).ToString());

                            } else if (origStartDate >= profile.SyncStart && copiedAi.Start < profile.SyncStart) {
                                Int16 newDaysInPast = (Int16)(profile.SyncStart.Date - copiedAi.Start.Date).TotalDays;
                                Ogcs.Extensions.MessageBox.Show("Sync profile affected: " + profile._ProfileName + "\r\n" +
                                    "An already synced appointment has been moved out of the synced date range.\r\n" +
                                    "In order this is synced, configuration has automatically been updated to " + (profile.DaysInThePast + newDaysInPast) + " days in the past.\r\n" +
                                    "After the next sync you may revert it to " + profile.DaysInThePast + ".", "Appointment moved out of synced date range",
                                    System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Exclamation);
                                Forms.Main.Instance.SetControlPropertyThreadSafe(Forms.Main.Instance.tbDaysInThePast, "Text", (profile.DaysInThePast + newDaysInPast).ToString());

                            } else if (origStartDate > profile.SyncEnd && copiedAi.Start <= profile.SyncEnd) {
                                Int16 newDaysInFuture = (Int16)(origStartDate - profile.SyncEnd.Date).TotalDays;
                                Ogcs.Extensions.MessageBox.Show("Sync profile affected: " + profile._ProfileName + "\r\n" +
                                    "An already synced appointment has been moved into the synced date range.\r\n" +
                                    "In order this is synced, configuration has automatically been updated to " + (profile.DaysInTheFuture + newDaysInFuture) + " days in the future.\r\n" +
                                    "After the next sync you may revert it to " + profile.DaysInTheFuture + ".", "Appointment moved into synced date range",
                                    System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Exclamation);
                                Forms.Main.Instance.SetControlPropertyThreadSafe(Forms.Main.Instance.tbDaysInTheFuture, "Text", (profile.DaysInTheFuture + newDaysInFuture).ToString());

                            } else if (origStartDate <= profile.SyncEnd && copiedAi.Start > profile.SyncEnd) {
                                Int16 newDaysInFuture = (Int16)(copiedAi.Start.Date - profile.SyncEnd.Date).TotalDays;
                                Ogcs.Extensions.MessageBox.Show("Sync profile affected: " + profile._ProfileName + "\r\n" +
                                    "An already synced appointment has been moved out of the synced date range.\r\n" +
                                    "In order this is synced, configuration has automatically been updated to " + (profile.DaysInTheFuture + newDaysInFuture) + " days in the future.\r\n" +
                                    "After the next sync you may revert it to " + profile.DaysInTheFuture + ".", "Appointment moved out of synced date range",
                                    System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Exclamation);
                                Forms.Main.Instance.SetControlPropertyThreadSafe(Forms.Main.Instance.tbDaysInTheFuture, "Text", (profile.DaysInTheFuture + newDaysInFuture).ToString());
                            }
                        }
                    }
                }

            } catch (System.Exception ex) {
                if (ex is System.Runtime.InteropServices.COMException && (
                    ex.GetErrorCode() == "0x8004010F" || //The message you specified cannot be found
                    ex.GetErrorCode() == "0x8004010A"))  //The operation cannot be performed because the object has been deleted
                {
                    log.Warn("Could not find Outlook item with entryID " + entryID + " for post-processing.");
                    Ogcs.Exception.LogAsFail(ref ex);
                }
                ex.Analyse("Failed to remove OGCS 'copied' property on copied item.");
            } finally {
                copiedAi = (AppointmentItem)Calendar.ReleaseObject(copiedAi);
            }
        }

        private void addOutlookCustomProperty(ref AppointmentItem copiedAi, String addKeyName, OlUserPropertyType keyType, object keyValue) {
            UserProperties ups = null;
            try {
                ups = copiedAi.UserProperties;
                if (ups[addKeyName] == null) {
                    try {
                        ups.Add(addKeyName, keyType);
                    } catch (System.Exception ex) {
                        Ogcs.Exception.Analyse(ex);
                        ups.Add(addKeyName, keyType, false);
                    }
                }
                ups[addKeyName].Value = keyValue;
            } catch (System.Exception) {
                log.Warn("Failed to add " + addKeyName);
                throw;
            } finally {
                ups = (UserProperties)Calendar.ReleaseObject(ups);
            }
        }

        private String deleteOutlookCustomProperty(ref AppointmentItem copiedAi, String propertyName) {
            UserProperties ups = null;
            UserProperty prop = null;
            String propertyValue = null;
            try {
                ups = copiedAi.UserProperties;
                prop = ups.Find(propertyName);
                if (prop != null) {
                    propertyValue = prop.Value.ToString();
                    prop.Delete();
                    log.Debug("Removed " + propertyName + " property.");
                }
            } finally {
                prop = (UserProperty)Calendar.ReleaseObject(prop);
                ups = (UserProperties)Calendar.ReleaseObject(ups);
            }
            return propertyValue;
        }
    }
}
