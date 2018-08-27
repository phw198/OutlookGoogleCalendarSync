using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Outlook;
using log4net;

namespace OutlookGoogleCalendarSync.OutlookOgcs {
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
                log.Debug("Adding listener for Explorer '" + newExplorer.Caption + "'");
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
        /// to revert any changes and repopulated values.
        /// </summary>
        private void beforeItemPaste(ref object ClipboardContent, MAPIFolder Target, ref bool Cancel) {
            log.Info("Item paste event caught.");
            try {
                Selection selection = ClipboardContent as Selection;
                log.Debug("We've got " + selection.Count + " items selected for copy.");
                
                foreach (Object item in selection) {
                    AppointmentItem copiedAi = null;
                    try {
                        if (item is AppointmentItem) {
                            copiedAi = item as AppointmentItem;
                        } else {
                            throw new ApplicationException("The item is not an appointment item.");
                        }
                        log.Debug(OutlookOgcs.Calendar.GetEventSummary(copiedAi));
                        String entryID = copiedAi.EntryID;
                        if (OutlookOgcs.CustomProperty.Exists(copiedAi, OutlookOgcs.CustomProperty.MetadataId.gEventID)) {
                            Dictionary<OutlookOgcs.CustomProperty.MetadataId, object> propertyBackup = cleanIDs(ref copiedAi);
                            System.Threading.Thread repopIDsThrd = new System.Threading.Thread(() => repopulateIDs(entryID, propertyBackup));
                            repopIDsThrd.Start();

                        } else {
                            log.Debug("This item isn't managed by OGCS.");
                            //But we still need to tag it as a "copied" item to avoid bad matches on Google events.
                            OutlookOgcs.CustomProperty.Add(ref copiedAi, OutlookOgcs.CustomProperty.MetadataId.locallyCopied, true.ToString());
                            copiedAi.Save();
                            System.Threading.Thread untagAsCopiedThrd = new System.Threading.Thread(() => untagAsCopied(entryID));
                            untagAsCopiedThrd.Start();
                        }

                    } catch (System.ApplicationException ex) {
                        log.Debug(ex.Message);

                    } catch (System.Exception ex) {
                        log.Warn("Not able to process copy and pasted event.");
                        OGCSexception.Analyse(ex);

                    } finally {
                        copiedAi = (AppointmentItem)OutlookOgcs.Calendar.ReleaseObject(copiedAi);
                    }
                }
            } catch (System.Exception ex) {
                OGCSexception.Analyse(ex);
            }
        }

        private Dictionary<OutlookOgcs.CustomProperty.MetadataId, object> cleanIDs(ref AppointmentItem copiedAi) {
            log.Info("Temporarily removing OGCS properties from copied Outlook appointment item.");

            Dictionary<OutlookOgcs.CustomProperty.MetadataId, object> propertyBackup = new Dictionary<OutlookOgcs.CustomProperty.MetadataId, object>();
            try {
                object backupValue = null;
                foreach (OutlookOgcs.CustomProperty.MetadataId metaDataId in Enum.GetValues(typeof(OutlookOgcs.CustomProperty.MetadataId))) {
                    log.Fine("Backing up " + metaDataId.ToString());
                    if (metaDataId == OutlookOgcs.CustomProperty.MetadataId.ogcsModified)
                        backupValue = OutlookOgcs.CustomProperty.GetOGCSlastModified(copiedAi);
                    else
                        backupValue = OutlookOgcs.CustomProperty.Get(copiedAi, metaDataId);
                    propertyBackup.Add(metaDataId, backupValue);
                }
                OutlookOgcs.CustomProperty.RemoveAll(ref copiedAi);
                OutlookOgcs.CustomProperty.Add(ref copiedAi, OutlookOgcs.CustomProperty.MetadataId.locallyCopied, true.ToString());
                copiedAi.Save();

            } catch (System.Exception ex) {
                log.Warn("Failed to clean OGCS properties from copied item.");
                OGCSexception.Analyse(ex);
            }
            return propertyBackup;
        }

        private void repopulateIDs(String entryID, Dictionary<OutlookOgcs.CustomProperty.MetadataId, object> propertyValues) {
            //Allow time for pasted item to complete
            System.Threading.Thread.Sleep(2000);
            log.Debug("Repopulating IDs to original copied Outlook item");

            AppointmentItem copiedAi = null;
            try {
                OutlookOgcs.Calendar.Instance.IOutlook.GetAppointmentByID(entryID, out copiedAi);
                if (copiedAi == null) {
                    throw new System.Exception("Could not find Outlook item with entryID " + entryID + " for post-processing.");
                }

                log.Debug(OutlookOgcs.Calendar.GetEventSummary(copiedAi));
                foreach (KeyValuePair<OutlookOgcs.CustomProperty.MetadataId, object> property in propertyValues) {
                    if (property.Value == null)
                        OutlookOgcs.CustomProperty.Remove(ref copiedAi, property.Key);
                    else {
                        if (property.Value is DateTime)
                            OutlookOgcs.CustomProperty.Add(ref copiedAi, property.Key, (DateTime)property.Value);
                        else
                            OutlookOgcs.CustomProperty.Add(ref copiedAi, property.Key, property.Value.ToString());
                    }
                }                
                copiedAi.Save();

            } catch (System.Exception ex) {
                log.Warn("Failed to repopulate OGCS properties back to copied item.");
                OGCSexception.Analyse(ex);
            } finally {
                copiedAi = (AppointmentItem)OutlookOgcs.Calendar.ReleaseObject(copiedAi);
            }
        }

        private void untagAsCopied(String entryID) {
            //Allow time for pasted item to complete
            System.Threading.Thread.Sleep(2000);
            log.Debug("Untagging copied Outlook item");

            AppointmentItem copiedAi = null;
            try {
                OutlookOgcs.Calendar.Instance.IOutlook.GetAppointmentByID(entryID, out copiedAi);
                if (copiedAi == null) {
                    throw new System.Exception("Could not find Outlook item with entryID " + entryID + " for post-processing.");
                }
                log.Debug(OutlookOgcs.Calendar.GetEventSummary(copiedAi));
                OutlookOgcs.CustomProperty.Remove(ref copiedAi, OutlookOgcs.CustomProperty.MetadataId.locallyCopied);
                copiedAi.Save();
            } catch (System.Exception ex) {
                log.Warn("Failed to remove OGCS 'copied' property on copied item.");
                OGCSexception.Analyse(ex);
            } finally {
                copiedAi = (AppointmentItem)OutlookOgcs.Calendar.ReleaseObject(copiedAi);
            }
        }
    }
}
