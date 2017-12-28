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

        public ExplorerWatcher(Explorers explorers) {
            log.Info("Setting up Explorer event watchers.");
            //Watch existing explorers
            for (int e = 1; e <= explorers.Count; e++) {
                watchForPasteEvents(explorers[e]);
            }
            //Watch for new explorers
            this.explorers = explorers;
            this.explorers.NewExplorer += new ExplorersEvents_NewExplorerEventHandler(explorers_NewExplorer);
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
        /// However, the clipboard is a reference to the copied item (so we can't change that)
        /// and the pasted object is not available yet until AFTER this function!
        /// Workaround is for a delayed background thread to post-process the pasted item.
        /// </summary>
        private void beforeItemPaste(ref object ClipboardContent, MAPIFolder Target, ref bool Cancel) {
            log.Info("Item paste event caught.");
            DateTime pasteEventCaught = DateTime.Now;

            try {
                Selection selection = ClipboardContent as Selection;
                log.Debug("We've got " + selection.Count + " items selected for copy.");

                List<String> gEventIDs = new List<string>();
                List<String> oApptIDs = new List<string>();
                String gEventID = "";
                String oApptID = "";

                foreach (Object item in selection) {
                    AppointmentItem copiedAi = null;
                    try {
                        if (item is AppointmentItem) {
                            copiedAi = item as AppointmentItem;
                        } else {
                            throw new ApplicationException("The item is not an appointment item.");
                        }
                        log.Debug(OutlookOgcs.Calendar.GetEventSummary(copiedAi));
                        if (OutlookOgcs.Calendar.GetOGCSproperty(copiedAi, OutlookOgcs.Calendar.MetadataId.gEventID, out gEventID)) {
                            gEventIDs.Add(gEventID);
                        } else {
                            log.Debug("This item isn't managed by OGCS.");
                            //But we still need to tag it as a "copied" item to avoid bad matches on Google events.
                            //We'll identify the Outlook item by GlobalApptID (if available) and EntryID
                            oApptID = OutlookOgcs.Calendar.Instance.IOutlook.GetGlobalApptID(copiedAi);
                            if (!oApptID.StartsWith(OutlookOgcs.Calendar.GlobalIdPattern)) oApptID = "";
                            oApptID += "~" + copiedAi.EntryID;
                            oApptIDs.Add(oApptID);
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

                if (gEventIDs.Count > 0) {
                    System.Threading.Thread cleanIDsThrd = new System.Threading.Thread(x => cleanIDs(gEventIDs, pasteEventCaught));
                    cleanIDsThrd.Start();
                }
                if (oApptIDs.Count > 0) {
                    System.Threading.Thread tagAsCopiedThrd = new System.Threading.Thread(x => tagAsCopied(oApptIDs, pasteEventCaught));
                    tagAsCopiedThrd.Start();
                }

            } catch (System.Exception ex) {
                OGCSexception.Analyse(ex);
            }
        }

        private void cleanIDs(List<String> gEventIDs, DateTime asOf) {
            //Allow time for pasted item to complete
            System.Threading.Thread.Sleep(2000);
            log.Debug("Cleaning IDs from copied Outlook items");

            try {
                DateTime asOfRounded = asOf.AddTicks(-(asOf.Ticks % TimeSpan.TicksPerMillisecond)); //Get rid of fractional milliseconds

                log.Debug("Finding " + gEventIDs.Count + " recently copied and pasted item(s) since " + asOf.ToString("dd/MM/yyyy hh:mm:ss.fff") + "...");
                foreach (String gEventID in gEventIDs) {
                    List<AppointmentItem> filtered = new List<AppointmentItem>();
                    try {
                        filtered = OutlookOgcs.Calendar.Instance.FilterCalendarEntries(OutlookOgcs.Calendar.Instance.UseOutlookCalendar.Items,
                            filterCategories: false, noDateFilter: true,
                            extraFilter: " AND [googleEventID] = '" + gEventID + "' AND [Modified] >= '" + asOfRounded.ToString(Settings.Instance.OutlookDateFormat) + "'");
                    } catch (System.Exception ex) {
                        log.Debug("Filter for Outlook items failed. Could be because googleEventID is not searchable.");
                        OGCSexception.Analyse(ex);
                    }   

                    if (filtered.Count > 1) {
                        log.Warn("We've got back " + filtered.Count + " items for " + gEventID + "! Only expected one - attempting to filter further...");
                        for (int i = filtered.Count - 1; i >= 0; i--) {
                            AppointmentItem ai = filtered[i];
                            log.Debug(OutlookOgcs.Calendar.GetEventSummary(ai));
                            if (ai.LastModificationTime < asOfRounded) {
                                log.Debug("Removed");
                                filtered.Remove(ai);
                                ai = (AppointmentItem)OutlookOgcs.Calendar.ReleaseObject(ai);
                            } else
                                log.Debug("Not removed.");
                        }
                    }

                    if (filtered.Count > 1) {
                        log.Error("We've still got " + filtered.Count + " items for " + gEventID + "! Impossible to determine which one was pasted.");
                        for (int i = 0; i < filtered.Count; i++) {
                            AppointmentItem ai = filtered[i];
                            ai = (AppointmentItem)OutlookOgcs.Calendar.ReleaseObject(ai);
                        }
                    } else if (filtered.Count == 1) {
                        AppointmentItem ai = null;
                        try {
                            ai = filtered[0];
                            log.Info("Removing OGCS properties from copied Outlook appointment item.");
                            log.Debug(OutlookOgcs.Calendar.GetEventSummary(ai));
                            OutlookOgcs.Calendar.RemoveOGCSproperties(ref ai);
                            OutlookOgcs.Calendar.AddOGCSproperty(ref ai, OutlookOgcs.Calendar.MetadataId.locallyCopied, true.ToString());
                            ai.Save();
                        } finally {
                            ai = (AppointmentItem)OutlookOgcs.Calendar.ReleaseObject(ai);
                        }
                    } else if (filtered.Count == 0) {
                        log.Warn("Could not find Outlook item with googleEventID " + gEventID + " for post-processing.");
                    }
                }
            } catch (System.Exception ex) {
                log.Warn("Failed to clean OGCS properties from copy and pasted items.");
                OGCSexception.Analyse(ex);
            }
        }

        private void tagAsCopied(List<String> oApptIDs, DateTime asOf) {
            //Allow time for pasted item to complete
            System.Threading.Thread.Sleep(2000);
            log.Debug("Tagging copied Outlook items");

            try {
                DateTime asOfRounded = asOf.AddTicks(-(asOf.Ticks % TimeSpan.TicksPerMillisecond)); //Get rid of fractional milliseconds

                log.Debug("Finding " + oApptIDs.Count + " recently copied and pasted item(s) since " + asOf.ToString("dd/MM/yyyy hh:mm:ss.fff") + "...");
                foreach (String oApptID in oApptIDs) {
                    try {
                        List<AppointmentItem> filtered = new List<AppointmentItem>();
                        filtered = OutlookOgcs.Calendar.Instance.FilterCalendarEntries(OutlookOgcs.Calendar.Instance.UseOutlookCalendar.Items,
                            filterCategories: false, noDateFilter: true,
                            extraFilter: " AND [Modified] >= '" + asOfRounded.AddSeconds(-1).ToString(Settings.Instance.OutlookDateFormat) + "'");

                        if (filtered.Count > 1) {
                            log.Warn("We've got back " + filtered.Count + " items. Filtering further...");
                            for (int i = filtered.Count - 1; i >= 0; i--) {
                                AppointmentItem ai = filtered[i];
                                log.Debug(OutlookOgcs.Calendar.GetEventSummary(ai));
                                Boolean isMatched = false;
                                if (ai.LastModificationTime < asOfRounded) {
                                    isMatched = false;
                                    log.Debug("Last modified: " + ai.LastModificationTime.ToString("dd/MM/yyyy hh:mm:ss"));
                                    log.Debug("Modified before paste event caught.");
                                } else {
                                    //Run the same checks as GoogleOgcs.ItemIDsMatch()
                                    String[] ids = oApptID.Split(new char[] { '~' }, StringSplitOptions.None);
                                    if (ids[0] == String.Empty) {
                                        isMatched = true;
                                        log.Debug("No GlobalID available, so continue matching on Entry ID");
                                    } else {
                                        log.Debug("Comparing GlobalIDs");
                                        String aiGlobalID = OutlookOgcs.Calendar.Instance.IOutlook.GetGlobalApptID(ai);
                                        isMatched = (aiGlobalID.StartsWith(OutlookOgcs.Calendar.GlobalIdPattern) &&
                                            ids[0].StartsWith(OutlookOgcs.Calendar.GlobalIdPattern) &&
                                            ids[0].Substring(72) == aiGlobalID.Substring(72));
                                    }
                                    if (isMatched) {
                                        log.Debug("Comparing EntryIDs");
                                        isMatched = (ids[1].Remove(ids[1].Length - 16) == ai.EntryID.Remove(ai.EntryID.Length - 16));
                                    }
                                }
                                if (!isMatched) {
                                    log.Debug("Removed");
                                    filtered.Remove(ai);
                                    ai = (AppointmentItem)OutlookOgcs.Calendar.ReleaseObject(ai);
                                } else
                                    log.Debug("Not removed.");
                            }
                        }

                        if (filtered.Count >= 1) {
                            log.Warn("We've still got " + filtered.Count + " items to mark as copied.");
                            for (int i = 0; i < filtered.Count; i++) {
                                AppointmentItem ai = null;
                                try {
                                    ai = filtered[0];
                                    log.Debug(OutlookOgcs.Calendar.GetEventSummary(ai));
                                    OutlookOgcs.Calendar.AddOGCSproperty(ref ai, OutlookOgcs.Calendar.MetadataId.locallyCopied, true.ToString());
                                    ai.Save();
                                } finally {
                                    ai = (AppointmentItem)OutlookOgcs.Calendar.ReleaseObject(ai);
                                }
                            }
                        } else if (filtered.Count == 0) {
                            log.Warn("Could not find any Outlook items for post-processing!");
                        }

                    } catch (System.Exception ex) {
                        log.Warn("Failed to set OGCS 'copied' property on copy and pasted item with GlobalID " + oApptID);
                        OGCSexception.Analyse(ex);
                    }
                }

            } catch (System.Exception ex) {
                log.Warn("Failed to set OGCS 'copied' property on copy and pasted item(s).");
                OGCSexception.Analyse(ex);
            }
        }
    }
}
