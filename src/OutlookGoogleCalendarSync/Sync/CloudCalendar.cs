﻿using Ogcs = OutlookGoogleCalendarSync;
using GcalData = Google.Apis.Calendar.v3.Data;
using log4net;
//using Microsoft.Office.Interop.Outlook;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace OutlookGoogleCalendarSync.Sync {
    public partial class Engine {
        protected internal class CloudCalendar {
            private static readonly ILog log = LogManager.GetLogger(typeof(CloudCalendar));

            /// <summary>
            /// The calendar settings profile currently being synced.
            /// </summary>
            public SettingsStore.Calendar Profile { internal set; get; }

            private int consecutiveSyncFails = 0;

            private static CloudCalendar instance;
            public static CloudCalendar Instance {
                get {
                    instance ??= new CloudCalendar();
                    return instance;
                }
                set {
                    instance = value;
                }
            }
            public CloudCalendar() {
                Profile = Sync.Engine.Instance.ActiveProfile as SettingsStore.Calendar;
            }

            public SyncResult ManualSynchronize() {
                //This function is just a shim to determine how the sync was triggered when looking at the call stack
                return Synchronize();
            }

            public SyncResult Synchronize() {
                this.Profile = Sync.Engine.Instance.ActiveProfile as SettingsStore.Calendar;

                Console console = Forms.Main.Instance.Console;
                console.Update("Finding Calendar Entries", Console.Markup.mag_right, newLine: false);

                List<Microsoft.Graph.Event> outlookEntries = null;
                List<GcalData.Event> googleEntries = null;
                if (!Ogcs.Google.Calendar.IsInstanceNull)
                    Ogcs.Google.Calendar.Instance.EphemeralProperties.Clear();
                Outlook.Calendar.Instance.EphemeralProperties.Clear();

                #region Read Outlook items
                console.Update("Scanning Outlook calendar...");
                outlookEntries = Outlook.Graph.Calendar.Instance.GetCalendarEntriesInRange(Sync.Engine.Calendar.Instance.Profile, false);
                console.Update(outlookEntries.Count + " Outlook calendar entries found.", Console.Markup.sectionEnd, newLine: false);

                if (Sync.Engine.Instance.CancellationPending) return SyncResult.UserCancelled;
                #endregion

                #region Read Google items
                SyncResult gotItems = Sync.Engine.Calendar.Instance.ReadGoogleItems(ref googleEntries);
                if (gotItems != SyncResult.OK) return gotItems;
                if (Sync.Engine.Instance.CancellationPending) return SyncResult.UserCancelled;
                #endregion
                
                Boolean success = true;
                String bubbleText = "";
                /*
                if (this.Profile.ExtirpateOgcsMetadata) {
                    return extirpateCustomProperties(outlookEntries, googleEntries);
                }

                //Reclaim orphans
                Ogcs.Google.Calendar.Instance.ReclaimOrphanCalendarEntries(ref googleEntries, ref outlookEntries);
                if (Sync.Engine.Instance.CancellationPending) return SyncResult.UserCancelled;

                Outlook.Calendar.Instance.ReclaimOrphanCalendarEntries(ref outlookEntries, ref googleEntries);
                if (Sync.Engine.Instance.CancellationPending) return SyncResult.UserCancelled;

                if (this.Profile.AddColours || this.Profile.SetEntriesColour) {
                    Outlook.Calendar.Categories.ValidateCategories();

                    if (this.Profile.ColourMaps.Count > 0) {
                        this.Profile.ColourMaps.ToList().ForEach(c => {
                            if (Outlook.Calendar.Categories.OutlookColour(c.Key) == null) {
                                if (Ogcs.Extensions.MessageBox.Show("There is a problem with your colour mapping configuration.\r\nColours may not get synced as intended.\r\nReview maps now for missing Outlook colours?",
                                    "Invalid colour map", System.Windows.Forms.MessageBoxButtons.YesNo, System.Windows.Forms.MessageBoxIcon.Error) == DialogResult.Yes)
                                    new Forms.ColourMap().ShowDialog();
                            }
                        });
                    }
                }
                */
                //Sync
                /*
                if (this.Profile.SyncDirection.Id != Direction.GoogleToOutlook.Id) {
                    success = outlookToGoogle(outlookEntries, googleEntries, ref bubbleText);
                    if (Sync.Engine.Instance.CancellationPending) return SyncResult.UserCancelled;
                }
                if (!success) return SyncResult.Fail;
                */
                if (this.Profile.SyncDirection.Id != Sync.Direction.OutlookToGoogle.Id) {
                    if (bubbleText != "") bubbleText += "\r\n";
                    success = googleToOutlook(googleEntries, outlookEntries, ref bubbleText);
                    if (Sync.Engine.Instance.CancellationPending) return SyncResult.UserCancelled;
                }
                if (bubbleText != "") {
                    log.Info(bubbleText.Replace("\r\n", ". "));
                    System.Text.RegularExpressions.Regex rgx = new System.Text.RegularExpressions.Regex(@"\D");
                    String changes = rgx.Replace(bubbleText, "").Trim('0');
                    if (Settings.Instance.ShowSystemNotifications &&
                        (!Settings.Instance.ShowSystemNotificationsIfChange || !String.IsNullOrEmpty(changes))) Forms.Main.Instance.NotificationTray.ShowBubbleInfo(bubbleText);
                }
                /*
                */
                return SyncResult.OK;
            }

            private Boolean googleToOutlook(List<GcalData.Event> googleEntries, List<Microsoft.Graph.Event> outlookEntries, ref String bubbleText) {
                log.Debug("Synchronising from Google to Outlook.");
                if (this.Profile.SyncDirection.Id == Sync.Direction.Bidirectional.Id)
                    Forms.Main.Instance.Console.Update("Syncing " + Sync.Direction.GoogleToOutlook.Name, Console.Markup.syncDirection, newLine: false);

                List<GcalData.Event> outlookEntriesToBeCreated = new List<GcalData.Event>(googleEntries);
                List<Microsoft.Graph.Event> outlookEntriesToBeDeleted = new List<Microsoft.Graph.Event>(outlookEntries);
                Dictionary<Microsoft.Graph.Event, GcalData.Event> entriesToBeCompared = new Dictionary<Microsoft.Graph.Event, GcalData.Event>();

                Console console = Forms.Main.Instance.Console;

                try {
                    Ogcs.Outlook.Graph.Calendar.IdentifyEventDifferences(ref outlookEntriesToBeCreated, ref outlookEntriesToBeDeleted, ref entriesToBeCompared);
                    if (Sync.Engine.Instance.CancellationPending) return false;
                } catch (System.Exception) {
                    console.Update("Unable to identify differences in Outlook calendar.", Console.Markup.error);
                    throw;
                }
                
                StringBuilder sb = new StringBuilder();
                console.BuildOutput(outlookEntriesToBeDeleted.Count + " Outlook calendar entries to be deleted.", ref sb, false);
                console.BuildOutput(outlookEntriesToBeCreated.Count + " Outlook calendar entries to be created.", ref sb, false);
                console.BuildOutput(entriesToBeCompared.Count + " calendar entries to be compared.", ref sb, false);
                console.Update(sb, Console.Markup.info, logit: true);

                //Protect against very first syncs which may trample pre-existing non-Google events in Outlook
                if (!this.Profile.DisableDelete && !this.Profile.ConfirmOnDelete &&
                    outlookEntriesToBeDeleted.Count == outlookEntries.Count && outlookEntries.Count > 1) {
                    if (Ogcs.Extensions.MessageBox.Show("All Outlook events are going to be deleted. Do you want to allow this?" +
                        "\r\nNote, " + outlookEntriesToBeCreated.Count + " events will then be created.", "Confirm mass deletion",
                        MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation) == DialogResult.No) {

                        while (outlookEntriesToBeDeleted.Count() > 0) {
                            Outlook.Calendar.ReleaseObject(outlookEntriesToBeDeleted.Last());
                            outlookEntriesToBeDeleted.Remove(outlookEntriesToBeDeleted.Last());
                        }
                    }
                }

                int entriesUpdated = 0;
                try {
                    #region Delete Outlook Entries
                    if (outlookEntriesToBeDeleted.Count > 0) {
                        console.Update("Deleting " + outlookEntriesToBeDeleted.Count + " Outlook calendar entries", Console.Markup.h2, newLine: false);
                        try {
                            Outlook.Graph.Calendar.Instance.DeleteCalendarEntries(outlookEntriesToBeDeleted);
                        } catch (UserCancelledSyncException ex) {
                            log.Info(ex.Message);
                            return false;
                        } catch (System.Exception) {
                            console.Update("Unable to delete obsolete entries in Google calendar.", Console.Markup.error);
                            throw;
                        }
                        log.Info("Done.");
                    }

                    if (Sync.Engine.Instance.CancellationPending) return false;
                    #endregion
                /*

                    #region Create Outlook Entries
                    if (outlookEntriesToBeCreated.Count > 0) {
                        console.Update("Creating " + outlookEntriesToBeCreated.Count + " Outlook calendar entries", Console.Markup.h2, newLine: false);
                        try {
                            Outlook.Calendar.Instance.CreateCalendarEntries(outlookEntriesToBeCreated);
                        } catch (UserCancelledSyncException ex) {
                            log.Info(ex.Message);
                            return false;
                        } catch (System.Exception) {
                            console.Update("Unable to add new entries into the Outlook Calendar.", Console.Markup.error);
                            throw;
                        }
                        log.Info("Done.");
                    }

                    if (Sync.Engine.Instance.CancellationPending) return false;
                    #endregion

                    #region Update Outlook Entries
                    if (entriesToBeCompared.Count > 0) {
                        console.Update("Comparing " + entriesToBeCompared.Count + " existing Outlook calendar entries", Console.Markup.h2, newLine: false);
                        try {
                            Outlook.Calendar.Instance.UpdateCalendarEntries(entriesToBeCompared, ref entriesUpdated);
                        } catch (UserCancelledSyncException ex) {
                            log.Info(ex.Message);
                            return false;
                        } catch (System.Exception) {
                            console.Update("Unable to update existing entries in the Outlook calendar.", Console.Markup.error);
                            throw;
                        }
                        console.Update(entriesUpdated + " entries updated.");
                    }

                    if (Sync.Engine.Instance.CancellationPending) return false;
                    #endregion
                */

                } finally {
                    bubbleText += "Outlook: " + outlookEntriesToBeCreated.Count + " created; " +
                        outlookEntriesToBeDeleted.Count + " deleted; " + entriesUpdated + " updated";

                    /*while (outlookEntriesToBeCreated.Count() > 0) {
                        Outlook.Calendar.ReleaseObject(outlookEntriesToBeCreated.Last());
                        outlookEntriesToBeCreated.Remove(outlookEntriesToBeCreated.Last());
                    }
                    while (outlookEntriesToBeDeleted.Count() > 0) {
                        Outlook.Calendar.ReleaseObject(outlookEntriesToBeDeleted.Last());
                        outlookEntriesToBeDeleted.Remove(outlookEntriesToBeDeleted.Last());
                    }
                    while (entriesToBeCompared.Count() > 0) {
                        Outlook.Calendar.ReleaseObject(entriesToBeCompared.Keys.Last());
                        entriesToBeCompared.Remove(entriesToBeCompared.Keys.Last());
                    }*/
                }
                return true;
            }


        }
    }
}