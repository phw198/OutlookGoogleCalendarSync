using Ogcs = OutlookGoogleCalendarSync;
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
                Ogcs.Google.Calendar.Instance.ExcludedByColour = new Dictionary<String, String>();
                Ogcs.Google.Calendar.Instance.ExcludedByConfig = new List<String>();
                Ogcs.Outlook.Graph.Calendar.Instance.CancelledOccurrences = new();
                if (!Ogcs.Google.Calendar.IsInstanceNull)
                    Ogcs.Google.Calendar.Instance.EphemeralProperties.Clear();
                if (!Ogcs.Outlook.Calendar.IsInstanceNull) 
                    Outlook.Calendar.Instance.EphemeralProperties.Clear();
                if (!Ogcs.Outlook.Graph.Calendar.IsInstanceNull)
                    Outlook.Graph.Calendar.Instance.EphemeralProperties.Clear();

                #region Read Outlook items
                console.Update($"Scanning Outlook calendar '{Sync.Engine.Calendar.Instance.Profile.UseOutlookCalendar.Name}'...");
                outlookEntries = Outlook.Graph.Calendar.Instance.GetCalendarEntriesInRange(Sync.Engine.Calendar.Instance.Profile, false);
                String consoleOutput = outlookEntries.Count + " Outlook calendar entries found.";
                if (Outlook.Graph.Recurrence.OutlookExceptions != null && Outlook.Graph.Recurrence.OutlookExceptions.Count > 0)
                    consoleOutput += "<br/>"+ (Outlook.Graph.Recurrence.OutlookExceptions.Count + Outlook.Graph.Calendar.Instance.CancelledOccurrences.Count) + " additional exceptions to recurring events.";
                console.Update(consoleOutput, Console.Markup.sectionEnd, newLine: false);

                if (Sync.Engine.Instance.CancellationPending) return SyncResult.UserCancelled;
                #endregion

                #region Read Google items
                SyncResult gotItems = Sync.Engine.Calendar.Instance.ReadGoogleItems(ref googleEntries);
                if (gotItems != SyncResult.OK) return gotItems;
                if (Sync.Engine.Instance.CancellationPending) return SyncResult.UserCancelled;
                #endregion

                console.Update("Outlook " + outlookEntries.Count + ", Google " + googleEntries.Count);

                Ogcs.Google.Calendar.ExportToCSV("Outputting all Events.", "google_events.csv", googleEntries);
                Outlook.Graph.Calendar.ExportToCSV("Outputting all Appointments.", "outlook_appointments.csv", outlookEntries);

                Boolean success = true;
                String bubbleText = "";

                if (this.Profile.ExtirpateOgcsMetadata) {
                    return extirpateCustomProperties(outlookEntries, googleEntries);
                }

                //Reclaim orphans
                Ogcs.Google.Graph.Calendar.ReclaimOrphanCalendarEntries(ref googleEntries, ref outlookEntries);
                if (Sync.Engine.Instance.CancellationPending) return SyncResult.UserCancelled;

                Outlook.Graph.Calendar.ReclaimOrphanCalendarEntries(ref outlookEntries, ref googleEntries);
                if (Sync.Engine.Instance.CancellationPending) return SyncResult.UserCancelled;

                /*
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
                if (this.Profile.SyncDirection.Id != Direction.GoogleToOutlook.Id) {
                    success = outlookToGoogle(outlookEntries, googleEntries, ref bubbleText);
                    if (Sync.Engine.Instance.CancellationPending) return SyncResult.UserCancelled;
                }
                if (!success) return SyncResult.Fail;
                if (this.Profile.SyncDirection.Id != Sync.Direction.OutlookToGoogle.Id) {
                    if (bubbleText != "") bubbleText += "\r\n";
                    success = googleToOutlook(googleEntries, outlookEntries, ref bubbleText);
                    if (Sync.Engine.Instance.CancellationPending) return SyncResult.UserCancelled;
                }
                if (!success) return SyncResult.Fail;
                if (bubbleText != "") {
                    log.Info(bubbleText.Replace("\r\n", ". "));
                    System.Text.RegularExpressions.Regex rgx = new System.Text.RegularExpressions.Regex(@"\D");
                    String changes = rgx.Replace(bubbleText, "").Trim('0');
                    if (Settings.Instance.ShowSystemNotifications &&
                        (!Settings.Instance.ShowSystemNotificationsIfChange || !String.IsNullOrEmpty(changes))) Forms.Main.Instance.NotificationTray.ShowBubbleInfo(bubbleText);
                }

                return SyncResult.OK;
            }

            private Boolean outlookToGoogle(List<Microsoft.Graph.Event> outlookEntries, List<GcalData.Event> googleEntries, ref String bubbleText) {
                log.Debug("Synchronising from Outlook to Google.");
                if (this.Profile.SyncDirection.Id == Sync.Direction.Bidirectional.Id)
                    Forms.Main.Instance.Console.Update("Syncing " + Sync.Direction.OutlookToGoogle.Name, Console.Markup.syncDirection, newLine: false);

                //  Make copies of each list of events (Not strictly needed)
                List<Microsoft.Graph.Event> googleEntriesToBeCreated = new(outlookEntries);
                List<GcalData.Event> googleEntriesToBeDeleted = new(googleEntries);
                Dictionary<Microsoft.Graph.Event, GcalData.Event> entriesToBeCompared = new();

                Console console = Forms.Main.Instance.Console;

                DateTime timeSection = DateTime.Now;
                try {
                    Ogcs.Google.Graph.Calendar.IdentifyEventDifferences(ref googleEntriesToBeCreated, ref googleEntriesToBeDeleted, ref entriesToBeCompared);
                    if (Sync.Engine.Instance.CancellationPending) return false;
                } catch (System.Exception) {
                    console.Update("Unable to identify differences in Google calendar.", Console.Markup.error);
                    throw;
                }
                TimeSpan sectionDuration = DateTime.Now - timeSection;
                if (sectionDuration.TotalSeconds > 30) {
                    log.Warn("That step took a long time! Issue #599");
                    new Telemetry.GA4Event.Event(Telemetry.GA4Event.Event.Name.debug)
                        .AddParameter(GA4.General.github_issue, 599)
                        .AddParameter("section", "Ogcs.Google.Calendar.Instance.IdentifyEventDifferences()")
                        .AddParameter("duration", sectionDuration.TotalSeconds)
                        .AddParameter("items", entriesToBeCompared.Count)
                        .Send();
                }

                StringBuilder sb = new StringBuilder();
                console.BuildOutput(googleEntriesToBeDeleted.Count + " Google calendar entries to be deleted.", ref sb, false);
                console.BuildOutput(googleEntriesToBeCreated.Count + " Google calendar entries to be created.", ref sb, false);
                console.BuildOutput(entriesToBeCompared.Count + " calendar entries to be compared.", ref sb, false);
                console.Update(sb, Console.Markup.info, logit: true);

                //Protect against very first syncs which may trample pre-existing non-Outlook events in Google
                if (!this.Profile.DisableDelete && !this.Profile.ConfirmOnDelete &&
                    googleEntriesToBeDeleted.Count == googleEntries.Count && googleEntries.Count > 1) {
                    if (Ogcs.Extensions.MessageBox.Show("All Google events are going to be deleted. Do you want to allow this?" +
                        "\r\nNote, " + googleEntriesToBeCreated.Count + " events will then be created.", "Confirm mass deletion",
                        MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation) == DialogResult.No) {
                        googleEntriesToBeDeleted = new();

                    } else if (this.Profile.SyncDirection.Id == Sync.Direction.OutlookToGoogle.Id &&
                        Ogcs.Extensions.MessageBox.Show("If you are syncing an Apple iCalendar from Outlook and get the 'mass deletion' warning for every sync, " +
                        "would you like to read up on a potential solution?", "iCal Syncing?", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes) {
                        Helper.OpenBrowser("https://github.com/phw198/OutlookGoogleCalendarSync/wiki/Syncing-Apple-iCalendar-in-Outlook-causes-'mass-deletion'-warnings");
                    }
                }

                int entriesUpdated = 0;
                try {
                    #region Delete Google Entries
                    if (googleEntriesToBeDeleted.Count > 0) {
                        console.Update("Deleting " + googleEntriesToBeDeleted.Count + " Google calendar entries", Console.Markup.h2, newLine: false);
                        try {
                            Ogcs.Google.Calendar.Instance.DeleteCalendarEntries(googleEntriesToBeDeleted);
                        } catch (UserCancelledSyncException ex) {
                            log.Info(ex.Message);
                            return false;
                        } catch (System.Exception ex) {
                            console.UpdateWithError("Unable to delete obsolete entries in Google calendar.", ex);
                            throw;
                        }
                        log.Info("Done.");
                    }

                    if (Sync.Engine.Instance.CancellationPending) return false;
                    #endregion

                    #region Create Google Entries
                    if (googleEntriesToBeCreated.Count > 0) {
                        console.Update("Creating " + googleEntriesToBeCreated.Count + " Google calendar entries", Console.Markup.h2, newLine: false);
                        try {
                            Ogcs.Google.Graph.Calendar.CreateCalendarEntries(googleEntriesToBeCreated);
                        } catch (UserCancelledSyncException ex) {
                            log.Info(ex.Message);
                            return false;
                        } catch (System.Exception ex) {
                            console.UpdateWithError("Unable to add new entries into the Google Calendar.", ex);
                            throw;
                        }
                        log.Info("Done.");
                    }

                    if (Sync.Engine.Instance.CancellationPending) return false;
                    #endregion

                    #region Update Google Entries
                    if (entriesToBeCompared.Count > 0) {
                        console.Update("Comparing " + entriesToBeCompared.Count + " existing Google calendar entries", Console.Markup.h2, newLine: false);
                        try {
                            Ogcs.Google.Graph.Calendar.UpdateCalendarEntries(entriesToBeCompared, ref entriesUpdated);
                        } catch (UserCancelledSyncException ex) {
                            log.Info(ex.Message);
                            return false;
                        } catch (System.Exception ex) {
                            console.UpdateWithError("Unable to update existing entries in the Google calendar.", ex);
                            throw;
                        }
                        console.Update(entriesUpdated + " entries updated.");
                    }

                    if (Sync.Engine.Instance.CancellationPending) return false;
                    #endregion

                } finally {
                    bubbleText = "Google: " + googleEntriesToBeCreated.Count + " created; " +
                        googleEntriesToBeDeleted.Count + " deleted; " + entriesUpdated + " updated";

                    if (this.Profile.SyncDirection.Id == Direction.OutlookToGoogle.Id) {
                        while (entriesToBeCompared.Count() > 0) {
                            Outlook.Calendar.ReleaseObject(entriesToBeCompared.Keys.Last());
                            entriesToBeCompared.Remove(entriesToBeCompared.Keys.Last());
                        }
                    }
                }
                return true;
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

                    #region Create Outlook Entries
                    if (outlookEntriesToBeCreated.Count > 0) {
                        console.Update("Creating " + outlookEntriesToBeCreated.Count + " Outlook calendar entries", Console.Markup.h2, newLine: false);
                        try {
                            Outlook.Graph.Calendar.Instance.CreateCalendarEntries(outlookEntriesToBeCreated);
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
                            Outlook.Graph.Calendar.Instance.UpdateCalendarEntries(entriesToBeCompared, ref entriesUpdated);
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

            private SyncResult extirpateCustomProperties(List<Microsoft.Graph.Event> outlookEntries, List<GcalData.Event> googleEntries) {
                SyncResult returnVal = SyncResult.Fail;
                Console console = Forms.Main.Instance.Console;
                try {
                    console.Update("Cleansing OGCS metadata from Outlook items...", Console.Markup.h2, newLine: false);
                    for (int o = 0; o < outlookEntries.Count; o++) {
                        Microsoft.Graph.Event ai = outlookEntries[o];
                        Outlook.Graph.CustomProperty.LogProperties(ai, log4net.Core.Level.Debug);
                        if (Outlook.Graph.CustomProperty.Extirpate(ref ai)) {
                            console.Update(Outlook.Graph.Calendar.GetEventSummary(ai, out String anonSummary), anonSummary, Console.Markup.calendar);
                        }
                        if (Sync.Engine.Instance.CancellationPending) return SyncResult.UserCancelled;
                    }

                    console.Update("Cleansing OGCS metadata from Google items...", Console.Markup.h2, newLine: false);
                    for (int g = 0; g < googleEntries.Count; g++) {
                        GcalData.Event ev = googleEntries[g];
                        Ogcs.Google.CustomProperty.LogProperties(ev, log4net.Core.Level.Debug);
                        if (Ogcs.Google.CustomProperty.Extirpate(ref ev)) {
                            console.Update(Ogcs.Google.Calendar.GetEventSummary(ev, out String anonSummary), anonSummary, Console.Markup.calendar);
                            Ogcs.Google.Calendar.Instance.UpdateCalendarEntry_save(ref ev);
                        }
                        if (Sync.Engine.Instance.CancellationPending) return SyncResult.UserCancelled;
                    }
                    returnVal = SyncResult.OK;
                    return returnVal;

                } catch (System.Exception ex) {
                    ex.Analyse("Failed to fully cleanse metadata!");
                    console.UpdateWithError(null, ex);
                    returnVal = SyncResult.Fail;
                    return returnVal;

                } finally {
                    if (Sync.Engine.Instance.CancellationPending) {
                        console.Update("Not letting this process run to completion is <b>strongly discouraged</b>.<br>" +
                            "If you are two-way syncing and use OGCS for normal syncing again, unexpected behaviour will ensue.<br>" +
                            "It is recommended to rerun the metadata cleanse to completion.", Console.Markup.warning);
                    } else if (returnVal == SyncResult.Fail) {
                        console.Update(
                            "It is recommended to rerun the metadata cleanse to <b>successful completion</b> before using OGCS for normal syncing again.<br>" +
                            "If this is not possible and you wish to continue using OGCS, please " +
                            "<a href='https://github.com/phw198/OutlookGoogleCalendarSync/issues' target='_blank'>raise an issue</a> on the GitHub project.", Console.Markup.warning);
                    }
                }
            }
        }
    }
}