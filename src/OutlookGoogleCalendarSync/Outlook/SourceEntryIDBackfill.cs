using Microsoft.Office.Interop.Outlook;
using System;

namespace OutlookGoogleCalendarSync.Outlook {
    public partial class Calendar {
        /// <summary>
        /// Backfill missing Outlook source EntryIDs for existing Google-linked appointments.
        /// Concept introduced in v3.0.6 to support pasted copies of appointments being treated as separate items instead of replacing their originals.
        /// </summary>
        public void BackfillSourceEntryIDs(SettingsStore.Calendar profile) {
            if (profile.SyncDirection.Id != Sync.Direction.Bidirectional.Id) return;

            String outlookCalendarId = profile.UseOutlookCalendar?.Id ?? String.Empty;
            String googleCalendarId = profile.UseGoogleCalendar?.Id ?? String.Empty;
            SettingsStore.SourceEntryIDBackfillState state = null;
            Boolean stateMatches = SettingsStore.SourceEntryIDBackfillState.TryParse(profile.SourceEntryIDBackfillProfile, out state) &&
                state.Matches(outlookCalendarId, googleCalendarId);

            if (SettingsStore.SourceEntryIDBackfillState.IsMalformedVersionedState(profile.SourceEntryIDBackfillProfile, state)) {
                String invalidStateWarning = "The source EntryID backfill state for profile '" + profile._ProfileName +
                    "' is invalid. No automatic backfill will run; close OGCS, back up settings.xml, and clear this profile's " +
                    "SourceEntryIDBackfillProfile setting to retry.";
                log.Error(invalidStateWarning);
                Forms.Main.Instance.Console.Update(invalidStateWarning, Console.Markup.warning);
                return;
            }

            if (stateMatches && state.Status == SettingsStore.SourceEntryIDBackfillStatus.Complete) return;

            if (stateMatches && state.Status == SettingsStore.SourceEntryIDBackfillStatus.Exhausted) {
                reportSourceEntryIDBackfillExhausted(profile);
                return;
            }

            if (stateMatches && state.Attempts >= SettingsStore.SourceEntryIDBackfillState.MaximumAttempts) {
                String exhaustedState = SettingsStore.SourceEntryIDBackfillState.Serialize(outlookCalendarId, googleCalendarId,
                    SettingsStore.SourceEntryIDBackfillStatus.Exhausted, SettingsStore.SourceEntryIDBackfillState.MaximumAttempts);
                profile.TryPersistSourceEntryIDBackfillProfile(exhaustedState);
                reportSourceEntryIDBackfillExhausted(profile);
                return;
            }

            Int32 attempt = SettingsStore.SourceEntryIDBackfillState.GetNextAttempt(state, stateMatches);
            String runningState = SettingsStore.SourceEntryIDBackfillState.Serialize(outlookCalendarId, googleCalendarId,
                SettingsStore.SourceEntryIDBackfillStatus.Running, attempt);
            if (!profile.TryPersistSourceEntryIDBackfillProfile(runningState)) {
                String settingsWarning = "Could not save the source EntryID backfill state for profile '" + profile._ProfileName +
                    "'. The backfill was not started; check that settings.xml is writable.";
                log.Warn(settingsWarning);
                Forms.Main.Instance.Console.Update(settingsWarning, Console.Markup.warning);
                return;
            }

            log.Info("Source EntryID backfill attempt " + attempt + " of " + SettingsStore.SourceEntryIDBackfillState.MaximumAttempts +
                " for profile '" + profile._ProfileName + "'.");

            MAPIFolder folder = null;
            Items items = null;
            Int32 stamped = 0;
            Int32 failedStamps = 0;
            Boolean passComplete = true;
            System.Exception folderException = null;
            try {
                folder = IOutlook.GetFolderByID(outlookCalendarId);
                items = folder.Items;
                Int32 itemCount = items.Count;
                for (Int32 index = 1; index <= itemCount; index++) {
                    if (Sync.Engine.Instance.CancellationPending) {
                        passComplete = false;
                        break;
                    }

                    Object item = null;
                    AppointmentItem ai = null;
                    try {
                        item = items[index];
                        ai = item as AppointmentItem;
                        if (ai != null && CustomProperty.PopulateMissingSourceEntryIDs(ai)) {
                            ai.Save();
                            stamped++;
                        }
                    } catch (System.Exception ex) {
                        ex.LogAsFail().Analyse("Failed to backfill the Outlook source EntryID.");
                        failedStamps++;
                        passComplete = false;
                    } finally {
                        item = ReleaseObject(item);
                        ai = null;
                    }
                }
            } catch (System.Exception ex) {
                folderException = ex;
                passComplete = false;
                ex.LogAsFail().Analyse("Failed to enumerate Outlook appointments for source EntryID backfill.");
            } finally {
                items = (Items)ReleaseObject(items);
                folder = (MAPIFolder)ReleaseObject(folder);
            }

            if (Sync.Engine.Instance.CancellationPending) passComplete = false;

            SettingsStore.SourceEntryIDBackfillStatus resultStatus = SettingsStore.SourceEntryIDBackfillState.GetResultStatus(passComplete, attempt);
            String resultState = SettingsStore.SourceEntryIDBackfillState.Serialize(outlookCalendarId, googleCalendarId, resultStatus, attempt);
            if (!profile.TryPersistSourceEntryIDBackfillProfile(resultState)) {
                String settingsWarning = "Could not save the result of source EntryID backfill for profile '" + profile._ProfileName +
                    "'. Check that settings.xml is writable; the next sync may retry this attempt.";
                log.Warn(settingsWarning);
                Forms.Main.Instance.Console.Update(settingsWarning, Console.Markup.warning);
            }

            log.Info("Backfilled source EntryIDs for " + stamped + " Outlook appointments.");
            if (failedStamps > 0) log.Warn(failedStamps + " Outlook appointments failed during source EntryID backfill.");
            if (resultStatus == SettingsStore.SourceEntryIDBackfillStatus.Exhausted)
                reportSourceEntryIDBackfillExhausted(profile);
            else if (!passComplete)
                log.Warn("Source EntryID backfill attempt " + attempt + " was incomplete; it will be retried on a later sync.");

            if (folderException != null)
                System.Runtime.ExceptionServices.ExceptionDispatchInfo.Capture(folderException).Throw();
        }

        /// <summary>
        /// Warn that automatic source EntryID backfill attempts are exhausted and explain how to retry them.
        /// </summary>
        private static void reportSourceEntryIDBackfillExhausted(SettingsStore.Calendar profile) {
            String warning = "Source EntryID backfill for profile '" + profile._ProfileName + "' did not complete after " +
                SettingsStore.SourceEntryIDBackfillState.MaximumAttempts + " attempts. It will not rescan automatically. To retry, close OGCS, " +
                "back up settings.xml, and clear this profile's SourceEntryIDBackfillProfile setting.";
            log.Warn(warning);
            Forms.Main.Instance.Console.Update(warning, Console.Markup.warning);
        }
    }
}
