/*
 * This file stores the retry state and profile persistence for the one-time source EntryID backfill.
 * The backfill fills missing Outlook EntryID markers for each Google calendar key set on existing appointments.
 * During later syncs, OGCS compares those markers with the current Outlook EntryID to recognize pasted copies,
 * so a copy can be treated as a separate item instead of being mistaken for or replacing its original.
 */
using System;
using System.Globalization;
using System.Linq;
using System.Runtime.Serialization;
using System.Text;
using System.Xml.Linq;

namespace OutlookGoogleCalendarSync.SettingsStore {
    /// <summary>
    /// Describes the persisted state of the one-time source EntryID backfill.
    /// </summary>
    internal enum SourceEntryIDBackfillStatus {
        Running,
        Complete,
        Exhausted
    }

    /// <summary>
    /// Encodes and validates retry state for a source EntryID backfill and calendar pair.
    /// </summary>
    internal sealed class SourceEntryIDBackfillState {
        internal const Int32 MaximumAttempts = 3;
        private const String Version = "v2";

        internal String OutlookCalendarId { get; private set; }
        internal String GoogleCalendarId { get; private set; }
        internal SourceEntryIDBackfillStatus Status { get; private set; }
        internal Int32 Attempts { get; private set; }

        /// <summary>
        /// Create a parsed backfill state for the specified calendar pair and attempt status.
        /// </summary>
        private SourceEntryIDBackfillState(String outlookCalendarId, String googleCalendarId,
            SourceEntryIDBackfillStatus status, Int32 attempts) {
            OutlookCalendarId = outlookCalendarId;
            GoogleCalendarId = googleCalendarId;
            Status = status;
            Attempts = attempts;
        }

        /// <summary>
        /// Serialize a backfill state into a versioned, delimiter-safe settings value.
        /// </summary>
        internal static String Serialize(String outlookCalendarId, String googleCalendarId,
            SourceEntryIDBackfillStatus status, Int32 attempts) {
            if (attempts < 1 || attempts > MaximumAttempts ||
                (status == SourceEntryIDBackfillStatus.Exhausted && attempts != MaximumAttempts))
                throw new ArgumentOutOfRangeException("attempts");

            return Version + "|" + Encode(outlookCalendarId) + "|" + Encode(googleCalendarId) + "|" +
                status.ToString() + "|" + attempts.ToString(CultureInfo.InvariantCulture);
        }

        /// <summary>
        /// Parse a versioned backfill settings value, returning false for legacy or malformed data.
        /// </summary>
        internal static Boolean TryParse(String value, out SourceEntryIDBackfillState state) {
            state = null;
            String[] parts = (value ?? String.Empty).Split('|');
            if (parts.Length != 5 || parts[0] != Version) return false;

            try {
                if (!Enum.TryParse(parts[3], false, out SourceEntryIDBackfillStatus status) ||
                    !Enum.IsDefined(typeof(SourceEntryIDBackfillStatus), status) ||
                    !Int32.TryParse(parts[4], NumberStyles.None, CultureInfo.InvariantCulture, out Int32 attempts) ||
                    attempts < 1 || attempts > MaximumAttempts ||
                    (status == SourceEntryIDBackfillStatus.Exhausted && attempts != MaximumAttempts))
                    return false;

                state = new SourceEntryIDBackfillState(Decode(parts[1]), Decode(parts[2]), status, attempts);
                return true;
            } catch (FormatException) {
                return false;
            }
        }

        /// <summary>
        /// Determine whether this state belongs to the specified Outlook and Google calendar pair.
        /// </summary>
        internal Boolean Matches(String outlookCalendarId, String googleCalendarId) {
            return OutlookCalendarId == (outlookCalendarId ?? String.Empty) &&
                GoogleCalendarId == (googleCalendarId ?? String.Empty);
        }

        /// <summary>
        /// Determine whether a versioned state is malformed rather than valid for another calendar pair.
        /// </summary>
        internal static Boolean IsMalformedVersionedState(String value, SourceEntryIDBackfillState state) {
            return (value ?? String.Empty).StartsWith(Version + "|", StringComparison.Ordinal) && state == null;
        }

        /// <summary>
        /// Select the next attempt number for the selected calendar pair.
        /// </summary>
        internal static Int32 GetNextAttempt(SourceEntryIDBackfillState state, Boolean stateMatches) {
            return stateMatches ? state.Attempts + 1 : 1;
        }

        /// <summary>
        /// Select the persisted status after a backfill pass completes or fails.
        /// </summary>
        internal static SourceEntryIDBackfillStatus GetResultStatus(Boolean passComplete, Int32 attempts) {
            if (attempts < 1 || attempts > MaximumAttempts)
                throw new ArgumentOutOfRangeException("attempts");
            if (passComplete) return SourceEntryIDBackfillStatus.Complete;
            return attempts == MaximumAttempts ? SourceEntryIDBackfillStatus.Exhausted : SourceEntryIDBackfillStatus.Running;
        }

        /// <summary>
        /// Encode a calendar ID as UTF-8 Base64 for storage in the delimited state value.
        /// </summary>
        private static String Encode(String value) {
            return Convert.ToBase64String(Encoding.UTF8.GetBytes(value ?? String.Empty));
        }

        /// <summary>
        /// Decode a UTF-8 Base64 calendar ID from the persisted state value.
        /// </summary>
        private static String Decode(String value) {
            return Encoding.UTF8.GetString(Convert.FromBase64String(value));
        }
    }

    public partial class Calendar {
        private String sourceEntryIDBackfillProfile;

        [DataMember] public String SourceEntryIDBackfillProfile {
            get { return sourceEntryIDBackfillProfile; }
            set {
                sourceEntryIDBackfillProfile = value;
                if (!Settings.Instance.Loading())
                    XMLManager.ExportElement(this, "SourceEntryIDBackfillProfile", value ?? "", Settings.ConfigFile);
            }
        }

        /// <summary>
        /// Persist the backfill state for this profile and verify the resulting settings XML value.
        /// </summary>
        internal Boolean TryPersistSourceEntryIDBackfillProfile(String value) {
            String previousValue = sourceEntryIDBackfillProfile;
            sourceEntryIDBackfillProfile = value ?? String.Empty;
            if (Settings.Instance.Loading()) return true;

            try {
                XMLManager.ExportElement(this, "SourceEntryIDBackfillProfile", sourceEntryIDBackfillProfile, Settings.ConfigFile);

                XNamespace ns = "http://schemas.datacontract.org/2004/07/OutlookGoogleCalendarSync";
                XElement settings = XDocument.Load(Settings.ConfigFile).Descendants(ns + "Settings").FirstOrDefault();
                XElement profile = settings?.Elements(ns + "Calendars").Elements(ns + "Calendar")
                    .FirstOrDefault(element => (String)element.Element(ns + "_ProfileName") == _ProfileName);
                if ((String)profile?.Element(ns + "SourceEntryIDBackfillProfile") == sourceEntryIDBackfillProfile)
                    return true;

                log.Error("Failed to verify source EntryID backfill state in settings for profile '" + _ProfileName + "'.");
            } catch (System.Exception ex) {
                ex.Analyse("Failed to persist source EntryID backfill state for profile '" + _ProfileName + "'.");
            }

            sourceEntryIDBackfillProfile = previousValue;
            return false;
        }
    }
}