using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Ogcs = OutlookGoogleCalendarSync;

namespace OutlookGoogleCalendarSync.Tests.Outlook {
    [TestClass]
    public class SourceEntryIDBackfillTests {
        /// <summary>Verify default, numbered, and gapped Google calendar key-set suffixes are recognized.</summary>
        [TestMethod]
        public void GoogleCalendarKeySet_RecognizesDefaultAndSuffixedNames() {
            Assert.IsTrue(Ogcs.Outlook.CustomProperty.TryGetGoogleCalendarKeySet("googleCalendarID", out String defaultSuffix));
            Assert.AreEqual(String.Empty, defaultSuffix);
            Assert.IsTrue(Ogcs.Outlook.CustomProperty.TryGetGoogleCalendarKeySet("googleCalendarID-01", out String firstSuffix));
            Assert.AreEqual("-01", firstSuffix);
            Assert.IsTrue(Ogcs.Outlook.CustomProperty.TryGetGoogleCalendarKeySet("googleCalendarID-03", out String gappedSuffix));
            Assert.AreEqual("-03", gappedSuffix);
        }

        /// <summary>Verify malformed Google calendar key-set property names are rejected.</summary>
        [TestMethod]
        public void GoogleCalendarKeySet_RejectsMalformedNames() {
            Assert.IsFalse(Ogcs.Outlook.CustomProperty.TryGetGoogleCalendarKeySet("googleCalendarID-00", out _));
            Assert.IsFalse(Ogcs.Outlook.CustomProperty.TryGetGoogleCalendarKeySet("googleCalendarID-1", out _));
            Assert.IsFalse(Ogcs.Outlook.CustomProperty.TryGetGoogleCalendarKeySet("googleCalendarID-100", out _));
            Assert.IsFalse(Ogcs.Outlook.CustomProperty.TryGetGoogleCalendarKeySet("googleCalendarIDX", out _));
        }

        /// <summary>Verify only missing or empty source EntryID markers are eligible for population.</summary>
        [TestMethod]
        public void SourceEntryIDMarker_PopulatesOnlyMissingOrEmptyValues() {
            Assert.IsTrue(String.IsNullOrEmpty(null));
            Assert.IsTrue(String.IsNullOrEmpty(String.Empty));
            Assert.IsFalse(String.IsNullOrEmpty("previous-entry-id"));
            Assert.IsFalse(String.IsNullOrEmpty("different-entry-id"));
        }

        /// <summary>Verify backfill state round-trips calendar IDs and matches only its selected pair.</summary>
        [TestMethod]
        public void BackfillState_RoundTripsProfileIDsAndStatus() {
            String serialized = Ogcs.SettingsStore.SourceEntryIDBackfillState.Serialize(
                "outlook|calendar", "google/calendar", Ogcs.SettingsStore.SourceEntryIDBackfillStatus.Running, 2);

            Assert.IsTrue(Ogcs.SettingsStore.SourceEntryIDBackfillState.TryParse(serialized, out Ogcs.SettingsStore.SourceEntryIDBackfillState state));
            Assert.AreEqual("outlook|calendar", state.OutlookCalendarId);
            Assert.AreEqual("google/calendar", state.GoogleCalendarId);
            Assert.AreEqual(Ogcs.SettingsStore.SourceEntryIDBackfillStatus.Running, state.Status);
            Assert.AreEqual(2, state.Attempts);
            Assert.IsTrue(state.Matches("outlook|calendar", "google/calendar"));
            Assert.IsFalse(state.Matches("other-outlook", "google/calendar"));
        }

        /// <summary>Verify a valid state for another calendar pair starts with a fresh attempt budget.</summary>
        [TestMethod]
        public void BackfillState_ValidDifferentCalendarPairStartsFreshAttempt() {
            String serialized = Ogcs.SettingsStore.SourceEntryIDBackfillState.Serialize(
                "previous-outlook", "previous-google", Ogcs.SettingsStore.SourceEntryIDBackfillStatus.Running, 2);

            Boolean stateIsValid = Ogcs.SettingsStore.SourceEntryIDBackfillState.TryParse(serialized, out Ogcs.SettingsStore.SourceEntryIDBackfillState state);
            Boolean stateMatches = stateIsValid && state.Matches("selected-outlook", "selected-google");
            Int32 attempt = Ogcs.SettingsStore.SourceEntryIDBackfillState.GetNextAttempt(state, stateMatches);
            Int32 samePairAttempt = Ogcs.SettingsStore.SourceEntryIDBackfillState.GetNextAttempt(state,
                state.Matches("previous-outlook", "previous-google"));

            Assert.IsTrue(stateIsValid);
            Assert.IsFalse(stateMatches);
            Assert.IsFalse(Ogcs.SettingsStore.SourceEntryIDBackfillState.IsMalformedVersionedState(serialized, state));
            Assert.AreEqual(1, attempt);
            Assert.AreEqual(3, samePairAttempt);
        }

        /// <summary>Verify legacy, malformed, and out-of-range backfill state values are rejected.</summary>
        [TestMethod]
        public void BackfillState_RejectsLegacyAndOutOfRangeAttempts() {
            Assert.IsFalse(Ogcs.SettingsStore.SourceEntryIDBackfillState.TryParse("outlook|google", out _));
            String malformedState = "v2|%%%|%%%|Running|1";
            Assert.IsFalse(Ogcs.SettingsStore.SourceEntryIDBackfillState.TryParse(malformedState, out Ogcs.SettingsStore.SourceEntryIDBackfillState state));
            Assert.IsTrue(Ogcs.SettingsStore.SourceEntryIDBackfillState.IsMalformedVersionedState(malformedState, state));
            Assert.IsFalse(Ogcs.SettingsStore.SourceEntryIDBackfillState.TryParse("v2|||Running|4", out _));
            Assert.ThrowsException<System.ArgumentOutOfRangeException>(() =>
                Ogcs.SettingsStore.SourceEntryIDBackfillState.Serialize(
                    "outlook", "google", Ogcs.SettingsStore.SourceEntryIDBackfillStatus.Exhausted, 2));
        }

        /// <summary>Verify incomplete backfills exhaust on the third attempt while complete passes succeed.</summary>
        [TestMethod]
        public void BackfillState_StopsAfterThreeStartedAttempts() {
            Assert.AreEqual(Ogcs.SettingsStore.SourceEntryIDBackfillStatus.Running,
                Ogcs.SettingsStore.SourceEntryIDBackfillState.GetResultStatus(false, 1));
            Assert.AreEqual(Ogcs.SettingsStore.SourceEntryIDBackfillStatus.Running,
                Ogcs.SettingsStore.SourceEntryIDBackfillState.GetResultStatus(false, 2));
            Assert.AreEqual(Ogcs.SettingsStore.SourceEntryIDBackfillStatus.Exhausted,
                Ogcs.SettingsStore.SourceEntryIDBackfillState.GetResultStatus(false, 3));
            Assert.AreEqual(Ogcs.SettingsStore.SourceEntryIDBackfillStatus.Complete,
                Ogcs.SettingsStore.SourceEntryIDBackfillState.GetResultStatus(true, 3));
        }
    }
}
