using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using Ogcs = OutlookGoogleCalendarSync;

namespace OutlookGoogleCalendarSync.Tests.SettingsStore {
    [TestClass]
    public class ReminderDndTests {
        [TestMethod]
        public void SameDayReminderDnd_IncludesWindowBoundaries() {
            Ogcs.SettingsStore.Calendar profile = CreateProfile(
                new DateTime(2026, 1, 1, 9, 0, 0),
                new DateTime(2026, 1, 1, 17, 0, 0));

            Assert.IsTrue(profile.IsReminderDNDActive(new DateTime(2026, 1, 1, 9, 0, 0)));
            Assert.IsTrue(profile.IsReminderDNDActive(new DateTime(2026, 1, 1, 13, 0, 0)));
            Assert.IsTrue(profile.IsReminderDNDActive(new DateTime(2026, 1, 1, 17, 0, 0)));
            Assert.IsFalse(profile.IsReminderDNDActive(new DateTime(2026, 1, 1, 8, 59, 0)));
            Assert.IsFalse(profile.IsReminderDNDActive(new DateTime(2026, 1, 1, 17, 1, 0)));
        }

        [TestMethod]
        public void OvernightReminderDnd_UsesExistingWindowSemantics() {
            Ogcs.SettingsStore.Calendar profile = CreateProfile(
                new DateTime(2026, 1, 1, 22, 0, 0),
                new DateTime(2026, 1, 1, 6, 0, 0));

            Assert.IsTrue(profile.IsReminderDNDActive(new DateTime(2026, 1, 1, 22, 0, 0)));
            Assert.IsTrue(profile.IsReminderDNDActive(new DateTime(2026, 1, 1, 23, 0, 0)));
            Assert.IsTrue(profile.IsReminderDNDActive(new DateTime(2026, 1, 2, 5, 0, 0)));
            Assert.IsTrue(profile.IsReminderDNDActive(new DateTime(2026, 1, 2, 6, 0, 0)));
            Assert.IsFalse(profile.IsReminderDNDActive(new DateTime(2026, 1, 1, 21, 59, 0)));
            Assert.IsFalse(profile.IsReminderDNDActive(new DateTime(2026, 1, 2, 6, 1, 0)));
        }

        [TestMethod]
        public void BidirectionalReminderDnd_IsNeverActive() {
            Ogcs.SettingsStore.Calendar profile = CreateProfile(
                new DateTime(2026, 1, 1, 9, 0, 0),
                new DateTime(2026, 1, 1, 17, 0, 0));
            profile.SyncDirection = Ogcs.Sync.Direction.Bidirectional;

            Assert.IsFalse(profile.IsReminderDNDActive(new DateTime(2026, 1, 1, 13, 0, 0)));
        }

        private static Ogcs.SettingsStore.Calendar CreateProfile(DateTime start, DateTime end) {
            return new Ogcs.SettingsStore.Calendar {
                SyncDirection = Ogcs.Sync.Direction.OutlookToGoogle,
                ReminderDND = true,
                ReminderDNDstart = start,
                ReminderDNDend = end
            };
        }
    }
}
