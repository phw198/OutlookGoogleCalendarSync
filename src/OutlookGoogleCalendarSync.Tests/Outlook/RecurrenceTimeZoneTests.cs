using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using Ogcs = OutlookGoogleCalendarSync;

namespace OutlookGoogleCalendarSync.Tests.Outlook {
    [TestClass]
    public class RecurrenceTimeZoneTests {
        private static string pacificTimeZone => Environment.OSVersion.Platform == PlatformID.Win32NT
            ? "Pacific Standard Time" : "America/Los_Angeles";

        [TestMethod]
        public void TimeInOutlookZone_PreviousDay_UsesTargetDateAndClock() {
            DateTimeOffset start = new DateTimeOffset(2026, 10, 25, 0, 30, 0, TimeSpan.Zero);

            DateTime converted = Ogcs.Outlook.Recurrence.TimeInOutlookZone(start, pacificTimeZone);

            Assert.AreEqual(new DateTime(2026, 10, 24, 17, 30, 0), converted);
        }

        [TestMethod]
        public void TimeInOutlookZone_NextDay_UsesTargetDateAndClock() {
            DateTimeOffset start = new DateTimeOffset(2026, 10, 25, 17, 30, 0, TimeSpan.Zero);

            DateTime converted = Ogcs.Outlook.Recurrence.TimeInOutlookZone(start,
                Environment.OSVersion.Platform == PlatformID.Win32NT ? "Tokyo Standard Time" : "Asia/Tokyo");

            Assert.AreEqual(new DateTime(2026, 10, 26, 2, 30, 0), converted);
        }

        [TestMethod]
        public void TimeInOutlookZone_RespectsOffsetAtEventInstant() {
            DateTimeOffset start = new DateTimeOffset(2026, 3, 8, 10, 30, 0, TimeSpan.Zero);

            DateTime converted = Ogcs.Outlook.Recurrence.TimeInOutlookZone(start, pacificTimeZone);

            Assert.AreEqual(new DateTime(2026, 3, 8, 3, 30, 0), converted);
        }

        [TestMethod]
        public void TimeInOutlookZone_EndTimeUsesStartZoneEvenWhenEventEndHasAnotherZone() {
            DateTimeOffset start = new DateTimeOffset(2026, 1, 8, 3, 0, 0, TimeSpan.FromHours(-5));
            DateTimeOffset end = new DateTimeOffset(2026, 1, 8, 14, 0, 0, TimeSpan.FromMinutes(270));
            string easternTimeZone = Environment.OSVersion.Platform == PlatformID.Win32NT
                ? "Eastern Standard Time" : "America/New_York";

            Assert.AreEqual(new DateTime(2026, 1, 8, 3, 0, 0),
                Ogcs.Outlook.Recurrence.TimeInOutlookZone(start, easternTimeZone));
            Assert.AreEqual(new DateTime(2026, 1, 8, 4, 30, 0),
                Ogcs.Outlook.Recurrence.TimeInOutlookZone(end, easternTimeZone));
        }

        [TestMethod]
        public void EndDate_UsesOffsetAtRecurrenceEndAcrossDaylightSavingTransitions() {
            DateTime beforeSpringForward = Ogcs.Google.Recurrence.EndDate("20260308T073000Z", "America/Los_Angeles");
            DateTime beforeFallBack = Ogcs.Google.Recurrence.EndDate("20261101T073000Z", "America/Los_Angeles");

            Assert.AreEqual(new DateTime(2026, 3, 7), beforeSpringForward);
            Assert.AreEqual(new DateTime(2026, 11, 1), beforeFallBack);
        }
    }
}
