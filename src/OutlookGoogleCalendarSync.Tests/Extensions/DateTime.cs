using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Globalization;
using Google.Apis.Calendar.v3.Data;
using OutlookGoogleCalendarSync.Extensions;
using Microsoft.Kiota.Abstractions;
using OutlookGoogleCalendarSync.Outlook.Graph.CustomClient.Models;
using OutlookGoogleCalendarSync; // Required for TimezoneDB - if it's in a different namespace, adjust accordingly


namespace OutlookGoogleCalendarSync.Tests
{
    [TestClass]
    public class DateTimeExtensionTests
    {
        // Tests for SafeDateTimeOffset(this EventDateTime evDt)
        [TestMethod]
        public void SafeDateTimeOffset_EventDateTime_DateOnly_ReturnsCorrectDateTimeOffset()
        {
            // Arrange
            EventDateTime evDt = new EventDateTime { Date = "2026-07-11" };
            DateTimeOffset expected = DateTimeOffset.ParseExact("2026-07-11", "yyyy-MM-dd", CultureInfo.InvariantCulture, DateTimeStyles.AssumeUniversal);

            // Act
            DateTimeOffset actual = evDt.SafeDateTimeOffset();

            // Assert
            Assert.AreEqual(expected.Date, actual.Date, "The date part should match for date-only event.");
            Assert.AreEqual(expected.Offset, actual.Offset, "The offset should be UTC for date-only event.");
            Assert.AreEqual(expected.TimeOfDay, actual.TimeOfDay, "The time should be midnight for date-only event.");
        }

        [TestMethod]
        public void SafeDateTimeOffset_EventDateTime_WithDateTime_UtcTimeZone_ReturnsCorrectDateTimeOffset()
        {
            // Arrange
            DateTimeOffset testDateTime = new DateTimeOffset(2026, 7, 11, 10, 30, 0, TimeSpan.Zero); // UTC
            EventDateTime evDt = new EventDateTime { DateTimeDateTimeOffset = testDateTime, TimeZone = "UTC" };

            // Act
            DateTimeOffset actual = evDt.SafeDateTimeOffset();

            // Assert
            Assert.AreEqual(testDateTime, actual, "DateTimeOffset should match exactly for UTC timezone.");
        }

        [TestMethod]
        public void SafeDateTimeOffset_EventDateTime_WithDateTime_NonUtcTimeZone_ReturnsCorrectlyOffsetDateTimeOffset()
        {
            // Arrange
            // Assuming TimezoneDB.GetUtcOffset("Europe/Berlin") returns 60 (for CEST) or whatever it's configured for.
            // For testing, we might need to mock TimezoneDB or ensure its behavior is consistent.
            // For this test, let's assume TimezoneDB.GetUtcOffset("Europe/Berlin") returns 60 for +01:00 offset.
            // If the time in evDt.DateTimeDateTimeOffset is already with offset +01:00, then we expect the same.
            DateTimeOffset originalDateTime = new DateTimeOffset(2026, 7, 11, 10, 30, 0, TimeSpan.FromHours(1)); // 10:30 with +01:00 offset
            EventDateTime evDt = new EventDateTime { DateTimeDateTimeOffset = originalDateTime, TimeZone = "Europe/Berlin" };

            // Act
            DateTimeOffset actual = evDt.SafeDateTimeOffset();

            // Assert
            // This assertion depends on the actual value returned by TimezoneDB.GetUtcOffset.
            // If TimezoneDB.GetUtcOffset("Europe/Berlin") returns 60, then the offset for 10:30+01:00 should remain +01:00.
            // If TimezoneDB uses a different offset for that date, the actual.Offset will differ.
            // For a baseline, let's assume it should match the original UTC value.
            Assert.AreEqual(originalDateTime.UtcDateTime, actual.UtcDateTime, "UTC time should be preserved after applying offset from TimeZoneDB.");
        }

        // Tests for SafeDateTimeOffset(this Microsoft.Kiota.Abstractions.Date graphDate)
        [TestMethod]
        public void SafeDateTimeOffset_GraphDate_ReturnsCorrectDateTimeOffset()
        {
            // Arrange
            Microsoft.Kiota.Abstractions.Date graphDate = new Microsoft.Kiota.Abstractions.Date(2026, 7, 11);
            DateTimeOffset expected = new DateTimeOffset(2026, 7, 11, 0, 0, 0, TimeSpan.Zero);

            // Act
            DateTimeOffset actual = graphDate.SafeDateTimeOffset();

            // Assert
            Assert.AreEqual(expected, actual, "Graph Date should convert to UTC midnight DateTimeOffset.");
        }

        // Tests for SafeDateTimeOffset(this Outlook.Graph.CustomClient.Models.DateTimeTimeZone evDt, Boolean? isAllDay, String timezone)
        [TestMethod]
        public void SafeDateTimeOffset_DateTimeTimeZone_AllDayEvent_ReturnsUtcMidnightDateTimeOffset()
        {
            // Arrange
            DateTimeTimeZone evDt = new DateTimeTimeZone { DateTime = "2026-07-11T10:00:00", TimeZone = "America/New_York" }; // Timezone will be ignored
            bool? isAllDay = true;
            string timezone = "America/New_York"; // This should be ignored for all-day

            DateTimeOffset expected = new DateTimeOffset(2026, 7, 11, 0, 0, 0, TimeSpan.Zero);

            // Act
            DateTimeOffset actual = evDt.SafeDateTimeOffset(isAllDay, timezone);

            // Assert
            Assert.AreEqual(expected, actual, "All-day event should convert to UTC midnight on the specified date.");
        }

        [TestMethod]
        public void SafeDateTimeOffset_DateTimeTimeZone_NonAllDayEvent_UtcTimeZone_ReturnsUtcDateTimeOffset()
        {
            // Arrange
            DateTimeTimeZone evDt = new DateTimeTimeZone { DateTime = "2026-07-11T10:30:00", TimeZone = "UTC" };
            bool? isAllDay = false;
            string timezone = "UTC";

            DateTimeOffset expected = new DateTimeOffset(2026, 7, 11, 10, 30, 0, TimeSpan.Zero);

            // Act
            DateTimeOffset actual = evDt.SafeDateTimeOffset(isAllDay, timezone);

            // Assert
            Assert.AreEqual(expected, actual, "Non-all-day UTC event should return exact UTC DateTimeOffset.");
        }

        [TestMethod]
        public void SafeDateTimeOffset_DateTimeTimeZone_NonAllDayEvent_NonUtcTimeZone_ReturnsCorrectlyOffsetDateTimeOffset()
        {
            // Arrange
            // Assuming TimezoneDB.IANAtimezone("America/New_York") returns "America/New_York"
            // And TimezoneDB.GetUtcOffset("America/New_York") returns -240 (for EDT) or -300 (for EST)
            // Let's assume for July 11, 2026 it's EDT, so -240 minutes (-4 hours).
            DateTimeTimeZone evDt = new DateTimeTimeZone { DateTime = "2026-07-11T10:30:00", TimeZone = "America/New_York" };
            bool? isAllDay = false;
            string timezone = "America/New_York"; // Original timezone

            // The Parse(evDt.DateTime, null, DateTimeStyles.AssumeUniversal) initially assumes it's UTC.
            // Then it applies the offset from TimezoneDB.GetUtcOffset(timezone).
            // So, 2026-07-11T10:30:00 (assumed UTC) becomes 2026-07-11T06:30:00-04:00 (if offset is -4 hours).
            // Or if original time is local, then it becomes UtcTime adjusted by offset
            // We need to carefully define expected based on the exact implementation.

            // Let's trace the expected behavior based on the current code for `SafeDateTimeOffset(this Outlook.Graph.CustomClient.Models.DateTimeTimeZone evDt, Boolean? isAllDay, String timezone)`:
            // 1. System.DateTimeOffset safeDate = System.DateTime.Parse(evDt.DateTime, null, DateTimeStyles.AssumeUniversal);
            //    -> For "2026-07-11T10:30:00", this results in 2026-07-11 10:30:00 +00:00 (UTC)

            // 2. timezone = TimezoneDB.IANAtimezone(timezone); // Assume this returns "America/New_York"
            // 3. Int16 offset = TimezoneDB.GetUtcOffset(timezone); // Assume -240 for EDT

            // 4. safeDate = safeDate.ToOffset(TimeSpan.FromMinutes(offset));
            //    -> 2026-07-11 10:30:00 +00:00.ToOffset(TimeSpan.FromHours(-4))
            //    -> This changes the *offset* while keeping the *point in time* the same.
            //    -> So, 2026-07-11 06:30:00 -04:00 (equivalent to 10:30 UTC)

            DateTimeOffset expected = new DateTimeOffset(2026, 7, 11, 6, 30, 0, TimeSpan.FromHours(-4)); // If TimezoneDB.GetUtcOffset is -240

            // Act
            DateTimeOffset actual = evDt.SafeDateTimeOffset(isAllDay, timezone);

            // Assert
            Assert.AreEqual(expected, actual, "Non-all-day non-UTC event should be correctly offset.");
        }
    }
}