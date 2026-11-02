using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.Linq;
using System.Xml.Linq;
using Ogcs = OutlookGoogleCalendarSync;

namespace OutlookGoogleCalendarSync.Tests.SettingsStore {
    [TestClass]
    public class UpgradeTests {
        private const string SettingsNamespace = "http://schemas.datacontract.org/2004/07/OutlookGoogleCalendarSync";

        [TestMethod]
        public void BidirectionalReminderDnd_LowercaseTrue_IsDisabled() {
            XDocument xml = CreateCalendarXml(Ogcs.Sync.Direction.Bidirectional.Id, "true");

            Ogcs.SettingsStore.Upgrade.upgradeBidirectionalReminderDnd(xml);

            Assert.AreEqual("false", GetReminderDndValue(xml));
        }

        [TestMethod]
        public void BidirectionalReminderDnd_TitlecaseTrue_IsDisabled() {
            XDocument xml = CreateCalendarXml(Ogcs.Sync.Direction.Bidirectional.Id, "True");

            Ogcs.SettingsStore.Upgrade.upgradeBidirectionalReminderDnd(xml);

            Assert.AreEqual("false", GetReminderDndValue(xml));
        }

        [TestMethod]
        public void BidirectionalReminderDnd_False_RemainsDisabled() {
            XDocument xml = CreateCalendarXml(Ogcs.Sync.Direction.Bidirectional.Id, "false");

            Ogcs.SettingsStore.Upgrade.upgradeBidirectionalReminderDnd(xml);

            Assert.AreEqual("false", GetReminderDndValue(xml));
        }

        [TestMethod]
        public void OneWayReminderDnd_True_RemainsEnabled() {
            XDocument xml = CreateCalendarXml(Ogcs.Sync.Direction.OutlookToGoogle.Id, "true");

            Ogcs.SettingsStore.Upgrade.upgradeBidirectionalReminderDnd(xml);

            Assert.AreEqual("true", GetReminderDndValue(xml));
        }

        private static XDocument CreateCalendarXml(int syncDirectionId, string reminderDnd) {
            XNamespace settingsNamespace = SettingsNamespace;
            return new XDocument(
                new XElement(settingsNamespace + "Settings",
                    new XElement(settingsNamespace + "Calendar",
                        new XElement(settingsNamespace + "SyncDirection",
                            new XElement(settingsNamespace + "Id", syncDirectionId)),
                        new XElement(settingsNamespace + "ReminderDND", reminderDnd))));
        }

        private static string GetReminderDndValue(XDocument xml) {
            XNamespace settingsNamespace = SettingsNamespace;
            return xml.Descendants(settingsNamespace + "ReminderDND").Single().Value;
        }
    }
}
