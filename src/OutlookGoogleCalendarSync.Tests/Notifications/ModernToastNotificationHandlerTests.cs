using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Text;
using Ogcs = OutlookGoogleCalendarSync;

namespace OutlookGoogleCalendarSync.Tests.Notifications {
    [TestClass]
    public class ModernToastNotificationHandlerTests {
        [TestMethod]
        public void BrowserUrl_RoundTripsWhenToastArgumentsEscapeBase64() {
            string browserUrl = "https://example.com/a?x=1&other=two+words";
            string encodedUrl = Uri.EscapeDataString(Convert.ToBase64String(Encoding.UTF8.GetBytes(browserUrl)));
            string toastArguments = "action=test&url=" + encodedUrl;
            Assert.IsTrue(encodedUrl.EndsWith("%3D"));

            var arguments = Ogcs.ModernToastNotificationHandler.getTriggerArguments(toastArguments);
            string decodedUrl = Ogcs.ModernToastNotificationHandler.getBrowserUrl(arguments);

            Assert.AreEqual(browserUrl, decodedUrl);
        }
    }
}
