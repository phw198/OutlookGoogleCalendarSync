using Microsoft.VisualStudio.TestTools.UnitTesting;
using OutlookGoogleCalendarSync.Extensions;
using OutlookGoogleCalendarSync.GraphExtension;
using MsGraph = OutlookGoogleCalendarSync.Outlook.Graph.CustomClient;

namespace OutlookGoogleCalendarSync.Tests.Extensions {
    [TestClass]
    public class MsGraphBodyTests {
        [TestMethod]
        public void RemoveHtmlTags_NullOrMultilineTags_ReturnsTextWithoutTags() {
            Assert.AreEqual("", ((string)null).RemoveHtmlTags());
            Assert.AreEqual("first\nsecond", "first\n<em\n>second</em>".RemoveHtmlTags());
        }

        [TestMethod]
        public void BodyContentForSync_ExchangeConvertedText_ReturnsCompleteDecodedText() {
            string longText = new string('x', 900);
            MsGraph.Models.ItemBody body = new MsGraph.Models.ItemBody {
                ContentType = MsGraph.Models.BodyType.Html,
                Content = "<html><head><!-- converted from text --></head><body><font><div class=\"PlainText\">&nbsp;<br>Some description &amp; more<br>" + longText + "</div></font></body></html>"
            };

            string result = body.BodyContentForSync();

            Assert.AreEqual(" \r\nSome description & more\r\n" + longText, result);
        }

        [TestMethod]
        public void BodyContentForSync_OrdinaryHtml_PreservesBodyMarkup() {
            MsGraph.Models.ItemBody body = new MsGraph.Models.ItemBody {
                ContentType = MsGraph.Models.BodyType.Html,
                Content = "<html><body><p>Formatted <b>description</b></p></body></html>"
            };

            Assert.AreEqual("<p>Formatted <b>description</b></p>", body.BodyContentForSync());
        }

        [TestMethod]
        public void BodyContentForSync_RawText_ReturnsText() {
            MsGraph.Models.ItemBody body = new MsGraph.Models.ItemBody {
                ContentType = MsGraph.Models.BodyType.Text,
                Content = "Some description in text?"
            };

            Assert.AreEqual("Some description in text?", body.BodyContentForSync());
        }

        [TestMethod]
        public void BodyContentForSync_NullOrEmptyBody_ReturnsEmptyString() {
            MsGraph.Models.ItemBody emptyBody = new MsGraph.Models.ItemBody { Content = "" };

            Assert.AreEqual("", ((MsGraph.Models.ItemBody)null).BodyContentForSync());
            Assert.AreEqual("", emptyBody.BodyContentForSync());
        }
    }
}
