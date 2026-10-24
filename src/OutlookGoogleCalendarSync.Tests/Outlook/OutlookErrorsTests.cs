using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Runtime.Remoting.Messaging;
using System.Runtime.Remoting.Proxies;
using System.Runtime.Serialization;
using Ogcs = OutlookGoogleCalendarSync;

namespace OutlookGoogleCalendarSync.Tests.Outlook {
    [TestClass]
    public class OutlookErrorsTests {
        [DataTestMethod]
        [DataRow(unchecked((int)0x8004011D))]
        [DataRow(unchecked((int)0xC204011D))]
        [DataRow(unchecked((int)0x80040115))]
        public void HandleComError_ExchangeConnectivityError_ClassifiesAsUnavailable(int hResult) {
            COMException exception = new COMException("Exchange unavailable", hResult);

            Assert.AreEqual(Ogcs.Outlook.Errors.ErrorType.ExchangeUnavailable,
                Ogcs.Outlook.Errors.HandleComError(exception));
            Assert.IsTrue(Ogcs.Outlook.Errors.LogAsFail(exception));
        }

        [TestMethod]
        public void HandleComError_UnrelatedComError_DoesNotClassifyAsExchangeUnavailable() {
            COMException exception = new COMException("Unexpected COM error", unchecked((int)0x80004005));

            Assert.AreEqual(Ogcs.Outlook.Errors.ErrorType.Unhandled,
                Ogcs.Outlook.Errors.HandleComError(exception));
        }

        [TestMethod]
        public void IsExchangeUnavailable_WrappedComException_ReturnsTrue() {
            AggregateException exception = new AggregateException(new ApplicationException(
                "Cannot access shared calendar",
                new COMException("Exchange unavailable", unchecked((int)0x8004011D))));

            Assert.IsTrue(Ogcs.Outlook.Errors.IsExchangeUnavailable(exception));
        }

        [TestMethod]
        public void IsExchangeUnavailable_SecondAggregateInnerException_ReturnsTrue() {
            AggregateException exception = new AggregateException(
                new InvalidOperationException("Unrelated"),
                new COMException("Exchange unavailable", unchecked((int)0xC204011D)));

            Assert.IsTrue(Ogcs.Outlook.Errors.IsExchangeUnavailable(exception));
        }

        [TestMethod]
        public void IsExchangeUnavailable_WrappedUnrecognizedComException_ReturnsFalse() {
            ApplicationException exception = new ApplicationException("Cannot access shared calendar",
                new COMException("Unexpected COM error", unchecked((int)0x80004005)));

            Assert.IsFalse(Ogcs.Outlook.Errors.IsExchangeUnavailable(exception));
        }

        [TestMethod]
        public void IsExchangeUnavailable_UnrelatedAggregate_DoesNotClassifyAsUnavailable() {
            AggregateException exception = new AggregateException(
                new InvalidOperationException("Unrelated failure"),
                new COMException("Unexpected COM error", unchecked((int)0x80004005)));

            Assert.IsFalse(Ogcs.Outlook.Errors.IsExchangeUnavailable(exception));
        }

        [TestMethod]
        public void HandleComError_ExchangeNetworkMessage_ClassifiesAsUnavailable() {
            COMException exception = new COMException(
                "Network problems are preventing connection to Microsoft Exchange.", unchecked((int)0x80004005));

            Assert.AreEqual(Ogcs.Outlook.Errors.ErrorType.ExchangeUnavailable,
                Ogcs.Outlook.Errors.HandleComError(exception));
        }

        [TestMethod]
        public void IsExchangeUnavailable_ResourceLimit_ReturnsFalse() {
            Assert.IsFalse(Ogcs.Outlook.Errors.IsExchangeUnavailable(
                new COMException("Too many open items", unchecked((int)0x80040305))));
        }

        [TestMethod]
        public void IsExchangeUnavailable_NonComExceptionWithSameHResult_ReturnsFalse() {
            Assert.IsFalse(Ogcs.Outlook.Errors.IsExchangeUnavailable(
                new ExternalException("Exchange unavailable", unchecked((int)0x8004011D))));
        }

        [TestMethod]
        public void CalendarInstance_ExchangeConnectivityComFailure_SetsUnavailableState() {
            FieldInfo instanceField = typeof(Ogcs.Outlook.Calendar).GetField("instance",
                BindingFlags.Static | BindingFlags.NonPublic);
            Object previousInstance = instanceField.GetValue(null);
            Boolean previousUnavailable = Ogcs.Outlook.Calendar.ExchangeUnavailable;
            Ogcs.Outlook.Calendar fakeCalendar = (Ogcs.Outlook.Calendar)FormatterServices.GetUninitializedObject(
                typeof(Ogcs.Outlook.Calendar));
            fakeCalendar.IOutlook = (Ogcs.Outlook.Interface)new ExchangeUnavailableInterfaceProxy().GetTransparentProxy();
            instanceField.SetValue(null, fakeCalendar);
            Ogcs.Outlook.Calendar.ExchangeUnavailable = false;

            try {
                COMException exception = Assert.ThrowsException<COMException>(() => {
                    Ogcs.Outlook.Calendar unused = Ogcs.Outlook.Calendar.Instance;
                });

                Assert.AreEqual(unchecked((int)0x8004011D), exception.ErrorCode);
                Assert.IsTrue(Ogcs.Outlook.Calendar.ExchangeUnavailable);
                Assert.IsFalse(Ogcs.Outlook.Calendar.Connecting);
            } finally {
                instanceField.SetValue(null, previousInstance);
                Ogcs.Outlook.Calendar.ExchangeUnavailable = previousUnavailable;
            }
        }

        private sealed class ExchangeUnavailableInterfaceProxy : RealProxy {
            public ExchangeUnavailableInterfaceProxy() : base(typeof(Ogcs.Outlook.Interface)) { }

            public override IMessage Invoke(IMessage message) {
                IMethodCallMessage call = (IMethodCallMessage)message;
                return new ReturnMessage(new COMException("Simulated Exchange unavailable", unchecked((int)0x8004011D)), call);
            }
        }
    }
}
