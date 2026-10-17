using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using Ogcs = OutlookGoogleCalendarSync;

namespace OutlookGoogleCalendarSync.Tests.Extensions {
    [TestClass]
    public class ExceptionTests {
        [TestMethod]
        public void LogAsFail_AggregateException_MarksAggregateAndAllInnerExceptions() {
            System.Exception inner = new InvalidOperationException("inner");
            AggregateException aggregate = new AggregateException(inner);

            System.Exception result = Ogcs.Exception.LogAsFail(aggregate);

            Assert.AreSame(aggregate, result);
            Assert.IsTrue(Ogcs.Exception.LoggingAsFail(aggregate));
            Assert.IsTrue(Ogcs.Exception.LoggingAsFail(inner));
        }

        [TestMethod]
        public void LogAsFail_NestedAggregateException_MarksEveryLevel() {
            System.Exception leaf = new InvalidOperationException("leaf");
            AggregateException nested = new AggregateException(leaf);
            AggregateException aggregate = new AggregateException(nested);

            Ogcs.Exception.LogAsFail(aggregate);

            Assert.IsTrue(Ogcs.Exception.LoggingAsFail(aggregate));
            Assert.IsTrue(Ogcs.Exception.LoggingAsFail(nested));
            Assert.IsTrue(Ogcs.Exception.LoggingAsFail(leaf));
        }

        [TestMethod]
        public void LogAsFail_OrdinaryException_MarksException() {
            System.Exception exception = new InvalidOperationException("exception");

            System.Exception result = Ogcs.Exception.LogAsFail(exception);

            Assert.AreSame(exception, result);
            Assert.IsTrue(Ogcs.Exception.LoggingAsFail(exception));
        }
    }
}
