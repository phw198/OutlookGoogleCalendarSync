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

        [TestMethod]
        public void IsTaskCancellation_DirectTaskCanceledException_ReturnsTrue() {
            System.Exception exception = new System.Threading.Tasks.TaskCanceledException();

            Assert.IsTrue(Ogcs.Exception.IsTaskCancellation(exception));
        }

        [TestMethod]
        public void IsTaskCancellation_HttpRequestExceptionWithTaskCanceledInnerException_ReturnsTrue() {
            System.Exception exception = new System.Net.Http.HttpRequestException(
                "request cancelled",
                new System.Threading.Tasks.TaskCanceledException());

            Assert.IsTrue(Ogcs.Exception.IsTaskCancellation(exception));
        }

        [TestMethod]
        public void IsTaskCancellation_AggregateExceptionWithNestedTaskCanceledException_ReturnsTrue() {
            System.Exception exception = new AggregateException(
                new System.Net.Http.HttpRequestException(
                    "request cancelled",
                    new System.Threading.Tasks.TaskCanceledException()));

            Assert.IsTrue(Ogcs.Exception.IsTaskCancellation(exception));
        }

        [TestMethod]
        public void IsTaskCancellation_HttpRequestExceptionWithoutTaskCanceledInnerException_ReturnsFalse() {
            System.Exception exception = new System.Net.Http.HttpRequestException(
                "request failed",
                new InvalidOperationException("failure"));

            Assert.IsFalse(Ogcs.Exception.IsTaskCancellation(exception));
        }

        [TestMethod]
        public void IsUserCancellation_DirectOperationCanceledException_WhenCancellationRequested_ReturnsTrue() {
            System.Exception exception = new OperationCanceledException();

            Assert.IsTrue(Ogcs.Exception.IsUserCancellation(exception, cancellationRequested: true));
        }

        [TestMethod]
        public void IsUserCancellation_DirectOperationCanceledExceptionWithoutRequest_ReturnsFalse() {
            System.Exception exception = new OperationCanceledException();

            Assert.IsFalse(Ogcs.Exception.IsUserCancellation(exception, cancellationRequested: false));
        }

        [TestMethod]
        public void IsUserCancellation_AggregateWithOnlyNestedCancellations_ReturnsTrue() {
            System.Exception exception = new AggregateException(
                new OperationCanceledException(),
                new AggregateException(new System.Threading.Tasks.TaskCanceledException()));

            Assert.IsTrue(Ogcs.Exception.IsUserCancellation(exception, cancellationRequested: true));
        }

        [TestMethod]
        public void IsUserCancellation_MixedAggregate_ReturnsFalse() {
            System.Exception exception = new AggregateException(
                new System.Threading.Tasks.TaskCanceledException(),
                new InvalidOperationException("non-cancellation failure"));

            Assert.IsFalse(Ogcs.Exception.IsUserCancellation(exception, cancellationRequested: true));
        }

        [TestMethod]
        public void IsUserCancellation_OperationCanceledExceptionWithFailureInner_ReturnsFalse() {
            System.Exception exception = new OperationCanceledException(
                "operation aborted",
                new InvalidOperationException("underlying failure"));

            Assert.IsFalse(Ogcs.Exception.IsUserCancellation(exception, cancellationRequested: true));
        }

        [TestMethod]
        public void IsUserCancellation_WrappedCancellation_WhenCancellationRequested_ReturnsTrue() {
            System.Exception exception = new System.Net.Http.HttpRequestException(
                "request cancelled",
                new System.Threading.Tasks.TaskCanceledException());

            Assert.IsTrue(Ogcs.Exception.IsUserCancellation(exception, cancellationRequested: true));
        }

        [TestMethod]
        public void IsUserCancellation_UserCancelledSyncException_ReturnsTrueWithoutTokenRequest() {
            System.Exception exception = new global::OutlookGoogleCalendarSync.UserCancelledSyncException("user declined to continue");

            Assert.IsTrue(Ogcs.Exception.IsUserCancellation(exception, cancellationRequested: false));
        }

        [TestMethod]
        public void IsTransientNetworkError_TaskCanceledExceptionWithoutCancellationRequest_ReturnsTrue() {
            System.Threading.CancellationTokenSource originalCancelToken = Ogcs.Sync.Engine.CancelToken;
            try {
                Ogcs.Sync.Engine.CancelToken = new System.Threading.CancellationTokenSource();

                Assert.IsTrue(Ogcs.Exception.IsTransientNetworkError(new System.Threading.Tasks.TaskCanceledException()));
            } finally {
                Ogcs.Sync.Engine.CancelToken = originalCancelToken;
            }
        }

        [TestMethod]
        public void IsTransientNetworkError_TaskCanceledExceptionWithCancellationRequest_ReturnsFalse() {
            System.Threading.CancellationTokenSource originalCancelToken = Ogcs.Sync.Engine.CancelToken;
            try {
                Ogcs.Sync.Engine.CancelToken = new System.Threading.CancellationTokenSource();
                Ogcs.Sync.Engine.CancelToken.Cancel();

                Assert.IsFalse(Ogcs.Exception.IsTransientNetworkError(new System.Threading.Tasks.TaskCanceledException()));
            } finally {
                Ogcs.Sync.Engine.CancelToken = originalCancelToken;
            }
        }
    }
}
