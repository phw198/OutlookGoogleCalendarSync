using System;

namespace OutlookGoogleCalendarSync {
    class UserCancelledSyncException : Exception {
        public UserCancelledSyncException() { }
        public UserCancelledSyncException(string message) : base(message) { }
        public UserCancelledSyncException(string message, Exception inner) : base(message, inner) { }
    }
}
