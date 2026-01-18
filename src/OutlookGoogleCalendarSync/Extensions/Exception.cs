using Ogcs = OutlookGoogleCalendarSync;
using log4net;
using System;
using System.ComponentModel;
using System.Linq;

namespace OutlookGoogleCalendarSync {
    class UserCancelledSyncException : System.Exception {
        public UserCancelledSyncException() { }
        public UserCancelledSyncException(string message) : base(message) { }
        public UserCancelledSyncException(string message, System.Exception inner) : base(message, inner) { }
    }

    public static class Exception {
        private static readonly ILog log = LogManager.GetLogger(typeof(Ogcs.Exception));

        public static void Analyse(this System.Exception ex, String warnDetail, Boolean includeStackTrace = false) {
            log.Warn(warnDetail);
            Analyse(ex, includeStackTrace: includeStackTrace);
        }
        public static void Analyse(this System.Exception ex, Boolean includeStackTrace = false) {
            log4net.Core.Level logLevel = log4net.Core.Level.Error;
            if (LoggingAsFail(ex) || Outlook.Errors.LogAsFail(ex)) {
                if (ex is ApplicationException) return;
                logLevel = Program.MyFailLevel;
            }

            log.ErrorOrFail(ex.GetType().FullName + ": " + ex.Message, logLevel);
            String locationDetails = "<Unknown File>";
            try {
                System.Diagnostics.StackTrace st = new System.Diagnostics.StackTrace(ex, true);
                foreach (System.Diagnostics.StackFrame sf in st.GetFrames()) {
                    String filename = sf.GetFileName();
                    if (string.IsNullOrEmpty(filename)) continue;
                    locationDetails = $"{sf.GetMethod().Name}() at offset {sf.GetNativeOffset()} in {System.IO.Path.GetFileName(filename)}:{sf.GetFileLineNumber()}:{sf.GetFileColumnNumber()}";
                    break;
                }
            } catch (System.Exception ex2) {
                log.Error("Unable to parse exception stack. " + ex2.Message);
            }
            String errorLocation = "; Location: " + ex.TargetSite?.Name + "() in " + locationDetails;
            int errorCode = getErrorCode(ex);
            log.ErrorOrFail("Code: 0x" + errorCode.ToString("X8") + "," + errorCode.ToString() + errorLocation, logLevel);

            if (ex.InnerException != null) {
                log.ErrorOrFail("InnerException:-", logLevel);
                Analyse(ex.InnerException, false);
            }
            if (includeStackTrace) {
                try {
                    log.ErrorOrFail("Exception stack trace " + ex.StackTrace.TrimStart(), logLevel);
                    string[] envStack = Environment.StackTrace.Split(new string[] { "\r\n" }, StringSplitOptions.RemoveEmptyEntries);
                    log.ErrorOrFail("Environment stack trace " + string.Join("\r\n", envStack.Skip(3)).TrimStart(), logLevel);
                } catch (System.Exception ex2) {
                    log.Error("Unable to include stack trace. " + ex2.Message);
                }
            }
        }

        public static String GetErrorCode(this System.Exception ex, UInt32 mask = 0xFFFFFFFF) {
            UInt32 maskedCode = (uint)(getErrorCode(ex) & mask);
            return "0x" + maskedCode.ToString("X8");
        }

        private static int getErrorCode(System.Exception ex) {
            try {
                Win32Exception w32ex = ex as Win32Exception;
                if (w32ex == null) {
                    w32ex = ex.InnerException as Win32Exception;
                }
                if (w32ex != null) {
                    return w32ex.ErrorCode;
                }
            } catch {
                log.Error("Failed to obtain Win32Exception.");
            }
            try {
                return System.Runtime.InteropServices.Marshal.GetHRForException(ex);
            } catch {
                log.Error("Failed to get HResult.");
            }
            return -1;
        }

        public static void AnalyseAggregate(this AggregateException agex, Boolean throwError = true) {
            foreach (System.Exception ex in agex.InnerExceptions) {
                if (ex is ApplicationException) {
                    if (!String.IsNullOrEmpty(ex.Message)) Forms.Main.Instance.Console.UpdateWithError(null, ex);
                    else log.Error(agex.Message);

                } else if (ex is global::Google.Apis.Auth.OAuth2.Responses.TokenResponseException) {
                    AnalyseTokenResponse(ex as global::Google.Apis.Auth.OAuth2.Responses.TokenResponseException, throwError);

                } else Analyse(ex);
            }
        }

        public static void AnalyseTokenResponse(this global::Google.Apis.Auth.OAuth2.Responses.TokenResponseException ex, Boolean throwError = true) {
            String instructions = "On the Settings > Google tab, please disconnect and re-authenticate your account.";

            log.Warn("Token response error: " + ex.Message);
            if (ex.Error.Error == "access_denied") {
                Forms.Main.Instance.Console.Update("Failed to obtain Calendar access from Google - it's possible your access has been revoked.<br/>" + instructions, Console.Markup.fail, notifyBubble: true);
                LogAsFail(ref ex);

            } else if ("invalid_client;unauthorized_client".Contains(ex.Error.Error)) {
                Forms.Main.Instance.Console.Update("Invalid authentication token. Account requires reauthorising.\r\n" + instructions, Console.Markup.fail, notifyBubble: true);
                LogAsFail(ref ex);

            } else if (ex.Error.Error == "invalid_grant") {
                Forms.Main.Instance.Console.Update("Google has revoked your authentication token. Account requires reauthorising.<br/>" + instructions, Console.Markup.fail, notifyBubble: true);
                LogAsFail(ref ex);

            } else {
                log.Warn("Unknown web exception.");
                Forms.Main.Instance.Console.UpdateWithError("Unable to communicate with Google. The following error occurred:", ex, notifyBubble: true);
            }
            if (throwError) throw ex;
        }

        public static String FriendlyMessage(this System.Exception ex) {
            if (ex is global::Google.GoogleApiException) {
                global::Google.GoogleApiException gaex = ex as global::Google.GoogleApiException;
                if (gaex.Error != null)
                    return gaex.Error.Message + " [" + gaex.Error.Code + "=" + gaex.HttpStatusCode + "]";
                else
                    return gaex.Message + " [" + gaex.HttpStatusCode + "]";
            } else {
                return ex.Message + (ex.InnerException != null && !(ex.InnerException is global::Google.GoogleApiException) ? "<br/>" + ex.InnerException.Message : "");
            }
        }

        #region Logging level for exception
        //FAIL is a lower level than ERROR and so will not trigger Error Reporting

        private enum LogLevel {
            ERROR,
            FAIL
        }
        private const String LogAs = "LogAs";

        /// <summary>
        /// Capture this exception as log4net FAIL (not ERROR) when logged
        /// </summary>
        public static void LogAsFail(ref System.Exception ex) {
            if (ex.Data.Contains(LogAs))
                ex.Data[LogAs] = Ogcs.Exception.LogLevel.FAIL;
            else
                ex.Data.Add(LogAs, Ogcs.Exception.LogLevel.FAIL);
        }
        /// <summary>
        /// Capture this exception as log4net FAIL (not ERROR) when logged
        /// </summary>
        public static System.Exception LogAsFail(this System.Exception ex) {
            LogAsFail(ref ex);
            return ex;
        }

        /// <summary>
        /// Capture this exception as log4net FAIL (not ERROR) when logged
        /// </summary>
        public static void LogAsFail(ref System.ApplicationException ex) {
            if (ex.Data.Contains(LogAs))
                ex.Data[LogAs] = Ogcs.Exception.LogLevel.FAIL;
            else
                ex.Data.Add(LogAs, Ogcs.Exception.LogLevel.FAIL);
        }
        
        /// <summary>
        /// Capture this exception as log4net FAIL (not ERROR) when logged
        /// </summary>
        public static void LogAsFail(ref System.Runtime.InteropServices.COMException ex) {
            if (ex.Data.Contains(LogAs))
                ex.Data[LogAs] = Ogcs.Exception.LogLevel.FAIL;
            else
                ex.Data.Add(LogAs, Ogcs.Exception.LogLevel.FAIL);
        }
        
        /// <summary>
        /// Capture this exception as log4net FAIL (not ERROR) when logged
        /// </summary>
        public static void LogAsFail(ref System.NullReferenceException ex) {
            if (ex.Data.Contains(LogAs))
                ex.Data[LogAs] = Ogcs.Exception.LogLevel.FAIL;
            else
                ex.Data.Add(LogAs, Ogcs.Exception.LogLevel.FAIL);
        }

        /// <summary>Capture this exception as log4net FAIL (not ERROR) when logged</summary>
        public static void LogAsFail(ref global::Google.Apis.Auth.OAuth2.Responses.TokenResponseException ex) {
            if (ex.Data.Contains(LogAs))
                ex.Data[LogAs] = Ogcs.Exception.LogLevel.FAIL;
            else
                ex.Data.Add(LogAs, Ogcs.Exception.LogLevel.FAIL);
        }

        /// <summary>Capture this exception as log4net FAIL (not ERROR) when logged</summary>
        public static void LogAsFail(ref global::Google.GoogleApiException ex) {
            if (ex.Data.Contains(LogAs))
                ex.Data[LogAs] = Ogcs.Exception.LogLevel.FAIL;
            else
                ex.Data.Add(LogAs, Ogcs.Exception.LogLevel.FAIL);
        }

        /// <summary>Capture this exception as log4net FAIL (not ERROR) when logged</summary>
        public static void LogAsFail(ref System.Net.WebException ex) {
            if (ex.Data.Contains(LogAs))
                ex.Data[LogAs] = Ogcs.Exception.LogLevel.FAIL;
            else
                ex.Data.Add(LogAs, Ogcs.Exception.LogLevel.FAIL);
        }

        /// <summary>
        /// Check if this exception has been set to log as log4net FAIL (not ERROR)
        /// </summary>
        public static Boolean LoggingAsFail(this System.Exception ex) {
            return (ex.Data.Contains(LogAs) && ex.Data[LogAs].ToString() == LogLevel.FAIL.ToString());
        }
        #endregion
    }
}
