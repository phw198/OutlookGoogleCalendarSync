using log4net;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Text;

namespace OutlookGoogleCalendarSync {
    class UserCancelledSyncException : Exception {
        public UserCancelledSyncException() { }
        public UserCancelledSyncException(string message) : base(message) { }
        public UserCancelledSyncException(string message, Exception inner) : base(message, inner) { }
    }

    class OGCSexception {
        private static readonly ILog log = LogManager.GetLogger(typeof(OGCSexception));

        public static void Analyse(System.Exception ex, Boolean includeStackTrace = false) {
            log.Error(ex.GetType().FullName +": "+ ex.Message);
            int errorCode = getErrorCode(ex);
            log.Error("Code: 0x" + errorCode.ToString("X8") +";"+ errorCode.ToString());
            if (includeStackTrace) log.Error(ex.StackTrace);
        }

        public static String GetErrorCode(System.Exception ex, UInt32 mask = 0xFFFFFFFF) {
            UInt32 maskedCode = (uint)(getErrorCode(ex) & mask);
            return "0x" + maskedCode.ToString("X8");
        }

        private static int getErrorCode(System.Exception ex) {
            try {
                var w32ex = ex as Win32Exception;
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

        #region DotNet Exceptions
        public static void AnalyseDotNetOpenAuth(DotNetOpenAuth.Messaging.ProtocolException ex) {
            Dictionary<String, String> errors = null;
            String webExceptionStr_orig = "";
            System.Net.WebException webException = null;
            //Process exact error
            try {
                webException = ex.InnerException as System.Net.WebException;
                webExceptionStr_orig = extractResponseString(webException);
            } catch (System.Exception subEx) {
                log.Error("Failed to retrieve WebException: " + subEx.Message);
                log.Debug(ex.Message);
                throw ex;
            }
            if (string.IsNullOrEmpty(webExceptionStr_orig)) {
                //Not an OAuthErrorMsg
                log.Error(webException.Message);
                throw ex;
            }
            try {
                /* Could treat this properly with JSON but would be another dll just to handle this situation.
                  * OAuthErrorMsg error =
                  * JsonConvert.DeserializeObject<OAuthErrorMsg>(ExtractResponseString(webException));
                  * var errorMessage = error.error_description; 
                  */
                //String webExceptionStr = "{\n  \"error\" : \"invalid_client\",\n  \"error_description\" : \"The OAuth client was not found.\"\n}";
                String webExceptionStr = webExceptionStr_orig.Replace("\"", "");
                webExceptionStr = webExceptionStr.TrimStart('{'); webExceptionStr = webExceptionStr.TrimEnd('}');
                errors = webExceptionStr.Split(new String[] { "\n" }, StringSplitOptions.RemoveEmptyEntries)
                    .Select(s => s.Split(':')).ToDictionary(x => x[0].Trim(), x => x[1].Trim().TrimEnd(','));

            } catch (System.Exception subEx) {
                log.Error("Failed to process exact WebException: " + subEx.Message);
                log.Debug(webExceptionStr_orig);
                throw ex;
            }

            if (errors.ContainsKey("error")) {
                String instructions = "On the Settings > Google tab, please disconnect and re-authenticate your account.";
                if ("invalid_client;unauthorized_client".Contains(errors["error"]))
                    throw new System.Exception("Invalid authentication token. Account requires reauthorising.\r\n" + instructions, ex);
                else if (errors["error"] == "invalid_grant")
                    throw new System.Exception("Google has revoked your authentication token. Account requires reauthorising.\r\n" + instructions, ex);
            }
            log.Debug("Unknown web exception.");
            throw ex;
        }

        private static String extractResponseString(System.Net.WebException webException) {
            if (webException == null || webException.Response == null)
                return null;

            var responseStream =
                webException.Response.GetResponseStream() as MemoryStream;

            if (responseStream == null)
                return null;

            var responseBytes = responseStream.ToArray();

            var responseString = Encoding.UTF8.GetString(responseBytes);
            return responseString;
        }
        #endregion
    }
}
