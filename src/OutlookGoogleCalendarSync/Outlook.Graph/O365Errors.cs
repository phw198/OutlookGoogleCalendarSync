using log4net;
using Microsoft.Graph;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Http;
using System.Threading;
using System.Threading.Tasks;
using MsGraph = OutlookGoogleCalendarSync.Outlook.Graph.CustomClient;
using Ogcs = OutlookGoogleCalendarSync;

namespace OutlookGoogleCalendarSync.Outlook.Graph {
    public enum ApiException {
        justContinue,
        backoffThenRetry,
        throwException,
        noODataResponse
    }

    public class O365Errors {
        private static readonly ILog log = LogManager.GetLogger(typeof(O365Errors));

        /// <summary>
        /// Locate and return the OData exception, if present.
        /// </summary>
        /// <param name="ex">The general exception object</param>
        /// <returns>The OData exception object, if available</returns>
        public static MsGraph.Models.ODataErrors.ODataError GetODataError(System.Exception ex) {
            return GetODataError(ex, out _);
        }

        /// <summary>
        /// Locate and return the OData exception, if present.
        /// </summary>
        /// <param name="ex">The general exception object</param>
        /// <param name="kiotaEx">The general exception cast to Kiota exception</param>
        /// <returns>The OData exception object, if available</returns>
        public static MsGraph.Models.ODataErrors.ODataError GetODataError(System.Exception ex, out Microsoft.Kiota.Abstractions.ApiException kiotaEx) {
            kiotaEx = null;
            MsGraph.Models.ODataErrors.ODataError oDataErr = null;
            string errorKey = new MsGraph.Models.ODataErrors.ODataError().GetType().FullName;

            if (ex is AggregateException aex) {
                ServiceException sex = null;
                sex = aex.InnerExceptions.FirstOrDefault(ie => ie is ServiceException) as ServiceException;
                if (sex != null) {
                    sex.Analyse("Why have we received a ServiceException, not Kiota ApiException/ODataError?");
                    return null;
                }
                kiotaEx = aex.InnerExceptions.FirstOrDefault(ie => ie is Microsoft.Kiota.Abstractions.ApiException) as Microsoft.Kiota.Abstractions.ApiException;
                return GetODataError(kiotaEx);

            } else if (ex is Microsoft.Kiota.Abstractions.ApiException) {
                kiotaEx = ex as Microsoft.Kiota.Abstractions.ApiException;
                if (kiotaEx.Data.Contains(errorKey) && kiotaEx.Data[errorKey] is MsGraph.Models.ODataErrors.ODataError)
                    oDataErr = kiotaEx.Data[errorKey] as MsGraph.Models.ODataErrors.ODataError;
                if (oDataErr == null) {
                    if (kiotaEx.ResponseHeaders.TryGetValue("request-id", out var values)) {
                        string requestId = values.FirstOrDefault();
                        GraphErrorInterceptor.GraphErrorStorage.TryGetValue(requestId ?? "", out Dictionary<string, object> graphError);
                        string rawJsonError = graphError?["json-error"]?.ToString();
                        if (!string.IsNullOrEmpty(rawJsonError)) {
                            Newtonsoft.Json.Linq.JObject jsonDoc = Newtonsoft.Json.Linq.JObject.Parse(rawJsonError);
                            Newtonsoft.Json.Linq.JToken errorToken = jsonDoc["error"];
                            if (errorToken != null) {
                                oDataErr = new() {
                                    Error = new() {
                                        Code = errorToken["code"]?.ToString() ?? "Unknown",
                                        Message = errorToken["message"]?.ToString() ?? rawJsonError
                                    }
                                };
                            }
                        }
                    }
                }
                if (oDataErr == null) {
                    ex.Analyse("Unable to determine OData exception details :-\\");
                }

            } else {
                ex.Analyse();
            }
            return oDataErr;
        }

        public static ApiException HandleAPIlimits(ref System.Exception ex) {
            Microsoft.Kiota.Abstractions.ApiException kiotaEx = null;
            MsGraph.Models.ODataErrors.ODataError oDataErr = GetODataError(ex, out kiotaEx);
            if (oDataErr == null) return ApiException.noODataResponse;

            //Analyse the Graph OData exception and then replace the aggregate exception with it
            ApiException retVal = handleAPIlimits(ref oDataErr, kiotaEx);
            ex = kiotaEx;
            return retVal;
        }
    
        private static ApiException handleAPIlimits(ref MsGraph.Models.ODataErrors.ODataError ex, Microsoft.Kiota.Abstractions.ApiException kiotaEx) {
            if (ex == null) return ApiException.noODataResponse;

            log.Fail(kiotaEx.FriendlyMessage());

            try {
                new Telemetry.GA4Event.Event(Telemetry.GA4Event.Event.Name.ogcs_error)
                    .AddParameter("api_graph_error", ex.Message)
                    .AddParameter("reason", ex.Error.Code)
                    .AddParameter("code", kiotaEx.ResponseStatusCode)
                    .AddParameter("message", ex.Error.Message)
                    .Send();
            } catch (System.Exception gaEx) {
                Ogcs.Exception.Analyse(gaEx);
            }

            if (kiotaEx.ResponseStatusCode == (int)System.Net.HttpStatusCode.Forbidden) {
                if (kiotaEx.Message.Contains("Check credentials and try again")) {
                    Forms.Main.Instance.Console.Update("You are not properly authenticated to Microsoft.<br/>" +
                        "On the Settings > Outlook tab, please disconnect and re-authenticate your account.", Console.Markup.error);
                    kiotaEx.Data.Add("OGCS", "Unauthenticated access to Microsoft account attempted. Authentication required.");
                }
                return ApiException.throwException;
            }

            log.Warn("Unhandled API exception.");
            return ApiException.throwException;
        }
    }


    /// <summary>
    /// This is only needed due to Kiota v1.17 and drawback #2 noted in .csproj file
    /// It's purpose is to intercept low-level network traffic and manually determine error code, message, etc
    /// Without it, Kiota cannot determine inherited properties and gives up, leaving just the HTTP responseCode available
    /// </summary>
    public class GraphErrorInterceptor : DelegatingHandler {
        //Example of response without interceptor:
        //{"The server returned an unexpected status code and the error registered for this code failed to deserialize: 404"}

        private static readonly ILog log = LogManager.GetLogger(typeof(GraphErrorInterceptor));

        // A thread-safe dictionary mapping request IDs to their last raw error body string
        public static readonly System.Collections.Concurrent.ConcurrentDictionary<string, System.Collections.Generic.Dictionary<string, object>> GraphErrorStorage = new();

        protected override async Task<HttpResponseMessage> SendAsync(HttpRequestMessage request, CancellationToken cancellationToken) {
            var response = await base.SendAsync(request, cancellationToken);
            
            // If Microsoft Graph returned an error code, grab the raw JSON payload
            if (!response.IsSuccessStatusCode && response.Content != null) {
                try {
                    // Read the raw JSON error string (e.g., {"error": {"code": "ResourceNotFound", "message": "..."}})
                    string jsonError = await response.Content.ReadAsStringAsync();
                    if (response.Headers.TryGetValues("request-id", out var values)) {
                        GraphErrorStorage[values.FirstOrDefault()] = new System.Collections.Generic.Dictionary<string, object>() {
                            { "json-error", jsonError },
                            { "timestamp", DateTimeOffset.UtcNow }
                        };
                        log.Fine($"Added new error key {values.FirstOrDefault()}");
                    }
                } catch (System.Exception ex) {
                    ex.Analyse("Unable to parse response stream for detailed error.");
                }

                try {
                    if (GraphErrorStorage.Count > 0) {
                        log.Fine($"GraphErrorStorage holds {GraphErrorStorage.Count} items");
                        List<string> expiredKeys = new(GraphErrorStorage
                            .Where(err => (DateTimeOffset)err.Value["timestamp"] < DateTimeOffset.UtcNow.AddMinutes(-1))
                            .Select(err => err.Key)
                            .ToList()
                        );
                        foreach (string key in expiredKeys) {
                            log.Fine($"Removing expired key {key}");
                            GraphErrorStorage.TryRemove(key, out _);
                        }
                    }
                } catch (System.Exception ex) {
                    ex.Analyse("Unable to purge GraphErrorStorage.");
                }
            }

            return response;
        }
    }
}
