using log4net;
using Microsoft.Graph;
using System;
using System.Linq;
using MsGraph = Microsoft.Graph.Models;
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
        public static MsGraph.ODataErrors.ODataError GetODataError(System.Exception ex) {
            MsGraph.ODataErrors.ODataError oDataErr = null;
            if (ex is AggregateException aex) {
                ServiceException sex = null;
                sex = aex.InnerExceptions.FirstOrDefault(ie => ie is ServiceException) as ServiceException;
                if (sex != null) {
                    sex.Analyse("Why have we received a ServiceException, not ODataError?");
                    return null;
                }
                oDataErr = aex.InnerExceptions.FirstOrDefault(ie => ie is MsGraph.ODataErrors.ODataError oDataErr) as MsGraph.ODataErrors.ODataError;
                if (oDataErr == null) {
                    log.Warn("No OData inner exception object found");
                    aex.Analyse();
                }
            } else if (ex is MsGraph.ODataErrors.ODataError) {
                oDataErr = ex as MsGraph.ODataErrors.ODataError;
            } else {
                ex.Analyse();
            }
            return oDataErr;
        }

        public static ApiException HandleAPIlimits(ref System.Exception ex) {
            MsGraph.ODataErrors.ODataError oDataErr = GetODataError(ex);
            if (oDataErr == null) return ApiException.noODataResponse;

            //Analyse the Graph OData exception and then replace the aggregate exception with it
            ApiException retVal = HandleAPIlimits(ref oDataErr);
            ex = oDataErr;
            return retVal;
        }
    
        public static ApiException HandleAPIlimits(ref MsGraph.ODataErrors.ODataError ex/*, Event ev*/) {
            if (ex == null) return ApiException.noODataResponse;

            log.Fail(ex.FriendlyMessage());

            try {
                new Telemetry.GA4Event.Event(Telemetry.GA4Event.Event.Name.ogcs_error)
                    .AddParameter("api_graph_error", ex.Message)
                    .AddParameter("reason", ex.Error.Code)
                    .AddParameter("code", ex.ResponseStatusCode)
                    .AddParameter("message", ex.Error.Message)
                    .Send();
            } catch (System.Exception gaEx) {
                Ogcs.Exception.Analyse(gaEx);
            }

            if (ex.ResponseStatusCode == (int)System.Net.HttpStatusCode.Forbidden) {
                if (ex.Message.Contains("Check credentials and try again")) {
                    Forms.Main.Instance.Console.Update("You are not properly authenticated to Microsoft.<br/>" +
                        "On the Settings > Outlook tab, please disconnect and re-authenticate your account.", Console.Markup.error);
                    ex.Data.Add("OGCS", "Unauthenticated access to Microsoft account attempted. Authentication required.");
                }
                return ApiException.throwException;
            }

            log.Warn("Unhandled API exception.");
            return ApiException.throwException;
        }
    }
}
