using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Identity.Client;
using log4net;

namespace OutlookGoogleCalendarSync.OutlookOgcs {
    public class Authenticator {
        private static readonly ILog log = LogManager.GetLogger(typeof(Authenticator));

        private Boolean authenticated = false;
        public System.Threading.CancellationTokenSource CancelTokenSource;
        private static readonly String _clientId = "3f85f044-607a-4139-bb2e-e12eac105f14";
        private static String clientId {
            get {
                //if (Settings.Instance.UsingPersonalAPIkeys()) {
                //    return Settings.Instance.PersonalClientIdentifier;
                //} else {
                return _clientId;
                //}
            }
        }
        private IPublicClientApplication oAuthApp;

        public Authenticator() {
            CancelTokenSource = new System.Threading.CancellationTokenSource();
            oAuthApp = PublicClientApplicationBuilder.Create(clientId).Build();
        }

        public void GetAuthenticated() {
            if (this.authenticated) return;

            Forms.Main.Instance.Console.Update("<span class='em em-key'></span>Authenticating with Outlook", Console.Markup.h2, newLine: false, verbose: true);

            System.Threading.Thread oAuth = new System.Threading.Thread(() => { spawnOauth(); });
            oAuth.Start();
            while (oAuth.IsAlive) {
                System.Windows.Forms.Application.DoEvents();
                System.Threading.Thread.Sleep(100);
            }
        }

        private void spawnOauth() {
            try {
                //Calling an async function from a static constructor needs to be called like this, else it deadlocks:-
                Task task = Task.Run(async () => { await getAuthenticated(Authenticator.clientId); });
                try {
                    task.Wait(CancelTokenSource.Token);
                } catch (System.OperationCanceledException) {
                    Forms.Main.Instance.Console.Update("Authorisation to allow OGCS to manage your Google calendar was cancelled.", Console.Markup.warning);
                } catch (System.Exception ex) {
                    OGCSexception.Analyse(ex);
                    Forms.Main.Instance.Console.UpdateWithError("Unable to authenticate with Microsoft. The following error occurred:", ex);
                }
            } catch (System.Exception ex) {
                log.Fail("Problem encountered in getCalendarClientSecrets()");
                Forms.Main.Instance.Console.UpdateWithError("Unable to authenticate with Google!", ex);
            }
        }

        private async Task<bool> getAuthenticated(String clientId) {
            String[] scopes = new string[] { "user.read" };
            String graphAPIEndpoint = "https://graph.microsoft.com/v1.0/me";

            IAccount firstAccount = (await oAuthApp.GetAccountsAsync()).FirstOrDefault();
            if (firstAccount == null)
                log.Warn("The user has not signed-in before or there is no account information in the cache.");

            AuthenticationResult authResult = null;
            try {
                authResult = await oAuthApp.AcquireTokenSilent(scopes, firstAccount).ExecuteAsync();
            } catch (MsalUiRequiredException msalSilentEx) {
                // This indicates you need to call AcquireTokenInteractive to acquire a token
                log.Warn($"Failed acquiring MS Graph token silently: {msalSilentEx.Message}");

                try {
                    authResult = await oAuthApp.AcquireTokenInteractive(scopes)
                        .WithAccount(firstAccount)
                        .WithPrompt(Prompt.SelectAccount)
                        .ExecuteAsync();
                } catch (MsalException msalInteractiveEx) {
                    log.Fail("Problem acquiring MS Graph token interactively.");
                    if (msalInteractiveEx.Message.Trim() == "User canceled authentication.")
                        CancelTokenSource.Cancel(true);
                    throw;
                } catch (System.Exception) {
                    log.Fail("Error during AcquireTokenInteractive()");
                    throw;
                }
            } catch (System.Exception ex) {
                log.Fail("Problem encountered in getAuthenticated()");
                Forms.Main.Instance.Console.UpdateWithError("Unable to authenticate with Microsoft!", ex);
                return false;
            }

            if (authResult == null) return false;

            String resultText = GetHttpContentWithToken(graphAPIEndpoint, authResult.AccessToken);
            Forms.Main.Instance.SetControlPropertyThreadSafe(Forms.Main.Instance.tbOutlookConnectedAcc, "Text", authResult.Account.Username);

            if (!String.IsNullOrEmpty(authResult.AccessToken) && authResult.ExpiresOn != null) {
                log.Info("Refresh and Access token successfully retrieved.");
                log.Debug("Access token expires " + authResult.ExpiresOn.ToLocalTime().ToString());
            }

            authenticated = true;
            Forms.Main.Instance.Console.Update("Handshake successful.", verbose: true);
            return authenticated;
        }

        /// <summary>
        /// Perform an HTTP GET request to a URL using an HTTP Authorization header
        /// </summary>
        /// <param name="url">The URL</param>
        /// <param name="token">The token</param>
        /// <returns>String containing the results of the GET operation</returns>
        private String GetHttpContentWithToken(String url, String token) {
            Extensions.OgcsWebClient wc = new Extensions.OgcsWebClient();
            try {
                wc.Headers.Add("Authorization", "Bearer " + token);
                String content = wc.DownloadString(url);
                return content;
            } catch (System.Exception ex) {
                return ex.ToString();
            }
        }
    }
}
