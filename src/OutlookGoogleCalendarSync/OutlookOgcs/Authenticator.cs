using log4net;
using Microsoft.Identity.Client;
using Microsoft.Identity.Client.Extensions.Msal;
using System;
using System.Linq;
using System.Threading.Tasks;

namespace OutlookGoogleCalendarSync.OutlookOgcs {
    public class Authenticator {
        private static readonly ILog log = LogManager.GetLogger(typeof(Authenticator));

        public System.Threading.CancellationTokenSource CancelTokenSource;
        public const String TokenFile = "Microsoft.Identity.Client.Extensions.Msal.TokenResponse-user";
        private String tokenFullPath;
        private Boolean tokenFileExists { get { return System.IO.File.Exists(tokenFullPath); } }
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

        private Boolean authenticated = false;
        private IPublicClientApplication oAuthApp;
        private AuthenticationResult authResult = null;

        public Authenticator() {
            CancelTokenSource = new System.Threading.CancellationTokenSource();
        }

        public void GetAuthenticated() {
            if (this.authenticated && authResult != null) return;

            Forms.Main.Instance.Console.Update("<span class='em em-key'></span>Authenticating with Microsoft", Console.Markup.h2, newLine: false, verbose: true);

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
                log.Fail("Problem encountered in spawnOauth()"); //getCalendarClientSecrets ***
                Forms.Main.Instance.Console.UpdateWithError("Unable to authenticate with Microsoft!", ex);
            }
        }

        private async Task<bool> getAuthenticated(String clientId) {
            log.Debug("Authenticating with Microsoft Graph service...");

            tokenFullPath = System.IO.Path.Combine(Program.UserFilePath, TokenFile);
            log.Debug("Microsoft credential file location: " + tokenFullPath);
            if (!tokenFileExists)
                log.Info("No Microsoft credentials file available - need user authorisation for OGCS to manage their calendar.");

            StorageCreationProperties storageProperties = new StorageCreationPropertiesBuilder(TokenFile, Program.UserFilePath).Build();

            oAuthApp = PublicClientApplicationBuilder.Create(clientId)
                .WithAuthority("https://login.microsoftonline.com/common")
                .WithRedirectUri("https://login.microsoftonline.com/common/oauth2/nativeclient")
                .Build();

            MsalCacheHelper cacheHelper = MsalCacheHelper.CreateAsync(storageProperties).Result;
            cacheHelper.RegisterCache(oAuthApp.UserTokenCache);
            
            String[] scopes = new string[] { "user.read" };

            IAccount firstAccount = (await oAuthApp.GetAccountsAsync()).FirstOrDefault();
            if (firstAccount == null)
                log.Warn("The user has not signed-in before or there is no account information in the cache.");

            try {
                authResult = await oAuthApp.AcquireTokenSilent(scopes, firstAccount).ExecuteAsync();
            } catch (MsalUiRequiredException msalSilentEx) {
                // This indicates the need to call AcquireTokenInteractive to acquire a token
                log.Warn($"Failed acquiring MS Graph token silently: {msalSilentEx.Message}");

                try {
                    authResult = await oAuthApp.AcquireTokenInteractive(scopes)
                        .WithAccount(firstAccount)
                        .WithPrompt(Prompt.SelectAccount)
                        .ExecuteAsync();

                    if (tokenFileExists)
                        log.Info("User has provided Graph authorisation and credential file saved.");

                    if (authResult != null)
                        Forms.Main.Instance.Console.Update("Handshake successful.", verbose: true);

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

            if (!String.IsNullOrEmpty(authResult.AccessToken) && authResult.ExpiresOn != null) {
                log.Info("Refresh and Access token successfully retrieved.");
                log.Debug("Access token expires " + authResult.ExpiresOn.ToLocalTime().ToString());
            }

            Forms.Main.Instance.SetControlPropertyThreadSafe(Forms.Main.Instance.tbOutlookConnectedAcc, "Text", authResult.Account.Username);
            getMSaccountEmail();
            authenticated = true;
            return authenticated;
        }

        private void getMSaccountEmail() {
            String resultText = GetHttpContentWithToken("https://graph.microsoft.com/v1.0/me");
        }
        
        /// <summary>
        /// Perform an HTTP GET request against a Graph URL using an HTTP Authorization bearer token header
        /// </summary>
        /// <param name="url">The Graph URL</param>
        /// <param name="token">The bearer token</param>
        /// <returns>String containing the results of the GET operation</returns>
        private String GetHttpContentWithToken(String url) {
            Extensions.OgcsWebClient wc = new Extensions.OgcsWebClient();
            try {
                wc.Headers.Add("Authorization", "Bearer " + authResult?.AccessToken);
                String content = wc.DownloadString(url);
                return content;
            } catch (System.Exception ex) {
                return ex.ToString();
            }
        }
    }
}
