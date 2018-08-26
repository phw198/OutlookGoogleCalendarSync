using Google.Apis.Auth.OAuth2;
using Google.Apis.Calendar.v3;
using Google.Apis.Calendar.v3.Data;
using Google.Apis.Util.Store;
using log4net;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace OutlookGoogleCalendarSync.GoogleOgcs {
    public class Authenticator {
        private static readonly ILog log = LogManager.GetLogger(typeof(Authenticator));

        private Boolean authenticated = false;
        public Boolean Authenticated { get { return authenticated; } }

        public const String TokenFile = "Google.Apis.Auth.OAuth2.Responses.TokenResponse-user";
        private String tokenFullPath;
        private Boolean tokenFileExists { get { return File.Exists(tokenFullPath); } }

        public System.Threading.CancellationTokenSource CancelTokenSource;

        private Boolean checkedOgcsUserStatus = false;
        private static String hashedGmailAccount = null;
        public static String HashedGmailAccount {
            get {
                if (string.IsNullOrEmpty(hashedGmailAccount)) {
                    if (!string.IsNullOrEmpty(Settings.Instance.GaccountEmail))
                        hashedGmailAccount = GetMd5(Settings.Instance.GaccountEmail, true);
                }
                return hashedGmailAccount;
            }
        }

        public Authenticator() {
            CancelTokenSource = new System.Threading.CancellationTokenSource();
        }

        public void GetAuthenticated() {
            if (this.authenticated) return;

            Forms.Main.Instance.Console.Update("<span class='em em-key'></span>Authenticating with Google", Console.Markup.h2, newLine: false, verbose: true);

            System.Threading.Thread oAuth = new System.Threading.Thread(() => { spawnOauth(); });
            oAuth.Start();
            while (oAuth.IsAlive) {
                System.Windows.Forms.Application.DoEvents();
                System.Threading.Thread.Sleep(100);
            }
        }

        private void spawnOauth() {
            ClientSecrets cs = getCalendarClientSecrets();
            //Calling an async function from a static constructor needs to be called like this, else it deadlocks:-
            var task = System.Threading.Tasks.Task.Run(async () => { await getAuthenticated(cs); });
            try {
                task.Wait(CancelTokenSource.Token);
            } catch (System.OperationCanceledException) {
                Forms.Main.Instance.Console.Update("Authorisation to allow OGCS to manage your Google calendar was cancelled.", Console.Markup.warning);
            } catch (System.Exception ex) {
                OGCSexception.Analyse(ex);
            }
        }

        private static ClientSecrets getCalendarClientSecrets() {
            ClientSecrets provider = new ClientSecrets();
            if (Settings.Instance.UsingPersonalAPIkeys()) {
                provider.ClientId = Settings.Instance.PersonalClientIdentifier;
                provider.ClientSecret = Settings.Instance.PersonalClientSecret;
            } else {
                ApiKeyring apiKeyring = new ApiKeyring();

                if (Settings.Instance.Subscribed != null && Settings.Instance.Subscribed != DateTime.Parse("01-Jan-2000")) {
                    if (apiKeyring.PickKey(ApiKeyring.KeyType.Subscriber) && apiKeyring.Key != null) {
                        provider.ClientId = apiKeyring.Key.ClientId;
                        provider.ClientSecret = apiKeyring.Key.ClientSecret;
                    } else {
                        provider.ClientId = "550071650559-44lnvhdu5liq5kftj5t8k0aasgei5g7t.apps.googleusercontent.com";
                        provider.ClientSecret = "MGUFapefXClJa2ysS4WNGS4k";
                    }
                } else {
                    if (apiKeyring.PickKey(ApiKeyring.KeyType.Standard) && apiKeyring.Key != null) {
                        provider.ClientId = apiKeyring.Key.ClientId;
                        provider.ClientSecret = apiKeyring.Key.ClientSecret;
                    } else {
                        provider.ClientId = "653617509806-2nq341ol8ejgqhh2ku4j45m7q2bgdimv.apps.googleusercontent.com";
                        provider.ClientSecret = "tAi-gZLWtasS58i8CcCwVwsq";
                    }
                }
            }
            return provider;
        }

        private async Task getAuthenticated(ClientSecrets cs) {
            log.Debug("Authenticating with Google calendar service...");

            FileDataStore tokenStore = new FileDataStore(Program.UserFilePath);
            tokenFullPath = Path.Combine(tokenStore.FolderPath, TokenFile);

            log.Debug("Google credential file location: " + tokenFullPath);
            if (!tokenFileExists)
                log.Info("No Google credentials file available - need user authorisation for OGCS to manage their calendar.");
            
            string[] scopes = new[] { "https://www.googleapis.com/auth/calendar", "email" };

            UserCredential credential = null;
            try {
                //This will open the authorisation process in a browser, if required
                credential = await GoogleWebAuthorizationBroker.AuthorizeAsync(cs, scopes, "user", CancelTokenSource.Token, tokenStore);
                if (tokenFileExists)
                    log.Debug("User has provided authorisation and credential file saved.");

            } catch (Google.Apis.Auth.OAuth2.Responses.TokenResponseException ex) {
                //OGCSexception.AnalyseTokenResponse(ex);
                if (ex.Error.Error == "access_denied") {
                    String noAuthGiven = "Sorry, but this application will not work if you don't allow it access to your Google Calendar :(";
                    log.Warn("User did not provide authorisation code. Sync will not be able to work.");
                    MessageBox.Show(noAuthGiven, "Authorisation not given", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    throw new ApplicationException(noAuthGiven);
                } else {
                    Forms.Main.Instance.Console.UpdateWithError("Unable to authenticate with Google. The following error occurred:", ex);
                }

            } catch (System.Exception ex) {
                OGCSexception.Analyse(ex);
                Forms.Main.Instance.Console.UpdateWithError("Unable to authenticate with Google. The following error occurred:", ex);
            }

            if (credential.Token.AccessToken != "" && credential.Token.RefreshToken != "") {
                log.Info("Refresh and Access token successfully retrieved.");
                log.Debug("Access token expires " + credential.Token.IssuedUtc.AddSeconds(credential.Token.ExpiresInSeconds.Value).ToLocalTime().ToString());
            }

            GoogleOgcs.Calendar.Instance.Service = new CalendarService(new Google.Apis.Services.BaseClientService.Initializer() { HttpClientInitializer = credential });

            if (credential.Token.IssuedUtc.AddSeconds(credential.Token.ExpiresInSeconds.Value) < DateTime.UtcNow.AddMinutes(-1)) {
                log.Debug("Access token needs refreshing.");
                //This will happen automatically when using the calendar service
                //But we need a valid token before we call getGaccountEmail() which doesn't use the service
                try {
                    GoogleOgcs.Calendar.Instance.Service.Settings.Get("useKeyboardShortcuts").Execute();
                } catch (System.Exception ex) {
                    if (ex is Google.Apis.Auth.OAuth2.Responses.TokenResponseException)
                        OGCSexception.AnalyseTokenResponse(ex as Google.Apis.Auth.OAuth2.Responses.TokenResponseException, false);
                    else {
                        OGCSexception.Analyse(ex);
                        Forms.Main.Instance.Console.Update("Unable to communicate with Google services.", Console.Markup.warning);
                    }
                    authenticated = false;
                    return;
                }
                log.Debug("Access token refreshed.");
            }

            getGaccountEmail(credential.Token.AccessToken);
            authenticated = true;
            Forms.Main.Instance.Console.Update("Handshake successful.", verbose: true);
        }

        public void Reset(Boolean reauthorise = true) {
            log.Info("Resetting Google Calendar authentication details.");
            Settings.Instance.AssignedClientIdentifier = "";
            Settings.Instance.GaccountEmail = "";
            Forms.Main.Instance.SetControlPropertyThreadSafe(Forms.Main.Instance.tbConnectedAcc, "Text", "Not connected");
            authenticated = false;
            if (tokenFileExists) File.Delete(tokenFullPath);
            if (!GoogleOgcs.Calendar.IsInstanceNull) {
                GoogleOgcs.Calendar.Instance.Authenticator = null;
                GoogleOgcs.Calendar.Instance.Service = null;
                if (reauthorise) {
                    GoogleOgcs.Calendar.Instance.Authenticator = new Authenticator();
                    GoogleOgcs.Calendar.Instance.Authenticator.GetAuthenticated();
                }
            }
        }

        private Int16 getEmailAttempts = 0;
        private void getGaccountEmail(String accessToken) {
            String jsonString = "";
            log.Debug("Retrieving email address associated with Google account.");
            try {
                System.Net.WebClient wc = new System.Net.WebClient();
                wc.Headers.Add("user-agent", Settings.Instance.Proxy.BrowserUserAgent);
                jsonString = wc.DownloadString("https://www.googleapis.com/plus/v1/people/me?fields=emails&access_token=" + accessToken);
                JObject jo = Newtonsoft.Json.Linq.JObject.Parse(jsonString);
                JToken jtEmail = jo["emails"].Where(e => e.Value<String>("type") == "account").First();
                String email = jtEmail.Value<String>("value");

                if (Settings.Instance.GaccountEmail != email) {
                    if (!String.IsNullOrEmpty(Settings.Instance.GaccountEmail))
                        log.Debug("Looks like the Google account username value has been tampering with? :-O");
                    Settings.Instance.GaccountEmail = email;
                    Forms.Main.Instance.SetControlPropertyThreadSafe(Forms.Main.Instance.tbConnectedAcc, "Text", email);
                    log.Debug("Updating Google account username: " + Settings.Instance.GaccountEmail_masked());
                }
                getEmailAttempts = 0;
            } catch (System.Net.WebException ex) {
                getEmailAttempts++;
                if (ex.InnerException != null) log.Error("Inner exception: "+ ex.InnerException.Message);
                if (ex.Response != null) {
                    log.Debug("Reading response.");
                    System.IO.Stream stream = ex.Response.GetResponseStream();
                    System.IO.StreamReader sr = new System.IO.StreamReader(stream);
                    log.Error(sr.ReadToEnd());
                }
                if (OGCSexception.GetErrorCode(ex) == "0x80131509") {
                    log.Warn(ex.Message);
                    System.Text.RegularExpressions.Regex rgx = new System.Text.RegularExpressions.Regex(@"\b(403|Forbidden|Prohibited|Insufficient Permission)\b", 
                        System.Text.RegularExpressions.RegexOptions.IgnoreCase);

                    if (rgx.IsMatch(ex.Message)) {
                        if (Settings.Instance.UsingPersonalAPIkeys()) {
                            String msg = "If you are using your own API keys, you must also enable the Google+ API.";
                            MessageBox.Show(msg, "Missing API Service", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Exclamation);
                            throw new System.ApplicationException(msg);
                        } else {
                            if (getEmailAttempts > 1) {
                                log.Error("Failed to retrieve Google account username.");
                                log.Debug("Using previously retrieved username: " + Settings.Instance.GaccountEmail_masked());
                            } else {
                                ApiKeyring.ChangeKeys();
                                return;
                            }
                        }
                    } else {
                        throw ex;
                    }
                }
                OGCSexception.Analyse(ex);
                if (ex.Message.ToLower().Contains("access denied")) {
                    Forms.Main.Instance.Console.Update("Failed to obtain Calendar access from Google - it's possible your access has been revoked."
                       + "<br/>Try disconnecting your Google account and reauthorising OGCS.", Console.Markup.error);
                } else if (ex.Message.ToLower().Contains("prohibited") && Settings.Instance.UsingPersonalAPIkeys()) {
                    Forms.Main.Instance.Console.Update("If you are using your own API keys, you must also enable the Google+ API.", Console.Markup.warning);
                }
                throw ex;

            } catch (System.Exception ex) {
                log.Debug("JSON: " + jsonString);
                log.Error("Failed to retrieve Google account username.");
                OGCSexception.Analyse(ex);
                log.Debug("Using previously retrieved username: " + Settings.Instance.GaccountEmail_masked());
            }
        }

        public static String GetMd5(String input, Boolean isEmailAddress = false) {
            log.Debug("Getting MD5 hash for '" + (isEmailAddress ? EmailAddress.MaskAddress(input) : input) + "'");

            System.Text.StringBuilder sb = new System.Text.StringBuilder();
            System.Security.Cryptography.MD5 md5 = System.Security.Cryptography.MD5.Create();

            try {
                byte[] inputBytes = System.Text.Encoding.ASCII.GetBytes(input);
                byte[] hash = md5.ComputeHash(inputBytes);

                //Convert byte array to hex string
                for (int i = 0; i < hash.Length; i++) {
                    sb.Append(hash[i].ToString("x2"));
                }
            } catch (System.Exception ex) {
                log.Error("Failed to create MD5 for '" + EmailAddress.MaskAddress(input) + "'");
                OGCSexception.Analyse(ex);
            }
            return sb.ToString();
        }

        #region OGCS user status
        public void OgcsUserStatus() {
            if (!checkedOgcsUserStatus) {
                UserSubscriptionCheck();
                userDonationCheck();
                checkedOgcsUserStatus = true;
            }
        }

        public Boolean UserSubscriptionCheck() {
            List<Event> result = new List<Event>();
            Events request = null;
            String pageToken = null;
            Int16 pageNum = 1;

            log.Debug("Retrieving all subscribers from past year.");
            try {
                do {
                    EventsResource.ListRequest lr = GoogleOgcs.Calendar.Instance.Service.Events.List("hahospj0gkekqentakho0vv224@group.calendar.google.com");

                    lr.PageToken = pageToken;
                    lr.SingleEvents = true;
                    lr.OrderBy = EventsResource.ListRequest.OrderByEnum.StartTime;
                    lr.Q = (Settings.Instance.GaccountEmail == null) ? "" : HashedGmailAccount;
                    request = lr.Execute();
                    log.Debug("Page " + pageNum + " received.");

                    if (request != null) {
                        pageToken = request.NextPageToken;
                        pageNum++;
                        if (request.Items != null) result.AddRange(request.Items);
                    }
                } while (pageToken != null);

                if (String.IsNullOrEmpty(Settings.Instance.GaccountEmail)) { //This gets retrieved via the above lr.Execute()
                    log.Warn("User's Google account username is not present - cannot check if they have subscribed.");
                    return false;
                }
            } catch (Google.Apis.Auth.OAuth2.Responses.TokenResponseException ex) {
                OGCSexception.AnalyseTokenResponse(ex);

            } catch (System.Exception ex) {
                log.Error(ex.Message);
                throw new ApplicationException("Failed to retrieve subscribers - cannot check if they have subscribed.");
            }

            log.Debug("Searching for subscription for: " + Settings.Instance.GaccountEmail_masked());
            List<Event> subscriptions = result.Where(x => x.Summary == HashedGmailAccount).ToList();
            if (subscriptions.Count == 0) {
                log.Fine("This user has never subscribed.");
                Settings.Instance.Subscribed = DateTime.Parse("01-Jan-2000");
                return false;
            } else {
                Boolean subscribed;
                Event subscription = subscriptions.Last();
                DateTime subscriptionStart = (subscription.Start.DateTime ?? DateTime.Parse(subscription.Start.Date)).Date;
                log.Debug("Last subscription date: " + subscriptionStart.ToString());
                Double subscriptionRemaining = (subscriptionStart.AddYears(1) - DateTime.Now.Date).TotalDays;
                if (subscriptionRemaining >= 0) {
                    if (subscriptionRemaining > 360)
                        Forms.Main.Instance.SyncNote(Forms.Main.SyncNotes.RecentSubscription, null);
                    if (subscriptionRemaining < 28)
                        Forms.Main.Instance.SyncNote(Forms.Main.SyncNotes.SubscriptionPendingExpire, subscriptionStart.AddYears(1));
                    subscribed = true;
                } else {
                    if (subscriptionRemaining > -14)
                        Forms.Main.Instance.SyncNote(Forms.Main.SyncNotes.SubscriptionExpired, subscriptionStart.AddYears(1));
                    subscribed = false;
                }

                DateTime prevSubscriptionStart = Settings.Instance.Subscribed;
                if (subscribed) {
                    log.Info("User has an active subscription.");
                    Settings.Instance.Subscribed = subscriptionStart;
                } else {
                    log.Info("User has no active subscription.");
                    Settings.Instance.Subscribed = DateTime.Parse("01-Jan-2000");
                }

                //Check for any unmigrated entries
                if (subscriptions.Where(s => s.ExtendedProperties != null && s.ExtendedProperties.Shared != null
                    && s.ExtendedProperties.Shared.ContainsKey("migrated") && s.ExtendedProperties.Shared["migrated"] == "true").Count() < subscriptions.Count())
                    Forms.Main.Instance.Console.CallGappScript("subscriber");

                if (prevSubscriptionStart != Settings.Instance.Subscribed) {
                    if (prevSubscriptionStart == DateTime.Parse("01-Jan-2000")            //No longer a subscriber
                        || Settings.Instance.Subscribed == DateTime.Parse("01-Jan-2000")) //New subscriber
                    {
                        ApiKeyring.ChangeKeys();
                    }
                }
                return subscribed;
            }
        }
        private Boolean userDonationCheck() {
            List<Event> result = new List<Event>();
            Events request = null;
            String pageToken = null;
            Int16 pageNum = 1;

            log.Debug("Retrieving all donors.");
            try {
                do {
                    EventsResource.ListRequest lr = GoogleOgcs.Calendar.Instance.Service.Events.List("toiqu5lfdklneh5aqq509jhhk8@group.calendar.google.com");

                    lr.PageToken = pageToken;
                    lr.SingleEvents = true;
                    lr.OrderBy = EventsResource.ListRequest.OrderByEnum.StartTime;
                    lr.Q = (Settings.Instance.GaccountEmail == null) ? "" : HashedGmailAccount;
                    request = lr.Execute();
                    log.Debug("Page " + pageNum + " received.");

                    if (request != null) {
                        pageToken = request.NextPageToken;
                        pageNum++;
                        if (request.Items != null) result.AddRange(request.Items);
                    }
                } while (pageToken != null);

                if (String.IsNullOrEmpty(Settings.Instance.GaccountEmail)) { //This gets retrieved via the above lr.Fetch()
                    log.Warn("User's Google account username is not present - cannot check if they have donated.");
                    return false;
                }

            } catch (System.ApplicationException ex) {
                throw ex;

            } catch (System.Exception ex) {
                log.Error("Failed to retrieve donors - cannot check if they have donated.");
                log.Error(ex.Message);
                return false;
            }

            log.Debug("Searching for donation from: " + Settings.Instance.GaccountEmail_masked());
            List<Event> donations = result.Where(x => x.Summary == HashedGmailAccount).ToList();
            if (donations.Count == 0) {
                log.Fine("No donation found for user.");
                Settings.Instance.Donor = false;
                return false;
            } else {
                log.Fine("User has kindly donated.");
                Settings.Instance.Donor = true;

                //Check for any unmigrated entries
                if (donations.Where(d => d.ExtendedProperties != null && d.ExtendedProperties.Shared != null
                    && d.ExtendedProperties.Shared.ContainsKey("migrated") && d.ExtendedProperties.Shared["migrated"] == "true").Count() < donations.Count())
                    Forms.Main.Instance.Console.CallGappScript("donor");

                return true;
            }
        }
        #endregion
    }
}
