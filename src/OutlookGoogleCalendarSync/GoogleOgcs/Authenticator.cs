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

        private const String tokenFile = "Google.Apis.Auth.OAuth2.Responses.TokenResponse-user";
        String tokenFullPath;
        Boolean tokenFileExists { get { return File.Exists(tokenFullPath); } }

        private Boolean checkedOgcsUserStatus = false;

        public Authenticator() {
            ClientSecrets cs = getCalendarClientSecrets();
            try {
                //Calling an async function from a static constructor needs to be called like this, else it deadlocks:-
                var task = System.Threading.Tasks.Task.Run(async () => { await getAuthenticated(cs); });
                task.Wait();
            } catch (System.Exception) {
                log.Error("Problem encountered instantiating Authenticator.");
                throw;
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
            tokenFullPath = Path.Combine(tokenStore.FolderPath, tokenFile);

            log.Debug("Google credential file location: " + tokenFullPath);
            if (!tokenFileExists)
                log.Info("No Google credentials file available - need user authorisation for OGCS to manage their calendar.");

            //string[] scopes = new[] { "https://www.googleapis.com/auth/calendar", "https://www.googleapis.com/auth/userinfo.email" };
            string[] scopes = new[] { "https://www.googleapis.com/auth/calendar", "email" };

            UserCredential credential = null;
            try {
                //This will open the authorisation process in a browser, if required
                credential = await GoogleWebAuthorizationBroker.AuthorizeAsync(cs, scopes, "user", System.Threading.CancellationToken.None, tokenStore);
                if (!tokenFileExists)
                    log.Debug("User has provided authentication code and credential file saved.");

            } catch (Google.Apis.Auth.OAuth2.Responses.TokenResponseException ex) {
                //OGCSexception.AnalyseTokenResponse(ex);
                if (ex.Error.Error == "access_denied") {
                    String noAuthGiven = "Sorry, but this application will not work if you don't allow it access to your Google Calendar :(";
                    log.Warn("User did not provide authorisation code. Sync will not be able to work.");
                    MessageBox.Show(noAuthGiven, "Authorisation not given", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    throw new ApplicationException(noAuthGiven);
                } else {
                    MainForm.Instance.AsyncLogboxout("Unable to authenticate with Google. The following error occurred:");
                    MainForm.Instance.AsyncLogboxout(ex.Message);
                }

            } catch (System.Exception ex) {
                OGCSexception.Analyse(ex);
                MainForm.Instance.AsyncLogboxout("Unable to authenticate with Google. The following error occurred:");
                MainForm.Instance.AsyncLogboxout(ex.Message);
            }
            
            if (credential.Token.AccessToken != "" && credential.Token.RefreshToken != "") {
                log.Info("Refresh and Access token successfully retrieved.");
                log.Debug("Access token expires " + credential.Token.Issued.AddSeconds(credential.Token.ExpiresInSeconds.Value).ToLocalTime().ToString());
            }

            GoogleCalendar.Instance.Service = new CalendarService(new Google.Apis.Services.BaseClientService.Initializer() { HttpClientInitializer = credential });

            if (credential.Token.Issued.AddSeconds(credential.Token.ExpiresInSeconds.Value) < DateTime.Now.AddMinutes(-1)) {
                log.Debug("Access token needs refreshing.");
                //This will happen automatically when using the calendar service
                //But we need a valid token before we call getGaccountEmail() which doesn't use the service
                GoogleCalendar.Instance.Service.Settings.Get("useKeyboardShortcuts").Execute();
                log.Debug("Access token refreshed.");
            }
            
            getGaccountEmail(credential.Token.AccessToken);
        }

        public void Reset() {
            log.Info("Resetting Google Calendar authentication details.");
            Settings.Instance.AssignedClientIdentifier = "";
            if (tokenFileExists) File.Delete(tokenFullPath);
            GoogleCalendar.Instance.Authenticator = new Authenticator();
        }

        private void getGaccountEmail(String accessToken) {
            String jsonString = "";
            log.Debug("Retrieving email address associated with Google account.");
            try {
                System.Net.WebClient wc = new System.Net.WebClient();
                wc.Headers.Add("user-agent", "Mozilla/5.0 (Windows NT 6.1; WOW64; rv:37.0) Gecko/20100101 Firefox/37.0");
                jsonString = wc.DownloadString("https://www.googleapis.com/plus/v1/people/me?fields=emails&access_token=" + accessToken);
                JObject jo = Newtonsoft.Json.Linq.JObject.Parse(jsonString);
                JToken jtEmail = jo["emails"].Where(e => e.Value<String>("type") == "account").First();
                String email = jtEmail.Value<String>("value");

                if (Settings.Instance.GaccountEmail != email) {
                    if (!String.IsNullOrEmpty(Settings.Instance.GaccountEmail))
                        log.Debug("Looks like the Google account username value has been tampering with? :-O");
                    Settings.Instance.GaccountEmail = email;
                    log.Debug("Updating Google account username: " + Settings.Instance.GaccountEmail_masked());
                }
            } catch (System.Net.WebException ex) {
                if (ex.Message.Contains("The remote server returned an error: (403) Forbidden.") || ex.Message == "Insufficient Permission") {
                    log.Warn(ex.Message);
                    String msg = ApiKeyring.ChangeKeys();
                    throw new System.ApplicationException(msg);
                }
                OGCSexception.Analyse(ex);
                if (ex.Message.ToLower().Contains("access denied")) {
                    MainForm.Instance.Logboxout("Failed to obtain Calendar access from Google - it's possible your access has been revoked."
                       + "\r\nTry disconnecting your Google account and reauthenticating.");
                }
                throw ex;

            } catch (System.Exception ex) {
                log.Debug("JSON: " + jsonString);
                log.Error("Failed to retrieve Google account username.");
                OGCSexception.Analyse(ex);
                log.Debug("Using previously retrieved username: " + Settings.Instance.GaccountEmail_masked());
            }
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
                    EventsResource.ListRequest lr = GoogleCalendar.Instance.Service.Events.List("pqeo689qhvpl1g09bcnma1uaoo@group.calendar.google.com");

                    lr.PageToken = pageToken;
                    lr.SingleEvents = true;
                    lr.OrderBy = EventsResource.ListRequest.OrderByEnum.StartTime;
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
            List<Event> subscriptions = result.Where(x => x.Summary.Equals(Settings.Instance.GaccountEmail)).ToList();
            if (subscriptions.Count == 0) {
                log.Fine("This user has never subscribed.");
                Settings.Instance.Subscribed = DateTime.Parse("01-Jan-2000");
                return false;
            } else {
                Boolean subscribed;
                Event subscription = subscriptions.Last();
                DateTime subscriptionStart = DateTime.Parse(subscription.Start.Date ?? subscription.Start.DateTime).Date;
                log.Debug("Last subscription date: " + subscriptionStart.ToString());
                Double subscriptionRemaining = (subscriptionStart.AddYears(1) - DateTime.Now.Date).TotalDays;
                if (subscriptionRemaining >= 0) {
                    if (subscriptionRemaining > 360)
                        MainForm.Instance.syncNote(MainForm.SyncNotes.RecentSubscription, null);
                    if (subscriptionRemaining < 28)
                        MainForm.Instance.syncNote(MainForm.SyncNotes.SubscriptionPendingExpire, subscriptionStart.AddYears(1));
                    subscribed = true;
                } else {
                    if (subscriptionRemaining > -14)
                        MainForm.Instance.syncNote(MainForm.SyncNotes.SubscriptionExpired, subscriptionStart.AddYears(1));
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
                    EventsResource.ListRequest lr = GoogleCalendar.Instance.Service.Events.List("codejbnbj3dp71bj63ingjii9g@group.calendar.google.com");

                    lr.PageToken = pageToken;
                    lr.SingleEvents = true;
                    lr.OrderBy = EventsResource.ListRequest.OrderByEnum.StartTime;
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
            List<Event> donations = result.Where(x => x.Summary.Equals(Settings.Instance.GaccountEmail)).ToList();
            if (donations.Count == 0) {
                log.Fine("No donation found for user.");
                Settings.Instance.Donor = false;
                return false;
            } else {
                log.Fine("User has kindly donated.");
                Settings.Instance.Donor = true;
                return true;
            }
        }
        #endregion
    }
}
