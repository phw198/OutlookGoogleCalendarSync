using Ogcs = OutlookGoogleCalendarSync;
using log4net;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

/* Google are beginning to sulk about API usage for this project, so we need to future proof.
 * There's no real limit on the number of Google developer projects that can use the Calendar API,
 * so the plan is to create a pool of them from among the user community who are willing to set up a project for general use.
 *  
 * NOTE: Incremental syncs is an option that will in theory reduce API demand. However, that's close to a full rewrite of the
 *       sync engine, so that's for another day!
 */

namespace OutlookGoogleCalendarSync.Google {
    public class ApiKeyring {
        private static readonly ILog log = LogManager.GetLogger(typeof(ApiKeyring));
        private const String keyringURL = "https://github.com/phw198/OutlookGoogleCalendarSync/raw/master/docs/keyring.md";

        public enum KeyType {
            Standard,
            Subscriber
        }
        private enum KeyStatus {
            ACTIVE,
            DEAD,
            DISABLED,
            FULL
        }

        private List<ApiKey> standardKeys = new List<ApiKey>();
        private List<ApiKey> subscriberKeys = new List<ApiKey>();

        private ApiKey key;
        public ApiKey Key { 
            get { return key; }
            set {
                key = value;
                Settings.Instance.AssignedClientIdentifier = key.ClientId;
                Settings.Instance.AssignedClientSecret = key.ClientSecret;
            }
        }
        
        public void PickKey(KeyType keyType) {
            log.Debug("Picking a " + keyType.ToString() + " key.");
            try {
                List<ApiKey> keyRing = (keyType == KeyType.Standard) ? standardKeys : subscriberKeys;
                if (keyRing == null || keyRing.Count == 0) {
                    log.Warn(keyType.ToString() + " keyring is empty!");
                    if (!string.IsNullOrEmpty(Settings.Instance.AssignedClientIdentifier) &&
                        !string.IsNullOrEmpty(Settings.Instance.AssignedClientSecret)) 
                    {
                        log.Debug("Using key from settings file.");
                        key = new ApiKey(Settings.Instance.AssignedClientIdentifier, Settings.Instance.AssignedClientSecret);
                    } else {
                        log.Debug("Reverting to default key.");
                        Key = new ApiKey.DefaultKey(keyType);
                        return;
                    }
                }

                if (!string.IsNullOrEmpty(Settings.Instance.AssignedClientIdentifier)) {
                    if (Settings.Instance.AssignedClientIdentifier == new ApiKey.DefaultKey(keyType).ClientId) {
                        log.Fine("Using default " + keyType.ToString() + " key.");
                        key = new ApiKey.DefaultKey(keyType);
                        return;
                    }
                    log.Fine("Checking non-default assigned API key is still on the keyring.");
                    ApiKey retrievedKey = keyRing.Find(k => k.ClientId == Settings.Instance.AssignedClientIdentifier);
                    if (retrievedKey == null) {
                        log.Warn("Could not find assigned key on keyring!");
                        if (standardKeys.Concat(subscriberKeys).Any(k => k.ClientId == Settings.Instance.AssignedClientIdentifier)) {
                            log.Warn("The key was been taken from the other keyring!");
                        }
                        if (keyRing == null || keyRing.Count == 0) {
                            log.Debug("Reverting to default key.");
                            Key = new ApiKey.DefaultKey(keyType);
                            return;
                        } else 
                            log.Debug("Picking a new key from the ring.");
                    } else {
                        if (retrievedKey.Status == KeyStatus.DEAD.ToString())
                            log.Warn("The assigned key can no longer be used. A new key must be assigned.");
                        else {
                            Key = retrievedKey;
                            return;
                        }
                    }
                }
                keyRing = keyRing.Where(k => k.Status == KeyStatus.ACTIVE.ToString()).ToList();
                Random rnd = new Random();
                int chosenKey = rnd.Next(0, keyRing.Count - 1);
                log.Fine("Picked random active key #" + chosenKey + 1);
                Key = keyRing[chosenKey];

            } catch (System.Exception ex) {
                log.Fail("Failed picking "+ keyType.ToString() +" API key. clientID=" + Settings.Instance.AssignedClientIdentifier);
                Ogcs.Exception.Analyse(ex);
                log.Debug("Reverting to default key.");
                Key = new ApiKey.DefaultKey(keyType);
            }
        }
        
        public ApiKeyring() {
            Forms.Main.Instance.Console.Update("Preparing to authenticate with Google.", verbose: true);

            List<ApiKey> allKeys = getKeyRing();
            if (allKeys == null) return;

            log.Debug(allKeys.Where(key => key.Status == KeyStatus.ACTIVE.ToString()).Count() + " keys are active.");

            standardKeys = allKeys.Where(key => key.Type == KeyType.Standard.ToString()).ToList();
            log.Fine(standardKeys.Count + " are standard keys.");
            subscriberKeys = allKeys.Where(key => key.Type == KeyType.Subscriber.ToString()).ToList();
            log.Fine(subscriberKeys.Count + " are subscriber keys.");
        }

        private static List<ApiKey> getKeyRing() {
            List<ApiKey> keyRing = new List<ApiKey>();

            log.Debug("Getting keyring.");
            string html = "";
            try {
                html = new Extensions.OgcsWebClient().DownloadString(keyringURL);
            } catch (System.Exception ex) {
                log.Error("Failed to retrieve keyring data: " + ex.Message);
            }
            if (!string.IsNullOrEmpty(html)) {
                html = html.Replace("\n", "");
                MatchCollection keyRecords = findText(html, @"\|(?<type>Standard)\|(?<projectName>.*?)\|(?<projectId>[a-z\-\d]+)\|(?<status>[A-Z]+)\|(?<clientId>.+?)\|(?<clientSecret>.*?)\|");
                if (keyRecords.Count == 0) {
                    log.Warn("Could not find table of keys.");
                    return keyRing;
                }
                foreach (Match record in keyRecords) {
                    keyRing.Add(new ApiKey(
                        type: record.Groups["type"].Value,
                        projectName: record.Groups["projectName"].Value,
                        projectId: record.Groups["projectId"].Value,
                        status: record.Groups["status"].Value,
                        clientId: record.Groups["clientId"].Value,
                        clientSecret: record.Groups["clientSecret"].Value
                    ));
                }
            }
            log.Debug("There are " + keyRing.Count + " keys.");
            return keyRing;
        }

        public static void ChangeKeys() {
            log.Info("Google API keys and refresh token are being updated.");
            System.Windows.Forms.OgcsMessageBox.Show("Your Google authorisation token needs updating.\r\n" +
                "The process to reauthorise access to your Google account will now begin...",
                "Authorisation token invalid", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Exclamation);
            Ogcs.Google.Calendar.Instance.Authenticator.Reset();
        }

        private static MatchCollection findText(string source, string pattern) {
            Regex rgx = new Regex(pattern, RegexOptions.IgnoreCase);
            return rgx.Matches(source);
        }
    }


    public class ApiKey {
        private static readonly ILog log = LogManager.GetLogger(typeof(ApiKey));
        
        public String Type { get; protected set; }
        public String ProjectName { get; protected set; }
        public String ProjectID { get; protected set; }
        public String Status { get; protected set; }
        public String ClientId { get; protected set; }
        public String ClientSecret { get; protected set; }

        public ApiKey(String type, String projectName, String projectId, String status, String clientId, String clientSecret) {
            try {
                log.Debug(projectName + " = " + status);
                this.Type = type;
                this.ProjectName = projectName;
                this.ProjectID = projectId;
                this.Status = status;
                this.ClientId = clientId;
                this.ClientSecret = clientSecret;
            } catch (System.Exception ex) {
                log.Error("Failed creating API key.");
                Ogcs.Exception.Analyse(ex);
                throw;
            }
        }

        public ApiKey(String clientId, String clientSecret) {
            this.ClientId = clientId;
            this.ClientSecret = clientSecret;
        }

        private ApiKey() { }

        public class DefaultKey : ApiKey {
            public DefaultKey(ApiKeyring.KeyType type) {
                this.Type = type.ToString();
                this.Status = "ACTIVE";
                if (type == ApiKeyring.KeyType.Standard) {
                    this.ProjectName = "OGCS Default";
                    this.ProjectID = "outlook-google-calendar-sync";
                    this.ClientId = "653617509806-2nq341ol8ejgqhh2ku4j45m7q2bgdimv.apps.googleusercontent.com";
                    this.ClientSecret = "tAi-gZLWtasS58i8CcCwVwsq";

                } else if (type == ApiKeyring.KeyType.Subscriber) {
                    this.ProjectName = "Premium OGCS";
                    this.ProjectID = "premium-ogcs";
                    this.ClientId = "550071650559-44lnvhdu5liq5kftj5t8k0aasgei5g7t.apps.googleusercontent.com";
                    this.ClientSecret = "MGUFapefXClJa2ysS4WNGS4k";
                }
            }
        }
    }
}
