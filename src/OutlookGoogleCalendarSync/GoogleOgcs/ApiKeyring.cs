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

namespace OutlookGoogleCalendarSync.GoogleOgcs {
    public class ApiKeyring {
        private static readonly ILog log = LogManager.GetLogger(typeof(ApiKeyring));
        private const String keyringURL = "https://github.com/phw198/OutlookGoogleCalendarSync/blob/master/docs/keyring.md";

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
        
        public Boolean PickKey(KeyType keyType) {
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
                        return true;
                    } else {
                        log.Debug("Reverting to default key.");
                        Settings.Instance.AssignedClientIdentifier = "";
                        return false;
                    }
                }

                if (!string.IsNullOrEmpty(Settings.Instance.AssignedClientIdentifier)) {
                    ApiKey retrievedKey = keyRing.Find(k => k.ClientId == Settings.Instance.AssignedClientIdentifier);
                    if (retrievedKey == null) {
                        log.Warn("Could not find assigned key on keyring!");
                        if (standardKeys.Concat(subscriberKeys).Any(k => k.ClientId == Settings.Instance.AssignedClientIdentifier)) {
                            log.Warn("The key was been taken from the other keyring!");
                        }
                        if (!string.IsNullOrEmpty(Settings.Instance.AssignedClientIdentifier) &&
                            !string.IsNullOrEmpty(Settings.Instance.AssignedClientSecret)) 
                        {
                            log.Debug("Using key from settings file.");
                            key = new ApiKey(Settings.Instance.AssignedClientIdentifier, Settings.Instance.AssignedClientSecret);
                            return true;
                        }
                    } else {
                        if (retrievedKey.Status == KeyStatus.DEAD.ToString())
                            log.Warn("The assigned key can no longer be used. A new key must be assigned.");
                        else {
                            key = retrievedKey;
                            return true;
                        }
                    }
                }
                keyRing = keyRing.Where(k => k.Status == KeyStatus.ACTIVE.ToString()).ToList();
                Random rnd = new Random();
                int chosenKey = rnd.Next(0, keyRing.Count - 1);
                log.Fine("Picked random active key #" + chosenKey + 1);
                Key = keyRing[chosenKey];
                return true;

            } catch (System.Exception ex) {
                log.Error("Failed picking API key. clientID=" + Settings.Instance.AssignedClientIdentifier);
                log.Error(ex.Message);
                Settings.Instance.AssignedClientIdentifier = "";
                return false;
            }
        }
        
        public ApiKeyring() {
            MainForm.Instance.Console.Update("Preparing to authenticate with Google.", verbose: true);

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
                html = new System.Net.WebClient().DownloadString(keyringURL);
            } catch (System.Exception ex) {
                log.Error("Failed to retrieve data: " + ex.Message);
            }
            if (!string.IsNullOrEmpty(html)) {
                html = html.Replace("\n", "");
                MatchCollection keyRecords = findText(html, @"<article class=\""markdown-body entry-content\"".*?<tbody>(<tr>.*?</tr>)</tbody></table>");
                if (keyRecords.Count == 0) {
                    log.Warn("Could you not find table of keys.");
                    return keyRing;
                }
                foreach (String record in keyRecords[0].Captures[0].Value.Split(new string[]{"<tr>"}, StringSplitOptions.None)) {
                    MatchCollection keyAttributes = findText(record, @"<td.*?>(.*?)</td>");
                    if (keyAttributes.Count > 0) {
                        try {
                            keyRing.Add(new ApiKey(keyAttributes));
                        } catch { }
                    }
                }
            }
            log.Debug("There are " + keyRing.Count + " keys.");
            return keyRing;
        }

        public static String ChangeKeys() {
            log.Info("Google API keys and refresh token are being updated.");
            String msg = "Your Google authorisation token needs updating.\r\n" +
                        "The process to reauthorise access to your Google account will now begin...";
            System.Windows.Forms.MessageBox.Show(msg, "Authorisation token invalid", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Exclamation);
            GoogleOgcs.Calendar.Instance.Authenticator.Reset();
            return msg;
        }

        private static MatchCollection findText(string source, string pattern) {
            Regex rgx = new Regex(pattern, RegexOptions.IgnoreCase);
            return rgx.Matches(source);
        }
    }


    public class ApiKey {
        private static readonly ILog log = LogManager.GetLogger(typeof(ApiKey));
        
        private String type;
        private String projectName;
        private String projectId;
        private String status;
        private String clientId;
        private String clientSecret;
        
        public String Type { get { return type; } }
        public String ProjectName { get { return projectName; } }
        public String ProjectID { get { return projectId; } }
        public String Status { get { return status; } }
        public String ClientId { get { return clientId; } }
        public String ClientSecret { get { return clientSecret; } }

        public ApiKey(MatchCollection keyAttributes) {
            //Table columns
            const int keyType = 1;
            const int keyProjectName = 2;
            const int keyProjectID = 3;
            const int keyStatus = 4;
            const int keyClientId = 5;
            const int keySecret = 6;

            try {
                log.Debug(keyAttributes[keyProjectName - 1].Groups[1] + " = " + keyAttributes[keyStatus - 1].Groups[1]);
                type = keyAttributes[keyType - 1].Groups[1].ToString();
                projectName = keyAttributes[keyProjectName - 1].Groups[1].ToString();
                projectId = keyAttributes[keyProjectID - 1].Groups[1].ToString();
                status = keyAttributes[keyStatus - 1].Groups[1].ToString();
                clientId = keyAttributes[keyClientId - 1].Groups[1].ToString();
                clientSecret = keyAttributes[keySecret - 1].Groups[1].ToString();
            } catch (System.Exception ex) {
                log.Error("Failed creating API key.");
                OGCSexception.Analyse(ex);
                throw ex;
            }
        }

        public ApiKey(String clientId, String clientSecret) {
            this.clientId = clientId;
            this.clientSecret = clientSecret;
        }
    }
}
