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

namespace OutlookGoogleCalendarSync {
    public class ApiKeyring {
        private static readonly ILog log = LogManager.GetLogger(typeof(ApiKeyring));
        private const String keyringURL = "https://outlookgooglecalendarsync.codeplex.com/wikipage?title=Keyring";

        public enum KeyType {
            Standard,
            Subscriber
        }
        private List<ApiKey> standardKeys = new List<ApiKey>();
        private List<ApiKey> subscriberKeys = new List<ApiKey>();

        private ApiKey key;
        public ApiKey Key { 
            get { return key; }
            set {
                key = value;
                Settings.Instance.AssignedClientIdentifier = key.ClientId;
            }
        }
        
        public Boolean PickKey(KeyType keyType) {
            try {
                List<ApiKey> keyRing = (keyType == KeyType.Standard) ? standardKeys : subscriberKeys;
                if (keyRing == null || keyRing.Count == 0) {
                    log.Debug(keyType.ToString() + " keyring is empty.");
                    Settings.Instance.AssignedClientIdentifier = "";
                    return false;
                }
                if (string.IsNullOrEmpty(Settings.Instance.AssignedClientIdentifier) &&
                    !string.IsNullOrEmpty(Settings.Instance.RefreshToken)) {
                    log.Debug("Legacy user with default API key.");
                    //Let them carry on with the default key, else everyone will have to re-authorise OGCS
                    //on a key with less quota
                    return false;
                }

                if (!string.IsNullOrEmpty(Settings.Instance.AssignedClientIdentifier)) {
                    Key = keyRing.Find(k => k.ClientId == Settings.Instance.AssignedClientIdentifier);
                    if (key != null) return true;
                    else log.Warn("Could not find assigned key!");
                }
                Random rnd = new Random();
                int chosenKey = rnd.Next(0, keyRing.Count - 1);
                log.Fine("Picked random key #" + chosenKey + 1);
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
            List<ApiKey> allKeys = getKeyRing();
            if (allKeys == null) return;

            List<ApiKey> activeKeys = new List<ApiKey>();
            activeKeys = allKeys.Where(key => key.Status == "ACTIVE").ToList();
            log.Debug(activeKeys.Count + " keys are active.");
            
            standardKeys = activeKeys.Where(key => key.Type == KeyType.Standard.ToString()).ToList();
            log.Fine(standardKeys.Count + " are standard keys.");
            subscriberKeys = activeKeys.Where(key => key.Type == KeyType.Subscriber.ToString()).ToList();
            log.Fine(subscriberKeys.Count + " are subscriber keys.");
        }

        private static List<ApiKey> getKeyRing() {
            List<ApiKey> keyRing = new List<ApiKey>();

            log.Debug("Getting keyring.");
            string html = "";
            try {
                html = new System.Net.WebClient().DownloadString(keyringURL);
            } catch (Exception ex) {
                log.Error("Failed to retrieve data: " + ex.Message);
            }
            if (!string.IsNullOrEmpty(html)) {
                html = html.Replace("\r\n", "");
                MatchCollection keyRecords = findText(html, @"<div class=\""wikidoc\""><table><tbody>(<tr>.*?</tr>)</tbody></table>");
                if (keyRecords.Count == 0) {
                    log.Warn("Could you not find table of keys.");
                    return keyRing;
                }
                foreach (String record in keyRecords[0].Captures[0].Value.Split(new string[]{"<tr>"}, StringSplitOptions.None)) {
                    MatchCollection keyAttributes = findText(record, @"<td>(.*?)</td>");
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
    }
}
