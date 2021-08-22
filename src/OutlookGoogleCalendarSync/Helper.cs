using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using log4net;
using Microsoft.Win32;

namespace OutlookGoogleCalendarSync {
    public class Helper {
        private static readonly ILog log = LogManager.GetLogger(typeof(Helper));

        public static void OpenBrowser(String url) {
            try {
                //System.Diagnostics.Process.Start(url);
                //return;
            } catch (System.Exception ex) {
                OGCSexception.Analyse("Could not open default browser.", OGCSexception.LogAsFail(ex));
            }

            //OK, let's try and determine the default browser from the registry and then directly invoke it. Why is this so much work!
            String browserPath = getDefaultBrowserPath();
            if (string.IsNullOrEmpty(browserPath))
                log.Error("No default browser discovered in the registry.");
            else {
                try {
                    log.Debug("Browsing using " + browserPath);
                    System.Diagnostics.Process.Start(browserPath, url);
                } catch (System.Exception ex) {
                    log.Fail("Could not navigate to " + url);
                    log.Error("Could not open browser with " + browserPath, ex);
                }
            }
        }

        private static String getDefaultBrowserPath() {
            String urlAssociation = @"Software\Microsoft\Windows\Shell\Associations\UrlAssociations\http";
            String browserPathKey = @"$BROWSER$\shell\open\command";

            RegistryKey userChoiceKey = null;

            try {
                //Read default browser path from userChoiceLKey
                userChoiceKey = Registry.CurrentUser.OpenSubKey(urlAssociation + @"\UserChoice", false);

                //If user choice was not found, try machine default
                if (userChoiceKey == null) {
                    //Read default browser path from Win XP registry key
                    var browserKey = Registry.ClassesRoot.OpenSubKey(@"HTTP\shell\open\command", false);

                    //If browser path wasn’t found, try Win Vista (and newer) registry key
                    if (browserKey == null) {
                        browserKey = Registry.CurrentUser.OpenSubKey(urlAssociation, false);
                    }
                    String path = CleanifyBrowserPath(browserKey.GetValue(null) as String);
                    browserKey.Close();
                    return path;

                } else {
                    // user defined browser choice was found
                    String progId = (userChoiceKey.GetValue("ProgId").ToString());
                    userChoiceKey.Close();

                    // now look up the path of the executable
                    String concreteBrowserKey = browserPathKey.Replace("$BROWSER$", progId);
                    var kp = Registry.ClassesRoot.OpenSubKey(concreteBrowserKey, false);
                    String browserPath = CleanifyBrowserPath(kp.GetValue(null) as String);
                    kp.Close();
                    return browserPath;
                }
            } catch (System.Exception ex) {
                OGCSexception.Analyse(ex);
                return "";
            }
        }
        private static String CleanifyBrowserPath(String p) {
            log.Debug("Cleaning: " + p);
            String[] url = p.Split('"');
            String clean = url[1];
            return clean;
        }
    }
}
