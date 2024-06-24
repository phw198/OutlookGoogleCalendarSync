﻿using Ogcs = OutlookGoogleCalendarSync;
using log4net;
using Microsoft.Win32;
using System;
using System.Linq;

namespace OutlookGoogleCalendarSync.Outlook {
    class Factory {
        private static readonly ILog log = LogManager.GetLogger(typeof(Factory));
        private static String outlookVersionFull;
        private static Int16 outlookVersion;

        private static String outlookVersionNameFull;
        public static String OutlookVersionNameFull {
            get {
                if (string.IsNullOrEmpty(outlookVersionNameFull)) {
                    getOutlookVersion();
                    Telemetry.Instance.OutlookVersion = outlookVersionFull;
                    Telemetry.Instance.OutlookVersionName = outlookVersionNameFull.Replace("Outlook", "");
                }
                return outlookVersionNameFull;
            }
        }
        private static OutlookVersionNames outlookVersionName;
        public static OutlookVersionNames OutlookVersionName {
            get {
                if (string.IsNullOrEmpty(outlookVersionFull)) getOutlookVersion();
                return outlookVersionName;
            }
        }
        public enum OutlookVersionNames : Int16 {
            Failed = -1,
            Unknown = 0,
            Outlook2003 = 11,
            Outlook2007 = 12,
            Outlook2010 = 14,
            Outlook2013 = 15,
            Outlook2016 = 16,
            //The following are faux numbers to distinguish v16 code base releases
            //https://docs.microsoft.com/en-us/office365/troubleshoot/installation/product-ids-supported-office-deployment-click-to-run
            ProPlusRetail = 17,
            ProPlus2019Retail = 18,
            O365ProPlusRetail = 19,
            O365HomePremRetail = 20,
            O365BusinessRetail = 21,
            HomeBusinessRetail = 22,
            HomeBusiness2019Retail = 23,
            HomeBusiness2021Retail = 24,
            HomeStudentRetail = 25,
            HomeStudent2019Retail = 26,
            OutlookRetail = 27,
            Outlook2019Retail = 28,
            Outlook2019Volume = 29,
            Outlook2021Volume = 30,
            Personal2019Retail = 31,
            Professional2019Retail = 32,
            Standard2019Volume = 33,
            Standard2021Volume = 34,
            ProPlus2019Volume = 35,
            ProPlus2021Volume = 36,
            ProPlus2021Retail = 37
        }

        private const Boolean testing2003 = false;

        public static Interface GetOutlookInterface() {
            if (OutlookVersionName >= OutlookVersionNames.Outlook2007) {
                return new OutlookNew();
            } else {
                return new OutlookOld();
            }
        }

        private static void getOutlookVersion() {
            //Attach just to get Outlook version - we don't know whether to provide New or Old interface yet
            Microsoft.Office.Interop.Outlook.Application oApp = null;
            Outlook.Calendar.AttachToOutlook(ref oApp);
            try {
                int attempts = 1;
                int maxAttempts = 3;
                while (attempts <= maxAttempts) {
                    try {
                        log.Fine("About to access Outlook oApp.version property...");
                        outlookVersionFull = oApp.Version;
                        attempts = maxAttempts + 1;
                    } catch (System.Runtime.InteropServices.COMException ex) {
                        Outlook.Errors.ErrorType error = Outlook.Errors.HandleComError(ex);
                        if (error == Outlook.Errors.ErrorType.PermissionFailure ||
                            error == Outlook.Errors.ErrorType.RpcRejected || 
                            error == Outlook.Errors.ErrorType.RpcServerUnavailable ||
                            error == Outlook.Errors.ErrorType.RpcFailed) //
                        {
                            log.Warn(ex.Message + " Attempt " + attempts + "/" + maxAttempts);
                            if (attempts == maxAttempts) {
                                String message = "Outlook has been unresponsive for " + maxAttempts * 10 + " seconds.\n" +
                                    "Please try running OGCS again later" +
                                    (Settings.Instance.StartOnStartup ? " or " + ((Settings.Instance.StartupDelay == 0) ? "set a" : "increase the") + " delay on startup." : ".");

                                throw new ApplicationException(message);
                            }
                            System.Threading.Thread.Sleep(10000);
                            attempts++;
                        } else throw;
                    }
                }

                log.Info("Outlook Version: " + outlookVersionFull);
#pragma warning disable 162 //Unreachable code
                if (testing2003) {
                    log.Info("*** 2003 TESTING ***");
                    outlookVersionFull = "11";
                }
#pragma warning restore 162
                outlookVersion = Convert.ToInt16(outlookVersionFull.Split(Convert.ToChar("."))[0]);
                getOutlookVersionName(outlookVersion, outlookVersionFull);

            } catch (System.Exception ex) {
                Outlook.Calendar.PoorlyOfficeInstall(ex);
            } finally {
                if (oApp != null) {
                    System.Runtime.InteropServices.Marshal.FinalReleaseComObject(oApp);
                    oApp = null;
                }
            }
        }

        private static void getOutlookVersionName(Int16 version, String versionFull) {
            try {
                outlookVersionNameFull = "Unknown-" + versionFull;
                outlookVersionName = OutlookVersionNames.Unknown;
                try {
                    outlookVersionName = (OutlookVersionNames)version;
                    outlookVersionNameFull = outlookVersionName.ToString();
                } catch (System.Exception ex) {
                    ex.Analyse("Failed determining Outlook client version.");
                    outlookVersionNameFull = "Failed-" + versionFull;
                    outlookVersionName = OutlookVersionNames.Failed;
                }

                try {
                    if (outlookVersionName == OutlookVersionNames.Outlook2016) { //The code base is the same from 2016 onwards (eg. 2019 and O365)
                        log.Debug("Checking for more accurate Outlook version for codebase v16.");

                        //Open registry as 32-bit first
                        RegistryKey baseKey = RegistryKey.OpenBaseKey(RegistryHive.LocalMachine, RegistryView.Registry32);
                        RegistryKey regKey = baseKey.OpenSubKey(@"SOFTWARE\Microsoft\Office\ClickToRun");
                        if (regKey == null || regKey.SubKeyCount == 0) {
                            log.Debug("Try accessing 64-bit registry key");
                            baseKey = RegistryKey.OpenBaseKey(RegistryHive.LocalMachine, RegistryView.Registry64);
                            regKey = baseKey.OpenSubKey(@"SOFTWARE\Microsoft\Office\ClickToRun");
                        }
                        if (regKey == null || regKey.SubKeyCount == 0) {
                            log.Debug("It's a 2016 MSI install.");
                        } else {
                            log.Fine("It's a Click-to-Run install.");
                            String regReleaseValue = null;
                            try {
                                if (regKey.GetSubKeyNames().Contains("Configuration")) {
                                    regKey = regKey.OpenSubKey("Configuration");
                                    if (regKey.GetValueNames().Contains("ProductReleaseIds")) {
                                        regReleaseValue = regKey.GetValue("ProductReleaseIds").ToString();
                                        OutlookVersionNames outlookVersionNameFor2016 = OutlookVersionNames.Unknown;
                                        foreach (String product in regReleaseValue.Split(',')) {
                                            if (Enum.TryParse(product, true, out outlookVersionNameFor2016)) {
                                                outlookVersionNameFull = outlookVersionNameFor2016.ToString();
                                                break;
                                            }
                                        }
                                        if (outlookVersionNameFor2016 == OutlookVersionNames.Unknown) {
                                            log.Error("Could not determine exact Outlook version with codebase v16. " + regReleaseValue);
                                        }
                                    } else {
                                        log.Warn("'ProductReleaseIds' value does not exist.");
                                        log.Debug(String.Join(",", regKey.GetValueNames()));
                                    }
                                } else {
                                    log.Warn("'Configuration' subdirectory does not exist.");
                                    log.Debug(String.Join(",", regKey.GetSubKeyNames()));
                                }

                                if (regReleaseValue == null)
                                    log.Error("Could not determine exact Outlook version with codebase v16.");

                            } catch (System.Exception ex) {
                                ex.Analyse("Failed determining Click-to-Run release.");
                            }
                        }
                    }
                } catch (System.Exception ex) {
                    ex.Analyse("Failed determining Outlook release name from registry for codebase v16.");
                }
            } finally {
                log.Info("Outlook product name: " + outlookVersionNameFull);
            }
        }
    }
}
