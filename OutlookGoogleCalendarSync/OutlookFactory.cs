using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Microsoft.Office.Interop.Outlook;
using log4net;

namespace OutlookGoogleCalendarSync {
    class OutlookFactory {
        private static readonly ILog log = LogManager.GetLogger(typeof(OutlookFactory));
        private static String outlookVersionFull;
        private static int outlookVersion;
        public static int OutlookVersion {
            get {
                if (string.IsNullOrEmpty(outlookVersionFull)) getOutlookVersion();
                return outlookVersion;
            }
        }
        private const Boolean testing2003 = false;

        public static OutlookInterface getOutlookInterface() {
            if (OutlookVersion >= 12) { //2007 or newer
                return new OutlookNew();
            } else {
                return new OutlookOld();
            }
        }

        private static void getOutlookVersion() {
            //Attach just to get Outlook version - we don't know whether to provide New or Old interface yet
            Microsoft.Office.Interop.Outlook.Application oApp = OutlookCalendar.AttachToOutlook();
            outlookVersionFull = oApp.Version;

            /* try {
                outlookVersionFull = oApp.Version;
            } catch (System.Exception ex) {
                if (ex.Message.Contains("RPC_E_CALL_REJECTED")) { //Issue #86: Call was rejected by callee.
                    log.Warn("The Outlook GUI is not quite ready yet - waiting until it is...");
                    int maxWait = 30;
                    while (maxWait > 0) {
                        if (maxWait < 30 && maxWait % 10 == 0) { log.Debug("Still waiting..."); }
                        try {
                            outlookVersionFull = oApp.Version;
                            maxWait = 0;
                        } catch { }
                        System.Threading.Thread.Sleep(1000);
                        maxWait--;
                    }
                    if (maxWait > 0) {
                        log.Error("Given up waiting for Outlook to respond.");
                        throw new ApplicationException("Outlook would not give timely response.");
                    }
                } else {
                    throw ex;
                }
            } finally {
                System.Runtime.InteropServices.Marshal.FinalReleaseComObject(oApp);
                oApp = null;
            }*/

            log.Info("Outlook Version: " + outlookVersionFull);
            if (testing2003) {
                log.Info("*** 2003 TESTING ***");
                outlookVersionFull = "11";
            }
            outlookVersion = Convert.ToInt16(outlookVersionFull.Split(Convert.ToChar("."))[0]);
            System.Runtime.InteropServices.Marshal.FinalReleaseComObject(oApp);
            oApp = null;
        }
    }
}
