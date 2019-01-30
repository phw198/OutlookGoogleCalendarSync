using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Microsoft.Office.Interop.Outlook;
using log4net;

namespace OutlookGoogleCalendarSync.OutlookOgcs {
    class Factory {
        private static readonly ILog log = LogManager.GetLogger(typeof(Factory));
        private static String outlookVersionFull;
        private static int outlookVersion;
        public static int OutlookVersion {
            get {
                if (string.IsNullOrEmpty(outlookVersionFull)) getOutlookVersion();
                return outlookVersion;
            }
        }
        private const Boolean testing2003 = false;

        public static OutlookOgcs.Interface GetOutlookInterface() {
            if (OutlookVersion >= 12) { //2007 or newer
                return new OutlookNew();
            } else {
                return new OutlookOld();
            }
        }

        private static void getOutlookVersion() {
            //Attach just to get Outlook version - we don't know whether to provide New or Old interface yet
            Microsoft.Office.Interop.Outlook.Application oApp = null;
            try {
                OutlookOgcs.Calendar.AttachToOutlook(ref oApp);
                outlookVersionFull = oApp.Version;

                log.Info("Outlook Version: " + outlookVersionFull);
                if (testing2003) {
                    #pragma warning disable 162 //Unreachable code
                    log.Info("*** 2003 TESTING ***");
                    outlookVersionFull = "11";
                    #pragma warning restore 162
                }
                outlookVersion = Convert.ToInt16(outlookVersionFull.Split(Convert.ToChar("."))[0]);

            } finally {
                if (oApp != null) {
                    System.Runtime.InteropServices.Marshal.FinalReleaseComObject(oApp);
                    oApp = null;
                }
            }
        }

        public static Boolean is2003() {
            return outlookVersionFull == "11";
        }
    }
}
