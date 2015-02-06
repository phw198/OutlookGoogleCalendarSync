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
        public static int outlookVersion;
        private const Boolean testing2003 = false;

        public static OutlookInterface getOutlookInterface() {
            getOutlookVersion();
            outlookVersion = Convert.ToInt16(outlookVersionFull.Split(Convert.ToChar("."))[0]);
            if (testing2003) outlookVersion = 11;
            if (outlookVersion >= 12) { //2007 or newer
                return new OutlookNew();
            } else {
                return new OutlookOld();
            }
            return null;
        }

        private static void getOutlookVersion() {
            Microsoft.Office.Interop.Outlook.Application oApp = new Microsoft.Office.Interop.Outlook.Application();
            log.Info("Outlook Version: " + oApp.Version);
            outlookVersionFull = oApp.Version;
        }
    }
}
