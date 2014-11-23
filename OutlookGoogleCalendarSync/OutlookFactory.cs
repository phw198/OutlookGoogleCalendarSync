using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Microsoft.Office.Interop.Outlook;

namespace OutlookGoogleCalendarSync {
    class OutlookFactory {
        private static String outlookVersionFull;
        public static int outlookVersion;
        private const Boolean testing2003 = false;

        public static OutlookInterface getOutlookInterface() {
            getOutlookVersion();
            outlookVersion = Convert.ToInt16(outlookVersionFull.Split(Convert.ToChar("."))[0]);
            if (testing2003) outlookVersion = 11;
            if (outlookVersion >= 14) { //2010 or newer
                return new OutlookNew();
            } else {
                if (MessageBox.Show("Support for Outlook 2007 and earlier is not properly supported yet.\r\n" +
                    "Click OK to continue...but please be patient with bugs!", "Under construction",
                    MessageBoxButtons.OKCancel) == DialogResult.OK) {
                    return new OutlookOld();
                } else {
                    System.Windows.Forms.Application.Exit();
                }
            }
            return null;
        }

        private static void getOutlookVersion() {
            Microsoft.Office.Interop.Outlook.Application oApp = new Microsoft.Office.Interop.Outlook.Application();
            outlookVersionFull = oApp.Version;
        }
    }
}
