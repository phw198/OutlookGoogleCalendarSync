using System;
using System.Windows.Forms;

namespace OutlookGoogleCalendarSync {
    /// <summary>
    /// Description of GoogleAuthorizationCode.
    /// </summary>
    public partial class frmGoogleAuthorizationCode : Form {
        public string authcode = "";

        public frmGoogleAuthorizationCode() {
            //
            // The InitializeComponent() call is required for Windows Forms designer support.
            //
            InitializeComponent();
        }

        private void btOK_Click(object sender, EventArgs e) {
            authcode = tbCode.Text;
        }
    }
}
