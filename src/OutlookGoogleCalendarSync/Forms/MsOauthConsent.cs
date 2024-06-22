using log4net;
using System;
using System.Windows.Forms;

namespace OutlookGoogleCalendarSync.Forms {
    public partial class MsOauthConsent : Form {
        private static readonly ILog log = LogManager.GetLogger(typeof(MsOauthConsent));
        public MsOauthConsent() {
            InitializeComponent();
        }

        private void rbBlocked_CheckedChanged(object sender, EventArgs e) {
            this.rbAdminGrant.Enabled = this.rbBlocked.Checked;
            this.rbJustificationGiven.Enabled = this.rbBlocked.Checked;
            this.rbEndOfRoad.Enabled = this.rbBlocked.Checked;

            if (this.rbBlocked.Checked)
                this.btOK.Enabled = false;
            else {
                this.rbAdminGrant.Checked = false;
                this.rbJustificationGiven.Checked = false;
                this.rbEndOfRoad.Checked = false;
            }
        }

        private void btOK_Click(object sender, EventArgs e) {
            Control optionSelected =
                this.rbDoubts.Checked ? this.rbDoubts :
                this.rbDunno.Checked ? this.rbDunno :
                this.rbBlocked.Checked ? 
                    (this.rbAdminGrant.Checked ? this.rbAdminGrant :
                    this.rbJustificationGiven.Checked ? this.rbJustificationGiven :
                    this.rbEndOfRoad.Checked ? this.rbEndOfRoad : null) :
                null;
            log.Info("Picked " + optionSelected.Name);
            new Telemetry.GA4Event.Event(Telemetry.GA4Event.Event.Name.debug)
                .AddParameter("MsOauthConsentCancelled", optionSelected?.Name)
                .Send();
        }

        private void rbDoubts_Click(object sender, EventArgs e) {
            this.btOK.Enabled = true;
        }

        private void rbJustificationGiven_Click(object sender, EventArgs e) {
            this.btOK.Enabled = true;
        }

        private void rbAdminGrant_Click(object sender, EventArgs e) {
            this.btOK.Enabled = true;
        }

        private void rbEndOfRoad_Click(object sender, EventArgs e) {
            this.btOK.Enabled = true;
        }

        private void rbDunno_Click(object sender, EventArgs e) {
            this.btOK.Enabled = true;
        }
    }
}
