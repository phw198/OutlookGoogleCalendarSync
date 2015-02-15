using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace OutlookGoogleCalendarSync {
    public partial class Splash : Form {
        public Splash() {
            InitializeComponent();
            lVersion.Text = "v" + Application.ProductVersion;
        }

        private void pbDonate_Click(object sender, EventArgs e) {
            System.Diagnostics.Process.Start("https://www.paypal.com/cgi-bin/webscr?cmd=_s-xclick&hosted_button_id=RT46CXQDSSYWJ");
            this.Close();
        }
    }
}
