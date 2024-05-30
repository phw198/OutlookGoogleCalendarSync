using Ogcs = OutlookGoogleCalendarSync;
using System.Windows.Forms;

namespace OutlookGoogleCalendarSync.Extensions {
    public partial class OgcsNumericUpDown : NumericUpDown {

        public ToolTip tooltip = new ToolTip();

        public override string Text {
            get { return base.Text; }
            set { base.Text = value; }
        }

        public override void UpButton() {
            base.UpButton();
            checkLimit();
        }

        public override void DownButton() {
            base.DownButton();
            checkLimit();
        }

        private void checkLimit() {
            try {
                if (base.ParentForm.Name != Forms.Main.Instance.Name) return;
                
                if (Settings.Instance.UsingPersonalAPIkeys()) {
                    if (!string.IsNullOrEmpty(this.tooltip.GetToolTip(this)))
                        this.tooltip.RemoveAll();
                } else {
                    if (this.Value == this.Maximum)
                        this.tooltip.Show("Limited to 1 year unless personal API keys are used. See 'Developer Options' on Google tab.", this);
                    else
                        this.tooltip.RemoveAll();
                }
            } catch (System.Exception ex) {
                ex.Analyse(this.Name);
            }
        }
    }
}
