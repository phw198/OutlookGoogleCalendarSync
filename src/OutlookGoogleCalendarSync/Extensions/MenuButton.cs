using System;
using System.ComponentModel;
using System.Drawing;
using System.Windows.Forms;

namespace OutlookGoogleCalendarSync.Extensions {
    public partial class MenuButton : Button {
        [DefaultValue(null), Browsable(true), DesignerSerializationVisibility(DesignerSerializationVisibility.Visible)]
        public ContextMenuStrip Menu { get; set; }

        [DefaultValue(20), Browsable(true), DesignerSerializationVisibility(DesignerSerializationVisibility.Visible)]
        public int SplitWidth { get; set; }

        [DefaultValue(true)]
        public Boolean MenuEnabled { get; set; }

        public MenuButton() {
            SplitWidth = 20;
            MenuEnabled = true;
        }

        protected override void OnMouseDown(MouseEventArgs mevent) {
            Rectangle splitRect = new Rectangle(this.Width - this.SplitWidth, 0, this.SplitWidth, this.Height);
            this.Tag = "";

            // Figure out if the button click was on the button itself or the menu split
            if (Menu != null && mevent.Button == MouseButtons.Left && splitRect.Contains(mevent.Location)) {
                if (!this.MenuEnabled) return;

                if (this.Menu.Visible) {
                    this.Tag = "CloseRequested";
                    this.Menu.Hide();
                } else {
                    Menu.Show(this, 0, this.Height);    // Shows menu under button
                    //Menu.Show(this, mevent.Location); // Shows menu at click location
                }
            } else {
                base.OnMouseDown(mevent);
            }
        }

        protected override void OnPaint(PaintEventArgs pEvent) {
            base.OnPaint(pEvent);

            if (this.Menu != null && this.SplitWidth > 0) {
                // Draw the arrow glyph on the right side of the button
                int arrowX = ClientRectangle.Width - (int)(SplitWidth - ((SplitWidth - 7)/2));
                int arrowY = ClientRectangle.Height / 2 - 1;

                var arrowBrush = (Enabled && MenuEnabled) ? SystemBrushes.ControlText : SystemBrushes.ButtonShadow;
                var arrows = new[] { new Point(arrowX, arrowY), new Point(arrowX + 7, arrowY), new Point(arrowX + 3, arrowY + 4) };
                pEvent.Graphics.FillPolygon(arrowBrush, arrows);

                // Draw a dashed separator on the left of the arrow
                int lineX = ClientRectangle.Width - this.SplitWidth;
                int lineYFrom = 6;
                int lineYTo = ClientRectangle.Height - 6;
                using (var separatorPen = new Pen(Brushes.DarkGray) { 
                    DashStyle = System.Drawing.Drawing2D.DashStyle.Dot 
                }) {
                    pEvent.Graphics.DrawLine(separatorPen, lineX, lineYFrom, lineX, lineYTo);
                }
            }
        }
    }

    public partial class ButtonContextMenuStrip : ContextMenuStrip {

        public ButtonContextMenuStrip(IContainer container) { }

        protected override void OnClosing(ToolStripDropDownClosingEventArgs e) {
            if (this.SourceControl is MenuButton) {
                MenuButton button = this.SourceControl as MenuButton;
                Rectangle dropdownRect = new Rectangle(button.Width - button.SplitWidth, 0, button.SplitWidth, button.Height);
                Point relativeMousePosition = button.PointToClient(Cursor.Position);
                if (dropdownRect.Contains(relativeMousePosition) && string.IsNullOrEmpty(button.Tag.ToString()))
                    e.Cancel = true;
                button.Tag = "";
            }
        }
    }
}
