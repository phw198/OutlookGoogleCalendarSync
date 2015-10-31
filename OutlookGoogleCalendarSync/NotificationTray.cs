using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace OutlookGoogleCalendarSync {
    class NotificationTray {
        private NotifyIcon icon;
        private Boolean exitEventFired = false;
        public Boolean Exited {
            get { return this.exitEventFired; }
        }

        public NotificationTray(NotifyIcon icon) {
            this.icon = icon;
            this.icon.ContextMenuStrip = new ContextMenuStrip();
            this.icon.ContextMenuStrip.Opening += ContextMenuStrip_Opening;
            this.icon.MouseClick += notifyIcon_Click;
            this.icon.MouseDoubleClick += notifyIcon_DoubleClick;
            this.icon.BalloonTipClicked += notifyIcon_BubbleClick;
            this.icon.Visible = true;
            buildMenu();
        }

        private void buildMenu() {
            this.icon.ContextMenuStrip.Items.Clear();

            ToolStripMenuItem cfg = toolStripMenuItemWithHandler("&Sync Now", "sync", syncItem_Click);
            cfg.Font = new System.Drawing.Font(cfg.Font, System.Drawing.FontStyle.Bold);
            this.icon.ContextMenuStrip.Items.Add(cfg);
            
            this.icon.ContextMenuStrip.Items.Add(toolStripMenuItemWithHandler("Sho&w", "show", showItem_Click));

            this.icon.ContextMenuStrip.Items.Add(new ToolStripSeparator());
            this.icon.ContextMenuStrip.Items.Add(toolStripMenuItemWithHandler("&Exit", "exit", exitItem_Click));
        }

        private ToolStripMenuItem toolStripMenuItemWithHandler(
            string displayText, string status, EventHandler eventHandler
        ) {
            var item = new ToolStripMenuItem(displayText);
            if (eventHandler != null) { item.Click += eventHandler; }
            switch (status) {
                case "sync": item.Image = Properties.Resources.sync; break;
                case "show": item.Image = Properties.Resources.cog; break;
                case "exit": item.Image = Properties.Resources.exit; break;
            }
            item.Name = status;
            return item;
        }

        public void UpdateItem(String itemName, String itemText) {
            foreach (ToolStripItem item in this.icon.ContextMenuStrip.Items) {
                if (item.Name == itemName) {
                    item.Text = itemText;
                    return;
                }
            }
        }

        #region Events
        private void ContextMenuStrip_Opening(object sender, System.ComponentModel.CancelEventArgs e) {
            e.Cancel = false;
            this.icon.ContextMenuStrip.Show();
        }

        private void syncItem_Click(object sender, EventArgs e) {
            MainForm.Instance.Sync_Requested();
        }

        private void showItem_Click(object sender, EventArgs e) {
            MainForm.Instance.MainFormShow();
        }
        
        private void notifyIcon_Click(object sender, MouseEventArgs e) { 
            if (e.Button == MouseButtons.Left)
                MainForm.Instance.MainFormShow(); 
        }
        private void notifyIcon_DoubleClick(object sender, MouseEventArgs e) {
            if (e.Button == MouseButtons.Left)
                MainForm.Instance.Sync_Requested();
        }

        private void notifyIcon_BubbleClick(object sender, EventArgs e) {
            NotifyIcon notifyIcon = (sender as NotifyIcon);
            if (notifyIcon.Tag != null && notifyIcon.Tag.ToString() == "ShowBubbleWhenMinimising") {
                Settings.Instance.ShowBubbleWhenMinimising = false;
                XMLManager.ExportElement("ShowBubbleWhenMinimising", false, Program.SettingsFile);
                notifyIcon.Tag = "";
            } else {
                MainForm.Instance.MainFormShow();
                MainForm.Instance.tabApp.SelectedTab = MainForm.Instance.tabPage_Sync;
            }
        }

        private void exitItem_Click(object sender, EventArgs e) {
            exitEventFired = true;
            MainForm.Instance.Close();
        }
        #endregion
    }
}
