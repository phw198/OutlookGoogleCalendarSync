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
            this.icon.DoubleClick += notifyIcon_DoubleClick;
            this.icon.Visible = true;
            buildMenu();
        }

        private void buildMenu() {
            this.icon.ContextMenuStrip.Items.Clear();

            this.icon.ContextMenuStrip.Items.Add(toolStripMenuItemWithHandler("&Sync Now", "sync", syncItem_Click));

            ToolStripMenuItem cfg = toolStripMenuItemWithHandler("Sho&w", "show", showItem_Click);
            cfg.Font = new System.Drawing.Font(cfg.Font, System.Drawing.FontStyle.Bold);
            this.icon.ContextMenuStrip.Items.Add(cfg);

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
            //buildMenu();

            //menuBuilder.BuildContextMenu(LyncLyteTrayIcon.ContextMenuStrip);
            //LyncLyteTrayIcon.ContextMenuStrip.Items.Add(menuBuilder.ToolStripMenuItemWithHandler("Connect to Lync Client", "connect", findLync_Click));
            //LyncLyteTrayIcon.ContextMenuStrip.Items.Add(menuBuilder.ToolStripMenuItemWithHandler("&Find LyncLyte(s)", "find", findLyte_Click));

            //ToolStripMenuItem cfg = menuBuilder.ToolStripMenuItemWithHandler("&Configure LyncLyte", "config", showConfig_Click);
            //cfg.Font = new System.Drawing.Font(cfg.Font, System.Drawing.FontStyle.Bold);
            //LyncLyteTrayIcon.ContextMenuStrip.Items.Add(cfg);

            //ToolStripMenuItem manualColours = new ToolStripMenuItem("&Set Colour");
            //manualColours.DropDownDirection = (ToolStripDropDownDirection.Left);
            //manualColours.DropDown.Items.AddRange(new ToolStripItem[] {
            //    menuBuilder.ToolStripMenuItemWithHandler("Available", "Available", setColourAvailable_Click),
            //    menuBuilder.ToolStripMenuItemWithHandler("Away", "Away", setColourAway_Click),
            //    menuBuilder.ToolStripMenuItemWithHandler("Busy", "Busy", setColourBusy_Click),
            //    menuBuilder.ToolStripMenuItemWithHandler("Do not disturb", "DND", setColourDND_Click),
            //    menuBuilder.ToolStripMenuItemWithHandler("Reset", setColourReset_Click)
            //});

            //LyncLyteTrayIcon.ContextMenuStrip.Items.AddRange(
            //    new ToolStripItem[] {
            //        manualColours
            //    });
        }

        private void syncItem_Click(object sender, EventArgs e) {
            MainForm.Instance.bSyncNow.PerformClick();
        }

        private void showItem_Click(object sender, EventArgs e) {
            MainForm.Instance.MainFormShow();
        }
        private void notifyIcon_DoubleClick(object sender, EventArgs e) { 
            MainForm.Instance.MainFormShow(); 
        }

        private void exitItem_Click(object sender, EventArgs e) {
            exitEventFired = true;
            MainForm.Instance.Close();
        }
        #endregion
    }
}
