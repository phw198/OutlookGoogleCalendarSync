using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using log4net;

namespace OutlookGoogleCalendarSync {
    public class NotificationTray {
        private static readonly ILog log = LogManager.GetLogger(typeof(NotificationTray));
        private NotifyIcon icon;
        public Object Tag {
            get { return icon.Tag; }
            set { icon.Tag = value; }
        }
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

            if (OutlookCalendar.OOMsecurityInfo) {
                ShowBubbleInfo("Your Outlook security settings may not be optimal.\r\n" +
                    "Click here for further details.", ToolTipIcon.Warning, "OOMsecurity");
            }
        }

        private void buildMenu() {
            this.icon.ContextMenuStrip.Items.Clear();

            ToolStripMenuItem cfg = toolStripMenuItemWithHandler("&Sync Now", "sync", syncItem_Click);
            cfg.Font = new System.Drawing.Font(cfg.Font, System.Drawing.FontStyle.Bold);
            this.icon.ContextMenuStrip.Items.Add(cfg);

            cfg = toolStripMenuItemWithHandler("&Auto Sync", "autoSync", null);
            cfg.DropDown.Items.AddRange(new ToolStripItem[] {
                toolStripMenuItemWithHandler("Enable", "autoSyncToggle", autoSyncToggle_Click),
                toolStripMenuItemWithHandler("Delay for 1 hour", "delay1hr", delaySync1Hr_Click),
                toolStripMenuItemWithHandler("Delay for 2 hours", "delay2hr", delaySync2Hr_Click),
                toolStripMenuItemWithHandler("Delay for 4 hours", "delay4hr", delaySync4Hr_Click),
                toolStripMenuItemWithHandler("Remove delay", "delayRemove", delaySyncRemove_Click)
            });
            this.icon.ContextMenuStrip.Items.Add(cfg);
            this.icon.ContextMenuStrip.Items.Add(new ToolStripSeparator());
            
            this.icon.ContextMenuStrip.Items.Add(toolStripMenuItemWithHandler("Sho&w", "show", showItem_Click));

            this.icon.ContextMenuStrip.Items.Add(new ToolStripSeparator());
            this.icon.ContextMenuStrip.Items.Add(toolStripMenuItemWithHandler("&Exit", "exit", ExitItem_Click));

            UpdateAutoSyncItems();
            UpdateItem("delayRemove", enabled: false);
        }

        private ToolStripMenuItem toolStripMenuItemWithHandler(
            string displayText, string name, EventHandler eventHandler
        ) {
            var item = new ToolStripMenuItem(displayText);
            if (eventHandler != null) { item.Click += eventHandler; }
            switch (name) {
                case "sync": item.Image = Properties.Resources.sync; break;
                case "autoSync": item.Image = Properties.Resources.delay; break;
                case "show": item.Image = Properties.Resources.cog; break;
                case "exit": item.Image = Properties.Resources.exit; break;
            }
            item.Name = name;
            return item;
        }

        public void UpdateItem(String itemName, String itemText = null, Boolean enabled = true) {
            ToolStripItem[] items = this.icon.ContextMenuStrip.Items.Find(itemName, true);
            if (items.Count() > 0) {
                ToolStripItem item = items.First();
                item.Text = itemText ?? item.Text;
                item.Enabled = enabled;
            } else {
                log.Warn("Could not find menu item with name \"" + itemName + "\"");
            }
            return;
        }

        public void UpdateAutoSyncItems() {
            Boolean autoSyncing = (MainForm.Instance.OgcsTimer == null 
                ? Settings.Instance.SyncInterval != 0 || Settings.Instance.OutlookPush
                : MainForm.Instance.OgcsTimer.Running());

            UpdateItem("autoSyncToggle", autoSyncing ? "Disable" : "Enable");
            UpdateItem("delay1hr", null, autoSyncing);
            UpdateItem("delay2hr", null, autoSyncing);
            UpdateItem("delay4hr", null, autoSyncing);
        }

        #region Events
        private void ContextMenuStrip_Opening(object sender, System.ComponentModel.CancelEventArgs e) {
            e.Cancel = false;
            this.icon.ContextMenuStrip.Show();
        }

        private void syncItem_Click(object sender, EventArgs e) {
            MainForm.Instance.Sync_Requested();
        }

        private void autoSyncToggle_Click(object sender, EventArgs e) {
            String menuItemText = (sender as ToolStripMenuItem).Text;
            MainForm.Instance.Logboxout("Automatic sync "+ menuItemText.ToLower() +"d.");
            if (menuItemText == "Enable") {
                if (Settings.Instance.SyncInterval == 0) {
                    log.Debug("Switching on automatic syncing - hourly.");
                    MainForm.Instance.cbIntervalUnit.SelectedItem = "Hours";
                    MainForm.Instance.tbInterval.Value = 1;
                    XMLManager.ExportElement("SyncInterval", 1, Program.SettingsFile);
                    XMLManager.ExportElement("SyncIntervalUnit", "Hours", Program.SettingsFile);
                }
                if (MainForm.Instance.OgcsTimer == null) MainForm.Instance.OgcsTimer = new SyncTimer();
                MainForm.Instance.OgcsTimer.Switch(true);
                MainForm.Instance.lNextSyncVal.Font = new System.Drawing.Font(MainForm.Instance.lNextSyncVal.Font, System.Drawing.FontStyle.Regular);
                if (Settings.Instance.OutlookPush) OutlookCalendar.Instance.RegisterForPushSync();
                UpdateAutoSyncItems();
            } else {
                if (MainForm.Instance.OgcsTimer == null) {
                    log.Warn("Auto sync timer not initialised.");
                    return;
                }
                MainForm.Instance.OgcsTimer.Switch(false);
                MainForm.Instance.lNextSyncVal.Font = new System.Drawing.Font(MainForm.Instance.lNextSyncVal.Font, System.Drawing.FontStyle.Strikeout);
                if (Settings.Instance.OutlookPush) OutlookCalendar.Instance.DeregisterForPushSync();
                UpdateAutoSyncItems();
            }
        }
        private void delaySync1Hr_Click(object sender, EventArgs e) {
            MainForm.Instance.Logboxout("Next sync delayed for 1 hour.");
            if (MainForm.Instance.OgcsTimer == null) {
                log.Warn("Auto sync timer not initialised.");
                return;
            }
            MainForm.Instance.OgcsTimer.SetNextSync(60, fromNow: true);
            OutlookCalendar.Instance.DeregisterForPushSync();
            UpdateItem("delayRemove", enabled: true);
        }
        private void delaySync2Hr_Click(object sender, EventArgs e) {
            MainForm.Instance.Logboxout("Next sync delayed for 2 hours.");
            if (MainForm.Instance.OgcsTimer == null) {
                log.Warn("Auto sync timer not initialised.");
                return;
            } 
            MainForm.Instance.OgcsTimer.SetNextSync(2 * 60, fromNow: true);
            OutlookCalendar.Instance.DeregisterForPushSync();
            UpdateItem("delayRemove", enabled: true);
        }
        private void delaySync4Hr_Click(object sender, EventArgs e) {
            MainForm.Instance.Logboxout("Next sync delayed for 4 hours.");
            if (MainForm.Instance.OgcsTimer == null) {
                log.Warn("Auto sync timer not initialised.");
                return;
            }
            MainForm.Instance.OgcsTimer.SetNextSync(4 * 60, fromNow: true);
            OutlookCalendar.Instance.DeregisterForPushSync();
            UpdateItem("delayRemove", enabled: true);
        }
        private void delaySyncRemove_Click(object sender, EventArgs e) {
            MainForm.Instance.Logboxout("Next sync delay removed.");
            if (MainForm.Instance.OgcsTimer == null) {
                log.Warn("Auto sync timer not initialised.");
                return;
            }
            MainForm.Instance.OgcsTimer.SetNextSync();
            if (Settings.Instance.OutlookPush) OutlookCalendar.Instance.RegisterForPushSync();
            UpdateItem("delayRemove", enabled: false);
        }

        private void showItem_Click(object sender, EventArgs e) {
            MainForm.Instance.MainFormShow();
        }
        
        private void notifyIcon_Click(object sender, MouseEventArgs e) { 
            if (e.Button == MouseButtons.Left)
                MainForm.Instance.MainFormShow(); 
        }
        private void notifyIcon_DoubleClick(object sender, MouseEventArgs e) {
            if (e.Button == MouseButtons.Left && !MainForm.Instance.SyncingNow)
                MainForm.Instance.Sync_Requested();
        }

        private void notifyIcon_BubbleClick(object sender, EventArgs e) {
            NotifyIcon notifyIcon = (sender as NotifyIcon);
            if (notifyIcon.Tag != null && notifyIcon.Tag.ToString() == "ShowBubbleWhenMinimising") {
                Settings.Instance.ShowBubbleWhenMinimising = false;
                XMLManager.ExportElement("ShowBubbleWhenMinimising", false, Program.SettingsFile);
                notifyIcon.Tag = "";

            } else if (notifyIcon.Tag != null && notifyIcon.Tag.ToString() == "OOMsecurity") {
                System.Diagnostics.Process.Start("https://github.com/phw198/OutlookGoogleCalendarSync/wiki/FAQs---Outlook-Security");
                notifyIcon.Tag = "";

            } else {
                MainForm.Instance.MainFormShow();
                MainForm.Instance.tabApp.SelectedTab = MainForm.Instance.tabPage_Sync;
            }
        }

        public void ExitItem_Click(object sender, EventArgs e) {
            exitEventFired = true;
            MainForm.Instance.Close();
        }
        #endregion

        public void ShowBubbleInfo(string message, ToolTipIcon iconType = ToolTipIcon.Info, String tagValue = "") {
            if (Settings.Instance.ShowBubbleTooltipWhenSyncing) {
                this.icon.ShowBalloonTip(
                    500,
                    "Outlook Google Calendar Sync",
                    message,
                    iconType
                );
            }
            this.icon.Tag = tagValue;
        }
    }
}
