using log4net;
using System;
using System.Linq;
using System.Windows.Forms;

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
            this.icon.Icon = Forms.Main.Instance.Icon;
            this.icon.Visible = true;
            this.icon.Text += (string.IsNullOrEmpty(Program.Title) ? "" : " - " + Program.Title);
            buildMenu();

            if (OutlookOgcs.Calendar.OOMsecurityInfo) {
                ShowBubbleInfo("Your Outlook security settings may not be optimal.\r\n" +
                    "Click here for further details.", ToolTipIcon.Warning, "OOMsecurity");
                Telemetry.Send(Analytics.Category.ogcs, Analytics.Action.setting, "OOMsecurity;SyncCount=" + Settings.Instance.CompletedSyncs);
            }
        }

        private void buildMenu() {
            this.icon.ContextMenuStrip.Items.Clear();

            ToolStripMenuItem cfg = toolStripMenuItemWithHandler("&Sync Now", "sync", null);
            Settings.Instance.Calendars.ForEach(cal => cfg.DropDown.Items.Add(toolStripMenuItemWithHandler(cal._ProfileName, "sync", syncItem_Click)));
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

            cfg.Font = new System.Drawing.Font(cfg.Font, System.Drawing.FontStyle.Bold);
            this.icon.ContextMenuStrip.Items.Add(toolStripMenuItemWithHandler("Sho&w", "show", showItem_Click));
            cfg.Font = new System.Drawing.Font(cfg.Font, System.Drawing.FontStyle.Regular);

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
            try {
                ToolStripItem[] items = this.icon.ContextMenuStrip.Items.Find(itemName, true);
                if (items.Count() > 0) {
                    ToolStripItem item = items.First();
                    item.Text = itemText ?? item.Text;
                    item.Enabled = enabled;
                } else {
                    log.Warn("Could not find menu item with name \"" + itemName + "\"");
                }
            } catch (System.Exception ex) {
                if (Forms.Main.Instance.IsDisposed) return;
                OGCSexception.Analyse(ex, true);
            }
        }

        public void AddProfileItem(String itemText) {
            try {
                ToolStripItem[] items = this.icon.ContextMenuStrip.Items.Find("sync", false);
                if (items.Count() > 0) {
                    ToolStripItem item = items.First();
                    if (item is ToolStripMenuItem) {
                        ToolStripMenuItem rootMenu = item as ToolStripMenuItem;
                        items = rootMenu.DropDown.Items.Find("sync", false);
                        if (items.Count(i => i.Text == itemText) > 0)
                            log.Warn("There already exists a menu item with the name: " + itemText);
                        else
                            rootMenu.DropDown.Items.Add(toolStripMenuItemWithHandler(itemText, "sync", syncItem_Click));
                    } else
                        log.Error("'Sync Now' item found does not contain a menu");
                } else
                    log.Error("Could not find root 'sync' item");
                    
            } catch (System.Exception ex) {
                if (Forms.Main.Instance.IsDisposed) return;
                OGCSexception.Analyse(ex, true);
            }
        }
        public void RenameProfileItem(String currentText, String newText) {
            try {
                ToolStripItem[] items = this.icon.ContextMenuStrip.Items.Find("sync", false);
                if (items.Count() > 0) {
                    ToolStripItem item = items.First();
                    if (item is ToolStripMenuItem) {
                        ToolStripMenuItem rootMenu = item as ToolStripMenuItem;
                        items = rootMenu.DropDown.Items.Find("sync", false);
                        items.ToList().Where(i => i.Text == currentText).ToList().ForEach(j => j.Text = newText);
                    } else
                        log.Error("'Sync Now' item found does not contain a menu");
                } else
                    log.Error("Could not find root 'sync' item");

            } catch (System.Exception ex) {
                if (Forms.Main.Instance.IsDisposed) return;
                OGCSexception.Analyse(ex, true);
            }
        }
        public void RemoveProfileItem(String itemText) {
            try {
                ToolStripItem[] items = this.icon.ContextMenuStrip.Items.Find("sync", false);
                if (items.Count() > 0) {
                    ToolStripItem item = items.First();
                    if (item is ToolStripMenuItem) {
                        ToolStripMenuItem rootMenu = item as ToolStripMenuItem;
                        items = rootMenu.DropDown.Items.Find("sync", false);
                        items.ToList().Where(i => i.Text == itemText).ToList().ForEach(j => rootMenu.DropDownItems.Remove(j));
                    } else
                        log.Error("'Sync Now' item found does not contain a menu");
                } else
                    log.Error("Could not find root 'sync' item");

            } catch (System.Exception ex) {
                if (Forms.Main.Instance.IsDisposed) return;
                OGCSexception.Analyse(ex, true);
            }
        }

        public void UpdateAutoSyncItems() {
            Boolean autoSyncing = Settings.Instance.Calendars.Any(c =>
                (c.OgcsTimer != null && c.OgcsTimer.Running()) ||
                (c.OgcsTimer == null && (c.SyncInterval != 0 || c.OutlookPush)));

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
            String menuItemText = (sender as ToolStripMenuItem).Text;
            SettingsStore.Calendar profile = Settings.Instance.Calendars.First(cal => cal._ProfileName == menuItemText);
            if (profile != null) {
                Sync.Engine.Instance.JobQueue.Add(new Sync.Engine.Job("NotificationTray", profile));
            } else {
                log.Error("Unable to find a profile by the name: " + menuItemText);
            }
        }

        private void autoSyncToggle_Click(object sender, EventArgs e) {
            String menuItemText = (sender as ToolStripMenuItem).Text;
            Forms.Main.Instance.Console.Update("Automatic sync(s) " + menuItemText.ToLower() + "d.");
            if (menuItemText == "Enable") {
                int cnt = 0;
                foreach (SettingsStore.Calendar cal in Settings.Instance.Calendars) {
                    if (cal.SyncInterval != 0) {
                        log.Info("Enabled sync for profile: " + cal._ProfileName);
                        cal.OgcsTimer.SetNextSync(1 + (3 * cnt), true);
                    }
                    if (cal.OutlookPush) cal.RegisterForPushSync();
                }
                Forms.Main.Instance.StrikeOutNextSyncVal(false);
                UpdateAutoSyncItems();

            } else {
                foreach (SettingsStore.Calendar cal in Settings.Instance.Calendars) {
                    log.Info("Disabled sync for profile: " + cal._ProfileName);
                    if (cal.OgcsTimer == null) {
                        log.Warn("Auto sync timer not initialised.");
                        continue;
                    }
                    cal.OgcsTimer.Activate(false);
                    if (cal.OutlookPush) cal.DeregisterForPushSync();
                }
                Forms.Main.Instance.StrikeOutNextSyncVal(true);
                UpdateAutoSyncItems();
            }
        }
        private void delaySync1Hr_Click(object sender, EventArgs e) {
            Forms.Main.Instance.Console.Update("Next sync delayed for 1 hour.");
            foreach (SettingsStore.Calendar cal in Settings.Instance.Calendars) {
                if (cal.OgcsTimer == null) continue;
                log.Info("Delaying sync for 1 hour: " + cal._ProfileName);
                cal.OgcsTimer.SetNextSync(60, fromNow: true);
                cal.DeregisterForPushSync();
            }
            UpdateItem("delayRemove", enabled: true);
        }
        private void delaySync2Hr_Click(object sender, EventArgs e) {
            Forms.Main.Instance.Console.Update("Next sync delayed for 2 hours.");
            foreach (SettingsStore.Calendar cal in Settings.Instance.Calendars) {
                if (cal.OgcsTimer == null) continue;
                log.Info("Delaying sync for 2 hours: " + cal._ProfileName);
                cal.OgcsTimer.SetNextSync(2 * 60, fromNow: true);
                cal.DeregisterForPushSync();
            }
            UpdateItem("delayRemove", enabled: true);
        }
        private void delaySync4Hr_Click(object sender, EventArgs e) {
            Forms.Main.Instance.Console.Update("Next sync delayed for 4 hours.");
            foreach (SettingsStore.Calendar cal in Settings.Instance.Calendars) {
                if (cal.OgcsTimer == null) continue;
                log.Info("Delaying sync for 4 hours: " + cal._ProfileName);
                cal.OgcsTimer.SetNextSync(4 * 60, fromNow: true);
                cal.DeregisterForPushSync();
            }
            UpdateItem("delayRemove", enabled: true);
        }
        private void delaySyncRemove_Click(object sender, EventArgs e) {
            Forms.Main.Instance.Console.Update("Next sync delay removed.");
            foreach (SettingsStore.Calendar cal in Settings.Instance.Calendars) {
                if (cal.OgcsTimer == null) continue;
                log.Info("Removing sync delay: " + cal._ProfileName);
                cal.OgcsTimer.SetNextSync();
                cal.RegisterForPushSync();
            }
            UpdateItem("delayRemove", enabled: false);
        }

        private void showItem_Click(object sender, EventArgs e) {
            Forms.Main.Instance.MainFormShow();
        }

        private void notifyIcon_Click(object sender, MouseEventArgs e) {
            if (e.Button == MouseButtons.Left) {
                Forms.Main.Instance.MainFormShow();
            }
        }
        private void notifyIcon_DoubleClick(object sender, MouseEventArgs e) {
            if (e.Button == MouseButtons.Left)
                Forms.Main.Instance.MainFormShow();
        }

        private void notifyIcon_BubbleClick(object sender, EventArgs e) {
            NotifyIcon notifyIcon = (sender as NotifyIcon);
            if (notifyIcon.Tag != null && notifyIcon.Tag.ToString() == "ShowBubbleWhenMinimising") {
                Settings.Instance.ShowBubbleWhenMinimising = false;
                XMLManager.ExportElement(Settings.Instance, "ShowBubbleWhenMinimising", false, Settings.ConfigFile);
                notifyIcon.Tag = "";

            } else if (notifyIcon.Tag != null && notifyIcon.Tag.ToString() == "OOMsecurity") {
                Helper.OpenBrowser("https://github.com/phw198/OutlookGoogleCalendarSync/wiki/FAQs---Outlook-Security");
                notifyIcon.Tag = "";

            } else {
                Forms.Main.Instance.TopMost = true;
                Forms.Main.Instance.MainFormShow();
                Forms.Main.Instance.TopMost = false;
                Forms.Main.Instance.tabApp.SelectedTab = Forms.Main.Instance.tabPage_Sync;
            }
        }

        public void ExitItem_Click(object sender, EventArgs e) {
            exitEventFired = true;
            Forms.Main.Instance.Close();
        }
        #endregion

        public void ShowBubbleInfo(string message, ToolTipIcon iconType = ToolTipIcon.None, String tagValue = "") {
            if (Settings.Instance.ShowBubbleTooltipWhenSyncing) {
                this.icon.ShowBalloonTip(
                    500,
                    "Outlook Google Calendar Sync" + (string.IsNullOrEmpty(Program.Title) ? "" : " - " + Program.Title),
                    message,
                    iconType
                );
            }
            this.icon.Tag = tagValue;
        }
    }
}
