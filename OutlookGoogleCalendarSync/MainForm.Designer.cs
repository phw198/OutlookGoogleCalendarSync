namespace OutlookGoogleCalendarSync {
    partial class MainForm {
        /// <summary>
        /// Designer variable used to keep track of non-visual components.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Disposes resources used by the form.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing) {
            if (disposing) {
                if (components != null) {
                    components.Dispose();
                }
            }
            base.Dispose(disposing);
        }

        /// <summary>
        /// This method is required for Windows Forms designer support.
        /// Do not change the method contents inside the source code editor. The Forms designer might
        /// not be able to load this method if it was changed manually.
        /// </summary>
        private void InitializeComponent() {
            this.components = new System.ComponentModel.Container();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(MainForm));
            this.tabApp = new System.Windows.Forms.TabControl();
            this.tabPage_Sync = new System.Windows.Forms.TabPage();
            this.cbVerboseOutput = new System.Windows.Forms.CheckBox();
            this.lNextSyncVal = new System.Windows.Forms.Label();
            this.lLastSyncVal = new System.Windows.Forms.Label();
            this.lNextSync = new System.Windows.Forms.Label();
            this.lLastSync = new System.Windows.Forms.Label();
            this.LogBox = new System.Windows.Forms.TextBox();
            this.bSyncNow = new System.Windows.Forms.Button();
            this.tabPage_Settings = new System.Windows.Forms.TabPage();
            this.lSettingInfo = new System.Windows.Forms.Label();
            this.tabAppSettings = new System.Windows.Forms.TabControl();
            this.tabOutlook = new System.Windows.Forms.TabPage();
            this.label11 = new System.Windows.Forms.Label();
            this.lOutlookCalendar = new System.Windows.Forms.Label();
            this.cbOutlookCalendars = new System.Windows.Forms.ComboBox();
            this.rbOutlookDefaultMB = new System.Windows.Forms.RadioButton();
            this.rbOutlookEWS = new System.Windows.Forms.RadioButton();
            this.rbOutlookAltMB = new System.Windows.Forms.RadioButton();
            this.ddMailboxName = new System.Windows.Forms.ComboBox();
            this.gbEWS = new System.Windows.Forms.GroupBox();
            this.txtEWSServerURL = new System.Windows.Forms.TextBox();
            this.label8 = new System.Windows.Forms.Label();
            this.txtEWSPass = new System.Windows.Forms.TextBox();
            this.label7 = new System.Windows.Forms.Label();
            this.txtEWSUser = new System.Windows.Forms.TextBox();
            this.label6 = new System.Windows.Forms.Label();
            this.tabGoogle = new System.Windows.Forms.TabPage();
            this.label13 = new System.Windows.Forms.Label();
            this.label12 = new System.Windows.Forms.Label();
            this.btResetGCal = new System.Windows.Forms.Button();
            this.lGoogleHelp = new System.Windows.Forms.Label();
            this.lGoogleCalendar = new System.Windows.Forms.Label();
            this.bGetGoogleCalendars = new System.Windows.Forms.Button();
            this.cbGoogleCalendars = new System.Windows.Forms.ComboBox();
            this.tabSyncOptions = new System.Windows.Forms.TabPage();
            this.gbSyncOptions_When = new System.Windows.Forms.GroupBox();
            this.cbOutlookPush = new System.Windows.Forms.CheckBox();
            this.cbIntervalUnit = new System.Windows.Forms.ComboBox();
            this.tbInterval = new System.Windows.Forms.NumericUpDown();
            this.label1 = new System.Windows.Forms.Label();
            this.tbDaysInTheFuture = new System.Windows.Forms.NumericUpDown();
            this.tbDaysInThePast = new System.Windows.Forms.NumericUpDown();
            this.lDaysInFuture = new System.Windows.Forms.Label();
            this.lDaysInPast = new System.Windows.Forms.Label();
            this.lDateRange = new System.Windows.Forms.Label();
            this.gbSyncOptions_What = new System.Windows.Forms.GroupBox();
            this.cbAddReminders = new System.Windows.Forms.CheckBox();
            this.lAttributes = new System.Windows.Forms.Label();
            this.cbAddAttendees = new System.Windows.Forms.CheckBox();
            this.cbAddDescription = new System.Windows.Forms.CheckBox();
            this.label15 = new System.Windows.Forms.Label();
            this.gbSyncOptions_How = new System.Windows.Forms.GroupBox();
            this.syncDirection = new System.Windows.Forms.ComboBox();
            this.lDirection = new System.Windows.Forms.Label();
            this.cbMergeItems = new System.Windows.Forms.CheckBox();
            this.cbDisableDeletion = new System.Windows.Forms.CheckBox();
            this.cbConfirmOnDelete = new System.Windows.Forms.CheckBox();
            this.tabAppBehaviour = new System.Windows.Forms.TabPage();
            this.gbProxy = new System.Windows.Forms.GroupBox();
            this.rbProxyNone = new System.Windows.Forms.RadioButton();
            this.rbProxyIE = new System.Windows.Forms.RadioButton();
            this.rbProxyCustom = new System.Windows.Forms.RadioButton();
            this.txtProxyPassword = new System.Windows.Forms.TextBox();
            this.cbProxyAuthRequired = new System.Windows.Forms.CheckBox();
            this.label9 = new System.Windows.Forms.Label();
            this.txtProxyPort = new System.Windows.Forms.TextBox();
            this.txtProxyUser = new System.Windows.Forms.TextBox();
            this.txtProxyServer = new System.Windows.Forms.TextBox();
            this.label10 = new System.Windows.Forms.Label();
            this.label5 = new System.Windows.Forms.Label();
            this.label4 = new System.Windows.Forms.Label();
            this.label14 = new System.Windows.Forms.Label();
            this.btLogLocation = new System.Windows.Forms.Button();
            this.label3 = new System.Windows.Forms.Label();
            this.cbLoggingLevel = new System.Windows.Forms.ComboBox();
            this.cbStartOnStartup = new System.Windows.Forms.CheckBox();
            this.cbShowBubbleTooltips = new System.Windows.Forms.CheckBox();
            this.cbMinimizeToTray = new System.Windows.Forms.CheckBox();
            this.cbStartInTray = new System.Windows.Forms.CheckBox();
            this.cbCreateFiles = new System.Windows.Forms.CheckBox();
            this.bSave = new System.Windows.Forms.Button();
            this.tabPage_About = new System.Windows.Forms.TabPage();
            this.label2 = new System.Windows.Forms.Label();
            this.pbDonate = new System.Windows.Forms.PictureBox();
            this.lAboutURL = new System.Windows.Forms.LinkLabel();
            this.lAboutMain = new System.Windows.Forms.Label();
            this.notifyIcon1 = new System.Windows.Forms.NotifyIcon(this.components);
            this.tabApp.SuspendLayout();
            this.tabPage_Sync.SuspendLayout();
            this.tabPage_Settings.SuspendLayout();
            this.tabAppSettings.SuspendLayout();
            this.tabOutlook.SuspendLayout();
            this.gbEWS.SuspendLayout();
            this.tabGoogle.SuspendLayout();
            this.tabSyncOptions.SuspendLayout();
            this.gbSyncOptions_When.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.tbInterval)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.tbDaysInTheFuture)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.tbDaysInThePast)).BeginInit();
            this.gbSyncOptions_What.SuspendLayout();
            this.gbSyncOptions_How.SuspendLayout();
            this.tabAppBehaviour.SuspendLayout();
            this.gbProxy.SuspendLayout();
            this.tabPage_About.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.pbDonate)).BeginInit();
            this.SuspendLayout();
            // 
            // tabApp
            // 
            this.tabApp.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom)
                        | System.Windows.Forms.AnchorStyles.Left)
                        | System.Windows.Forms.AnchorStyles.Right)));
            this.tabApp.Controls.Add(this.tabPage_Sync);
            this.tabApp.Controls.Add(this.tabPage_Settings);
            this.tabApp.Controls.Add(this.tabPage_About);
            this.tabApp.Location = new System.Drawing.Point(12, 12);
            this.tabApp.Multiline = true;
            this.tabApp.Name = "tabApp";
            this.tabApp.SelectedIndex = 0;
            this.tabApp.Size = new System.Drawing.Size(495, 568);
            this.tabApp.TabIndex = 0;
            // 
            // tabPage_Sync
            // 
            this.tabPage_Sync.Controls.Add(this.cbVerboseOutput);
            this.tabPage_Sync.Controls.Add(this.lNextSyncVal);
            this.tabPage_Sync.Controls.Add(this.lLastSyncVal);
            this.tabPage_Sync.Controls.Add(this.lNextSync);
            this.tabPage_Sync.Controls.Add(this.lLastSync);
            this.tabPage_Sync.Controls.Add(this.LogBox);
            this.tabPage_Sync.Controls.Add(this.bSyncNow);
            this.tabPage_Sync.Location = new System.Drawing.Point(4, 22);
            this.tabPage_Sync.Name = "tabPage_Sync";
            this.tabPage_Sync.Padding = new System.Windows.Forms.Padding(3);
            this.tabPage_Sync.Size = new System.Drawing.Size(487, 542);
            this.tabPage_Sync.TabIndex = 0;
            this.tabPage_Sync.Text = "Sync";
            this.tabPage_Sync.UseVisualStyleBackColor = true;
            // 
            // cbVerboseOutput
            // 
            this.cbVerboseOutput.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.cbVerboseOutput.AutoSize = true;
            this.cbVerboseOutput.Location = new System.Drawing.Point(383, 484);
            this.cbVerboseOutput.Name = "cbVerboseOutput";
            this.cbVerboseOutput.Size = new System.Drawing.Size(98, 17);
            this.cbVerboseOutput.TabIndex = 5;
            this.cbVerboseOutput.Text = "Verbose output";
            this.cbVerboseOutput.UseVisualStyleBackColor = true;
            this.cbVerboseOutput.CheckedChanged += new System.EventHandler(this.cbVerboseOutput_CheckedChanged);
            // 
            // lNextSyncVal
            // 
            this.lNextSyncVal.Location = new System.Drawing.Point(271, 28);
            this.lNextSyncVal.Name = "lNextSyncVal";
            this.lNextSyncVal.Size = new System.Drawing.Size(216, 26);
            this.lNextSyncVal.TabIndex = 4;
            this.lNextSyncVal.Text = "Unknown";
            // 
            // lLastSyncVal
            // 
            this.lLastSyncVal.Location = new System.Drawing.Point(22, 28);
            this.lLastSyncVal.Name = "lLastSyncVal";
            this.lLastSyncVal.Size = new System.Drawing.Size(224, 26);
            this.lLastSyncVal.TabIndex = 3;
            this.lLastSyncVal.Text = "N/A";
            // 
            // lNextSync
            // 
            this.lNextSync.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lNextSync.Location = new System.Drawing.Point(252, 14);
            this.lNextSync.Name = "lNextSync";
            this.lNextSync.Size = new System.Drawing.Size(232, 14);
            this.lNextSync.TabIndex = 2;
            this.lNextSync.Text = "Next scheduled:-";
            // 
            // lLastSync
            // 
            this.lLastSync.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lLastSync.Location = new System.Drawing.Point(5, 14);
            this.lLastSync.Name = "lLastSync";
            this.lLastSync.Size = new System.Drawing.Size(251, 14);
            this.lLastSync.TabIndex = 2;
            this.lLastSync.Text = "Last successful:-";
            // 
            // LogBox
            // 
            this.LogBox.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom)
                        | System.Windows.Forms.AnchorStyles.Left)
                        | System.Windows.Forms.AnchorStyles.Right)));
            this.LogBox.Font = new System.Drawing.Font("Courier New", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.LogBox.Location = new System.Drawing.Point(3, 57);
            this.LogBox.Multiline = true;
            this.LogBox.Name = "LogBox";
            this.LogBox.ScrollBars = System.Windows.Forms.ScrollBars.Both;
            this.LogBox.Size = new System.Drawing.Size(478, 421);
            this.LogBox.TabIndex = 1;
            // 
            // bSyncNow
            // 
            this.bSyncNow.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.bSyncNow.Location = new System.Drawing.Point(3, 484);
            this.bSyncNow.Name = "bSyncNow";
            this.bSyncNow.Size = new System.Drawing.Size(98, 31);
            this.bSyncNow.TabIndex = 0;
            this.bSyncNow.Tag = "0";
            this.bSyncNow.Text = "Start Sync";
            this.bSyncNow.UseVisualStyleBackColor = true;
            this.bSyncNow.Click += new System.EventHandler(this.sync_Click);
            // 
            // tabPage_Settings
            // 
            this.tabPage_Settings.Controls.Add(this.lSettingInfo);
            this.tabPage_Settings.Controls.Add(this.tabAppSettings);
            this.tabPage_Settings.Controls.Add(this.bSave);
            this.tabPage_Settings.Location = new System.Drawing.Point(4, 22);
            this.tabPage_Settings.Name = "tabPage_Settings";
            this.tabPage_Settings.Padding = new System.Windows.Forms.Padding(3);
            this.tabPage_Settings.Size = new System.Drawing.Size(487, 542);
            this.tabPage_Settings.TabIndex = 1;
            this.tabPage_Settings.Text = "Settings";
            this.tabPage_Settings.UseVisualStyleBackColor = true;
            // 
            // lSettingInfo
            // 
            this.lSettingInfo.Anchor = System.Windows.Forms.AnchorStyles.Bottom;
            this.lSettingInfo.AutoSize = true;
            this.lSettingInfo.ForeColor = System.Drawing.SystemColors.MenuHighlight;
            this.lSettingInfo.Location = new System.Drawing.Point(131, 495);
            this.lSettingInfo.Name = "lSettingInfo";
            this.lSettingInfo.Size = new System.Drawing.Size(176, 26);
            this.lSettingInfo.TabIndex = 27;
            this.lSettingInfo.Text = "Settings will take effect immediately,\r\nbut to make them persist, hit Save.";
            this.lSettingInfo.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // tabAppSettings
            // 
            this.tabAppSettings.Alignment = System.Windows.Forms.TabAlignment.Left;
            this.tabAppSettings.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom)
                        | System.Windows.Forms.AnchorStyles.Left)
                        | System.Windows.Forms.AnchorStyles.Right)));
            this.tabAppSettings.Controls.Add(this.tabOutlook);
            this.tabAppSettings.Controls.Add(this.tabGoogle);
            this.tabAppSettings.Controls.Add(this.tabSyncOptions);
            this.tabAppSettings.Controls.Add(this.tabAppBehaviour);
            this.tabAppSettings.DrawMode = System.Windows.Forms.TabDrawMode.OwnerDrawFixed;
            this.tabAppSettings.ItemSize = new System.Drawing.Size(35, 75);
            this.tabAppSettings.Location = new System.Drawing.Point(6, 7);
            this.tabAppSettings.Multiline = true;
            this.tabAppSettings.Name = "tabAppSettings";
            this.tabAppSettings.SelectedIndex = 0;
            this.tabAppSettings.Size = new System.Drawing.Size(475, 470);
            this.tabAppSettings.SizeMode = System.Windows.Forms.TabSizeMode.Fixed;
            this.tabAppSettings.TabIndex = 20;
            this.tabAppSettings.DrawItem += new System.Windows.Forms.DrawItemEventHandler(this.tabAppSettings_DrawItem);
            // 
            // tabOutlook
            // 
            this.tabOutlook.BackColor = System.Drawing.Color.White;
            this.tabOutlook.Controls.Add(this.label11);
            this.tabOutlook.Controls.Add(this.lOutlookCalendar);
            this.tabOutlook.Controls.Add(this.cbOutlookCalendars);
            this.tabOutlook.Controls.Add(this.rbOutlookDefaultMB);
            this.tabOutlook.Controls.Add(this.rbOutlookEWS);
            this.tabOutlook.Controls.Add(this.rbOutlookAltMB);
            this.tabOutlook.Controls.Add(this.ddMailboxName);
            this.tabOutlook.Controls.Add(this.gbEWS);
            this.tabOutlook.Location = new System.Drawing.Point(79, 4);
            this.tabOutlook.Name = "tabOutlook";
            this.tabOutlook.Padding = new System.Windows.Forms.Padding(3);
            this.tabOutlook.Size = new System.Drawing.Size(392, 462);
            this.tabOutlook.TabIndex = 0;
            this.tabOutlook.Text = "  Outlook";
            // 
            // label11
            // 
            this.label11.Anchor = System.Windows.Forms.AnchorStyles.Top;
            this.label11.AutoSize = true;
            this.label11.Font = new System.Drawing.Font("Arial Black", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label11.ForeColor = System.Drawing.SystemColors.MenuHighlight;
            this.label11.Location = new System.Drawing.Point(67, 13);
            this.label11.Name = "label11";
            this.label11.Size = new System.Drawing.Size(268, 15);
            this.label11.TabIndex = 26;
            this.label11.Text = "Select the Outlook Calendar to Synchronise";
            this.label11.TextAlign = System.Drawing.ContentAlignment.TopCenter;
            // 
            // lOutlookCalendar
            // 
            this.lOutlookCalendar.AutoSize = true;
            this.lOutlookCalendar.Location = new System.Drawing.Point(20, 203);
            this.lOutlookCalendar.Name = "lOutlookCalendar";
            this.lOutlookCalendar.Size = new System.Drawing.Size(81, 13);
            this.lOutlookCalendar.TabIndex = 25;
            this.lOutlookCalendar.Text = "Select calendar";
            // 
            // cbOutlookCalendars
            // 
            this.cbOutlookCalendars.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left)
                        | System.Windows.Forms.AnchorStyles.Right)));
            this.cbOutlookCalendars.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cbOutlookCalendars.FormattingEnabled = true;
            this.cbOutlookCalendars.Location = new System.Drawing.Point(107, 200);
            this.cbOutlookCalendars.Name = "cbOutlookCalendars";
            this.cbOutlookCalendars.Size = new System.Drawing.Size(277, 21);
            this.cbOutlookCalendars.TabIndex = 24;
            this.cbOutlookCalendars.SelectedIndexChanged += new System.EventHandler(this.cbOutlookCalendar_SelectedIndexChanged);
            // 
            // rbOutlookDefaultMB
            // 
            this.rbOutlookDefaultMB.AutoSize = true;
            this.rbOutlookDefaultMB.Checked = true;
            this.rbOutlookDefaultMB.Location = new System.Drawing.Point(20, 47);
            this.rbOutlookDefaultMB.Name = "rbOutlookDefaultMB";
            this.rbOutlookDefaultMB.Size = new System.Drawing.Size(98, 17);
            this.rbOutlookDefaultMB.TabIndex = 18;
            this.rbOutlookDefaultMB.TabStop = true;
            this.rbOutlookDefaultMB.Text = "Default Mailbox";
            this.rbOutlookDefaultMB.UseVisualStyleBackColor = true;
            this.rbOutlookDefaultMB.CheckedChanged += new System.EventHandler(this.rbOutlookDefaultMB_CheckedChanged);
            // 
            // rbOutlookEWS
            // 
            this.rbOutlookEWS.AutoSize = true;
            this.rbOutlookEWS.Enabled = false;
            this.rbOutlookEWS.Location = new System.Drawing.Point(20, 93);
            this.rbOutlookEWS.Name = "rbOutlookEWS";
            this.rbOutlookEWS.Size = new System.Drawing.Size(143, 17);
            this.rbOutlookEWS.TabIndex = 19;
            this.rbOutlookEWS.Text = "Exchange Web Services";
            this.rbOutlookEWS.UseVisualStyleBackColor = true;
            this.rbOutlookEWS.CheckedChanged += new System.EventHandler(this.rbOutlookEWS_CheckedChanged);
            // 
            // rbOutlookAltMB
            // 
            this.rbOutlookAltMB.AutoSize = true;
            this.rbOutlookAltMB.Location = new System.Drawing.Point(20, 70);
            this.rbOutlookAltMB.Name = "rbOutlookAltMB";
            this.rbOutlookAltMB.Size = new System.Drawing.Size(114, 17);
            this.rbOutlookAltMB.TabIndex = 17;
            this.rbOutlookAltMB.Text = "Alternative Mailbox";
            this.rbOutlookAltMB.UseVisualStyleBackColor = true;
            this.rbOutlookAltMB.CheckedChanged += new System.EventHandler(this.rbOutlookAltMB_CheckedChanged);
            // 
            // ddMailboxName
            // 
            this.ddMailboxName.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left)
                        | System.Windows.Forms.AnchorStyles.Right)));
            this.ddMailboxName.FormattingEnabled = true;
            this.ddMailboxName.Location = new System.Drawing.Point(134, 70);
            this.ddMailboxName.Name = "ddMailboxName";
            this.ddMailboxName.Size = new System.Drawing.Size(250, 21);
            this.ddMailboxName.TabIndex = 16;
            this.ddMailboxName.SelectedIndexChanged += new System.EventHandler(this.ddMailboxName_SelectedIndexChanged);
            // 
            // gbEWS
            // 
            this.gbEWS.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left)
                        | System.Windows.Forms.AnchorStyles.Right)));
            this.gbEWS.Controls.Add(this.txtEWSServerURL);
            this.gbEWS.Controls.Add(this.label8);
            this.gbEWS.Controls.Add(this.txtEWSPass);
            this.gbEWS.Controls.Add(this.label7);
            this.gbEWS.Controls.Add(this.txtEWSUser);
            this.gbEWS.Controls.Add(this.label6);
            this.gbEWS.Location = new System.Drawing.Point(49, 97);
            this.gbEWS.Name = "gbEWS";
            this.gbEWS.Size = new System.Drawing.Size(335, 96);
            this.gbEWS.TabIndex = 23;
            this.gbEWS.TabStop = false;
            this.gbEWS.Text = "EWS Settings";
            // 
            // txtEWSServerURL
            // 
            this.txtEWSServerURL.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left)
                        | System.Windows.Forms.AnchorStyles.Right)));
            this.txtEWSServerURL.Location = new System.Drawing.Point(63, 67);
            this.txtEWSServerURL.Name = "txtEWSServerURL";
            this.txtEWSServerURL.Size = new System.Drawing.Size(265, 20);
            this.txtEWSServerURL.TabIndex = 13;
            this.txtEWSServerURL.TextChanged += new System.EventHandler(this.txtEWSServerURL_TextChanged);
            // 
            // label8
            // 
            this.label8.Location = new System.Drawing.Point(6, 70);
            this.label8.Name = "label8";
            this.label8.Size = new System.Drawing.Size(58, 17);
            this.label8.TabIndex = 12;
            this.label8.Text = "Server:";
            // 
            // txtEWSPass
            // 
            this.txtEWSPass.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left)
                        | System.Windows.Forms.AnchorStyles.Right)));
            this.txtEWSPass.Location = new System.Drawing.Point(63, 41);
            this.txtEWSPass.Name = "txtEWSPass";
            this.txtEWSPass.Size = new System.Drawing.Size(265, 20);
            this.txtEWSPass.TabIndex = 11;
            this.txtEWSPass.TextChanged += new System.EventHandler(this.txtEWSPass_TextChanged);
            // 
            // label7
            // 
            this.label7.Location = new System.Drawing.Point(6, 44);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(58, 18);
            this.label7.TabIndex = 10;
            this.label7.Text = "Password:";
            // 
            // txtEWSUser
            // 
            this.txtEWSUser.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left)
                        | System.Windows.Forms.AnchorStyles.Right)));
            this.txtEWSUser.Location = new System.Drawing.Point(63, 15);
            this.txtEWSUser.Name = "txtEWSUser";
            this.txtEWSUser.Size = new System.Drawing.Size(265, 20);
            this.txtEWSUser.TabIndex = 9;
            this.txtEWSUser.TextChanged += new System.EventHandler(this.txtEWSUser_TextChanged);
            // 
            // label6
            // 
            this.label6.Location = new System.Drawing.Point(6, 18);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(35, 17);
            this.label6.TabIndex = 8;
            this.label6.Text = "Email:";
            // 
            // tabGoogle
            // 
            this.tabGoogle.Controls.Add(this.label13);
            this.tabGoogle.Controls.Add(this.label12);
            this.tabGoogle.Controls.Add(this.btResetGCal);
            this.tabGoogle.Controls.Add(this.lGoogleHelp);
            this.tabGoogle.Controls.Add(this.lGoogleCalendar);
            this.tabGoogle.Controls.Add(this.bGetGoogleCalendars);
            this.tabGoogle.Controls.Add(this.cbGoogleCalendars);
            this.tabGoogle.Location = new System.Drawing.Point(79, 4);
            this.tabGoogle.Name = "tabGoogle";
            this.tabGoogle.Padding = new System.Windows.Forms.Padding(3);
            this.tabGoogle.Size = new System.Drawing.Size(392, 462);
            this.tabGoogle.TabIndex = 1;
            this.tabGoogle.Text = "  Google";
            this.tabGoogle.UseVisualStyleBackColor = true;
            // 
            // label13
            // 
            this.label13.AutoSize = true;
            this.label13.Location = new System.Drawing.Point(153, 96);
            this.label13.MaximumSize = new System.Drawing.Size(200, 0);
            this.label13.Name = "label13";
            this.label13.Size = new System.Drawing.Size(164, 26);
            this.label13.TabIndex = 28;
            this.label13.Text = "Reset the Google account being synchronised with\r\n";
            // 
            // label12
            // 
            this.label12.Anchor = System.Windows.Forms.AnchorStyles.Top;
            this.label12.AutoSize = true;
            this.label12.Font = new System.Drawing.Font("Arial Black", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label12.ForeColor = System.Drawing.SystemColors.MenuHighlight;
            this.label12.Location = new System.Drawing.Point(68, 13);
            this.label12.Name = "label12";
            this.label12.Size = new System.Drawing.Size(263, 15);
            this.label12.TabIndex = 27;
            this.label12.Text = "Select the Google Calendar to Synchronise";
            this.label12.TextAlign = System.Drawing.ContentAlignment.TopCenter;
            // 
            // btResetGCal
            // 
            this.btResetGCal.BackColor = System.Drawing.Color.Transparent;
            this.btResetGCal.ForeColor = System.Drawing.Color.Red;
            this.btResetGCal.Location = new System.Drawing.Point(34, 98);
            this.btResetGCal.Name = "btResetGCal";
            this.btResetGCal.Size = new System.Drawing.Size(115, 23);
            this.btResetGCal.TabIndex = 10;
            this.btResetGCal.Text = "Reset Account";
            this.btResetGCal.UseVisualStyleBackColor = false;
            this.btResetGCal.Click += new System.EventHandler(this.btResetGCal_Click);
            // 
            // lGoogleHelp
            // 
            this.lGoogleHelp.AutoSize = true;
            this.lGoogleHelp.Location = new System.Drawing.Point(155, 41);
            this.lGoogleHelp.MaximumSize = new System.Drawing.Size(200, 0);
            this.lGoogleHelp.Name = "lGoogleHelp";
            this.lGoogleHelp.Size = new System.Drawing.Size(200, 39);
            this.lGoogleHelp.TabIndex = 9;
            this.lGoogleHelp.Text = "If this is the first time, you\'ll need to authorise the app to connect.\r\nDoesn\'t " +
                "take long - just follow the steps :)";
            // 
            // lGoogleCalendar
            // 
            this.lGoogleCalendar.Location = new System.Drawing.Point(11, 151);
            this.lGoogleCalendar.Name = "lGoogleCalendar";
            this.lGoogleCalendar.Size = new System.Drawing.Size(81, 14);
            this.lGoogleCalendar.TabIndex = 8;
            this.lGoogleCalendar.Text = "Select calendar";
            // 
            // bGetGoogleCalendars
            // 
            this.bGetGoogleCalendars.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.bGetGoogleCalendars.Location = new System.Drawing.Point(34, 51);
            this.bGetGoogleCalendars.Name = "bGetGoogleCalendars";
            this.bGetGoogleCalendars.Size = new System.Drawing.Size(115, 23);
            this.bGetGoogleCalendars.TabIndex = 7;
            this.bGetGoogleCalendars.Text = "Retrieve Calendars";
            this.bGetGoogleCalendars.UseVisualStyleBackColor = true;
            this.bGetGoogleCalendars.Click += new System.EventHandler(this.GetMyGoogleCalendars_Click);
            // 
            // cbGoogleCalendars
            // 
            this.cbGoogleCalendars.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left)
                        | System.Windows.Forms.AnchorStyles.Right)));
            this.cbGoogleCalendars.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cbGoogleCalendars.FormattingEnabled = true;
            this.cbGoogleCalendars.Location = new System.Drawing.Point(97, 148);
            this.cbGoogleCalendars.Name = "cbGoogleCalendars";
            this.cbGoogleCalendars.Size = new System.Drawing.Size(281, 21);
            this.cbGoogleCalendars.TabIndex = 6;
            this.cbGoogleCalendars.SelectedIndexChanged += new System.EventHandler(this.cbGoogleCalendars_SelectedIndexChanged);
            // 
            // tabSyncOptions
            // 
            this.tabSyncOptions.Controls.Add(this.gbSyncOptions_When);
            this.tabSyncOptions.Controls.Add(this.gbSyncOptions_What);
            this.tabSyncOptions.Controls.Add(this.label15);
            this.tabSyncOptions.Controls.Add(this.gbSyncOptions_How);
            this.tabSyncOptions.Location = new System.Drawing.Point(79, 4);
            this.tabSyncOptions.Name = "tabSyncOptions";
            this.tabSyncOptions.Size = new System.Drawing.Size(392, 462);
            this.tabSyncOptions.TabIndex = 2;
            this.tabSyncOptions.Text = "  Sync Options";
            this.tabSyncOptions.UseVisualStyleBackColor = true;
            // 
            // gbSyncOptions_When
            // 
            this.gbSyncOptions_When.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left)
                        | System.Windows.Forms.AnchorStyles.Right)));
            this.gbSyncOptions_When.Controls.Add(this.cbOutlookPush);
            this.gbSyncOptions_When.Controls.Add(this.cbIntervalUnit);
            this.gbSyncOptions_When.Controls.Add(this.tbInterval);
            this.gbSyncOptions_When.Controls.Add(this.label1);
            this.gbSyncOptions_When.Controls.Add(this.tbDaysInTheFuture);
            this.gbSyncOptions_When.Controls.Add(this.tbDaysInThePast);
            this.gbSyncOptions_When.Controls.Add(this.lDaysInFuture);
            this.gbSyncOptions_When.Controls.Add(this.lDaysInPast);
            this.gbSyncOptions_When.Controls.Add(this.lDateRange);
            this.gbSyncOptions_When.Font = new System.Drawing.Font("Arial Black", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.gbSyncOptions_When.ForeColor = System.Drawing.SystemColors.MenuHighlight;
            this.gbSyncOptions_When.Location = new System.Drawing.Point(15, 151);
            this.gbSyncOptions_When.Name = "gbSyncOptions_When";
            this.gbSyncOptions_When.Size = new System.Drawing.Size(366, 116);
            this.gbSyncOptions_When.TabIndex = 41;
            this.gbSyncOptions_When.TabStop = false;
            this.gbSyncOptions_When.Text = "When";
            // 
            // cbOutlookPush
            // 
            this.cbOutlookPush.AutoSize = true;
            this.cbOutlookPush.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.cbOutlookPush.ForeColor = System.Drawing.SystemColors.ControlText;
            this.cbOutlookPush.Location = new System.Drawing.Point(42, 92);
            this.cbOutlookPush.Name = "cbOutlookPush";
            this.cbOutlookPush.Size = new System.Drawing.Size(191, 17);
            this.cbOutlookPush.TabIndex = 42;
            this.cbOutlookPush.Text = "Push Outlook changes immediately";
            this.cbOutlookPush.UseVisualStyleBackColor = true;
            this.cbOutlookPush.CheckedChanged += new System.EventHandler(this.cbOutlookPush_CheckedChanged);
            // 
            // cbIntervalUnit
            // 
            this.cbIntervalUnit.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cbIntervalUnit.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.cbIntervalUnit.ForeColor = System.Drawing.SystemColors.ControlText;
            this.cbIntervalUnit.FormattingEnabled = true;
            this.cbIntervalUnit.Items.AddRange(new object[] {
            "Minutes",
            "Hours"});
            this.cbIntervalUnit.Location = new System.Drawing.Point(118, 65);
            this.cbIntervalUnit.Name = "cbIntervalUnit";
            this.cbIntervalUnit.Size = new System.Drawing.Size(84, 21);
            this.cbIntervalUnit.TabIndex = 41;
            this.cbIntervalUnit.SelectedIndexChanged += new System.EventHandler(this.cbIntervalUnit_SelectedIndexChanged);
            // 
            // tbInterval
            // 
            this.tbInterval.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.tbInterval.ForeColor = System.Drawing.SystemColors.ControlText;
            this.tbInterval.Location = new System.Drawing.Point(74, 66);
            this.tbInterval.Maximum = new decimal(new int[] {
            120,
            0,
            0,
            0});
            this.tbInterval.Name = "tbInterval";
            this.tbInterval.Size = new System.Drawing.Size(40, 20);
            this.tbInterval.TabIndex = 40;
            this.tbInterval.TextAlign = System.Windows.Forms.HorizontalAlignment.Right;
            this.tbInterval.Value = new decimal(new int[] {
            1,
            0,
            0,
            0});
            this.tbInterval.ValueChanged += new System.EventHandler(this.tbMinuteOffsets_ValueChanged);
            // 
            // label1
            // 
            this.label1.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label1.ForeColor = System.Drawing.SystemColors.ControlText;
            this.label1.Location = new System.Drawing.Point(10, 68);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(66, 14);
            this.label1.TabIndex = 39;
            this.label1.Text = "Interval:";
            this.label1.TextAlign = System.Drawing.ContentAlignment.TopRight;
            // 
            // tbDaysInTheFuture
            // 
            this.tbDaysInTheFuture.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.tbDaysInTheFuture.ForeColor = System.Drawing.SystemColors.ControlText;
            this.tbDaysInTheFuture.Location = new System.Drawing.Point(74, 42);
            this.tbDaysInTheFuture.Maximum = new decimal(new int[] {
            365,
            0,
            0,
            0});
            this.tbDaysInTheFuture.Name = "tbDaysInTheFuture";
            this.tbDaysInTheFuture.Size = new System.Drawing.Size(40, 20);
            this.tbDaysInTheFuture.TabIndex = 38;
            this.tbDaysInTheFuture.TextAlign = System.Windows.Forms.HorizontalAlignment.Right;
            this.tbDaysInTheFuture.Value = new decimal(new int[] {
            7,
            0,
            0,
            0});
            this.tbDaysInTheFuture.ValueChanged += new System.EventHandler(this.tbDaysInTheFuture_ValueChanged);
            // 
            // tbDaysInThePast
            // 
            this.tbDaysInThePast.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.tbDaysInThePast.ForeColor = System.Drawing.SystemColors.ControlText;
            this.tbDaysInThePast.Location = new System.Drawing.Point(74, 18);
            this.tbDaysInThePast.Maximum = new decimal(new int[] {
            365,
            0,
            0,
            0});
            this.tbDaysInThePast.Name = "tbDaysInThePast";
            this.tbDaysInThePast.Size = new System.Drawing.Size(40, 20);
            this.tbDaysInThePast.TabIndex = 36;
            this.tbDaysInThePast.TextAlign = System.Windows.Forms.HorizontalAlignment.Right;
            this.tbDaysInThePast.Value = new decimal(new int[] {
            1,
            0,
            0,
            0});
            this.tbDaysInThePast.ValueChanged += new System.EventHandler(this.tbDaysInThePast_ValueChanged);
            // 
            // lDaysInFuture
            // 
            this.lDaysInFuture.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lDaysInFuture.ForeColor = System.Drawing.SystemColors.ControlText;
            this.lDaysInFuture.Location = new System.Drawing.Point(115, 44);
            this.lDaysInFuture.Name = "lDaysInFuture";
            this.lDaysInFuture.Size = new System.Drawing.Size(104, 20);
            this.lDaysInFuture.TabIndex = 34;
            this.lDaysInFuture.Text = "days in the future";
            // 
            // lDaysInPast
            // 
            this.lDaysInPast.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lDaysInPast.ForeColor = System.Drawing.SystemColors.ControlText;
            this.lDaysInPast.Location = new System.Drawing.Point(115, 21);
            this.lDaysInPast.Name = "lDaysInPast";
            this.lDaysInPast.Size = new System.Drawing.Size(87, 18);
            this.lDaysInPast.TabIndex = 35;
            this.lDaysInPast.Text = "days in the past";
            // 
            // lDateRange
            // 
            this.lDateRange.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lDateRange.ForeColor = System.Drawing.SystemColors.ControlText;
            this.lDateRange.Location = new System.Drawing.Point(10, 21);
            this.lDateRange.Name = "lDateRange";
            this.lDateRange.Size = new System.Drawing.Size(66, 14);
            this.lDateRange.TabIndex = 37;
            this.lDateRange.Text = "Date range:";
            this.lDateRange.TextAlign = System.Drawing.ContentAlignment.TopRight;
            // 
            // gbSyncOptions_What
            // 
            this.gbSyncOptions_What.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left)
                        | System.Windows.Forms.AnchorStyles.Right)));
            this.gbSyncOptions_What.Controls.Add(this.cbAddReminders);
            this.gbSyncOptions_What.Controls.Add(this.lAttributes);
            this.gbSyncOptions_What.Controls.Add(this.cbAddAttendees);
            this.gbSyncOptions_What.Controls.Add(this.cbAddDescription);
            this.gbSyncOptions_What.Font = new System.Drawing.Font("Arial Black", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.gbSyncOptions_What.ForeColor = System.Drawing.SystemColors.MenuHighlight;
            this.gbSyncOptions_What.Location = new System.Drawing.Point(12, 273);
            this.gbSyncOptions_What.Name = "gbSyncOptions_What";
            this.gbSyncOptions_What.Size = new System.Drawing.Size(369, 96);
            this.gbSyncOptions_What.TabIndex = 39;
            this.gbSyncOptions_What.TabStop = false;
            this.gbSyncOptions_What.Text = "What";
            // 
            // cbAddReminders
            // 
            this.cbAddReminders.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.cbAddReminders.ForeColor = System.Drawing.SystemColors.ActiveCaptionText;
            this.cbAddReminders.Location = new System.Drawing.Point(45, 73);
            this.cbAddReminders.Name = "cbAddReminders";
            this.cbAddReminders.Size = new System.Drawing.Size(80, 17);
            this.cbAddReminders.TabIndex = 32;
            this.cbAddReminders.Text = "Reminders";
            this.cbAddReminders.UseVisualStyleBackColor = true;
            this.cbAddReminders.CheckedChanged += new System.EventHandler(this.CbAddRemindersCheckedChanged);
            // 
            // lAttributes
            // 
            this.lAttributes.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lAttributes.ForeColor = System.Drawing.SystemColors.ActiveCaptionText;
            this.lAttributes.Location = new System.Drawing.Point(6, 19);
            this.lAttributes.Name = "lAttributes";
            this.lAttributes.Size = new System.Drawing.Size(120, 14);
            this.lAttributes.TabIndex = 33;
            this.lAttributes.Text = "Attributes to include:-";
            // 
            // cbAddAttendees
            // 
            this.cbAddAttendees.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.cbAddAttendees.ForeColor = System.Drawing.SystemColors.ActiveCaptionText;
            this.cbAddAttendees.Location = new System.Drawing.Point(45, 54);
            this.cbAddAttendees.Name = "cbAddAttendees";
            this.cbAddAttendees.Size = new System.Drawing.Size(80, 17);
            this.cbAddAttendees.TabIndex = 30;
            this.cbAddAttendees.Text = "Attendees";
            this.cbAddAttendees.UseVisualStyleBackColor = true;
            this.cbAddAttendees.CheckedChanged += new System.EventHandler(this.cbAddAttendees_CheckedChanged);
            // 
            // cbAddDescription
            // 
            this.cbAddDescription.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.cbAddDescription.ForeColor = System.Drawing.SystemColors.ActiveCaptionText;
            this.cbAddDescription.Location = new System.Drawing.Point(45, 36);
            this.cbAddDescription.Name = "cbAddDescription";
            this.cbAddDescription.Size = new System.Drawing.Size(80, 17);
            this.cbAddDescription.TabIndex = 31;
            this.cbAddDescription.Text = "Description";
            this.cbAddDescription.UseVisualStyleBackColor = true;
            this.cbAddDescription.CheckedChanged += new System.EventHandler(this.CbAddDescriptionCheckedChanged);
            // 
            // label15
            // 
            this.label15.AutoSize = true;
            this.label15.Font = new System.Drawing.Font("Arial Black", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label15.ForeColor = System.Drawing.SystemColors.MenuHighlight;
            this.label15.Location = new System.Drawing.Point(121, 13);
            this.label15.Name = "label15";
            this.label15.Size = new System.Drawing.Size(153, 15);
            this.label15.TabIndex = 35;
            this.label15.Text = "Synchronisation Options";
            this.label15.TextAlign = System.Drawing.ContentAlignment.TopCenter;
            // 
            // gbSyncOptions_How
            // 
            this.gbSyncOptions_How.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left)
                        | System.Windows.Forms.AnchorStyles.Right)));
            this.gbSyncOptions_How.Controls.Add(this.syncDirection);
            this.gbSyncOptions_How.Controls.Add(this.lDirection);
            this.gbSyncOptions_How.Controls.Add(this.cbMergeItems);
            this.gbSyncOptions_How.Controls.Add(this.cbDisableDeletion);
            this.gbSyncOptions_How.Controls.Add(this.cbConfirmOnDelete);
            this.gbSyncOptions_How.Font = new System.Drawing.Font("Arial Black", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.gbSyncOptions_How.ForeColor = System.Drawing.SystemColors.MenuHighlight;
            this.gbSyncOptions_How.Location = new System.Drawing.Point(15, 36);
            this.gbSyncOptions_How.Name = "gbSyncOptions_How";
            this.gbSyncOptions_How.Size = new System.Drawing.Size(366, 109);
            this.gbSyncOptions_How.TabIndex = 40;
            this.gbSyncOptions_How.TabStop = false;
            this.gbSyncOptions_How.Text = "How";
            // 
            // syncDirection
            // 
            this.syncDirection.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.syncDirection.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.syncDirection.FormattingEnabled = true;
            this.syncDirection.Location = new System.Drawing.Point(61, 22);
            this.syncDirection.Name = "syncDirection";
            this.syncDirection.Size = new System.Drawing.Size(299, 21);
            this.syncDirection.TabIndex = 37;
            this.syncDirection.SelectedIndexChanged += new System.EventHandler(this.syncDirection_SelectedIndexChanged);
            // 
            // lDirection
            // 
            this.lDirection.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lDirection.ForeColor = System.Drawing.SystemColors.ControlText;
            this.lDirection.Location = new System.Drawing.Point(8, 25);
            this.lDirection.Name = "lDirection";
            this.lDirection.Size = new System.Drawing.Size(55, 14);
            this.lDirection.TabIndex = 38;
            this.lDirection.Text = "Direction:";
            this.lDirection.TextAlign = System.Drawing.ContentAlignment.TopRight;
            // 
            // cbMergeItems
            // 
            this.cbMergeItems.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.cbMergeItems.ForeColor = System.Drawing.SystemColors.ControlText;
            this.cbMergeItems.Location = new System.Drawing.Point(42, 49);
            this.cbMergeItems.Name = "cbMergeItems";
            this.cbMergeItems.Size = new System.Drawing.Size(152, 17);
            this.cbMergeItems.TabIndex = 36;
            this.cbMergeItems.Text = "Merge with existing entries";
            this.cbMergeItems.UseVisualStyleBackColor = true;
            this.cbMergeItems.CheckedChanged += new System.EventHandler(this.cbMergeItems_CheckedChanged);
            // 
            // cbDisableDeletion
            // 
            this.cbDisableDeletion.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.cbDisableDeletion.ForeColor = System.Drawing.SystemColors.ControlText;
            this.cbDisableDeletion.Location = new System.Drawing.Point(42, 67);
            this.cbDisableDeletion.Name = "cbDisableDeletion";
            this.cbDisableDeletion.Size = new System.Drawing.Size(106, 17);
            this.cbDisableDeletion.TabIndex = 33;
            this.cbDisableDeletion.Text = "Disable deletions";
            this.cbDisableDeletion.UseVisualStyleBackColor = true;
            this.cbDisableDeletion.CheckedChanged += new System.EventHandler(this.cbDisableDeletion_CheckedChanged);
            // 
            // cbConfirmOnDelete
            // 
            this.cbConfirmOnDelete.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.cbConfirmOnDelete.ForeColor = System.Drawing.SystemColors.ControlText;
            this.cbConfirmOnDelete.Location = new System.Drawing.Point(42, 85);
            this.cbConfirmOnDelete.Name = "cbConfirmOnDelete";
            this.cbConfirmOnDelete.Size = new System.Drawing.Size(111, 17);
            this.cbConfirmOnDelete.TabIndex = 34;
            this.cbConfirmOnDelete.Text = "Confirm deletions";
            this.cbConfirmOnDelete.UseVisualStyleBackColor = true;
            this.cbConfirmOnDelete.CheckedChanged += new System.EventHandler(this.cbConfirmOnDelete_CheckedChanged);
            // 
            // tabAppBehaviour
            // 
            this.tabAppBehaviour.Controls.Add(this.gbProxy);
            this.tabAppBehaviour.Controls.Add(this.label14);
            this.tabAppBehaviour.Controls.Add(this.btLogLocation);
            this.tabAppBehaviour.Controls.Add(this.label3);
            this.tabAppBehaviour.Controls.Add(this.cbLoggingLevel);
            this.tabAppBehaviour.Controls.Add(this.cbStartOnStartup);
            this.tabAppBehaviour.Controls.Add(this.cbShowBubbleTooltips);
            this.tabAppBehaviour.Controls.Add(this.cbMinimizeToTray);
            this.tabAppBehaviour.Controls.Add(this.cbStartInTray);
            this.tabAppBehaviour.Controls.Add(this.cbCreateFiles);
            this.tabAppBehaviour.Location = new System.Drawing.Point(79, 4);
            this.tabAppBehaviour.Name = "tabAppBehaviour";
            this.tabAppBehaviour.Size = new System.Drawing.Size(392, 462);
            this.tabAppBehaviour.TabIndex = 3;
            this.tabAppBehaviour.Text = "  Application Behaviour";
            this.tabAppBehaviour.UseVisualStyleBackColor = true;
            // 
            // gbProxy
            // 
            this.gbProxy.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left)
                        | System.Windows.Forms.AnchorStyles.Right)));
            this.gbProxy.Controls.Add(this.rbProxyNone);
            this.gbProxy.Controls.Add(this.rbProxyIE);
            this.gbProxy.Controls.Add(this.rbProxyCustom);
            this.gbProxy.Controls.Add(this.txtProxyPassword);
            this.gbProxy.Controls.Add(this.cbProxyAuthRequired);
            this.gbProxy.Controls.Add(this.label9);
            this.gbProxy.Controls.Add(this.txtProxyPort);
            this.gbProxy.Controls.Add(this.txtProxyUser);
            this.gbProxy.Controls.Add(this.txtProxyServer);
            this.gbProxy.Controls.Add(this.label10);
            this.gbProxy.Controls.Add(this.label5);
            this.gbProxy.Controls.Add(this.label4);
            this.gbProxy.Font = new System.Drawing.Font("Arial Black", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.gbProxy.ForeColor = System.Drawing.SystemColors.MenuHighlight;
            this.gbProxy.Location = new System.Drawing.Point(16, 168);
            this.gbProxy.Name = "gbProxy";
            this.gbProxy.Size = new System.Drawing.Size(364, 200);
            this.gbProxy.TabIndex = 37;
            this.gbProxy.TabStop = false;
            this.gbProxy.Text = "Proxy Setting";
            this.gbProxy.Leave += new System.EventHandler(this.gbProxy_Leave);
            // 
            // rbProxyNone
            // 
            this.rbProxyNone.AutoSize = true;
            this.rbProxyNone.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.rbProxyNone.ForeColor = System.Drawing.SystemColors.ControlText;
            this.rbProxyNone.Location = new System.Drawing.Point(20, 22);
            this.rbProxyNone.Name = "rbProxyNone";
            this.rbProxyNone.Size = new System.Drawing.Size(68, 17);
            this.rbProxyNone.TabIndex = 1;
            this.rbProxyNone.Tag = "None";
            this.rbProxyNone.Text = "No Proxy";
            this.rbProxyNone.UseVisualStyleBackColor = true;
            this.rbProxyNone.CheckedChanged += new System.EventHandler(this.rbProxyCustom_CheckedChanged);
            // 
            // rbProxyIE
            // 
            this.rbProxyIE.AutoSize = true;
            this.rbProxyIE.Checked = true;
            this.rbProxyIE.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.rbProxyIE.ForeColor = System.Drawing.SystemColors.ControlText;
            this.rbProxyIE.Location = new System.Drawing.Point(20, 45);
            this.rbProxyIE.Name = "rbProxyIE";
            this.rbProxyIE.Size = new System.Drawing.Size(157, 17);
            this.rbProxyIE.TabIndex = 2;
            this.rbProxyIE.TabStop = true;
            this.rbProxyIE.Tag = "IE";
            this.rbProxyIE.Text = "Inherit from Internet Explorer";
            this.rbProxyIE.UseVisualStyleBackColor = true;
            this.rbProxyIE.CheckedChanged += new System.EventHandler(this.rbProxyCustom_CheckedChanged);
            // 
            // rbProxyCustom
            // 
            this.rbProxyCustom.AutoSize = true;
            this.rbProxyCustom.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.rbProxyCustom.ForeColor = System.Drawing.SystemColors.ControlText;
            this.rbProxyCustom.Location = new System.Drawing.Point(20, 68);
            this.rbProxyCustom.Name = "rbProxyCustom";
            this.rbProxyCustom.Size = new System.Drawing.Size(104, 17);
            this.rbProxyCustom.TabIndex = 3;
            this.rbProxyCustom.Tag = "Custom";
            this.rbProxyCustom.Text = "Custom Setttings";
            this.rbProxyCustom.UseVisualStyleBackColor = true;
            this.rbProxyCustom.CheckedChanged += new System.EventHandler(this.rbProxyCustom_CheckedChanged);
            // 
            // txtProxyPassword
            // 
            this.txtProxyPassword.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left)
                        | System.Windows.Forms.AnchorStyles.Right)));
            this.txtProxyPassword.Enabled = false;
            this.txtProxyPassword.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtProxyPassword.ForeColor = System.Drawing.SystemColors.ControlText;
            this.txtProxyPassword.Location = new System.Drawing.Point(94, 167);
            this.txtProxyPassword.Name = "txtProxyPassword";
            this.txtProxyPassword.Size = new System.Drawing.Size(259, 20);
            this.txtProxyPassword.TabIndex = 8;
            // 
            // cbProxyAuthRequired
            // 
            this.cbProxyAuthRequired.AutoSize = true;
            this.cbProxyAuthRequired.Enabled = false;
            this.cbProxyAuthRequired.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.cbProxyAuthRequired.ForeColor = System.Drawing.SystemColors.ControlText;
            this.cbProxyAuthRequired.Location = new System.Drawing.Point(38, 120);
            this.cbProxyAuthRequired.Name = "cbProxyAuthRequired";
            this.cbProxyAuthRequired.Size = new System.Drawing.Size(140, 17);
            this.cbProxyAuthRequired.TabIndex = 6;
            this.cbProxyAuthRequired.Text = "Authentication Required";
            this.cbProxyAuthRequired.UseVisualStyleBackColor = true;
            this.cbProxyAuthRequired.CheckedChanged += new System.EventHandler(this.cbProxyAuthRequired_CheckedChanged);
            // 
            // label9
            // 
            this.label9.AutoSize = true;
            this.label9.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label9.ForeColor = System.Drawing.SystemColors.ControlText;
            this.label9.Location = new System.Drawing.Point(35, 144);
            this.label9.Name = "label9";
            this.label9.Size = new System.Drawing.Size(58, 13);
            this.label9.TabIndex = 0;
            this.label9.Text = "Username:";
            // 
            // txtProxyPort
            // 
            this.txtProxyPort.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.txtProxyPort.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtProxyPort.ForeColor = System.Drawing.SystemColors.ControlText;
            this.txtProxyPort.Location = new System.Drawing.Point(302, 92);
            this.txtProxyPort.Name = "txtProxyPort";
            this.txtProxyPort.Size = new System.Drawing.Size(51, 20);
            this.txtProxyPort.TabIndex = 5;
            // 
            // txtProxyUser
            // 
            this.txtProxyUser.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left)
                        | System.Windows.Forms.AnchorStyles.Right)));
            this.txtProxyUser.Enabled = false;
            this.txtProxyUser.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtProxyUser.ForeColor = System.Drawing.SystemColors.ControlText;
            this.txtProxyUser.Location = new System.Drawing.Point(94, 143);
            this.txtProxyUser.Name = "txtProxyUser";
            this.txtProxyUser.Size = new System.Drawing.Size(259, 20);
            this.txtProxyUser.TabIndex = 7;
            // 
            // txtProxyServer
            // 
            this.txtProxyServer.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left)
                        | System.Windows.Forms.AnchorStyles.Right)));
            this.txtProxyServer.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtProxyServer.ForeColor = System.Drawing.SystemColors.ControlText;
            this.txtProxyServer.Location = new System.Drawing.Point(94, 92);
            this.txtProxyServer.Name = "txtProxyServer";
            this.txtProxyServer.Size = new System.Drawing.Size(174, 20);
            this.txtProxyServer.TabIndex = 4;
            // 
            // label10
            // 
            this.label10.AutoSize = true;
            this.label10.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label10.ForeColor = System.Drawing.SystemColors.ControlText;
            this.label10.Location = new System.Drawing.Point(35, 170);
            this.label10.Name = "label10";
            this.label10.Size = new System.Drawing.Size(56, 13);
            this.label10.TabIndex = 0;
            this.label10.Text = "Password:";
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label5.ForeColor = System.Drawing.SystemColors.ControlText;
            this.label5.Location = new System.Drawing.Point(276, 95);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(29, 13);
            this.label5.TabIndex = 0;
            this.label5.Text = "Port:";
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label4.ForeColor = System.Drawing.SystemColors.ControlText;
            this.label4.Location = new System.Drawing.Point(35, 95);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(41, 13);
            this.label4.TabIndex = 0;
            this.label4.Text = "Server:";
            // 
            // label14
            // 
            this.label14.Anchor = System.Windows.Forms.AnchorStyles.Top;
            this.label14.AutoSize = true;
            this.label14.Font = new System.Drawing.Font("Arial Black", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label14.ForeColor = System.Drawing.SystemColors.MenuHighlight;
            this.label14.Location = new System.Drawing.Point(136, 13);
            this.label14.Name = "label14";
            this.label14.Size = new System.Drawing.Size(139, 15);
            this.label14.TabIndex = 36;
            this.label14.Text = "Application Behaviour";
            this.label14.TextAlign = System.Drawing.ContentAlignment.TopCenter;
            // 
            // btLogLocation
            // 
            this.btLogLocation.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.btLogLocation.Location = new System.Drawing.Point(302, 136);
            this.btLogLocation.Name = "btLogLocation";
            this.btLogLocation.Size = new System.Drawing.Size(80, 23);
            this.btLogLocation.TabIndex = 19;
            this.btLogLocation.Text = "Open Log";
            this.btLogLocation.UseVisualStyleBackColor = true;
            this.btLogLocation.Click += new System.EventHandler(this.btLogLocation_Click);
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(13, 141);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(70, 13);
            this.label3.TabIndex = 18;
            this.label3.Text = "Logging level";
            // 
            // cbLoggingLevel
            // 
            this.cbLoggingLevel.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left)
                        | System.Windows.Forms.AnchorStyles.Right)));
            this.cbLoggingLevel.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cbLoggingLevel.FormattingEnabled = true;
            this.cbLoggingLevel.Items.AddRange(new object[] {
            "Off",
            "Fatal",
            "Error",
            "Warn",
            "Info",
            "Debug",
            "Fine",
            "All"});
            this.cbLoggingLevel.Location = new System.Drawing.Point(86, 137);
            this.cbLoggingLevel.Name = "cbLoggingLevel";
            this.cbLoggingLevel.Size = new System.Drawing.Size(210, 21);
            this.cbLoggingLevel.TabIndex = 17;
            this.cbLoggingLevel.Tag = "l";
            this.cbLoggingLevel.SelectedIndexChanged += new System.EventHandler(this.cbLoggingLevel_SelectedIndexChanged);
            // 
            // cbStartOnStartup
            // 
            this.cbStartOnStartup.AutoSize = true;
            this.cbStartOnStartup.Location = new System.Drawing.Point(16, 43);
            this.cbStartOnStartup.Name = "cbStartOnStartup";
            this.cbStartOnStartup.Size = new System.Drawing.Size(88, 17);
            this.cbStartOnStartup.TabIndex = 16;
            this.cbStartOnStartup.Text = "Start on login";
            this.cbStartOnStartup.UseVisualStyleBackColor = true;
            this.cbStartOnStartup.CheckedChanged += new System.EventHandler(this.cbStartOnStartup_CheckedChanged);
            // 
            // cbShowBubbleTooltips
            // 
            this.cbShowBubbleTooltips.Location = new System.Drawing.Point(16, 96);
            this.cbShowBubbleTooltips.Name = "cbShowBubbleTooltips";
            this.cbShowBubbleTooltips.Size = new System.Drawing.Size(259, 24);
            this.cbShowBubbleTooltips.TabIndex = 14;
            this.cbShowBubbleTooltips.Text = "Show Bubble Tooltip in Taskbar when Syncing";
            this.cbShowBubbleTooltips.UseVisualStyleBackColor = true;
            this.cbShowBubbleTooltips.CheckedChanged += new System.EventHandler(this.cbShowBubbleTooltipsCheckedChanged);
            // 
            // cbMinimizeToTray
            // 
            this.cbMinimizeToTray.Location = new System.Drawing.Point(16, 77);
            this.cbMinimizeToTray.Name = "cbMinimizeToTray";
            this.cbMinimizeToTray.Size = new System.Drawing.Size(104, 24);
            this.cbMinimizeToTray.TabIndex = 12;
            this.cbMinimizeToTray.Text = "Minimize to Tray";
            this.cbMinimizeToTray.UseVisualStyleBackColor = true;
            this.cbMinimizeToTray.CheckedChanged += new System.EventHandler(this.cbMinimizeToTrayCheckedChanged);
            // 
            // cbStartInTray
            // 
            this.cbStartInTray.Location = new System.Drawing.Point(16, 58);
            this.cbStartInTray.Name = "cbStartInTray";
            this.cbStartInTray.Size = new System.Drawing.Size(104, 24);
            this.cbStartInTray.TabIndex = 13;
            this.cbStartInTray.Text = "Start in Tray";
            this.cbStartInTray.UseVisualStyleBackColor = true;
            this.cbStartInTray.CheckedChanged += new System.EventHandler(this.cbStartInTrayCheckedChanged);
            // 
            // cbCreateFiles
            // 
            this.cbCreateFiles.Location = new System.Drawing.Point(16, 115);
            this.cbCreateFiles.Name = "cbCreateFiles";
            this.cbCreateFiles.Size = new System.Drawing.Size(235, 24);
            this.cbCreateFiles.TabIndex = 15;
            this.cbCreateFiles.Text = "Create CSV files of calendar entries";
            this.cbCreateFiles.UseVisualStyleBackColor = true;
            this.cbCreateFiles.CheckedChanged += new System.EventHandler(this.cbCreateFiles_CheckedChanged);
            // 
            // bSave
            // 
            this.bSave.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.bSave.Location = new System.Drawing.Point(392, 493);
            this.bSave.Name = "bSave";
            this.bSave.Size = new System.Drawing.Size(75, 31);
            this.bSave.TabIndex = 8;
            this.bSave.Text = "Save";
            this.bSave.UseVisualStyleBackColor = true;
            this.bSave.Click += new System.EventHandler(this.Save_Click);
            // 
            // tabPage_About
            // 
            this.tabPage_About.Controls.Add(this.label2);
            this.tabPage_About.Controls.Add(this.pbDonate);
            this.tabPage_About.Controls.Add(this.lAboutURL);
            this.tabPage_About.Controls.Add(this.lAboutMain);
            this.tabPage_About.Location = new System.Drawing.Point(4, 22);
            this.tabPage_About.Name = "tabPage_About";
            this.tabPage_About.Padding = new System.Windows.Forms.Padding(3);
            this.tabPage_About.Size = new System.Drawing.Size(487, 542);
            this.tabPage_About.TabIndex = 2;
            this.tabPage_About.Text = "About";
            this.tabPage_About.UseVisualStyleBackColor = true;
            // 
            // label2
            // 
            this.label2.Anchor = System.Windows.Forms.AnchorStyles.Top;
            this.label2.AutoSize = true;
            this.label2.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label2.Location = new System.Drawing.Point(153, 48);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(181, 13);
            this.label2.TabIndex = 4;
            this.label2.Text = "Outlook Google Calendar Sync";
            // 
            // pbDonate
            // 
            this.pbDonate.Anchor = System.Windows.Forms.AnchorStyles.Top;
            this.pbDonate.Cursor = System.Windows.Forms.Cursors.Hand;
            this.pbDonate.Image = global::OutlookGoogleCalendarSync.Properties.Resources.paypalDonate;
            this.pbDonate.Location = new System.Drawing.Point(206, 311);
            this.pbDonate.Name = "pbDonate";
            this.pbDonate.Size = new System.Drawing.Size(75, 23);
            this.pbDonate.TabIndex = 3;
            this.pbDonate.TabStop = false;
            this.pbDonate.Click += new System.EventHandler(this.pbDonate_Click);
            // 
            // lAboutURL
            // 
            this.lAboutURL.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Left | System.Windows.Forms.AnchorStyles.Right)));
            this.lAboutURL.Location = new System.Drawing.Point(6, 396);
            this.lAboutURL.Name = "lAboutURL";
            this.lAboutURL.Size = new System.Drawing.Size(475, 23);
            this.lAboutURL.TabIndex = 2;
            this.lAboutURL.TabStop = true;
            this.lAboutURL.Text = "http://outlookgooglecalendarsync.codeplex.com/";
            this.lAboutURL.TextAlign = System.Drawing.ContentAlignment.TopCenter;
            this.lAboutURL.LinkClicked += new System.Windows.Forms.LinkLabelLinkClickedEventHandler(this.lAboutURL_LinkClicked);
            // 
            // lAboutMain
            // 
            this.lAboutMain.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left)
                        | System.Windows.Forms.AnchorStyles.Right)));
            this.lAboutMain.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.lAboutMain.Location = new System.Drawing.Point(30, 32);
            this.lAboutMain.Name = "lAboutMain";
            this.lAboutMain.Padding = new System.Windows.Forms.Padding(15);
            this.lAboutMain.Size = new System.Drawing.Size(426, 318);
            this.lAboutMain.TabIndex = 1;
            this.lAboutMain.Text = resources.GetString("lAboutMain.Text");
            this.lAboutMain.TextAlign = System.Drawing.ContentAlignment.TopCenter;
            // 
            // notifyIcon1
            // 
            this.notifyIcon1.Icon = ((System.Drawing.Icon)(resources.GetObject("notifyIcon1.Icon")));
            this.notifyIcon1.Text = "Outlook Google Calendar Sync";
            this.notifyIcon1.Click += new System.EventHandler(this.NotifyIcon1_Click);
            // 
            // MainForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(519, 592);
            this.Controls.Add(this.tabApp);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MinimumSize = new System.Drawing.Size(535, 630);
            this.Name = "MainForm";
            this.SizeGripStyle = System.Windows.Forms.SizeGripStyle.Show;
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Outlook Google Calendar Sync";
            this.Resize += new System.EventHandler(this.mainFormResize);
            this.tabApp.ResumeLayout(false);
            this.tabPage_Sync.ResumeLayout(false);
            this.tabPage_Sync.PerformLayout();
            this.tabPage_Settings.ResumeLayout(false);
            this.tabPage_Settings.PerformLayout();
            this.tabAppSettings.ResumeLayout(false);
            this.tabOutlook.ResumeLayout(false);
            this.tabOutlook.PerformLayout();
            this.gbEWS.ResumeLayout(false);
            this.gbEWS.PerformLayout();
            this.tabGoogle.ResumeLayout(false);
            this.tabGoogle.PerformLayout();
            this.tabSyncOptions.ResumeLayout(false);
            this.tabSyncOptions.PerformLayout();
            this.gbSyncOptions_When.ResumeLayout(false);
            this.gbSyncOptions_When.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.tbInterval)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.tbDaysInTheFuture)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.tbDaysInThePast)).EndInit();
            this.gbSyncOptions_What.ResumeLayout(false);
            this.gbSyncOptions_How.ResumeLayout(false);
            this.tabAppBehaviour.ResumeLayout(false);
            this.tabAppBehaviour.PerformLayout();
            this.gbProxy.ResumeLayout(false);
            this.gbProxy.PerformLayout();
            this.tabPage_About.ResumeLayout(false);
            this.tabPage_About.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.pbDonate)).EndInit();
            this.ResumeLayout(false);

        }
        private System.Windows.Forms.LinkLabel lAboutURL;
        private System.Windows.Forms.TabPage tabPage_About;
        private System.Windows.Forms.NotifyIcon notifyIcon1;
        private System.Windows.Forms.Label lAboutMain;
        private System.Windows.Forms.TextBox LogBox;
        private System.Windows.Forms.Button bSave;
        private System.Windows.Forms.TabPage tabPage_Sync;
        private System.Windows.Forms.Label lLastSync;
        private System.Windows.Forms.Label lNextSync;
        private System.Windows.Forms.Label lNextSyncVal;
        private System.Windows.Forms.Label lLastSyncVal;
        private System.Windows.Forms.CheckBox cbVerboseOutput;
        public System.Windows.Forms.TabControl tabApp;
        public System.Windows.Forms.TabPage tabPage_Settings;
        private System.Windows.Forms.PictureBox pbDonate;
        private System.Windows.Forms.Label label2;
        public System.Windows.Forms.Button bSyncNow;
        private System.Windows.Forms.TabControl tabAppSettings;
        private System.Windows.Forms.TabPage tabOutlook;
        private System.Windows.Forms.RadioButton rbOutlookEWS;
        public System.Windows.Forms.RadioButton rbOutlookDefaultMB;
        private System.Windows.Forms.RadioButton rbOutlookAltMB;
        private System.Windows.Forms.ComboBox ddMailboxName;
        private System.Windows.Forms.TabPage tabGoogle;
        private System.Windows.Forms.Label lOutlookCalendar;
        private System.Windows.Forms.ComboBox cbOutlookCalendars;
        private System.Windows.Forms.GroupBox gbEWS;
        private System.Windows.Forms.TextBox txtEWSServerURL;
        private System.Windows.Forms.Label label8;
        private System.Windows.Forms.TextBox txtEWSPass;
        private System.Windows.Forms.Label label7;
        private System.Windows.Forms.TextBox txtEWSUser;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.Button btResetGCal;
        private System.Windows.Forms.Label lGoogleHelp;
        private System.Windows.Forms.Label lGoogleCalendar;
        private System.Windows.Forms.Button bGetGoogleCalendars;
        private System.Windows.Forms.ComboBox cbGoogleCalendars;
        private System.Windows.Forms.TabPage tabSyncOptions;
        private System.Windows.Forms.TabPage tabAppBehaviour;
        private System.Windows.Forms.Label label11;
        private System.Windows.Forms.Label label13;
        private System.Windows.Forms.Label label12;
        private System.Windows.Forms.Label label15;
        private System.Windows.Forms.Button btLogLocation;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.ComboBox cbLoggingLevel;
        private System.Windows.Forms.CheckBox cbStartOnStartup;
        private System.Windows.Forms.CheckBox cbShowBubbleTooltips;
        private System.Windows.Forms.CheckBox cbMinimizeToTray;
        private System.Windows.Forms.CheckBox cbStartInTray;
        private System.Windows.Forms.CheckBox cbCreateFiles;
        private System.Windows.Forms.GroupBox gbSyncOptions_When;
        private System.Windows.Forms.CheckBox cbOutlookPush;
        private System.Windows.Forms.ComboBox cbIntervalUnit;
        private System.Windows.Forms.NumericUpDown tbInterval;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.NumericUpDown tbDaysInTheFuture;
        private System.Windows.Forms.NumericUpDown tbDaysInThePast;
        private System.Windows.Forms.Label lDaysInFuture;
        private System.Windows.Forms.Label lDaysInPast;
        private System.Windows.Forms.Label lDateRange;
        private System.Windows.Forms.GroupBox gbSyncOptions_What;
        private System.Windows.Forms.CheckBox cbAddReminders;
        private System.Windows.Forms.Label lAttributes;
        private System.Windows.Forms.CheckBox cbAddAttendees;
        private System.Windows.Forms.CheckBox cbAddDescription;
        private System.Windows.Forms.GroupBox gbSyncOptions_How;
        private System.Windows.Forms.ComboBox syncDirection;
        private System.Windows.Forms.Label lDirection;
        private System.Windows.Forms.CheckBox cbMergeItems;
        private System.Windows.Forms.CheckBox cbDisableDeletion;
        private System.Windows.Forms.CheckBox cbConfirmOnDelete;
        private System.Windows.Forms.Label label14;
        private System.Windows.Forms.GroupBox gbProxy;
        private System.Windows.Forms.RadioButton rbProxyIE;
        private System.Windows.Forms.RadioButton rbProxyCustom;
        private System.Windows.Forms.TextBox txtProxyPassword;
        private System.Windows.Forms.CheckBox cbProxyAuthRequired;
        private System.Windows.Forms.Label label9;
        private System.Windows.Forms.TextBox txtProxyPort;
        private System.Windows.Forms.TextBox txtProxyUser;
        private System.Windows.Forms.TextBox txtProxyServer;
        private System.Windows.Forms.Label label10;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.RadioButton rbProxyNone;
        private System.Windows.Forms.Label lSettingInfo;
    }
}
